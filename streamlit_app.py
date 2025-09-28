import io, zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Enterprise Structure Generator", page_icon="📊", layout="wide")
st.title("Enterprise Structure Generator — Excel + draw.io (identifier-safe, no lost hangers)")

st.markdown("""
Upload up to **9 Oracle export ZIPs** (any order):
- `Manage General Ledger` → **GL_PRIMARY_LEDGER.csv**
- `Manage Legal Entities` → **XLE_ENTITY_PROFILE.csv**
- `Assign Legal Entities` → **ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv**
- *(optional backup)* Journal config detail → **ORA_GL_JOURNAL_CONFIG_DETAIL.csv**
- `Manage Business Units` → **FUN_BUSINESS_UNIT.csv**
- `Manage Cost Organizations` → **CST_COST_ORGANIZATION.csv**
- `Manage Cost Organization Relationships` → **CST_COST_ORG_BOOK.csv**
- `Manage Inventory Organizations` → **INV_ORGANIZATION_PARAMETER.csv**
- `Cost Org ↔ Inventory Org relationships` → **ORA_CST_COST_ORG_INV.csv**
""")

uploads = st.file_uploader("Drop your ZIPs here", type="zip", accept_multiple_files=True)

# ---------- helpers ----------
def read_csv_from_zip(zf, name):
    if name not in zf.namelist():
        return None
    with zf.open(name) as fh:
        return pd.read_csv(fh, dtype=str)

def pick_col(df, candidates):
    cols = list(df.columns)
    for c in candidates:
        if c in cols:
            return c
    lower_map = {c.lower(): c for c in cols}
    for c in candidates:
        if c.lower() in lower_map:
            return lower_map[c.lower()]
    for c in candidates:
        for existing in cols:
            if c.lower() in existing.lower():
                return existing
    return None

def _blankify(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    df = df.copy().fillna("")
    obj_cols = list(df.select_dtypes(include=["object"]).columns)
    for c in obj_cols:
        s = df[c]
        mask = s.apply(lambda x: isinstance(x, str) and x.strip().lower() == "nan")
        if mask.any():
            df.loc[mask, c] = ""
    return df

if not uploads:
    st.info("Upload your ZIPs to generate the Excel & diagram.")
else:
    # ------------ Collectors ------------
    ledger_names = set()
    legal_entity_names = set()
    ledger_to_idents = {}            # ledger -> {LE identifier}
    ident_to_le_name = {}            # identifier -> preferred name (ObjectName first, XLE fallback)
    bu_rows = []                     # BU rows for Tab 1

    # Cost/IO data
    costorg_rows = []                # [{Name, LegalEntityIdentifier, JoinKey}]
    books_by_joinkey = {}            # JoinKey -> {CostBookCode}
    invorg_rows = []                 # [{Code, Name, LEIdent, BUName, PCBU, Mfg}]
    invorg_rel = {}                  # InvOrgCode -> CostOrgJoinKey

    # Diagnostics
    unresolved_ident_pairs = []      # [(Ledger, Identifier)]
    name_collisions = []             # [(Ledger, DisplayName, [idents...])]

    # ------------ Scan uploads ------------
    for up in uploads:
        try:
            z = zipfile.ZipFile(up)
        except Exception as e:
            st.error(f"Could not open `{up.name}` as a ZIP: {e}")
            continue

        # Ledgers
        df = read_csv_from_zip(z, "GL_PRIMARY_LEDGER.csv")
        if df is not None:
            col = pick_col(df, ["ORA_GL_PRIMARY_LEDGER_CONFIG.Name"])
            if col:
                ledger_names |= set(df[col].dropna().map(str).str.strip())
            else:
                st.warning("`GL_PRIMARY_LEDGER.csv` missing `ORA_GL_PRIMARY_LEDGER_CONFIG.Name`.")

        # Legal Entities (XLE master — fallback names & unassigned LEs)
        df = read_csv_from_zip(z, "XLE_ENTITY_PROFILE.csv")
        if df is not None:
            name_col  = pick_col(df, ["Name", "LegalEntityName", "EntityName"])
            ident_col = pick_col(df, ["LegalEntityIdentifier", "LEIdentifier"])
            if name_col:
                legal_entity_names |= set(df[name_col].dropna().map(str).str.strip())
            if name_col and ident_col:
                for _, r in df[[name_col, ident_col]].dropna(how="all").iterrows():
                    nm  = str(r[name_col]).strip()
                    ident = str(r[ident_col]).strip()
                    # Only set if we don't already have ObjectName for this ident
                    if ident and nm and ident not in ident_to_le_name:
                        ident_to_le_name[ident] = nm
            else:
                st.warning(f"`XLE_ENTITY_PROFILE.csv` missing needed columns. Found: {list(df.columns)}")

        # Ledger ↔ LE identifier
        df = read_csv_from_zip(z, "ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv")
        if df is not None:
            led_col   = pick_col(df, ["GL_LEDGER.Name", "LedgerName"])
            ident_col = pick_col(df, ["LegalEntityIdentifier", "LEIdentifier"])
            if led_col and ident_col:
                for _, r in df[[led_col, ident_col]].dropna(how="all").iterrows():
                    led = str(r[led_col]).strip()
                    ident = str(r[ident_col]).strip()
                    if led and ident:
                        ledger_to_idents.setdefault(led, set()).add(ident)
            else:
                st.warning(f"`ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv` missing needed columns. Found: {list(df.columns)}")

        # Prefer ObjectName for identifier → name
        df = read_csv_from_zip(z, "ORA_GL_JOURNAL_CONFIG_DETAIL.csv")
        if df is not None:
            ident_col = pick_col(df, ["LegalEntityIdentifier", "LEIdentifier"])
            obj_col   = pick_col(df, ["ObjectName", "Name"])
            if ident_col and obj_col:
                for _, r in df[[ident_col, obj_col]].dropna(how="all").iterrows():
                    ident = str(r[ident_col]).strip()
                    obj   = str(r[obj_col]).strip()
                    if ident and obj:
                        ident_to_le_name[ident] = obj  # ObjectName takes precedence
            else:
                st.warning(f"`ORA_GL_JOURNAL_CONFIG_DETAIL.csv` missing needed columns. Found: {list(df.columns)}")

        # Business Units (for Tab 1)
        df = read_csv_from_zip(z, "FUN_BUSINESS_UNIT.csv")
        if df is not None:
            bu_col  = pick_col(df, ["Name"])
            led_col = pick_col(df, ["PrimaryLedgerName"])
            le_col  = pick_col(df, ["LegalEntityName"])
            if bu_col and led_col and le_col:
                for _, r in df[[bu_col, led_col, le_col]].dropna(how="all").iterrows():
                    bu_rows.append({
                        "Name": str(r[bu_col]).strip(),
                        "PrimaryLedgerName": str(r[led_col]).strip(),
                        "LegalEntityName": str(r[le_col]).strip()
                    })
            else:
                st.warning(f"`FUN_BUSINESS_UNIT.csv` missing needed columns. Found: {list(df.columns)}")

        # Cost Orgs (MASTER)
        df = read_csv_from_zip(z, "CST_COST_ORGANIZATION.csv")
        if df is not None:
            name_col   = pick_col(df, ["Name"])
            ident_col  = pick_col(df, ["LegalEntityIdentifier", "LEIdentifier"])
            join_col   = pick_col(df, ["OrgInformation2"])  # join to BOOKS + IO relationships
            if name_col and ident_col and join_col:
                for _, r in df[[name_col, ident_col, join_col]].dropna(how="all").iterrows():
                    name  = str(r[name_col]).strip()
                    ident = str(r[ident_col]).strip()
                    joink = str(r[join_col]).strip()
                    costorg_rows.append({"Name": name, "LegalEntityIdentifier": ident, "JoinKey": joink})
            else:
                st.warning(f"`CST_COST_ORGANIZATION.csv` missing needed columns (Name, LegalEntityIdentifier, OrgInformation2). Found: {list(df.columns)}")

        # Cost Books — JoinKey(CostOrgCode) -> {CostBookCode}
        df = read_csv_from_zip(z, "CST_COST_ORG_BOOK.csv")
        if df is not None:
            key_col  = pick_col(df, ["ORA_CST_ACCT_COST_ORG.CostOrgCode", "CostOrgCode"])
            book_col = pick_col(df, ["CostBookCode"])
            if key_col and book_col:
                for _, r in df[[key_col, book_col]].dropna(how="all").iterrows():
                    joink = str(r[key_col]).strip()
                    book  = str(r[book_col]).strip()
                    if joink and book:
                        books_by_joinkey.setdefault(joink, set()).add(book)
            else:
                st.warning(f"`CST_COST_ORG_BOOK.csv` missing needed columns (CostOrgCode, CostBookCode). Found: {list(df.columns)}")

        # Inventory Orgs (MASTER)
        df = read_csv_from_zip(z, "INV_ORGANIZATION_PARAMETER.csv")
        if df is not None:
            code_col  = pick_col(df, ["OrganizationCode"])
            name_col  = pick_col(df, ["Name", "OrganizationName"])
            le_col    = pick_col(df, ["LegalEntityIdentifier", "LEIdentifier"])
            bu_col    = pick_col(df, ["BusinessUnitName"])
            pcbu_col  = pick_col(df, ["ProfitCenterBuName"])
            mfg_col   = pick_col(df, ["MfgPlantFlag"])
            if code_col and name_col:
                for _, r in df.dropna(how="all").iterrows():
                    invorg_rows.append({
                        "Code": str(r.get(code_col, "")).strip(),
                        "Name": str(r.get(name_col, "")).strip(),
                        "LEIdent": str(r.get(le_col, "")).strip(),
                        "BUName": str(r.get(bu_col, "")).strip(),
                        "PCBU": str(r.get(pcbu_col, "")).strip(),
                        "Mfg": "Yes" if str(r.get(mfg_col, "")).strip().upper() == "Y" else ""
                    })
            else:
                st.warning(f"`INV_ORGANIZATION_PARAMETER.csv` missing needed columns. Found: {list(df.columns)}")

        # Cost Org ↔ Inventory Org relationships
        df = read_csv_from_zip(z, "ORA_CST_COST_ORG_INV.csv")
        if df is not None:
            inv_col  = pick_col(df, ["OrganizationCode", "InventoryOrganizationCode"])
            co_col   = pick_col(df, ["ORA_CST_ACCT_COST_ORG.CostOrgCode", "CostOrgCode"])
            if inv_col and co_col:
                for _, r in df[[inv_col, co_col]].dropna(how="all").iterrows():
                    inv_code, co_code = str(r[inv_col]).strip(), str(r[co_col]).strip()
                    if inv_code and co_code:
                        invorg_rel[inv_code] = co_code
            else:
                st.warning(f"`ORA_CST_COST_ORG_INV.csv` missing needed columns (OrganizationCode, CostOrgCode). Found: {list(df.columns)}")

    # ------------ Derived maps & display-name disambiguation ------------
    def name_for_ident(ident: str) -> str:
        nm = (ident_to_le_name.get(ident, "") or "").strip()
        return nm if nm else f"[LE {ident}]"

    ident_to_ledgers = {}
    for led, idents in ledger_to_idents.items():
        for ident in idents:
            ident_to_ledgers.setdefault(ident, set()).add(led)

    ledger_to_names = {}
    for led, idents in ledger_to_idents.items():
        for ident in idents:
            nm = name_for_ident(ident)
            ledger_to_names.setdefault(led, []).append((ident, nm))

    # Disambiguate same display name under the same ledger
    display_name_by_ident_per_ledger = {}
    for led, pairs in ledger_to_names.items():
        buckets = {}
        for ident, nm in pairs:
            buckets.setdefault(nm, []).append(ident)
        for nm, ids in buckets.items():
            if len(ids) == 1:
                display_name_by_ident_per_ledger[(led, ids[0])] = nm
            else:
                name_collisions.append((led, nm, sorted(ids)))
                for ident in ids:
                    display_name_by_ident_per_ledger[(led, ident)] = f"{nm} [LE {ident}]"

    # For EG-1 style back-fill
    le_to_ledgers = {}
    for led, pairs in ledger_to_names.items():
        for ident, nm in pairs:
            if nm.startswith("[LE "):
                continue
            le_to_ledgers.setdefault(nm, set()).add(led)

    # Also build name -> idents (for unassigned unique mapping)
    name_to_idents = {}
    for ident, nm in ident_to_le_name.items():
        if nm:
            name_to_idents.setdefault(nm, set()).add(ident)

    # ===================================================
    # Tab 1: Ledger – Legal Entity – Identifier – Business Unit
    # ===================================================
    rows1 = []
    seen_triples = set()
    seen_ledgers_with_bu = set()
    seen_les_with_bu = set()

    # Reverse name→one ident (best-effort, first seen)
    reverse_name_to_ident = {}
    for ident, nm in ident_to_le_name.items():
        if nm and nm not in reverse_name_to_ident:
            reverse_name_to_ident[nm] = ident

    # 1) BU-driven rows with unique back-fill
    for r in bu_rows:
        bu  = r["Name"]
        led = r["PrimaryLedgerName"]
        le  = r["LegalEntityName"]

        led = led if led in ledger_names else ""
        le  = le if le else ""

        if not led and le and le in le_to_ledgers and len(le_to_ledgers[le]) == 1:
            led = next(iter(le_to_ledgers[le]))
        if not le and led and led in ledger_to_names and len(ledger_to_names[led]) == 1:
            le = ledger_to_names[led][0][1]

        ident = reverse_name_to_ident.get(le, "")
        disp  = display_name_by_ident_per_ledger.get((led, ident), le)

        rows1.append({
            "Ledger Name": led,
            "Legal Entity": disp,
            "Legal Entity Identifier": ident,
            "Business Unit": bu
        })
        seen_triples.add((led, ident, bu))
        if led: seen_ledgers_with_bu.add(led)
        if le:  seen_les_with_bu.add(le)

    # 2) Ledger–LE pairs with no BU (by IDENT)
    seen_pairs = {(a, b) for (a, b, _) in seen_triples}
    for led, idents in ledger_to_idents.items():
        if not idents:
            if led not in seen_ledgers_with_bu:
                rows1.append({"Ledger Name": led, "Legal Entity": "", "Legal Entity Identifier": "", "Business Unit": ""})
            continue
        for ident in sorted(idents):
            if (led, ident) not in seen_pairs:
                disp = display_name_by_ident_per_ledger.get((led, ident), name_for_ident(ident))
                rows1.append({
                    "Ledger Name": led,
                    "Legal Entity": disp,
                    "Legal Entity Identifier": ident,
                    "Business Unit": ""
                })

    # 3) Orphan ledgers (exist in master, not mapped, no BU)
    mapped_ledgers = set(ledger_to_idents.keys())
    for led in sorted(ledger_names - mapped_ledgers - seen_ledgers_with_bu):
        rows1.append({"Ledger Name": led, "Legal Entity": "", "Legal Entity Identifier": "", "Business Unit": ""})

    # 4) Unassigned LEs (no ledger, no BU); fill ident if uniquely known for that name
    les_known = set(ident_to_le_name.values()) | set(legal_entity_names)
    les_in_bu = {r["LegalEntityName"] for r in bu_rows if r.get("LegalEntityName")}
    names_in_map = {nm for pairs in ledger_to_names.values() for _, nm in pairs}
    unassigned_les = sorted(les_known - les_in_bu - names_in_map)
    for le in unassigned_les:
        cand_ids = name_to_idents.get(le, set())
        ident = next(iter(cand_ids)) if len(cand_ids) == 1 else ""
        rows1.append({"Ledger Name": "", "Legal Entity": le, "Legal Entity Identifier": ident, "Business Unit": ""})

    # ---------- Build DF with EFFECTIVE KEY (fix for hanging LEs) ----------
    df1 = pd.DataFrame(rows1)

    def _ekey(row):
        ident = (row.get("Legal Entity Identifier") or "").strip()
        if ident:
            return f"ID::{ident}"
        name = (row.get("Legal Entity") or "").strip()
        return f"NAME::{name}"

    df1["__LEKey"] = df1.apply(_ekey, axis=1)
    df1 = df1.drop_duplicates(subset=["Ledger Name", "__LEKey", "Business Unit"]).reset_index(drop=True)

    # Sort (push blanks)
    df1["__LedgerEmpty"] = (df1["Ledger Name"] == "").astype(int)
    df1 = (
        df1.sort_values(
            ["__LedgerEmpty", "Ledger Name", "Legal Entity", "Business Unit"],
            ascending=[True, True, True, True]
        )
        .drop(columns=["__LedgerEmpty", "__LEKey"])
        .reset_index(drop=True)
    )
    df1.insert(0, "Assignment", range(1, len(df1) + 1))

    # ===================================================
    # Tab 2: Ledger – LE – Cost Org – Cost Book – Inventory Org – ProfitCenter BU – Management BU – Mfg Plant
    # ===================================================
    rows2 = []

    co_name_by_joinkey = {r["JoinKey"]: r["Name"] for r in costorg_rows if r.get("JoinKey")}

    for inv in invorg_rows:
        code = inv.get("Code", "")
        name = inv.get("Name", "")
        le_ident = inv.get("LEIdent", "")
        le_name  = name_for_ident(le_ident) if le_ident else ""
        leds     = ident_to_ledgers.get(le_ident, set()) if le_ident else set()

        co_key  = invorg_rel.get(code, "")
        co_name = co_name_by_joinkey.get(co_key, "") if co_key else ""
        books   = "; ".join(sorted(books_by_joinkey.get(co_key, []))) if co_key else ""

        if leds:
            for led in sorted(leds):
                rows2.append({
                    "Ledger Name": led,
                    "Legal Entity": le_name,
                    "Cost Organization": co_name,
                    "Cost Book": books,
                    "Inventory Org": name,
                    "Profit Center BU": inv.get("PCBU", ""),
                    "Management BU": inv.get("BUName", ""),
                    "Manufacturing Plant": inv.get("Mfg", "")
                })
        else:
            rows2.append({
                "Ledger Name": "",
                "Legal Entity": le_name,
                "Cost Organization": co_name,
                "Cost Book": books,
                "Inventory Org": name,
                "Profit Center BU": inv.get("PCBU", ""),
                "Management BU": inv.get("BUName", ""),
                "Manufacturing Plant": inv.get("Mfg", "")
            })

    df2 = pd.DataFrame(rows2).drop_duplicates().reset_index(drop=True)

    if not df2.empty:
        df2["__LedgerEmpty"] = (df2["Ledger Name"].fillna("") == "").astype(int)
        df2["__COEmpty"]     = (df2["Cost Organization"].fillna("") == "").astype(int)
        df2 = (
            df2.sort_values(
                ["__LedgerEmpty", "Ledger Name", "Legal Entity", "__COEmpty", "Cost Organization", "Inventory Org"],
                ascending=[True, True, True, True, True, True]
            )
            .drop(columns=["__LedgerEmpty", "__COEmpty"])
            .reset_index(drop=True)
        )
    df2.insert(0, "Assignment", range(1, len(df2) + 1))

    # ----------- normalize blanks -----------
    df1 = _blankify(df1)
    df2 = _blankify(df2)

    # ----------------- Diagnostics -----------------
    for led, idents in ledger_to_idents.items():
        for ident in idents:
            if name_for_ident(ident).startswith("[LE "):
                unresolved_ident_pairs.append((led, ident))

    diag_rows = []
    for led, ident in unresolved_ident_pairs:
        diag_rows.append({
            "Issue": "Unresolved LE Identifier (name missing in XLE/backup)",
            "Ledger": led,
            "LegalEntityIdentifier": ident,
            "Used Name": f"[LE {ident}]"
        })
    for led, nm, id_list in name_collisions:
        diag_rows.append({
            "Issue": f"Duplicate LE display name under ledger (split across identifiers)",
            "Ledger": led,
            "LegalEntityIdentifier": "; ".join(id_list),
            "Used Name": nm
        })
    df_diag = pd.DataFrame(diag_rows) if diag_rows else pd.DataFrame(
        [{"Issue": "No issues detected", "Ledger": "", "LegalEntityIdentifier": "", "Used Name": ""}]
    )

    # ------------ Excel Output ------------
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        df1.to_excel(writer, index=False, sheet_name="Ledger_LE_BU_Assignments")
        df2.to_excel(writer, index=False, sheet_name="Ledger_LE_CostOrg_IOs")
        df_diag.to_excel(writer, index=False, sheet_name="Diagnostics")

    st.success(f"Built {len(df1)} Tab-1 rows and {len(df2)} Tab-2 rows. (No lost hangers; identifier-safe.)")
    st.dataframe(df1.head(25), use_container_width=True, height=300)
    st.dataframe(df2.head(25), use_container_width=True, height=320)
    st.dataframe(df_diag.head(25), use_container_width=True, height=220)

    st.download_button(
        "⬇️ Download Excel (EnterpriseStructure.xlsx)",
        data=excel_buf.getvalue(),
        file_name="EnterpriseStructure.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ===================== DRAW.IO DIAGRAM BLOCK (Hard-coded 3 swimlanes per LE) =====================
if (
    "df1" in locals() and isinstance(df1, pd.DataFrame) and not df1.empty and
    "df2" in locals() and isinstance(df2, pd.DataFrame)
):
    import xml.etree.ElementTree as ET
    import zlib, base64, uuid

    def _make_drawio_xml(df_bu: pd.DataFrame, df_tab2: pd.DataFrame) -> str:
        # ---------------- layout constants ----------------
        W, H         = 180, 48
        LANE_OFFSET  = 260      # distance from LE center to lane centers
        BU_SPREAD    = 170
        CO_SPREAD    = 200
        DIO_SPREAD   = 170
        LEDGER_GAP   = 380
        CLUSTER_GAP  = 420
        LEFT_PAD     = 260

        # vertical rows
        Y_LEDGER     = 150
        Y_LE         = 320
        Y_BU         = 480
        Y_CO         = 640
        Y_CB         = 800
        Y_IO         = 960

        # styles
        S_LEDGER = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE6E6;strokeColor=#C86868;fontSize=12;"
        S_LE     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE2C2;strokeColor=#A66000;fontSize=12;"
        S_BU     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFF1B3;strokeColor=#B38F00;fontSize=12;"
        S_CO     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#C2F0C2;strokeColor=#008000;fontSize=12;"
        S_CB     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#A0D080;strokeColor=#004d00;fontSize=12;"
        S_IO     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#C2E0F9;strokeColor=#004080;fontSize=12;"

        S_EDGE   = "endArrow=block;rounded=1;edgeStyle=orthogonalEdgeStyle;orthogonal=1;jettySize=auto;" \
                   "strokeColor=#666666;exitX=0.5;exitY=0;entryX=0.5;entryY=1;"

        # --- normalize data ---
        df_bu = df_bu.fillna("").copy()
        df_tab2 = df_tab2.fillna("").copy()

        # ledger->LE mapping
        ledger_to_les = {}
        for _, r in df_bu.iterrows():
            L, E = r["Ledger Name"], r["Legal Entity"]
            if L and E:
                ledger_to_les.setdefault(L, set()).add(E)

        # collect BUs by LE
        bu_by_le = {}
        for _, r in df_bu.iterrows():
            L, E, B = r["Ledger Name"], r["Legal Entity"], r["Business Unit"]
            if L and E and B:
                bu_by_le.setdefault((L, E), []).append(B)

        # collect COs & children
        co_by_le, cb_by_co, io_by_co = {}, {}, {}
        for _, r in df_tab2.iterrows():
            L, E, C, CB, IO = r["Ledger Name"], r["Legal Entity"], r["Cost Organization"], r["Cost Book"], r["Inventory Org"]
            if L and E and C:
                co_by_le.setdefault((L, E), []).append(C)
                if CB:
                    cb_by_co.setdefault((L, E, C), []).append(CB)
                if IO:
                    io_by_co.setdefault((L, E, C), []).append(IO)

        # direct IOs (not under CO)
        dio_by_le = {}
        for _, r in df_tab2.iterrows():
            L, E, C, IO = r["Ledger Name"], r["Legal Entity"], r["Cost Organization"], r["Inventory Org"]
            if L and E and (not C) and IO:
                dio_by_le.setdefault((L, E), []).append(IO)

        # --- assign coordinates ---
        next_x = LEFT_PAD
        led_x, le_x, bu_x, co_x, cb_x, io_x, dio_x = {}, {}, {}, {}, {}, {}, {}

        for L in sorted(ledger_to_les.keys()):
            les = sorted(ledger_to_les[L])
            le_positions = []
            for E in les:
                cx = next_x
                le_x[(L, E)] = cx
                # lanes relative to LE
                bu_lane = cx - LANE_OFFSET
                co_lane = cx
                dio_lane = cx + LANE_OFFSET
                # spread BUs
                for j, b in enumerate(sorted(bu_by_le.get((L, E), []))):
                    bu_x[(L, E, b)] = bu_lane + j * BU_SPREAD
                # spread COs
                for j, c in enumerate(sorted(co_by_le.get((L, E), []))):
                    co_x[(L, E, c)] = co_lane + j * CO_SPREAD
                    for k, cb in enumerate(sorted(cb_by_co.get((L, E, c), []))):
                        cb_x[(L, E, c, cb)] = co_x[(L, E, c)] - 120 + k * 160
                    for m, io in enumerate(sorted(io_by_co.get((L, E, c), []))):
                        io_x[(L, E, c, io)] = co_x[(L, E, c)] + m * 160
                # spread Direct IOs
                for j, io in enumerate(sorted(dio_by_le.get((L, E), []))):
                    dio_x[(L, E, io)] = dio_lane + j * DIO_SPREAD

                le_positions.append(cx)
                next_x += LEDGER_GAP
            if le_positions:
                led_x[L] = sum(le_positions)//len(le_positions)
            else:
                led_x[L] = next_x
            next_x += CLUSTER_GAP

        # --- XML skeleton ---
        mxfile  = ET.Element("mxfile", attrib={"host": "app.diagrams.net"})
        diagram = ET.SubElement(mxfile, "diagram", attrib={"id": str(uuid.uuid4()), "name": "Enterprise Structure"})
        model   = ET.SubElement(diagram, "mxGraphModel", attrib={
            "dx":"1284","dy":"682","grid":"1","gridSize":"10","page":"1","pageWidth":"1920","pageHeight":"1080","background":"#ffffff"
        })
        root    = ET.SubElement(model, "root")
        ET.SubElement(root, "mxCell", attrib={"id":"0"})
        ET.SubElement(root, "mxCell", attrib={"id":"1","parent":"0"})

        def add_vertex(label, style, x, y, w=W, h=H):
            vid = uuid.uuid4().hex[:8]
            c = ET.SubElement(root,"mxCell",attrib={"id":vid,"value":label,"style":style,"vertex":"1","parent":"1"})
            ET.SubElement(c,"mxGeometry",attrib={"x":str(int(x)),"y":str(int(y)),"width":str(w),"height":str(h),"as":"geometry"})
            return vid

        def add_edge(src, tgt):
            eid = uuid.uuid4().hex[:8]
            c = ET.SubElement(root,"mxCell",attrib={"id":eid,"value":"","style":S_EDGE,"edge":"1","parent":"1","source":src,"target":tgt})
            ET.SubElement(c,"mxGeometry",attrib={"relative":"1","as":"geometry"})

        id_map = {}
        # ledgers
        for L in sorted(led_x.keys()):
            id_map[("L", L)] = add_vertex(L, S_LEDGER, led_x[L], Y_LEDGER)
        # LEs
        for (L,E), x in le_x.items():
            id_map[("E",L,E)] = add_vertex(E, S_LE, x, Y_LE)
            add_edge(id_map[("E",L,E)], id_map[("L",L)])
        # BUs
        for (L,E,b), x in bu_x.items():
            id_map[("B",L,E,b)] = add_vertex(b, S_BU, x, Y_BU)
            add_edge(id_map[("B",L,E,b)], id_map[("E",L,E)])
        # COs
        for (L,E,c), x in co_x.items():
            id_map[("C",L,E,c)] = add_vertex(c, S_CO, x, Y_CO)
            add_edge(id_map[("C",L,E,c)], id_map[("E",L,E)])
        # CBs
        for (L,E,c,cb), x in cb_x.items():
            id_map[("CB",L,E,c,cb)] = add_vertex(cb, S_CB, x, Y_CB)
            add_edge(id_map[("CB",L,E,c,cb)], id_map[("C",L,E,c)])
        # IOs under CO
        for (L,E,c,io), x in io_x.items():
            id_map[("IO",L,E,c,io)] = add_vertex(io, S_IO, x, Y_IO)
            add_edge(id_map[("IO",L,E,c,io)], id_map[("C",L,E,c)])
        # Direct IOs
        for (L,E,io), x in dio_x.items():
            id_map[("DIO",L,E,io)] = add_vertex(io, S_IO, x, Y_IO)
            add_edge(id_map[("DIO",L,E,io)], id_map[("E",L,E)])

        # legend
        def add_legend(x=20, y=20):
            panel_w, panel_h = 220, 160
            add_vertex("", "rounded=1;fillColor=#FFFFFF;strokeColor=#CBD5E1;", x, y, panel_w, panel_h)
            items = [
                ("Ledger", "#FFE6E6"),
                ("Legal Entity", "#FFE2C2"),
                ("Business Unit", "#FFF1B3"),
                ("Cost Org", "#C2F0C2"),
                ("Cost Book", "#A0D080"),
                ("Inventory Org", "#C2E0F9")
            ]
            for i,(lbl,col) in enumerate(items):
                box_y = y+28 + i*22
                box = add_vertex("", f"rounded=1;fillColor={col};strokeColor=#666666;", x+12, box_y, 18, 12)
                text = add_vertex(lbl, "text;align=left;verticalAlign=middle;fontSize=12;", x+36, box_y-4, 150, 20)

        add_legend()

        return ET.tostring(mxfile, encoding="utf-8", method="xml").decode("utf-8")

    def _drawio_url_from_xml(xml: str) -> str:
        raw = zlib.compress(xml.encode("utf-8"), level=9)[2:-4]
        b64 = base64.b64encode(raw).decode("ascii")
        return f"https://app.diagrams.net/?title=EnterpriseStructure.drawio#R{b64}"

    _xml = _make_drawio_xml(df1, df2)

    st.download_button(
        "⬇️ Download diagram (.drawio)",
        data=_xml.encode("utf-8"),
        file_name="EnterpriseStructure.drawio",
        mime="application/xml",
        use_container_width=True
    )
    st.markdown(f"[🔗 Open in draw.io (preview)]({_drawio_url_from_xml(_xml)})")

