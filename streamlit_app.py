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

# ===================== DRAW.IO DIAGRAM BLOCK (Swimlanes + Hybrid BU + Min IO Gap + Umbrella Gap + Low Elbows) =====================
if (
    "df1" in locals() and isinstance(df1, pd.DataFrame) and not df1.empty and
    "df2" in locals() and isinstance(df2, pd.DataFrame)
):
    import xml.etree.ElementTree as ET
    import zlib, base64, uuid

    def _make_drawio_xml(df_bu: pd.DataFrame, df_tab2: pd.DataFrame) -> str:
        # ----- geometry -----
        W, H = 180, 48

        # Rows
        Y_LEDGER = 150
        Y_LE     = 320
        Y_BU     = 480
        Y_CO     = 640
        Y_CB     = 800
        Y_IO     = 960

        # Low elbows (closer to the child row, but still between rows)
        def low_elbow(y_child, y_parent, bias=0.75):
            return int(y_parent + (y_child - y_parent) * bias)

        ELBOW_LE_TO_LED = low_elbow(Y_LE, Y_LEDGER)   # LE -> Ledger
        ELBOW_BU_TO_LE  = low_elbow(Y_BU, Y_LE)       # BU -> LE
        ELBOW_CO_TO_LE  = low_elbow(Y_CO, Y_LE)       # CO -> LE
        ELBOW_CB_TO_CO  = low_elbow(Y_CB, Y_CO)       # Cost Book -> CO
        ELBOW_IO_TO_CO  = low_elbow(Y_IO, Y_CO)       # IO (under CO) -> CO
        ELBOW_DIO_TO_LE = low_elbow(Y_IO, Y_LE)       # Direct IO -> LE

        # Lanes (relative to LE center)
        BOOK_OFFSET_LEFT = 140        # cost books sit left of CO
        BU_LANE_LEFT     = BOOK_OFFSET_LEFT
        DIO_LANE_RIGHT   = BOOK_OFFSET_LEFT

        # Spacing inside lanes (respect min gap so boxes never touch)
        MIN_GAP = 40
        BU_SPREAD_BASE   = 190
        CO_SPREAD_BASE   = 220
        DIO_SPREAD_BASE  = 190
        IO_UNDER_CO_BASE = 170
        BOOK_SPREAD_BASE = 160

        def spread(base): return max(base, W + MIN_GAP)

        # Inter-block layout
        LEDGER_BLOCK_GAP = 120
        CLUSTER_GAP      = 360
        LEFT_PAD         = 260

        # NEW: Minimum spacing between neighboring LE “umbrellas”
        MIN_UMBRELLA_GAP = 120

        # ----- styles -----
        S_LEDGER = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE6E6;strokeColor=#C86868;fontSize=12;"
        S_LE     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE2C2;strokeColor=#A66000;fontSize=12;"
        S_BU     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFF1B3;strokeColor=#B38F00;fontSize=12;"
        S_CO     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#C2F0C2;strokeColor=#008000;fontSize=12;"
        S_CB     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#A0D080;strokeColor=#004d00;fontSize=12;"
        S_IO     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#C2E0F9;strokeColor=#004080;fontSize=12;"
        S_IO_PLT = "rounded=1;whiteSpace=wrap;html=1;fillColor=#C2E0F9;strokeColor=#1F4D7A;strokeWidth=2;fontSize=12;"

        S_EDGE = ("endArrow=block;rounded=1;edgeStyle=orthogonalEdgeStyle;orthogonal=1;"
                  "jettySize=auto;strokeColor=#666666;exitX=0.5;exitY=0;entryX=0.5;entryY=1;")

        # ----- normalize source data -----
        df_bu = df_bu.fillna("").copy()
        df2   = df_tab2.fillna("").copy()

        # Build ledger → {LE} from BOTH tabs
        ledger_to_les = {}
        for _, r in pd.concat([
            df_bu[["Ledger Name", "Legal Entity"]],
            df2[["Ledger Name", "Legal Entity"]],
        ]).drop_duplicates().iterrows():
            L, E = r["Ledger Name"], r["Legal Entity"]
            if L and E:
                ledger_to_les.setdefault(L, set()).add(E)

        # BU by LE
        bu_by_le = {}
        for _, r in df_bu.iterrows():
            L, E, B = r["Ledger Name"], r["Legal Entity"], r["Business Unit"]
            if L and E and B:
                bu_by_le.setdefault((L, E), set()).add(B)

        # CO, Books, IO-under-CO
        co_by_le, cb_by_co, io_by_co = {}, {}, {}
        for _, r in df2.iterrows():
            L, E, C = r["Ledger Name"], r["Legal Entity"], r["Cost Organization"]
            CB, IO, MFG = r["Cost Book"], r["Inventory Org"], (r.get("Manufacturing Plant","") or "")
            if L and E and C:
                co_by_le.setdefault((L, E), set()).add(C)
                if CB:
                    for bk in [b.strip() for b in CB.split(";") if b.strip()]:
                        cb_by_co.setdefault((L, E, C), set()).add(bk)
                if IO:
                    rec = {"Name": IO, "Mfg": str(MFG).strip().lower() in ("yes","y","true","1")}
                    io_by_co.setdefault((L, E, C), [])
                    if all(x["Name"] != rec["Name"] for x in io_by_co[(L, E, C)]):
                        io_by_co[(L, E, C)].append(rec)

        # Direct IOs (no CO)
        dio_by_le = {}
        for _, r in df2.iterrows():
            L, E, C, IO = r["Ledger Name"], r["Legal Entity"], r["Cost Organization"], r["Inventory Org"]
            MFG = (r.get("Manufacturing Plant","") or "")
            if L and E and not C and IO:
                rec = {"Name": IO, "Mfg": str(MFG).strip().lower() in ("yes","y","true","1")}
                dio_by_le.setdefault((L, E), [])
                if all(x["Name"] != rec["Name"] for x in dio_by_le[(L, E)]):
                    dio_by_le[(L, E)].append(rec)

        # helpers
        def centered_positions(center_x, n, base_spread):
            s = spread(base_spread)
            if n <= 0: return []
            if n == 1: return [center_x]
            start = center_x - (s * (n - 1)) / 2.0
            return [start + i * s for i in range(n)]

        # ----- place nodes -----
        next_x = LEFT_PAD
        led_x  = {}
        le_x   = {}
        bu_x   = {}
        co_x   = {}
        cb_x   = {}
        io_x   = {}   # (L,E,C,io) -> (x, is_mfg)
        dio_x  = {}   # (L,E,io)   -> (x, is_mfg)

        for L in sorted(ledger_to_les.keys()):
            centers = []
            prev_umbrella_max_x = None

            for E in sorted(ledger_to_les[L]):
                cx = next_x  # tentative center

                # Determine lane centers
                co_center  = cx                               # CO lane = center
                bu_center  = cx - BU_LANE_LEFT                # BU lane left
                dio_center = cx + DIO_LANE_RIGHT              # direct IO right

                # Hybrid rule: if no COs & no Direct IOs, center BUs under LE
                has_co  = len(co_by_le.get((L, E), [])) > 0
                has_dio = len(dio_by_le.get((L, E), [])) > 0
                if not has_co and not has_dio:
                    bu_center = cx

                # BUs
                bu_list = sorted(bu_by_le.get((L, E), []))
                bu_positions = centered_positions(bu_center, len(bu_list), BU_SPREAD_BASE)
                for x, b in zip(bu_positions, bu_list):
                    bu_x[(L, E, b)] = x

                # COs (+ books left, IOs under)
                co_list = sorted(co_by_le.get((L, E), []))
                co_positions = centered_positions(co_center, len(co_list), CO_SPREAD_BASE)
                for x, c in zip(co_positions, co_list):
                    co_x[(L, E, c)] = x
                    # Books
                    books = sorted(cb_by_co.get((L, E, c), []))
                    for k, bk in enumerate(books):
                        cb_x[(L, E, c, bk)] = x - BOOK_OFFSET_LEFT + k * spread(BOOK_SPREAD_BASE)
                    # IO under this CO
                    ios = sorted(io_by_co.get((L, E, c), []), key=lambda d: d["Name"])
                    io_positions = centered_positions(x, len(ios), IO_UNDER_CO_BASE)
                    for xio, rec in zip(io_positions, ios):
                        io_x[(L, E, c, rec["Name"])] = (xio, rec["Mfg"])

                # Direct IOs
                dios = sorted(dio_by_le.get((L, E), []), key=lambda d: d["Name"])
                dio_positions = centered_positions(dio_center, len(dios), DIO_SPREAD_BASE)
                for x, rec in zip(dio_positions, dios):
                    dio_x[(L, E, rec["Name"])] = (x, rec["Mfg"])

                # Tentative LE center
                le_x[(L, E)] = cx

                # Compute umbrella span for this LE
                xs = [cx]
                xs += bu_positions
                xs += co_positions
                xs += [v[0] for k, v in dio_x.items() if k[:2] == (L, E)]
                for c in co_list:
                    xs += [cb_x[(L, E, c, bk)] for bk in sorted(cb_by_co.get((L, E, c), []))]
                    xs += [io_x[(L, E, c, rec["Name"])][0] for rec in sorted(io_by_co.get((L, E, c), []), key=lambda d: d["Name"])]

                min_x = min(xs) - W/2
                max_x = max(xs) + W/2

                # Ensure umbrella gap vs. previous LE
                if prev_umbrella_max_x is not None and min_x < prev_umbrella_max_x + MIN_UMBRELLA_GAP:
                    shift = (prev_umbrella_max_x + MIN_UMBRELLA_GAP) - min_x
                    # Shift LE center and all of its children in this ledger/LE
                    le_x[(L, E)] = cx + shift
                    # BU
                    for k in list(bu_x):
                        if k[0] == L and k[1] == E:
                            bu_x[k] += shift
                    # CO
                    for k in list(co_x):
                        if k[0] == L and k[1] == E:
                            co_x[k] += shift
                    # Cost Books
                    for k in list(cb_x):
                        if k[0] == L and k[1] == E:
                            cb_x[k] += shift
                    # IO under CO
                    for k in list(io_x):
                        if k[0] == L and k[1] == E:
                            io_x[k] = (io_x[k][0] + shift, io_x[k][1])
                    # Direct IOs
                    for k in list(dio_x):
                        if k[0] == L and k[1] == E:
                            dio_x[k] = (dio_x[k][0] + shift, dio_x[k][1])

                    # Recompute span after shift
                    xs = [le_x[(L, E)]]
                    xs += [bu_x[(L, E, b)] for b in bu_list]
                    xs += [co_x[(L, E, c)] for c in co_list]
                    xs += [v[0] for k, v in dio_x.items() if k[:2] == (L, E)]
                    for c in co_list:
                        xs += [cb_x[(L, E, c, bk)] for bk in sorted(cb_by_co.get((L, E, c), []))]
                        xs += [io_x[(L, E, c, rec["Name"])][0] for rec in sorted(io_by_co.get((L, E, c), []), key=lambda d: d["Name"])]
                    min_x = min(xs) - W/2
                    max_x = max(xs) + W/2

                # Advance placement cursors
                prev_umbrella_max_x = max_x
                next_x = max_x + LEDGER_BLOCK_GAP
                centers.append(le_x[(L, E)])

            led_x[L] = int(sum(centers)/len(centers)) if centers else next_x
            next_x += CLUSTER_GAP

        # ----- XML skeleton -----
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

        def add_edge_with_elbow(src_id, tgt_id, src_center_x, tgt_center_x, elbow_y):
            eid = uuid.uuid4().hex[:8]
            c = ET.SubElement(root, "mxCell", attrib={
                "id": eid, "value": "", "style": S_EDGE, "edge": "1", "parent": "1",
                "source": src_id, "target": tgt_id
            })
            g = ET.SubElement(c, "mxGeometry", attrib={"relative": "1", "as": "geometry"})
            arr = ET.SubElement(g, "Array", attrib={"as": "points"})
            ET.SubElement(arr, "mxPoint", attrib={"x": str(int(src_center_x)), "y": str(int(elbow_y))})
            ET.SubElement(arr, "mxPoint", attrib={"x": str(int(tgt_center_x)), "y": str(int(elbow_y))})

        def cx(x_left):  # left x -> center x
            return int(x_left + W/2)

        id_map = {}
        # Ledgers
        for L in sorted(led_x.keys()):
            id_map[("L", L)] = add_vertex(L, S_LEDGER, led_x[L], Y_LEDGER)

        # LEs
        for (L, E), x in le_x.items():
            id_map[("E", L, E)] = add_vertex(E, S_LE, x, Y_LE)
            add_edge_with_elbow(id_map[("E", L, E)], id_map[("L", L)], cx(x), cx(led_x[L]), ELBOW_LE_TO_LED)

        # BUs
        for (L, E, b), x in bu_x.items():
            id_map[("B", L, E, b)] = add_vertex(b, S_BU, x, Y_BU)
            add_edge_with_elbow(id_map[("B", L, E, b)], id_map[("E", L, E)], cx(x), cx(le_x[(L, E)]), ELBOW_BU_TO_LE)

        # COs
        for (L, E, c), x in co_x.items():
            id_map[("C", L, E, c)] = add_vertex(c, S_CO, x, Y_CO)
            add_edge_with_elbow(id_map[("C", L, E, c)], id_map[("E", L, E)], cx(x), cx(le_x[(L, E)]), ELBOW_CO_TO_LE)

        # Cost Books
        for (L, E, c, bk), x in cb_x.items():
            id_map[("CB", L, E, c, bk)] = add_vertex(bk, S_CB, x, Y_CB)
            add_edge_with_elbow(id_map[("CB", L, E, c, bk)], id_map[("C", L, E, c)], cx(x), cx(co_x[(L, E, c)]), ELBOW_CB_TO_CO)

        # IOs under CO
        for (L, E, c, name), (x, is_mfg) in io_x.items():
            style = S_IO_PLT if is_mfg else S_IO
            label = f"🏭 {name}" if is_mfg else name
            id_map[("IO", L, E, c, name)] = add_vertex(label, style, x, Y_IO)
            add_edge_with_elbow(id_map[("IO", L, E, c, name)], id_map[("C", L, E, c)], cx(x), cx(co_x[(L, E, c)]), ELBOW_IO_TO_CO)

        # Direct IOs
        for (L, E, name), (x, is_mfg) in dio_x.items():
            style = S_IO_PLT if is_mfg else S_IO
            label = f"🏭 {name}" if is_mfg else name
            id_map[("DIO", L, E, name)] = add_vertex(label, style, x, Y_IO)
            add_edge_with_elbow(id_map[("DIO", L, E, name)], id_map[("E", L, E)], cx(x), cx(le_x[(L, E)]), ELBOW_DIO_TO_LE)

        # Legend
        def add_legend(x=20, y=20):
            _ = add_vertex("", "rounded=1;fillColor=#FFFFFF;strokeColor=#CBD5E1;", x, y, 250, 184)
            items = [
                ("Ledger", "#FFE6E6", None),
                ("Legal Entity", "#FFE2C2", None),
                ("Business Unit", "#FFF1B3", None),
                ("Cost Org", "#C2F0C2", None),
                ("Cost Book", "#A0D080", None),
                ("Inventory Org", "#C2E0F9", None),
                ("Manufacturing Plant (IO)", "#C2E0F9", "bold"),
            ]
            for i, (lbl, col, flavor) in enumerate(items):
                style = f"rounded=1;fillColor={col};strokeColor=#666666;"
                if flavor == "bold":
                    style = "rounded=1;fillColor=#C2E0F9;strokeColor=#1F4D7A;strokeWidth=2;"
                add_vertex("", style, x+12, y+28+i*22, 18, 12)
                add_vertex(lbl, "text;align=left;verticalAlign=middle;fontSize=12;", x+36, y+24+i*22, 200, 20)

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
# ===================== END DRAW.IO BLOCK =====================



