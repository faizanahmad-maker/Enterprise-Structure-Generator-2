# --- Enterprise Structure Generator: Excel + draw.io (unassigned LEs/BUs handled) ---

import io, zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Enterprise Structure Generator", page_icon="📊", layout="wide")
st.title("Enterprise Structure Generator — Excel + draw.io")

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
    """Return the first matching column from `candidates` (exact > case-insensitive > substring)."""
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
    df = df.copy()
    df = df.fillna("")
    obj_cols = list(df.select_dtypes(include=["object"]).columns)
    for c in obj_cols:
        s = df[c]
        mask = s.apply(lambda x: isinstance(x, str) and x.strip().lower() == "nan")
        if mask.any():
            df.loc[mask, c] = ""
    return df

def _norm_key(x: str) -> str:
    s = str(x or "").strip().lower()
    return " ".join(s.split())

if not uploads:
    st.info("Upload your ZIPs to generate the Excel & diagram.")
else:
    # ------------ Collectors ------------
    ledger_names = set()
    legal_entity_names = set()
    ledger_to_idents = {}            # ledger -> {LE identifier}
    ident_to_le_name = {}            # LE identifier -> LE name
    bu_rows = []                     # BU rows (for Tab 1 only)

    # Cost Orgs (MASTER)
    costorg_rows = []                # [{Name, LegalEntityIdentifier, JoinKey}]
    costorg_name_to_joinkeys = {}    # Name -> {JoinKey}
    # Cost Books: JoinKey -> {CostBookCode}
    books_by_joinkey = {}

    # Inventory Orgs (MASTER)
    invorg_rows = []                 # [{Code, Name, LEIdent, BUName, PCBU, Mfg}]
    # IO↔CostOrg relationships: InvOrgCode -> CostOrgJoinKey
    invorg_rel = {}

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

        # Legal Entities
        df = read_csv_from_zip(z, "XLE_ENTITY_PROFILE.csv")
        if df is not None:
            name_col  = pick_col(df, ["Name"])
            if name_col:
                legal_entity_names |= set(df[name_col].dropna().map(str).str.strip())
            else:
                st.warning(f"`XLE_ENTITY_PROFILE.csv` missing `Name`. Found: {list(df.columns)}")

        # Ledger ↔ LE identifier
        df = read_csv_from_zip(z, "ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv")
        if df is not None:
            led_col   = pick_col(df, ["GL_LEDGER.Name"])
            ident_col = pick_col(df, ["LegalEntityIdentifier"])
            if led_col and ident_col:
                for _, r in df[[led_col, ident_col]].dropna(how="all").iterrows():
                    led = str(r[led_col]).strip()
                    ident = str(r[ident_col]).strip()
                    if led and ident:
                        ledger_to_idents.setdefault(led, set()).add(ident)
            else:
                st.warning(f"`ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv` missing needed columns. Found: {list(df.columns)}")

        # Backup map for identifier -> LE name
        df = read_csv_from_zip(z, "ORA_GL_JOURNAL_CONFIG_DETAIL.csv")
        if df is not None:
            ident_col = pick_col(df, ["LegalEntityIdentifier"])
            obj_col   = pick_col(df, ["ObjectName"])
            if ident_col and obj_col:
                for _, r in df[[ident_col, obj_col]].dropna(how="all").iterrows():
                    ident = str(r[ident_col]).strip()
                    obj   = str(r[obj_col]).strip()
                    if ident and obj and ident not in ident_to_le_name:
                        ident_to_le_name[ident] = obj

        # Business Units (keep raw values; we'll backfill smartly later)
        df = read_csv_from_zip(z, "FUN_BUSINESS_UNIT.csv")
        if df is not None:
            bu_col  = pick_col(df, ["Name"])
            led_col = pick_col(df, ["PrimaryLedgerName"])
            le_col  = pick_col(df, ["LegalEntityName"])
            if bu_col and led_col and le_col:
                for _, r in df[[bu_col, led_col, le_col]].dropna(how="all").iterrows():
                    bu_rows.append({
                        "Name": str(r.get(bu_col, "")).strip(),
                        "PrimaryLedgerName": str(r.get(led_col, "")).strip(),
                        "LegalEntityName": str(r.get(le_col, "")).strip()
                    })
            else:
                st.warning(f"`FUN_BUSINESS_UNIT.csv` missing needed columns. Found: {list(df.columns)}")

        # Cost Orgs (MASTER)
        df = read_csv_from_zip(z, "CST_COST_ORGANIZATION.csv")
        if df is not None:
            name_col   = pick_col(df, ["Name"])
            ident_col  = pick_col(df, ["LegalEntityIdentifier"])
            join_col   = pick_col(df, ["OrgInformation2"])  # join to BOOKS + IO relationships
            if name_col and ident_col and join_col:
                for _, r in df[[name_col, ident_col, join_col]].dropna(how="all").iterrows():
                    name  = str(r[name_col]).strip()
                    ident = str(r[ident_col]).strip()
                    joink = str(r[join_col]).strip()
                    costorg_rows.append({"Name": name, "LegalEntityIdentifier": ident, "JoinKey": joink})
                    if name and joink:
                        costorg_name_to_joinkeys.setdefault(name, set()).add(joink)
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

    # ------------ Derived maps ------------
    ident_to_ledgers = {}
    for led, idents in ledger_to_idents.items():
        for ident in idents:
            ident_to_ledgers.setdefault(ident, set()).add(led)

    ledger_to_le_names = {}
    for led, idents in ledger_to_idents.items():
        for ident in idents:
            le_name = ident_to_le_name.get(ident, "").strip()
            if le_name:
                ledger_to_le_names.setdefault(led, set()).add(le_name)

    known_pairs = set()
    for led, idents in ledger_to_idents.items():
        for ident in idents:
            le_name = ident_to_le_name.get(ident, "").strip()
            if le_name:
                known_pairs.add((led, le_name))

    # ===================================================
    # Tab 1: Core Enterprise Structure (WITH hanging LEs)
    # ===================================================
    from collections import defaultdict

    # 1) LE display names union
    le_names_master = set(legal_entity_names)
    le_names_from_ident = {v for v in ident_to_le_name.values() if v}
    all_le_display = set()
    all_le_display |= {n for n in le_names_master if n}
    all_le_display |= {n for n in le_names_from_ident if n}

    norm2display = {}
    for n in all_le_display:
        k = _norm_key(n)
        if k and k not in norm2display:
            norm2display[k] = n

    # 2) LE→Ledgers via identifier mapping
    le_to_ledgers = defaultdict(set)
    for led, idents in ledger_to_idents.items():
        for ident in idents:
            nm = ident_to_le_name.get(ident, "")
            if nm:
                le_to_ledgers[_norm_key(nm)].add(led)

    # 3) Build rows
    rows1 = []
    seen_pairs_norm = set()     # (ledger, norm(LE))
    seen_ledgers_with_bu = set()
    seen_les_with_bu_norm = set()

    # BU-driven rows (+ smart backfill)
    for r in bu_rows:
        bu  = str(r.get("Name", "")).strip()
        led = str(r.get("PrimaryLedgerName", "")).strip()
        le  = str(r.get("LegalEntityName", "")).strip()

        # Backfill ledger from LE if missing and unique
        if not led and le:
            ks = le_to_ledgers.get(_norm_key(le), set())
            if len(ks) == 1:
                led = next(iter(ks))

        # Backfill LE from ledger if missing and unique
        if not le and led:
            candidates = [disp for k, disp in norm2display.items() if led in le_to_ledgers.get(k, set())]
            if len(candidates) == 1:
                le = candidates[0]

        rows1.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": bu})

        if led:
            seen_ledgers_with_bu.add(led)
        if le:
            seen_les_with_bu_norm.add(_norm_key(le))
        if led and le:
            seen_pairs_norm.add((led, _norm_key(le)))

    # Add ledger–LE pairs without BU (from mapping)
    for led, idents in ledger_to_idents.items():
        for ident in idents:
            le_name = ident_to_le_name.get(ident, "").strip()
            if not le_name:
                continue
            if (led, _norm_key(le_name)) not in seen_pairs_norm:
                rows1.append({"Ledger Name": led, "Legal Entity": le_name, "Business Unit": ""})

    # Orphan ledgers
    mapped_ledgers = set(ledger_to_le_names.keys())
    for led in sorted(ledger_names - mapped_ledgers - seen_ledgers_with_bu):
        rows1.append({"Ledger Name": led, "Legal Entity": "", "Business Unit": ""})

    # Orphan (hanging) LEs
    all_le_norm = {_norm_key(n) for n in all_le_display if n}
    covered_le_norm = {k for (_, k) in seen_pairs_norm} | seen_les_with_bu_norm
    for k in sorted(all_le_norm - covered_le_norm):
        display = norm2display.get(k, "")
        ledgers_for_le = le_to_ledgers.get(k, set())
        led_guess = next(iter(ledgers_for_le)) if len(ledgers_for_le) == 1 else ""
        rows1.append({"Ledger Name": led_guess, "Legal Entity": display, "Business Unit": ""})

    df1 = pd.DataFrame(rows1).drop_duplicates().reset_index(drop=True)
    df1["__LedgerEmpty"] = (df1["Ledger Name"].fillna("") == "").astype(int)
    df1 = (
        df1.sort_values(
            ["__LedgerEmpty", "Ledger Name", "Legal Entity", "Business Unit"],
            ascending=[True, True, True, True]
        )
        .drop(columns="__LedgerEmpty")
        .reset_index(drop=True)
    )
    df1.insert(0, "Assignment", range(1, len(df1) + 1))

    # ===================================================
    # Tab 2: Ledger – LE – Cost Org – Cost Book – Inventory Org – PCB / Mgt BU / Mfg
    # ===================================================
    rows2 = []

    # Helper: Cost Org name from JoinKey
    co_name_by_joinkey = {r["JoinKey"]: r["Name"] for r in costorg_rows if r.get("JoinKey")}

    for inv in invorg_rows:
        code = inv.get("Code", "")
        name = inv.get("Name", "")
        le_ident = inv.get("LEIdent", "")
        le_name  = ident_to_le_name.get(le_ident, "") if le_ident else ""
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

    # ----------- turn all NaNs/‘nan’ into blanks for both tabs -----------
    df1 = _blankify(df1)
    df2 = _blankify(df2)

    # ------------ Excel Output ------------
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        df1.to_excel(writer, index=False, sheet_name="Core Enterprise Structure")
        df2.to_excel(writer, index=False, sheet_name="Inventory Org Structure")

    st.success(f"Built {len(df1)} Core rows and {len(df2)} Inventory Org rows (hanging LEs handled).")
    st.dataframe(df1.head(25), use_container_width=True, height=280)
    st.dataframe(df2.head(25), use_container_width=True, height=320)

    st.download_button(
        "⬇️ Download Excel (EnterpriseStructure.xlsx)",
        data=excel_buf.getvalue(),
        file_name="EnterpriseStructure.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ===================== DRAW.IO DIAGRAM (unassigned on far-right) =====================
    if (
        "df1" in locals() and isinstance(df1, pd.DataFrame) and not df1.empty and
        "df2" in locals() and isinstance(df2, pd.DataFrame)
    ):
        import xml.etree.ElementTree as ET
        import zlib, base64, uuid

        def _make_drawio_xml(df_bu: pd.DataFrame, df_tab2: pd.DataFrame) -> str:
            # --- layout & spacing ---
            W, H           = 180, 48
            X_STEP         = 230
            IO_GAP         = 40
            IO_STEP        = max(180, W + IO_GAP)

            LEDGER_PAD     = 320
            PAD_GROUP      = 60
            LEFT_PAD       = 260
            RIGHT_PAD      = 300  # extra space before the far-right parking lot

            Y_LEDGER   = 150
            Y_LE       = 310
            Y_BU       = 470
            Y_CO       = 630
            Y_CB       = 790
            Y_IO       = 1060

            # --- styles ---
            S_LEDGER = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE6E6;strokeColor=#C86868;fontSize=12;"
            S_LE     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE2C2;strokeColor=#A66000;fontSize=12;"
            S_BU     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFF1B3;strokeColor=#B38F00;fontSize=12;"
            S_CO     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#E2F7E2;strokeColor=#3D8B3D;fontSize=12;"
            S_CB     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#7FBF7F;strokeColor=#2F7D2F;fontSize=12;"
            S_IO     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#D6EFFF;strokeColor=#2F71A8;fontSize=12;"
            S_IO_PLT = "rounded=1;whiteSpace=wrap;html=1;fillColor=#D6EFFF;strokeColor=#1F4D7A;strokeWidth=2;fontSize=12;"

            S_EDGE   = ("endArrow=block;rounded=1;edgeStyle=orthogonalEdgeStyle;orthogonal=1;"
                        "jettySize=auto;strokeColor=#666666;exitX=0.5;exitY=0;entryX=0.5;entryY=1;")
            S_HDR    = "text;align=left;verticalAlign=middle;fontSize=13;fontStyle=1;"

            # --- normalize input ---
            df_bu = df_bu[["Ledger Name", "Legal Entity", "Business Unit"]].copy()
            for c in df_bu.columns:
                df_bu[c] = df_bu[c].fillna("").map(str).str.strip()

            df = df_tab2[[
                "Ledger Name","Legal Entity","Cost Organization","Cost Book",
                "Inventory Org","Manufacturing Plant"
            ]].copy()
            for c in df.columns:
                df[c] = df[c].fillna("").map(str).str.strip()

            ledgers_all = sorted({*df_bu["Ledger Name"].unique(), *df["Ledger Name"].unique()} - {""})

            # --- maps ---
            le_map, bu_map, co_map = {}, {}, {}
            cb_by_co = {}   # (L,E,C) -> [book,...]
            io_by_co = {}   # (L,E,C) -> [{"Name":..., "Mfg":...}, ...]

            tmp = pd.concat([df_bu[["Ledger Name","Legal Entity"]],
                             df[["Ledger Name","Legal Entity"]]]).drop_duplicates()
            for _, r in tmp.iterrows():
                L, E = r["Ledger Name"], r["Legal Entity"]
                if L and E:
                    le_map.setdefault(L, set()).add(E)

            for _, r in df_bu.iterrows():
                L, E, B = r["Ledger Name"], r["Legal Entity"], r["Business Unit"]
                if L and E and B:
                    bu_map.setdefault((L,E), set()).add(B)

            for _, r in df.iterrows():
                L, E, C = r["Ledger Name"], r["Legal Entity"], r["Cost Organization"]
                if L and E and C:
                    co_map.setdefault((L,E), set()).add(C)

            for _, r in df.iterrows():
                L, E, C = r["Ledger Name"], r["Legal Entity"], r["Cost Organization"]
                B, IO, MFG = r["Cost Book"], r["Inventory Org"], r["Manufacturing Plant"]
                if L and E and C and B:
                    for bk in [b.strip() for b in B.split(";") if b.strip()]:
                        cb_by_co.setdefault((L,E,C), []).append(bk)
                if L and E and C and IO:
                    io_by_co.setdefault((L,E,C), [])
                    rec = {"Name": IO, "Mfg": (MFG or "")}
                    if all(x["Name"] != IO for x in io_by_co[(L,E,C)]):  # de-dup
                        io_by_co[(L,E,C)].append(rec)

            # --- x coordinates with guaranteed ledger padding ---
            next_x = LEFT_PAD
            led_x, le_x, bu_x, co_x, cb_x, io_x = {}, {}, {}, {}, {}, {}

            for L in ledgers_all:
                ledger_x_used = []
                les = sorted(le_map.get(L, []))

                if not les:
                    led_x[L] = next_x
                    ledger_x_used.append(next_x)
                    next_x = next_x + LEDGER_PAD
                else:
                    for le in les:
                        buses = sorted(bu_map.get((L, le), []))
                        cos   = sorted(co_map.get((L, le), []))

                        if not buses and not cos:
                            le_x[(L, le)] = next_x; ledger_x_used.append(next_x); next_x += X_STEP
                        else:
                            for b in buses:
                                if b not in bu_x:
                                    bu_x[b] = next_x; ledger_x_used.append(next_x); next_x += X_STEP
                            for c in cos:
                                if c not in co_x:
                                    co_x[c] = next_x; ledger_x_used.append(next_x); next_x += X_STEP

                            xs = [bu_x[b] for b in buses] + [co_x[c] for c in cos]
                            le_center = int(sum(xs)/len(xs)) if xs else next_x
                            le_x[(L, le)] = le_center
                            ledger_x_used.append(le_center)

                        # Books left of CO; IO under CO
                        for c in cos:
                            base = co_x[c]

                            books = sorted(dict.fromkeys(cb_by_co.get((L, le, c), [])))
                            for i, bk in enumerate(books, start=1):
                                x_pos = base - i*X_STEP
                                cb_x[(L, le, c, bk)] = x_pos
                                ledger_x_used.append(x_pos)

                            ios = sorted(io_by_co.get((L, le, c), []), key=lambda k: k["Name"])
                            n = len(ios)
                            if n == 1:
                                io_x[(L, le, c, ios[0]["Name"])] = base
                                ledger_x_used.append(base)
                            elif n > 1:
                                start = base - ((n - 1) * IO_STEP) // 2
                                for j, io in enumerate(ios):
                                    x_pos = start + j*IO_STEP
                                    io_x[(L, le, c, io["Name"])] = x_pos
                                    ledger_x_used.append(x_pos)

                    xs_led = [le_x[(L, le)] for le in les]
                    led_x[L] = int(sum(xs_led)/len(xs_led)) if xs_led else next_x
                    ledger_x_used.append(led_x[L])

                    max_x_used = max(ledger_x_used) if ledger_x_used else next_x
                    next_x = max(next_x + PAD_GROUP, max_x_used + LEDGER_PAD)

            # --- FAR-RIGHT parking lot: unassigned LEs and BUs from df_bu ---
            unassigned_les = sorted(
                {r["Legal Entity"] for _, r in df_bu.iterrows()
                 if (r["Legal Entity"] and not r["Ledger Name"])}
            )
            # BU is considered unassigned if its ledger or LE is blank
            unassigned_bus = sorted(
                {r["Business Unit"] for _, r in df_bu.iterrows()
                 if (r["Business Unit"] and (not r["Ledger Name"] or not r["Legal Entity"]))}
            )

            start_parking_x = (max(led_x.values()) + RIGHT_PAD) if led_x else (next_x + RIGHT_PAD)
            px = start_parking_x  # x cursor for parking lane
            le_parking_x = {}
            bu_parking_x = {}

            # Allocate LEs first, then BUs
            for e in unassigned_les:
                le_parking_x[e] = px
                px += X_STEP
            # small gap between LE and BU columns
            px += PAD_GROUP
            for b in unassigned_bus:
                bu_parking_x[b] = px
                px += X_STEP

            # --- XML skeleton ---
            mxfile  = ET.Element("mxfile", attrib={"host": "app.diagrams.net"})
            diagram = ET.SubElement(mxfile, "diagram", attrib={"id": str(uuid.uuid4()), "name": "Enterprise Structure"})
            model   = ET.SubElement(diagram, "mxGraphModel", attrib={
                "dx": "1284", "dy": "682", "grid": "1", "gridSize": "10",
                "page": "1", "pageWidth": "1920", "pageHeight": "1080",
                "background": "#ffffff"
            })
            root    = ET.SubElement(model, "root")
            ET.SubElement(root, "mxCell", attrib={"id": "0"})
            ET.SubElement(root, "mxCell", attrib={"id": "1", "parent": "0"})

            def add_vertex(label, style, x, y):
                vid = uuid.uuid4().hex[:8]
                c = ET.SubElement(root, "mxCell", attrib={
                    "id": vid, "value": label, "style": style, "vertex": "1", "parent": "1"})
                ET.SubElement(c, "mxGeometry", attrib={
                    "x": str(int(x)), "y": str(int(y)), "width": str(W), "height": str(H), "as": "geometry"})
                return vid

            def add_edge(src, tgt):
                eid = uuid.uuid4().hex[:8]
                c = ET.SubElement(root, "mxCell", attrib={
                    "id": eid, "value": "", "style": S_EDGE, "edge": "1", "parent": "1",
                    "source": src, "target": tgt})
                ET.SubElement(c, "mxGeometry", attrib={"relative": "1", "as": "geometry"})

            # --- vertices ---
            id_map = {}
            for L in ledgers_all:
                id_map[("L", L)] = add_vertex(L, S_LEDGER, led_x[L], Y_LEDGER)

                for le in sorted(le_map.get(L, [])):
                    id_map[("E", L, le)] = add_vertex(le, S_LE, le_x[(L, le)], Y_LE)

                    for b in sorted(bu_map.get((L, le), [])):
                        id_map[("B", L, le, b)] = add_vertex(b, S_BU, bu_x[b], Y_BU)

                    for c in sorted(co_map.get((L, le), [])):
                        id_map[("C", L, le, c)] = add_vertex(c, S_CO, co_x[c], Y_CO)

                        for bk in sorted(set(cb_by_co.get((L, le, c), []))):
                            id_map[("CB", L, le, c, bk)] = add_vertex(bk, S_CB, cb_x[(L, le, c, bk)], Y_CB)

                        for io in sorted(io_by_co.get((L, le, c), []), key=lambda k: k["Name"]):
                            label = f"🏭 {io['Name']}" if str(io["Mfg"]).lower() == "yes" else io["Name"]
                            style = S_IO_PLT if str(io["Mfg"]).lower() == "yes" else S_IO
                            id_map[("IO", L, le, c, io["Name"])] = add_vertex(label, style, io_x[(L, le, c, io["Name"])], Y_IO)

            # Far-right parking LEs/BUs
            for e, x in le_parking_x.items():
                id_map[("E", "UNASSIGNED", e)] = add_vertex(e, S_LE, x, Y_LE)
            for b, x in bu_parking_x.items():
                id_map[("B", "UNASSIGNED", b)] = add_vertex(b, S_BU, x, Y_BU)

            # --- edges (top-center → bottom-center) ---
            for L in ledgers_all:
                for le in sorted(le_map.get(L, [])):
                    if ("E", L, le) in id_map:
                        add_edge(id_map[("E", L, le)], id_map[("L", L)])

                    for b in sorted(bu_map.get((L, le), [])):
                        k = ("B", L, le, b)
                        if k in id_map:
                            add_edge(id_map[k], id_map[("E", L, le)])

                    for c in sorted(co_map.get((L, le), [])):
                        kc = ("C", L, le, c)
                        if kc in id_map:
                            add_edge(id_map[kc], id_map[("E", L, le)])

                            for bk in sorted(set(cb_by_co.get((L, le, c), []))):
                                kcb = ("CB", L, le, c, bk)
                                if kcb in id_map:
                                    add_edge(id_map[kcb], id_map[kc])

                            for io in io_by_co.get((L, le, c), []):
                                kio = ("IO", L, le, c, io["Name"])
                                if kio in id_map:
                                    add_edge(id_map[kio], id_map[kc])

            # Edges inside parking lot: BU → LE if they share same LE name
            # (if BU row had an LE name)
            for _, r in df_bu.iterrows():
                led, le, bu = r["Ledger Name"], r["Legal Entity"], r["Business Unit"]
                if bu and (not led or not le) and le:
                    kB = ("B", "UNASSIGNED", bu)
                    kE = ("E", "UNASSIGNED", le)
                    if (kB in id_map) and (kE in id_map):
                        add_edge(id_map[kB], id_map[kE])

            # --- legend ---
            def add_legend(x=20, y=20):
                def swatch(lbl, color, offset, stroke="#666666", bold=False):
                    style = f"rounded=1;fillColor={color};strokeColor={stroke};"
                    if bold: style += "strokeWidth=2;"
                    box = ET.SubElement(root, "mxCell", attrib={
                        "id": uuid.uuid4().hex[:8], "value": "",
                        "style": style, "vertex": "1", "parent": "1"})
                    ET.SubElement(box, "mxGeometry", attrib={
                        "x": str(x+12), "y": str(y+offset), "width": "18", "height": "12", "as": "geometry"})
                    txt = ET.SubElement(root, "mxCell", attrib={
                        "id": uuid.uuid4().hex[:8], "value": lbl,
                        "style": "text;align=left;verticalAlign=middle;fontSize=12;",
                        "vertex": "1", "parent": "1"})
                    ET.SubElement(txt, "mxGeometry", attrib={
                        "x": str(x+36), "y": str(y+offset-4), "width": "250", "height": "20", "as": "geometry"})
                swatch("Ledger", "#FFE6E6", 36)
                swatch("Legal Entity", "#FFE2C2", 62)
                swatch("Business Unit", "#FFF1B3", 88)
                swatch("Cost Org", "#E2F7E2", 114)
                swatch("Cost Book (left of CO)", "#7FBF7F", 140)
                swatch("Inventory Org (under CO)", "#D6EFFF", 166, stroke="#2F71A8")
                swatch("Manufacturing Plant (IO)", "#D6EFFF", 192, stroke="#1F4D7A", bold=True)
                # parking lot note
                txt = ET.SubElement(root, "mxCell", attrib={
                    "id": uuid.uuid4().hex[:8], "value": "Unassigned LEs/BUs → far-right",
                    "style": "text;align=left;verticalAlign=middle;fontSize=11;",
                    "vertex": "1", "parent": "1"})
                ET.SubElement(txt, "mxGeometry", attrib={
                    "x": str(x+12), "y": str(y+220), "width": "260", "height": "20", "as": "geometry"})

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
