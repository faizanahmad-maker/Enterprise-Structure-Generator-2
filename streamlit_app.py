import io, zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Enterprise Structure Generator", page_icon="📊", layout="wide")
st.title("Enterprise Structure Generator — 3 Tabs + Primary Cost Book (Excel + draw.io)")

st.markdown("""
Upload your Oracle export ZIPs (any order). The app will auto-detect the files:
- **Manage Primary Ledgers** → GL_PRIMARY_LEDGER.csv, ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv, (backup) ORA_GL_JOURNAL_CONFIG_DETAIL.csv
- **Manage Legal Entity** → XLE_ENTITY_PROFILE.csv
- **Manage Business Unit** → FUN_BUSINESS_UNIT.csv
- **Manage Cost Organizations** → CST_COST_ORGANIZATION.csv
- **Cost Organization Relationships** → CST_COST_ORG_BOOK.csv, ORA_CST_COST_ORG_INV.csv
- **Manage Inventory Organizations** → INV_ORGANIZATION_PARAMETER.csv
""")

uploads = st.file_uploader("Drop ZIPs here", type="zip", accept_multiple_files=True)

# ---------- helpers ----------
def read_csv_from_zip(zf, name):
    if name not in zf.namelist():
        return None
    with zf.open(name) as fh:
        return pd.read_csv(fh, dtype=str)

def pick_col(df, candidates):
    """Return the first matching column from `candidates` (exact > case-insensitive > substring)."""
    if df is None: return None
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
    st.info("Upload the ZIPs to generate the Excel & diagram.")
else:
    # ------------ Collectors ------------
    ledger_names = set()
    legal_entity_names = set()
    ledger_to_idents = {}            # ledger -> {LE identifier}
    ident_to_le_name = {}            # LE identifier -> LE name
    bu_rows = []                     # BU rows for Tab 1 only

    # Cost Orgs (MASTER)
    costorg_rows = []                # [{Name, LegalEntityIdentifier, JoinKey}]
    co_code_to_name = {}             # CostOrg join code -> Name
    co_name_to_joinkeys = {}         # Name -> {JoinKey}

    # Cost Books: for each CostOrg joinkey -> {CostBookCode: is_primary}
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
            col = pick_col(df, ["ORA_GL_PRIMARY_LEDGER_CONFIG.Name", "Name"])
            if col:
                ledger_names |= set(df[col].dropna().map(str).str.strip())
            else:
                st.warning("`GL_PRIMARY_LEDGER.csv` missing ledger name column.")

        # Legal Entities
        df = read_csv_from_zip(z, "XLE_ENTITY_PROFILE.csv")
        if df is not None:
            name_col  = pick_col(df, ["Name"])
            ident_col = pick_col(df, ["LegalEntityIdentifier"])
            if name_col and ident_col:
                for _, r in df[[name_col, ident_col]].dropna(how="all").iterrows():
                    le_name = str(r[name_col]).strip()
                    le_ident = str(r[ident_col]).strip()
                    if le_name:
                        legal_entity_names.add(le_name)
                    if le_ident and le_name:
                        ident_to_le_name[le_ident] = le_name
            else:
                st.warning(f"`XLE_ENTITY_PROFILE.csv` missing needed columns. Found: {list(df.columns)}")

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
            ident_col  = pick_col(df, ["LegalEntityIdentifier"])
            join_col   = pick_col(df, ["OrgInformation2", "OrganizationCode"])
            if name_col and ident_col and join_col:
                for _, r in df[[name_col, ident_col, join_col]].dropna(how="all").iterrows():
                    name  = str(r[name_col]).strip()
                    ident = str(r[ident_col]).strip()
                    joink = str(r[join_col]).strip()
                    costorg_rows.append({"Name": name, "LegalEntityIdentifier": ident, "JoinKey": joink})
                    if joink and name:
                        co_code_to_name[joink] = name
                    if name and joink:
                        co_name_to_joinkeys.setdefault(name, set()).add(joink)
            else:
                st.warning(f"`CST_COST_ORGANIZATION.csv` missing needed columns (Name, LegalEntityIdentifier, OrgInformation2). Found: {list(df.columns)}")

        # Cost Books — JoinKey(CostOrgCode) -> {book: is_primary}
        df = read_csv_from_zip(z, "CST_COST_ORG_BOOK.csv")
        if df is not None:
            key_col  = pick_col(df, ["ORA_CST_ACCT_COST_ORG.CostOrgCode", "CostOrgCode"])
            book_col = pick_col(df, ["CostBookCode"])
            prim_col = pick_col(df, ["PrimaryBookFlag"])
            if key_col and book_col and prim_col:
                for _, r in df[[key_col, book_col, prim_col]].dropna(how="all").iterrows():
                    joink = str(r[key_col]).strip()
                    book  = str(r[book_col]).strip()
                    prim  = str(r[prim_col]).strip().upper() == "Y"
                    if joink and book:
                        d = books_by_joinkey.setdefault(joink, {})
                        # OR across duplicates
                        d[book] = d.get(book, False) or prim
            else:
                st.warning(f"`CST_COST_ORG_BOOK.csv` missing needed columns (CostOrgCode, CostBookCode, PrimaryBookFlag). Found: {list(df.columns)}")

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

    # CostOrg Name -> {Book: is_primary}
    co_name_to_book_primary = {}
    for joink, bookmap in books_by_joinkey.items():
        co_name = co_code_to_name.get(joink, "")
        if not co_name:
            continue
        m = co_name_to_book_primary.setdefault(co_name, {})
        for bk, prim in bookmap.items():
            m[bk] = m.get(bk, False) or bool(prim)

    # ===================================================
    # Tab 1: Core Enterprise Structure (Ledger – Legal Entity – Business Unit)
    # ===================================================
    rows1, seen_triples, seen_ledgers_with_bu = [], set(), set()

    # Emit BU-driven rows (strict)
    for r in bu_rows:
        bu  = r["Name"]
        led = r["PrimaryLedgerName"]
        le  = r["LegalEntityName"]
        rows1.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": bu})
        seen_triples.add((led, le, bu))
        if led:
            seen_ledgers_with_bu.add(led)

    # Add ledger–LE pairs from mapping that have no BU
    seen_pairs = {(a, b) for (a, b, _) in seen_triples}
    for led, le in sorted(known_pairs):
        if (led, le) not in seen_pairs:
            rows1.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": ""})

    # Orphan ledgers (exist, but no mapping & no BU)
    mapped_ledgers = set(ledger_to_le_names.keys())
    for led in sorted(ledger_names - mapped_ledgers - seen_ledgers_with_bu):
        rows1.append({"Ledger Name": led, "Legal Entity": "", "Business Unit": ""})

    df1 = pd.DataFrame(rows1).drop_duplicates().reset_index(drop=True)
    df1["__LedgerEmpty"] = (df1["Ledger Name"] == "").astype(int)
    df1 = (
        df1.sort_values(
            ["__LedgerEmpty", "Ledger Name", "Legal Entity", "Business Unit"],
            ascending=[True, True, True, True]
        )
        .drop(columns="__LedgerEmpty")
        .reset_index(drop=True)
    )
    df1.insert(0, "Assignment", range(1, len(df1) + 1))
    df1 = _blankify(df1)

    # ===================================================
    # Tab 2: Inventory Org Structure (Ledger – LE – (Cost Org?) – Inventory Org – PC BU – Mgmt BU – Mfg Plant)
    # ===================================================
    rows2 = []
    for inv in invorg_rows:
        code = inv.get("Code", "")
        name = inv.get("Name", "")
        le_ident = inv.get("LEIdent", "")
        le_name  = ident_to_le_name.get(le_ident, "") if le_ident else ""
        leds     = ident_to_ledgers.get(le_ident, set()) if le_ident else set()
        co_key   = invorg_rel.get(code, "")
        co_name  = co_code_to_name.get(co_key, "") if co_key else ""

        if leds:
            for led in sorted(leds):
                rows2.append({
                    "Ledger Name": led,
                    "Legal Entity": le_name,
                    "Cost Organization": co_name,
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
    df2 = _blankify(df2)

    # ===================================================
    # Tab 3: Costing Structure (Ledger – LE – Cost Org – Cost Book – Primary?)
    # ===================================================
    rows3 = []
    # Build: for each Cost Org, from its LE ident get leds; then books with primary flag
    for r in costorg_rows:
        co_name = r["Name"]
        le_ident = r["LegalEntityIdentifier"]
        le_name = ident_to_le_name.get(le_ident, "")
        leds = sorted(ident_to_ledgers.get(le_ident, set())) or [""]
        book_map = co_name_to_book_primary.get(co_name, {})
        if not book_map:
            # still emit a row to show the CO even if no books present
            for led in leds:
                rows3.append({
                    "Ledger Name": led,
                    "Legal Entity": le_name,
                    "Cost Organization": co_name,
                    "Cost Book": "",
                    "Primary Cost Book": ""
                })
        else:
            for led in leds:
                for bk, is_primary in sorted(book_map.items(), key=lambda kv: kv[0]):
                    rows3.append({
                        "Ledger Name": led,
                        "Legal Entity": le_name,
                        "Cost Organization": co_name,
                        "Cost Book": bk,
                        "Primary Cost Book": "Yes" if is_primary else "No"
                    })

    df3 = pd.DataFrame(rows3).drop_duplicates().reset_index(drop=True)
    if not df3.empty:
        df3["__LedgerEmpty"] = (df3["Ledger Name"].fillna("") == "").astype(int)
        df3 = (
            df3.sort_values(
                ["__LedgerEmpty", "Ledger Name", "Legal Entity", "Cost Organization", "Cost Book"],
                ascending=[True, True, True, True, True]
            )
            .drop(columns=["__LedgerEmpty"])
            .reset_index(drop=True)
        )
    df3.insert(0, "Assignment", range(1, len(df3) + 1))
    df3 = _blankify(df3)

    # ------------ Excel Output (3 tabs) ------------
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        df1.to_excel(writer, index=False, sheet_name="Core Enterprise Structure")
        df2.to_excel(writer, index=False, sheet_name="Inventory Org Structure")
        df3.to_excel(writer, index=False, sheet_name="Costing Structure")

    st.success(f"Built: Tab1 {len(df1)} rows • Tab2 {len(df2)} rows • Tab3 {len(df3)} rows.")
    st.dataframe(df1.head(20), use_container_width=True, height=240)
    st.dataframe(df2.head(20), use_container_width=True, height=280)
    st.dataframe(df3.head(20), use_container_width=True, height=280)

    st.download_button(
        "⬇️ Download Excel (EnterpriseStructure.xlsx)",
        data=excel_buf.getvalue(),
        file_name="EnterpriseStructure.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ===================== DRAW.IO DIAGRAM (direct-IO trunk pushed right + cost books restored) =====================
if (
    "df1" in locals() and isinstance(df1, pd.DataFrame) and not df1.empty and
    "df2" in locals() and isinstance(df2, pd.DataFrame)
):
    import xml.etree.ElementTree as ET
    import zlib, base64, uuid
    import pandas as pd
    from types import MappingProxyType

    # --- Try to find a Costing tab automatically (Tab 3). Prefer explicit df_cost; otherwise scan locals. ---
    df_cost = locals().get("df_cost", None)
    if (df_cost is None) or not isinstance(df_cost, pd.DataFrame) or df_cost.empty:
        for _name, _obj in list(locals().items()):
            if isinstance(_obj, pd.DataFrame) and not _obj.empty:
                cols = {c.strip().lower() for c in _obj.columns}
                # Heuristic signature for the Costing Structure sheet
                if ("cost organization" in cols or "cost org" in cols) and ("cost book" in cols):
                    df_cost = _obj.copy()
                    break

    def _make_drawio_xml(df_bu: pd.DataFrame, df_io: pd.DataFrame, df_costing: pd.DataFrame | None) -> str:
        # ---------- Geometry / layout ----------
        W, H = 180, 48
        Y_LEDGER, Y_LE, Y_BU, Y_CO, Y_CB, Y_IO = 150, 320, 480, 640, 800, 960

        def elbow(y_child, y_parent, bias=0.75):
            return int(y_parent + (y_child - y_parent) * bias)

        ELBOW_LE_TO_LED = elbow(Y_LE, Y_LEDGER)  # LE→Ledger bus height
        ELBOW_BU_TO_LE  = elbow(Y_BU, Y_LE)      # BU/Direct-IO horizontal
        ELBOW_CO_TO_LE  = elbow(Y_CO, Y_LE)      # CO horizontal
        ELBOW_IO_TO_CO  = elbow(Y_IO, Y_CO)      # CO-owned IO horizontal split

        # spacing
        MIN_GAP = 70
        def spread(base): return max(base, W + MIN_GAP)
        BU_SPREAD_BASE, CO_SPREAD_BASE    = 210, 230
        IO_UNDER_CO_BASE, BOOK_SPREAD_BASE = 190, 170
        LEDGER_BLOCK_GAP, CLUSTER_GAP, LEFT_PAD = 120, 420, 260
        MIN_UMBRELLA_GAP = 140

        # Direct-IO basin sits to the right of CO lane
        DIO_BASIN_GAP    = 200
        # NEW: push the direct-IO trunk a bit further right than the midpoint of those IO boxes
        TRUNK_RIGHT_BIAS = 90   # tweak to taste

        # ---------- Styles ----------
        S_LEDGER = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE6E6;strokeColor=#C86868;fontSize=12;"
        S_LE     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE2C2;strokeColor=#A66000;fontSize=12;"
        S_BU     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFF1B3;strokeColor=#B38F00;fontSize=12;"
        S_CO     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#C2F0C2;strokeColor=#008000;fontSize=12;"
        S_CB     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#A0D080;strokeColor=#004d00;fontSize=12;"
        S_CB_P   = "rounded=1;whiteSpace=wrap;html=1;fillColor=#A0D080;strokeColor=#004d00;strokeWidth=2;fontSize=12;"
        S_IO     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#C2E0F9;strokeColor=#004080;fontSize=12;"
        S_IO_PLT = "rounded=1;whiteSpace=wrap;html=1;fillColor=#C2E0F9;strokeColor=#1F4D7A;strokeWidth=2;fontSize=12;"
        S_EDGE   = ("endArrow=block;rounded=1;edgeStyle=orthogonalEdgeStyle;orthogonal=1;"
                    "jettySize=auto;strokeColor=#666666;exitX=0.5;exitY=0;entryX=0.5;entryY=1;")

        # ---------- Helpers ----------
        def pick(df, candidates):
            if df is None: return None
            for c in candidates:
                if c in df.columns: return c
                for col in df.columns:
                    if col.lower() == c.lower(): return col
            return None

        # ---------- Normalize inputs ----------
        df_bu = df_bu[["Ledger Name","Legal Entity","Business Unit"]].copy().fillna("").astype(str)

        # IO tab (Tab 2)
        LCOL = pick(df_io, ["Ledger Name","Ledger"])
        ECOL = pick(df_io, ["Legal Entity","LegalEntity"])
        COCOL= pick(df_io, ["Cost Organization","Cost Org","CostOrganization"])
        IOCOL= pick(df_io, ["Inventory Org","Inventory Organization","InventoryOrg","Inventory Organization Name"])
        MFGCOL=pick(df_io, ["Manufacturing Plant","Mfg","ManufacturingPlant","IsManufacturingPlant"])
        IO_BKC   = pick(df_io, ["Cost Book","CostBook","Cost Books"])             # optional in IO tab
        IO_BKPC  = pick(df_io, ["Primary Cost Book","PrimaryBook","Primary CostBook"])
        df_io = df_io[[x for x in [LCOL,ECOL,COCOL,IOCOL,MFGCOL,IO_BKC,IO_BKPC] if x is not None]].copy().fillna("").astype(str)
        df_io.rename(columns={LCOL:"Ledger Name", ECOL:"Legal Entity", COCOL:"Cost Organization",
                              IOCOL:"Inventory Org", MFGCOL:"Manufacturing Plant"}, inplace=True)
        if IO_BKC:  df_io.rename(columns={IO_BKC:"Cost Book"}, inplace=True)
        if IO_BKPC: df_io.rename(columns={IO_BKPC:"Primary Cost Book"}, inplace=True)

        # Costing tab (Tab 3) — optional but preferred for books/primary flag
        if df_costing is not None and not df_costing.empty:
            cLCOL = pick(df_costing, ["Ledger Name","Ledger"])
            cECOL = pick(df_costing, ["Legal Entity","LegalEntity"])
            cCO   = pick(df_costing, ["Cost Organization","Cost Org","CostOrganization"])
            cBKC  = pick(df_costing, ["Cost Book","CostBook"])
            cBKPC = pick(df_costing, ["Primary Cost Book","PrimaryBook","Primary Flag","PrimaryBookFlag"])
            df_costing = df_costing[[x for x in [cLCOL,cECOL,cCO,cBKC,cBKPC] if x is not None]].copy().fillna("").astype(str)
            df_costing.rename(columns={cLCOL:"Ledger Name", cECOL:"Legal Entity", cCO:"Cost Organization",
                                       cBKC:"Cost Book"}, inplace=True)
            if cBKPC: df_costing.rename(columns={cBKPC:"Primary Cost Book"}, inplace=True)
        else:
            df_costing = None

        ledgers_all = sorted({*df_bu["Ledger Name"].unique(), *df_io["Ledger Name"].unique()} - {""})

        # ---------- Build maps ----------
        le_map, bu_map, co_map = {}, {}, {}
        io_by_co, dio_by_le = {}, {}
        cb_by_co, cb_primary = {}, {}

        tmp = pd.concat([df_bu[["Ledger Name","Legal Entity"]], df_io[["Ledger Name","Legal Entity"]]]).drop_duplicates()
        for _, r in tmp.iterrows():
            L, E = r["Ledger Name"], r["Legal Entity"]
            if L and E: le_map.setdefault(L, set()).add(E)

        for _, r in df_bu.iterrows():
            L,E,B = r["Ledger Name"], r["Legal Entity"], r["Business Unit"]
            if L and E and B: bu_map.setdefault((L,E), set()).add(B)

        for _, r in df_io.iterrows():
            L,E,C = r.get("Ledger Name",""), r.get("Legal Entity",""), r.get("Cost Organization","")
            if L and E and C: co_map.setdefault((L,E), set()).add(C)

        for _, r in df_io.iterrows():
            L,E,C = r.get("Ledger Name",""), r.get("Legal Entity",""), r.get("Cost Organization","")
            IO,MFG = r.get("Inventory Org",""), r.get("Manufacturing Plant","")
            if L and E and IO:
                rec = {"Name": IO, "Mfg": (MFG or "")}
                if C:
                    io_by_co.setdefault((L,E,C), [])
                    if all(x["Name"] != IO for x in io_by_co[(L,E,C)]):
                        io_by_co[(L,E,C)].append(rec)
                else:
                    dio_by_le.setdefault((L,E), [])
                    if all(x["Name"] != IO for x in dio_by_le[(L,E)]):
                        dio_by_le[(L,E)].append(rec)

        # Books from IO tab (if present)
        if "Cost Book" in df_io.columns and df_io["Cost Book"].astype(str).str.strip().any():
            for _, r in df_io.iterrows():
                L,E,C = r.get("Ledger Name",""), r.get("Legal Entity",""), r.get("Cost Organization","")
                if not (L and E and C): continue
                if r.get("Cost Book","").strip():
                    books = [b.strip() for b in str(r["Cost Book"]).split(";") if b.strip()]
                    for bk in books:
                        cb_by_co.setdefault((L,E,C), [])
                        if bk not in cb_by_co[(L,E,C)]: cb_by_co[(L,E,C)].append(bk)
                        if "Primary Cost Book" in df_io.columns:
                            raw = str(r.get("Primary Cost Book","")).strip().lower()
                            is_primary = raw in ("yes","y","true","1")
                            cb_primary[(L,E,C,bk)] = is_primary

        # Books from Costing tab (preferred / overrides)
        if df_costing is not None:
            for _, r in df_costing.iterrows():
                L,E,C = r.get("Ledger Name",""), r.get("Legal Entity",""), r.get("Cost Organization","")
                bk    = r.get("Cost Book","").strip()
                if not (L and E and C and bk): continue
                cb_by_co.setdefault((L,E,C), [])
                if bk not in cb_by_co[(L,E,C)]: cb_by_co[(L,E,C)].append(bk)
                if "Primary Cost Book" in df_costing.columns:
                    raw = str(r.get("Primary Cost Book","")).strip().lower()
                    is_primary = raw in ("yes","y","true","1","primary")
                    cb_primary[(L,E,C,bk)] = is_primary

        # ---------- Placement ----------
        next_x = LEFT_PAD
        led_x, le_x, bu_x, co_x, io_x, dio_x, cb_x = {}, {}, {}, {}, {}, {}, {}

        def centered_positions(center_x, n, base_spread):
            s = spread(base_spread)
            if n <= 0: return []
            if n == 1: return [int(center_x)]
            start = center_x - (s * (n - 1)) / 2.0
            return [int(start + i * s) for i in range(n)]

        def cx(x_left): return int(x_left + W/2)

        prev_umbrella_max_x = None
        for L in ledgers_all:
            les = sorted(le_map.get(L, []))
            centers = []
            for E in les:
                cx_le_val = next_x
                le_x[(L,E)] = cx_le_val
                centers.append(cx_le_val)

                cos = sorted(co_map.get((L,E), []))
                has_direct_io = bool(dio_by_le.get((L,E), []))
                has_io_under_co = any(io_by_co.get((L,E,c)) for c in cos)
                has_any_co_or_io = bool(cos) or has_io_under_co or has_direct_io

                bu_center = cx_le_val if not has_any_co_or_io else (cx_le_val - 150)
                bu_list = sorted(bu_map.get((L,E), []))
                for x,b in zip(centered_positions(bu_center, len(bu_list), BU_SPREAD_BASE), bu_list):
                    bu_x[(L,E,b)] = x

                for x,c in zip(centered_positions(cx_le_val, len(cos), CO_SPREAD_BASE), cos):
                    co_x[(L,E,c)] = x
                    # books left of CO
                    books = sorted(dict.fromkeys(cb_by_co.get((L,E,c), [])))
                    for i,bk in enumerate(books, start=1):
                        cb_x[(L,E,c,bk)] = x - i * BOOK_SPREAD_BASE
                    # IOs under CO
                    ios = sorted(io_by_co.get((L,E,c), []), key=lambda d: d["Name"])
                    for xio, rec in zip(centered_positions(x, len(ios), IO_UNDER_CO_BASE), ios):
                        io_x[(L,E,c,rec["Name"])] = (xio, rec["Mfg"])

                # direct IOs
                dlist = sorted(dio_by_le.get((L,E), []), key=lambda d: d["Name"])
                if dlist:
                    xs_ref = [cx_le_val] + [co_x[(L,E,c)] for c in cos]
                    for c in cos:
                        xs_ref += [io_x[(L,E,c,r["Name"])][0] for r in io_by_co.get((L,E,c), [])]
                    right_edge = max(xs_ref) if xs_ref else cx_le_val
                    basin_center = right_edge + DIO_BASIN_GAP + W/2
                    for xio, rec in zip(centered_positions(basin_center, len(dlist), IO_UNDER_CO_BASE), dlist):
                        dio_x[(L,E,rec["Name"])] = (xio, rec["Mfg"])

                # umbrella guard
                xs_span = [cx_le_val]
                xs_span += [bu_x[(L,E,b)] for b in bu_list]
                xs_span += [co_x[(L,E,c)] for c in cos]
                for c in cos:
                    xs_span += [io_x[(L,E,c,r["Name"])][0] for r in io_by_co.get((L,E,c), [])]
                    xs_span += [cb_x[(L,E,c,bk)] for bk in cb_by_co.get((L,E,c), [])]
                xs_span += [v[0] for k,v in dio_x.items() if k[:2]==(L,E)]

                min_x = (min(xs_span) - W/2) if xs_span else cx_le_val - W/2
                max_x_ = (max(xs_span) + W/2) if xs_span else cx_le_val + W/2

                if prev_umbrella_max_x is not None and min_x < prev_umbrella_max_x + MIN_UMBRELLA_GAP:
                    shift = (prev_umbrella_max_x + MIN_UMBRELLA_GAP) - min_x
                    le_x[(L,E)] += shift
                    def shift_layer(d, cond):
                        for k in list(d.keys()):
                            if cond(k):
                                if d in (io_x, dio_x):
                                    d[k] = (d[k][0] + shift, d[k][1])
                                else:
                                    d[k] = d[k] + shift
                    shift_layer(bu_x, lambda k: k[0]==L and k[1]==E)
                    shift_layer(co_x, lambda k: k[0]==L and k[1]==E)
                    shift_layer(io_x, lambda k: k[0]==L and k[1]==E)
                    shift_layer(dio_x,lambda k: k[0]==L and k[1]==E)
                    shift_layer(cb_x, lambda k: k[0]==L and k[1]==E)
                    max_x_ += shift

                prev_umbrella_max_x = max_x_
                next_x = max_x_ + LEDGER_BLOCK_GAP

            led_x[L] = int(sum(centers)/len(centers)) if centers else next_x
            next_x += CLUSTER_GAP

        # ensure direct-IO basin doesn't bump CO-owned IOs
        for L in ledgers_all:
            for E in sorted(le_map.get(L, [])):
                co_ios = []
                for c in sorted(co_map.get((L,E), [])):
                    co_ios += [io_x[(L,E,c,r["Name"])][0] for r in io_by_co.get((L,E,c), [])]
                if not co_ios: continue
                max_co_right = max(x + W/2 for x in co_ios)
                dio_ios = [v[0] for k,v in dio_x.items() if k[:2]==(L,E)]
                if not dio_ios: continue
                min_dio_left = min(x - W/2 for x in dio_ios)
                if min_dio_left < max_co_right + MIN_GAP:
                    shift = (max_co_right + MIN_GAP) - min_dio_left
                    for k in list(dio_x.keys()):
                        if k[0]==L and k[1]==E:
                            dio_x[k] = (dio_x[k][0] + shift, dio_x[k][1])

        # ---------- XML ----------
        mxfile  = ET.Element("mxfile", attrib={"host":"app.diagrams.net"})
        diagram = ET.SubElement(mxfile, "diagram", attrib={"id":str(uuid.uuid4()), "name":"Enterprise Structure"})
        model   = ET.SubElement(diagram, "mxGraphModel", attrib={
            "dx":"1284","dy":"682","grid":"1","gridSize":"10",
            "page":"1","pageWidth":"1920","pageHeight":"1080",
            "background":"#ffffff"
        })
        root    = ET.SubElement(model, "root")
        ET.SubElement(root, "mxCell", attrib={"id":"0"})
        ET.SubElement(root, "mxCell", attrib={"id":"1","parent":"0"})

        def add_vertex(label, style, x, y, w=W, h=H):
            vid = uuid.uuid4().hex[:8]
            c = ET.SubElement(root, "mxCell", attrib={"id":vid,"value":label,"style":style,"vertex":"1","parent":"1"})
            ET.SubElement(c, "mxGeometry", attrib={"x":str(int(x)),"y":str(int(y)),"width":str(w),"height":str(h),"as":"geometry"})
            return vid
        def add_edge_points(src_id, tgt_id, points):
            eid = uuid.uuid4().hex[:8]
            c = ET.SubElement(root, "mxCell", attrib={"id":eid,"value":"","style":S_EDGE,"edge":"1","parent":"1","source":src_id,"target":tgt_id})
            g = ET.SubElement(c, "mxGeometry", attrib={"relative":"1","as":"geometry"})
            arr = ET.SubElement(g, "Array", attrib={"as":"points"})
            for (px, py) in points: ET.SubElement(arr, "mxPoint", attrib={"x":str(int(px)),"y":str(int(py))})
        def add_edge_with_elbow(src_id, tgt_id, src_center_x, tgt_center_x, elbow_y):
            add_edge_points(src_id, tgt_id, [(src_center_x, elbow_y), (tgt_center_x, elbow_y)])
        def cx(x_left): return int(x_left + W/2)

        id_map = {}
        # Ledgers
        for L in ledgers_all:
            id_map[("L",L)] = add_vertex(L, S_LEDGER, led_x[L], Y_LEDGER)
        # LEs
        for (L,E), x in le_x.items():
            id_map[("E",L,E)] = add_vertex(E, S_LE, x, Y_LE)
            add_edge_with_elbow(id_map[("E",L,E)], id_map[("L",L)], cx(x), cx(led_x[L]), ELBOW_LE_TO_LED)
        # BUs
        for (L,E,b), x in bu_x.items():
            id_map[("B",L,E,b)] = add_vertex(b, S_BU, x, Y_BU)
            add_edge_with_elbow(id_map[("B",L,E,b)], id_map[("E",L,E)], cx(x), cx(le_x[(L,E)]), ELBOW_BU_TO_LE)
        # COs
        for (L,E,c), x in co_x.items():
            id_map[("C",L,E,c)] = add_vertex(c, S_CO, x, Y_CO)
            add_edge_with_elbow(id_map[("C",L,E,c)], id_map[("E",L,E)], cx(x), cx(le_x[(L,E)]), ELBOW_CO_TO_LE)
        # Books (left of CO) — bold if primary
        for (L,E,c,bk), x in cb_x.items():
            style = S_CB_P if cb_primary.get((L,E,c,bk), False) else S_CB
            key = ("CB",L,E,c,bk)
            id_map[key] = add_vertex(bk, style, x, Y_CB)
            add_edge_with_elbow(id_map[key], id_map[("C",L,E,c)], cx(x), cx(co_x[(L,E,c)]), ELBOW_CB_TO_CO)
        # IOs under CO
        for (L,E,c,name), (x, is_mfg) in io_x.items():
            style = S_IO_PLT if str(is_mfg).lower() in ("yes","y","true","1") else S_IO
            label = f"🏭 {name}" if style == S_IO_PLT else name
            id_map[("IO",L,E,c,name)] = add_vertex(label, style, x, Y_IO)
            add_edge_with_elbow(id_map[("IO",L,E,c,name)], id_map[("C",L,E,c)], cx(x), cx(co_x[(L,E,c)]), ELBOW_IO_TO_CO)

        # Direct-to-LE IOs: trunk at (midpoint + bias), align at BU elbow then IO split
        dio_trunk_x = {}
        for (L,E), lst in dio_by_le.items():
            if not lst: continue
            xs = [dio_x[(L,E,r["Name"])][0] for r in lst]
            dio_trunk_x[(L,E)] = int(sum(xs)/len(xs)) + TRUNK_RIGHT_BIAS

        for (L,E,name), (x, is_mfg) in dio_x.items():
            style = S_IO_PLT if str(is_mfg).lower() in ("yes","y","true","1") else S_IO
            label = f"🏭 {name}" if style == S_IO_PLT else name
            key = ("DIO",L,E,name)
            id_map[key] = add_vertex(label, style, x, Y_IO)

            le_center_x = cx(le_x[(L,E)])
            trunk_x = dio_trunk_x.get((L,E), max(le_center_x, x) + TRUNK_RIGHT_BIAS)

            points = [
                (trunk_x, ELBOW_IO_TO_CO),  # same horizontal as CO-owned IO split
                (trunk_x, ELBOW_BU_TO_LE),  # same horizontal as BU elbow
                (le_center_x, ELBOW_BU_TO_LE)
            ]
            add_edge_points(id_map[key], id_map[("E",L,E)], points)

        # legend (smaller panel; won’t clip first ledger)
        def add_legend(x=12, y=12):
            _ = add_vertex("", "rounded=1;fillColor=#FFFFFF;strokeColor=#CBD5E1;", x, y, 164, 168)
            items = [
                ("Ledger", "#FFE6E6", None),
                ("Legal Entity", "#FFE2C2", None),
                ("Business Unit", "#FFF1B3", None),
                ("Cost Org", "#C2F0C2", None),
                ("Cost Book", "#A0D080", None),
                ("Primary Cost Book", "#A0D080", "bold"),
                ("Inventory Org", "#C2E0F9", None),
                ("Manufacturing Plant (IO)", "#C2E0F9", "io_bold"),
            ]
            yoff = 26
            for i, (lbl, col, flavor) in enumerate(items):
                if flavor == "bold":
                    style = "rounded=1;fillColor=#A0D080;strokeColor=#004d00;strokeWidth=2;"
                elif flavor == "io_bold":
                    style = "rounded=1;fillColor=#C2E0F9;strokeColor=#1F4D7A;strokeWidth=2;"
                else:
                    style = f"rounded=1;fillColor={col};strokeColor=#666666;"
                add_vertex("", style, x+10, y+yoff+i*18, 14, 9)
                add_vertex(lbl, "text;align=left;verticalAlign=middle;fontSize=11;", x+30, y+yoff-5+i*18, 120, 16)

        add_legend()
        return ET.tostring(mxfile, encoding="utf-8", method="xml").decode("utf-8")

    def _drawio_url_from_xml(xml: str) -> str:
        raw = zlib.compress(xml.encode("utf-8"), level=9)[2:-4]
        b64 = base64.b64encode(raw).decode("ascii")
        return f"https://app.diagrams.net/?title=EnterpriseStructure.drawio#R{b64}"

    _xml = _make_drawio_xml(df1, df2, df_cost)

    st.download_button(
        "⬇️ Download diagram (.drawio)",
        data=_xml.encode("utf-8"),
        file_name="EnterpriseStructure.drawio",
        mime="application/xml",
        use_container_width=True
    )
    st.markdown(f"[🔗 Open in draw.io (preview)]({_drawio_url_from_xml(_xml)})")

