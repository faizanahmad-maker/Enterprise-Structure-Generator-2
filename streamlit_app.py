import io, zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Enterprise Structure Generator", page_icon="📊", layout="wide")
st.title("Enterprise Structure Generator — Excel + draw.io (3 tabs)")

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
    for c in df.select_dtypes(include=["object"]).columns:
        mask = df[c].apply(lambda x: isinstance(x, str) and x.strip().lower() == "nan")
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
    bu_rows = []                     # BU rows

    # Cost Orgs
    costorg_rows = []                # [{Name, LegalEntityIdentifier, JoinKey}]
    costorg_name_to_joinkeys = {}    # Name -> {JoinKey}

    # Cost Books (with primary flag)  joinkey -> list[(book_code, is_primary)]
    books_by_joinkey = {}

    # Inventory Orgs
    invorg_rows = []                 # [{Code, Name, LEIdent, BUName, PCBU, Mfg}]
    invorg_rel = {}                  # InvOrgCode -> CostOrgJoinKey

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

        # Legal Entities (names)
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

        # Backup: identifier -> LE name (ObjectName)
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

        # Business Units
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

        # Cost Orgs
        df = read_csv_from_zip(z, "CST_COST_ORGANIZATION.csv")
        if df is not None:
            name_col   = pick_col(df, ["Name"])
            ident_col  = pick_col(df, ["LegalEntityIdentifier"])
            join_col   = pick_col(df, ["OrgInformation2"])
            if name_col and ident_col and join_col:
                for _, r in df[[name_col, ident_col, join_col]].dropna(how="all").iterrows():
                    name  = str(r[name_col]).strip()
                    ident = str(r[ident_col]).strip()
                    joink = str(r[join_col]).strip()
                    costorg_rows.append({"Name": name, "LegalEntityIdentifier": ident, "JoinKey": joink})
                    if name and joink:
                        costorg_name_to_joinkeys.setdefault(name, set()).add(joink)
            else:
                st.warning(f"`CST_COST_ORGANIZATION.csv` missing needed columns (Name, LegalEntityIdentifier, OrgInformation2).")

        # Cost Books (with Primary flag)
        df = read_csv_from_zip(z, "CST_COST_ORG_BOOK.csv")
        if df is not None:
            key_col   = pick_col(df, ["ORA_CST_ACCT_COST_ORG.CostOrgCode", "CostOrgCode"])
            book_col  = pick_col(df, ["CostBookCode"])
            prim_col  = pick_col(df, ["PrimaryBookFlag", "PrimaryFlag", "Primary"])
            if key_col and book_col:
                for _, r in df.dropna(how="all").iterrows():
                    joink = str(r.get(key_col, "")).strip()
                    book  = str(r.get(book_col, "")).strip()
                    rawp  = str(r.get(prim_col, "")).strip().upper() if prim_col else ""
                    is_primary = rawp in {"Y","YES","1","TRUE"}
                    if joink and book:
                        books_by_joinkey.setdefault(joink, []).append((book, is_primary))
            else:
                st.warning(f"`CST_COST_ORG_BOOK.csv` missing needed columns (CostOrgCode, CostBookCode).")

        # Inventory Orgs
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
                st.warning(f"`INV_ORGANIZATION_PARAMETER.csv` missing needed columns.")

        # Cost Org ↔ Inv Org
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
                st.warning(f"`ORA_CST_COST_ORG_INV.csv` missing needed columns (OrganizationCode, CostOrgCode).")

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

    known_pairs = {(led, ident_to_le_name.get(ident, "").strip())
                   for led, idents in ledger_to_idents.items()
                   for ident in idents if ident_to_le_name.get(ident, "").strip()}

    # ===================================================
    # Tab 1: Core Enterprise Structure (WITH hanging LEs)
    # ===================================================
    from collections import defaultdict

    le_names_master = set(legal_entity_names)
    le_names_from_ident = {v for v in ident_to_le_name.values() if v}
    all_le_display = {n for n in le_names_master | le_names_from_ident if n}

    norm2display = {}
    for n in all_le_display:
        k = _norm_key(n)
        if k and k not in norm2display:
            norm2display[k] = n

    # LE -> ledgers (via identifier mapping)
    le_to_ledgers = defaultdict(set)
    for led, idents in ledger_to_idents.items():
        for ident in idents:
            nm = ident_to_le_name.get(ident, "")
            if nm:
                le_to_ledgers[_norm_key(nm)].add(led)

    rows1 = []
    seen_pairs_norm = set()
    seen_ledgers_with_bu = set()
    seen_les_with_bu_norm = set()

    # BU-driven rows + smart backfill (unique-only)
    for r in bu_rows:
        bu  = str(r.get("Name", "")).strip()
        led = str(r.get("PrimaryLedgerName", "")).strip()
        le  = str(r.get("LegalEntityName", "")).strip()

        if not led and le:
            ks = le_to_ledgers.get(_norm_key(le), set())
            if len(ks) == 1:
                led = next(iter(ks))
        if not le and led:
            candidates = [disp for k, disp in norm2display.items() if led in le_to_ledgers.get(k, set())]
            if len(candidates) == 1:
                le = candidates[0]

        rows1.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": bu})
        if led: seen_ledgers_with_bu.add(led)
        if le:  seen_les_with_bu_norm.add(_norm_key(le))
        if led and le: seen_pairs_norm.add((led, _norm_key(le)))

    # Ledger–LE pairs without BU
    for led, idents in ledger_to_idents.items():
        for ident in idents:
            le_name = ident_to_le_name.get(ident, "").strip()
            if le_name and (led, _norm_key(le_name)) not in seen_pairs_norm:
                rows1.append({"Ledger Name": led, "Legal Entity": le_name, "Business Unit": ""})

    # Orphan ledgers
    mapped_ledgers = set(ledger_to_le_names.keys())
    for led in sorted(ledger_names - mapped_ledgers - seen_ledgers_with_bu):
        rows1.append({"Ledger Name": led, "Legal Entity": "", "Business Unit": ""})

    # Orphan LEs (hanging)
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
        df1.sort_values(["__LedgerEmpty", "Ledger Name", "Legal Entity", "Business Unit"],
                        ascending=[True, True, True, True])
           .drop(columns="__LedgerEmpty")
           .reset_index(drop=True)
    )
    df1.insert(0, "Assignment", range(1, len(df1) + 1))

    # ===================================================
    # Tab 2: Inventory Org Structure  (NO Cost Book columns)
    # ===================================================
    rows2 = []
    co_name_by_joinkey = {r["JoinKey"]: r["Name"] for r in costorg_rows if r.get("JoinKey")}

    for inv in invorg_rows:
        code = inv.get("Code", "")
        name = inv.get("Name", "")
        le_ident = inv.get("LEIdent", "")
        le_name  = ident_to_le_name.get(le_ident, "") if le_ident else ""
        leds     = ident_to_ledgers.get(le_ident, set()) if le_ident else set()

        co_key  = invorg_rel.get(code, "")
        co_name = co_name_by_joinkey.get(co_key, "") if co_key else ""

        base_row = {
            "Legal Entity": le_name,
            "Cost Organization": co_name,                    # blank when no CO
            "Inventory Org": name,
            "Manufacturing Plant": inv.get("Mfg", ""),
            "Profit Center BU": inv.get("PCBU", ""),
            "Management BU": inv.get("BUName", ""),
        }

        if leds:
            for led in sorted(leds):
                row = {"Ledger Name": led}
                row.update(base_row)
                rows2.append(row)
        else:
            row = {"Ledger Name": ""}
            row.update(base_row)
            rows2.append(row)

    df2 = pd.DataFrame(rows2).drop_duplicates().reset_index(drop=True)
    if not df2.empty:
        df2["__LedgerEmpty"] = (df2["Ledger Name"].fillna("") == "").astype(int)
        df2["__COEmpty"]     = (df2["Cost Organization"].fillna("") == "").astype(int)
        df2 = (
            df2.sort_values(
                ["__LedgerEmpty", "Ledger Name", "Legal Entity", "__COEmpty",
                 "Cost Organization", "Inventory Org"],
                ascending=[True, True, True, True, True, True]
            )
            .drop(columns=["__LedgerEmpty", "__COEmpty"])
            .reset_index(drop=True)
        )
    df2.insert(0, "Assignment", range(1, len(df2) + 1))

    # ===================================================
    # Tab 3: Costing Structure (Ledger • LE • Cost Org • Cost Book • Primary?)
    # ===================================================
    rows3 = []
    for co in costorg_rows:
        co_name = co.get("Name", "")
        le_ident = co.get("LegalEntityIdentifier", "")
        joink   = co.get("JoinKey", "")
        le_name = ident_to_le_name.get(le_ident, "") if le_ident else ""

        books = books_by_joinkey.get(joink, [])  # list of (book, is_primary)
        leds  = ident_to_ledgers.get(le_ident, set()) if le_ident else set()

        # If no books present, still list the Cost Org with blank Cost Book
        if not books:
            base = {
                "Legal Entity": le_name,
                "Cost Organization": co_name,
                "Cost Book": "",
                "Primary Cost Book": ""
            }
            if leds:
                for led in sorted(leds):
                    r = {"Ledger Name": led}
                    r.update(base)
                    rows3.append(r)
            else:
                r = {"Ledger Name": ""}
                r.update(base)
                rows3.append(r)
            continue

        # With books → one row per book (and per ledger if multiple)
        for (bk, is_primary) in sorted(books, key=lambda x: (x[0], not x[1])):
            base = {
                "Legal Entity": le_name,
                "Cost Organization": co_name,
                "Cost Book": bk,
                "Primary Cost Book": "Yes" if is_primary else "No"
            }
            if leds:
                for led in sorted(leds):
                    r = {"Ledger Name": led}
                    r.update(base)
                    rows3.append(r)
            else:
                r = {"Ledger Name": ""}
                r.update(base)
                rows3.append(r)

    df3 = pd.DataFrame(rows3).drop_duplicates().reset_index(drop=True)
    if not df3.empty:
        df3["__LedgerEmpty"] = (df3["Ledger Name"].fillna("") == "").astype(int)
        df3 = (
            df3.sort_values(
                ["__LedgerEmpty", "Ledger Name", "Legal Entity",
                 "Cost Organization", "Cost Book", "Primary Cost Book"],
                ascending=[True, True, True, True, True, False]
            )
            .drop(columns=["__LedgerEmpty"])
            .reset_index(drop=True)
        )
    df3.insert(0, "Assignment", range(1, len(df3) + 1))

    # ----------- clean NaNs -----------
    df1, df2, df3 = _blankify(df1), _blankify(df2), _blankify(df3)

    # ------------ Excel Output (tab order kept as you want) ------------
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        df1.to_excel(writer, index=False, sheet_name="Core Enterprise Structure")
        df2.to_excel(writer, index=False, sheet_name="Inventory Org Structure")   # Tab 2 (no Cost Book)
        df3.to_excel(writer, index=False, sheet_name="Costing Structure")         # Tab 3

    st.success(f"Built {len(df1)} Core rows, {len(df2)} Inventory rows, {len(df3)} Costing rows.")
    st.dataframe(df1.head(20), use_container_width=True, height=260)
    st.dataframe(df2.head(20), use_container_width=True, height=260)
    st.dataframe(df3.head(20), use_container_width=True, height=260)

    st.download_button(
        "⬇️ Download Excel (EnterpriseStructure.xlsx)",
        data=excel_buf.getvalue(),
        file_name="EnterpriseStructure.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ===================== DRAW.IO DIAGRAM (CO straight down + GLOBAL MIN SPACING + guided trunk) =====================
if (
    "df1" in locals() and isinstance(df1, pd.DataFrame) and not df1.empty and
    "df2" in locals() and isinstance(df2, pd.DataFrame) and
    "df3" in locals() and isinstance(df3, pd.DataFrame)
):
    import xml.etree.ElementTree as ET
    import zlib, base64, uuid

    def _make_drawio_xml(df_bu: pd.DataFrame, df_io: pd.DataFrame, df_costing: pd.DataFrame) -> str:
        # ---------- Geometry ----------
        W, H = 180, 48
        Y_LEDGER, Y_LE, Y_BU, Y_CO, Y_CB, Y_IO = 150, 320, 480, 640, 800, 960

        def elbow(y_child, y_parent, bias=0.75):
            return int(y_parent + (y_child - y_parent) * bias)

        ELBOW_LE_TO_LED = elbow(Y_LE, Y_LEDGER)
        ELBOW_BU_TO_LE  = elbow(Y_BU, Y_LE)
        ELBOW_CO_TO_LE  = elbow(Y_CO, Y_LE)
        ELBOW_IO_TO_CO  = elbow(Y_IO, Y_CO)
        ELBOW_CB_TO_CO  = elbow(Y_CB, Y_CO)

        # spacing
        MIN_GAP = 70
        def spread(base): return max(base, W + MIN_GAP)
        BU_SPREAD_BASE, CO_SPREAD_BASE = 210, 230
        IO_UNDER_CO_BASE = 220
        LEDGER_BLOCK_GAP, CLUSTER_GAP, LEFT_PAD = 120, 420, 260
        MIN_UMBRELLA_GAP = 140

        # GLOBAL spacing per LE, per layer:
        MIN_GLOBAL_SPACING = 200  # <<< raise if you want even more air

        # lane offsets (CO stays vertical)
        BU_LANE_OFFSET  = 180
        CO_LANE_OFFSET  = 0
        DIO_LANE_OFFSET = 420

        # books (vertical column to the left of CO)
        BOOK_X_OFFSET     = 220
        BOOK_VERTICAL_GAP = 64

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

        def cx(x_left): return int(x_left + W/2)

        def centers(center_x, n, base):
            s = spread(base)
            if n <= 0: return []
            if n == 1: return [int(center_x)]
            start = center_x - (s*(n-1))/2.0
            return [int(start + i*s) for i in range(n)]

        def enforce_spacing_sorted(xs, min_spacing):
            if not xs: return xs
            xs_sorted = sorted(xs)
            for i in range(1, len(xs_sorted)):
                if xs_sorted[i] - xs_sorted[i-1] < min_spacing:
                    xs_sorted[i] = xs_sorted[i-1] + min_spacing
            return xs_sorted

        # ---------- Normalize inputs ----------
        df_bu = df_bu[["Ledger Name","Legal Entity","Business Unit"]].copy().fillna("").astype(str)

        # Tab 2 (no Cost Book)
        LCOL = pick(df_io, ["Ledger Name","Ledger"])
        ECOL = pick(df_io, ["Legal Entity","LegalEntity"])
        COCOL= pick(df_io, ["Cost Organization","Cost Org","CostOrganization"])
        IOCOL= pick(df_io, ["Inventory Org","Inventory Organization","InventoryOrg"])
        MFGCOL=pick(df_io, ["Manufacturing Plant","Mfg","ManufacturingPlant","IsManufacturingPlant"])
        df_io = df_io[[x for x in [LCOL,ECOL,COCOL,IOCOL,MFGCOL] if x is not None]].copy().fillna("").astype(str)
        df_io.rename(columns={LCOL:"Ledger Name", ECOL:"Legal Entity", COCOL:"Cost Organization",
                              IOCOL:"Inventory Org", MFGCOL:"Manufacturing Plant"}, inplace=True)

        # Tab 3 (Costing)
        cLCOL = pick(df_costing, ["Ledger Name","Ledger"])
        cECOL = pick(df_costing, ["Legal Entity","LegalEntity"])
        cCO   = pick(df_costing, ["Cost Organization","Cost Org","CostOrganization"])
        cBKC  = pick(df_costing, ["Cost Book","CostBook"])
        cBKPC = pick(df_costing, ["Primary Cost Book","PrimaryBook","Primary Flag","PrimaryBookFlag"])
        df_costing = df_costing[[x for x in [cLCOL,cECOL,cCO,cBKC,cBKPC] if x is not None]].copy().fillna("").astype(str)
        df_costing.rename(columns={cLCOL:"Ledger Name", cECOL:"Legal Entity", cCO:"Cost Organization",
                                   cBKC:"Cost Book"}, inplace=True)
        if cBKPC: df_costing.rename(columns={cBKPC:"Primary Cost Book"}, inplace=True)

        ledgers_all = sorted({*df_bu["Ledger Name"].unique(), *df_io["Ledger Name"].unique()} - {""})

        # ---------- Build maps ----------
        from collections import defaultdict
        le_map = defaultdict(set)
        bu_map = defaultdict(list)
        co_map = defaultdict(list)
        io_by_co = defaultdict(list)
        dio_by_le = defaultdict(list)
        cb_by_co = defaultdict(list)
        cb_primary = {}

        tmp = pd.concat([df_bu[["Ledger Name","Legal Entity"]], df_io[["Ledger Name","Legal Entity"]]]).drop_duplicates()
        for _, r in tmp.iterrows():
            L,E = r["Ledger Name"], r["Legal Entity"]
            if L and E: le_map[L].add(E)

        for _, r in df_bu.iterrows():
            L,E,B = r["Ledger Name"], r["Legal Entity"], r["Business Unit"]
            if L and E and B: bu_map[(L,E)].append(B)

        for _, r in df_io.iterrows():
            L,E,C = r["Ledger Name"], r["Legal Entity"], r["Cost Organization"]
            if L and E and C and C not in co_map[(L,E)]: co_map[(L,E)].append(C)

        for _, r in df_io.iterrows():
            L,E,C = r["Ledger Name"], r["Legal Entity"], r["Cost Organization"]
            IO,MFG = r["Inventory Org"], r["Manufacturing Plant"]
            if not (L and E and IO): continue
            rec = {"Name": IO, "Mfg": (MFG or "")}
            if C:
                if all(x["Name"] != IO for x in io_by_co[(L,E,C)]): io_by_co[(L,E,C)].append(rec)
            else:
                if all(x["Name"] != IO for x in dio_by_le[(L,E)]): dio_by_le[(L,E)].append(rec)

        for _, r in df_costing.iterrows():
            L,E,C = r.get("Ledger Name",""), r.get("Legal Entity",""), r.get("Cost Organization","")
            bk    = r.get("Cost Book","").strip()
            if not (L and E and C and bk): continue
            if bk not in cb_by_co[(L,E,C)]: cb_by_co[(L,E,C)].append(bk)
            if "Primary Cost Book" in df_costing.columns:
                raw = str(r.get("Primary Cost Book","")).strip().lower()
                cb_primary[(L,E,C,bk)] = raw in ("yes","y","true","1","primary")

        # ---------- Placement ----------
        next_x = LEFT_PAD
        led_x, le_x, bu_x, co_x = {}, {}, {}, {}
        io_x, dio_x, cb_xy = {}, {}, {}

        def co_cluster_halfwidth(L,E,C):
            ios = io_by_co[(L,E,C)]
            io_half = (max(1, len(ios)) * IO_UNDER_CO_BASE)/2 + W/2
            left_half = W/2 + (BOOK_X_OFFSET if cb_by_co[(L,E,C)] else 0)
            return max(left_half, io_half)

        prev_umbrella_max_x = None
        for L in ledgers_all:
            les = sorted(le_map[L])
            le_centers = []
            for E in les:
                le_pos = next_x
                le_x[(L,E)] = le_pos
                le_centers.append(le_pos)

                bu_list = sorted(set(bu_map[(L,E)]))
                cos     = list(co_map[(L,E)])

                has_bu  = bool(bu_list)
                has_co  = bool(cos)
                has_dio = bool(dio_by_le[(L,E)])

                bu_center  = le_pos if (has_bu and not (has_co or has_dio)) else (le_pos - BU_LANE_OFFSET if has_bu else le_pos)
                co_center  = le_pos  # CO straight down
                dio_center = le_pos + DIO_LANE_OFFSET if has_dio else None

                # BUs
                for x,b in zip(centers(bu_center, len(bu_list), BU_SPREAD_BASE), bu_list):
                    bu_x[(L,E,b)] = x

                # COs
                if has_co:
                    placed = []
                    for idx, C in enumerate(sorted(cos)):
                        half = co_cluster_halfwidth(L,E,C)
                        if idx == 0:
                            xC = co_center
                        else:
                            prev = placed[-1]
                            need = prev["half"] + half + MIN_GAP
                            xC = int(prev["x"] + need)
                        placed.append({"C":C, "x":xC, "half":half})
                        co_x[(L,E,C)] = xC

                        # IOs under this CO
                        ios = sorted(io_by_co[(L,E,C)], key=lambda d: d["Name"])
                        xs = centers(xC, len(ios), IO_UNDER_CO_BASE)
                        xs = enforce_spacing_sorted(xs, MIN_GAP)  # local tidy
                        for xio, rec in zip(xs, ios):
                            io_x[(L,E,C,rec["Name"])] = (xio, rec["Mfg"])

                        # Books (vertical to the left)
                        for i, bk in enumerate(sorted(cb_by_co[(L,E,C)])):
                            cb_xy[(L,E,C,bk)] = (xC - BOOK_X_OFFSET, Y_CB + i*BOOK_VERTICAL_GAP)

                # Direct IOs
                if has_dio:
                    dlist = sorted(dio_by_le[(L,E)], key=lambda d: d["Name"])
                    xs = centers(dio_center, len(dlist), IO_UNDER_CO_BASE)
                    xs = enforce_spacing_sorted(xs, MIN_GAP)
                    for xio, rec in zip(xs, dlist):
                        dio_x[(L,E,rec["Name"])] = (xio, rec["Mfg"])

                # umbrella guard
                xs_span = [le_pos]
                xs_span += [bu_x[(L,E,b)] for b in bu_list]
                for C in cos:
                    xs_span.append(co_x[(L,E,C)])
                    xs_span += [io_x[(L,E,C,r["Name"])][0] for r in io_by_co[(L,E,C)]]
                    xs_span += [cb_xy[(L,E,C,bk)][0] for bk in cb_by_co[(L,E,C)] if (L,E,C,bk) in cb_xy]
                xs_span += [v[0] for k,v in dio_x.items() if k[:2]==(L,E)]

                min_x = (min(xs_span) - W/2) if xs_span else le_pos - W/2
                max_x_ = (max(xs_span) + W/2) if xs_span else le_pos + W/2

                if prev_umbrella_max_x is not None and min_x < prev_umbrella_max_x + MIN_UMBRELLA_GAP:
                    shift = (prev_umbrella_max_x + MIN_UMBRELLA_GAP) - min_x
                    le_x[(L,E)] += shift
                    for k in list(bu_x.keys()):
                        if k[0]==L and k[1]==E: bu_x[k] += shift
                    for k in list(co_x.keys()):
                        if k[0]==L and k[1]==E: co_x[k] += shift
                    for k in list(io_x.keys()):
                        if k[0]==L and k[1]==E: io_x[k] = (io_x[k][0] + shift, io_x[k][1])
                    for k in list(cb_xy.keys()):
                        if k[0]==L and k[1]==E: cb_xy[k] = (cb_xy[k][0] + shift, cb_xy[k][1])
                    for k in list(dio_x.keys()):
                        if k[0]==L and k[1]==E: dio_x[k] = (dio_x[k][0] + shift, dio_x[k][1])
                    max_x_ += shift

                prev_umbrella_max_x = max_x_
                next_x = max_x_ + LEDGER_BLOCK_GAP

            # provisional ledger center for this block
            if le_centers:
                led_x[L] = int(sum(le_x[(L,E)] for E in les) / len(les))
            else:
                led_x[L] = next_x
            next_x += CLUSTER_GAP

        # ---------- GLOBAL MIN SPACING per LE & per LAYER ----------
        # Utility: take (key -> x) dict, return right-shifted map with minimum spacing
        def layer_global_spacing(update_fn, xs_with_keys):
            if not xs_with_keys: return
            # sort by x, enforce global spacing
            xs_sorted = enforce_spacing_sorted([x for _, x in xs_with_keys], MIN_GLOBAL_SPACING)
            for (k, _), new_x in zip(sorted(xs_with_keys, key=lambda t: t[1]), xs_sorted):
                update_fn(k, new_x)

        for L in ledgers_all:
            for E in sorted(le_map[L]):
                # BU layer
                bu_keys = [(k, bu_x[k]) for k in bu_x if k[0]==L and k[1]==E]
                layer_global_spacing(lambda k, nx: bu_x.__setitem__(k, nx), bu_keys)

                # CO layer
                co_keys = [(k, co_x[k]) for k in co_x if k[0]==L and k[1]==E]
                layer_global_spacing(lambda k, nx: co_x.__setitem__(k, nx), co_keys)

                # IO layer (CO-owned IOs + direct IOs together)
                io_keys = [((k), io_x[k][0]) for k in io_x if k[0]==L and k[1]==E]
                dio_keys= [((k), dio_x[k][0]) for k in dio_x if k[0]==L and k[1]==E]
                all_io  = io_keys + dio_keys
                def _upd_io(k, nx):
                    if len(k)==4 and k in io_x:
                        io_x[k] = (nx, io_x[k][1])
                    elif len(k)==3 and k in dio_x:
                        dio_x[k] = (nx, dio_x[k][1])
                layer_global_spacing(_upd_io, all_io)

        # final re-center ledgers
        for L in ledgers_all:
            les = sorted(le_map[L])
            if les:
                led_x[L] = int(sum(le_x[(L,E)] for E in les) / len(les))

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
            for (px, py) in points:
                ET.SubElement(arr, "mxPoint", attrib={"x":str(int(px)),"y":str(int(py))})

        def add_edge_with_elbow(src_id, tgt_id, src_center_x, tgt_center_x, elbow_y):
            add_edge_points(src_id, tgt_id, [(src_center_x, elbow_y), (tgt_center_x, elbow_y)])

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
        # Books (vertical, left of CO)
        for (L,E,c,bk), (xbk, ybk) in cb_xy.items():
            style = S_CB_P if cb_primary.get((L,E,c,bk), False) else S_CB
            id_map[("CB",L,E,c,bk)] = add_vertex(bk, style, xbk, ybk)
            add_edge_with_elbow(id_map[("CB",L,E,c,bk)], id_map[("C",L,E,c)], cx(xbk), cx(co_x[(L,E,c)]), ELBOW_CB_TO_CO)
        # IOs under CO
        for (L,E,c,name), (x, is_mfg) in io_x.items():
            style = S_IO_PLT if str(is_mfg).lower() in ("yes","y","true","1") else S_IO
            label = f"🏭 {name}" if style == S_IO_PLT else name
            v = add_vertex(label, style, x, Y_IO)
            add_edge_with_elbow(v, id_map[("C",L,E,c)], cx(x), cx(co_x[(L,E,c)]), ELBOW_IO_TO_CO)

        # Direct IOs with shared guided trunk
        TRUNK_RIGHT_BIAS = 90
        dio_trunk_x = {}
        for L in ledgers_all:
            for E in sorted(le_map[L]):
                xs = [pos[0] for (k,pos) in dio_x.items() if k[0]==L and k[1]==E]
                dio_trunk_x[(L,E)] = (int(sum(xs)/len(xs)) if xs else cx(le_x[(L,E)])) + TRUNK_RIGHT_BIAS

        for (L,E,name), (x, is_mfg) in dio_x.items():
            style = S_IO_PLT if str(is_mfg).lower() in ("yes","y","true","1") else S_IO
            label = f"🏭 {name}" if style == S_IO_PLT else name
            v = add_vertex(label, style, x, Y_IO)
            le_center_x = cx(le_x[(L,E)])
            trunk_x = dio_trunk_x[(L,E)]
            add_edge_points(
                v, id_map[("E",L,E)],
                [(trunk_x, ELBOW_IO_TO_CO),
                 (trunk_x, ELBOW_BU_TO_LE),
                 (le_center_x, ELBOW_BU_TO_LE)]
            )

        # Legend
        def add_legend(x=12, y=12):
            _ = add_vertex("", "rounded=1;fillColor=#FFFFFF;strokeColor=#CBD5E1;", x, y, 180, 176)
            items = [
                ("Ledger", S_LEDGER),
                ("Legal Entity", S_LE),
                ("Business Unit", S_BU),
                ("Cost Org", S_CO),
                ("Cost Book", S_CB),
                ("Primary Cost Book", S_CB_P),
                ("Inventory Org", S_IO),
                ("Manufacturing Plant (IO)", S_IO_PLT),
            ]
            yoff = 26
            for i,(lbl, style) in enumerate(items):
                add_vertex("", style, x+10, y+yoff+i*18, 14, 9)
                add_vertex(lbl, "text;align=left;verticalAlign=middle;fontSize=11;", x+30, y+yoff-5+i*18, 140, 16)

        add_legend()
        return ET.tostring(mxfile, encoding="utf-8", method="xml").decode("utf-8")

    def _drawio_url_from_xml(xml: str) -> str:
        raw = zlib.compress(xml.encode("utf-8"), level=9)[2:-4]
        b64 = base64.b64encode(raw).decode("ascii")
        return f"https://app.diagrams.net/?title=EnterpriseStructure.drawio#R{b64}"

    _xml = _make_drawio_xml(df1, df2, df3)
    st.download_button(
        "⬇️ Download diagram (.drawio)",
        data=_xml.encode("utf-8"),
        file_name="EnterpriseStructure.drawio",
        mime="application/xml",
        use_container_width=True
    )
    st.markdown(f"[🔗 Open in draw.io (preview)]({_drawio_url_from_xml(_xml)})")

