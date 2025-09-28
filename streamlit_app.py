import io, zipfile
import pandas as pd
import streamlit as st

# ---------- App Header ----------
st.set_page_config(page_title="Enterprise Structure Generator — Sheets Only", page_icon="📊", layout="wide")
st.title("Enterprise Structure Generator — Excel Only (BUs, Cost Orgs, Books, Inventory Orgs)")

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
- `Manage Cost Org — Inventory Orgs` → **ORA_CST_COST_ORG_INV.csv** (or similarly named export mapping IO→CostOrg code)
""")

uploads = st.file_uploader("Drop your ZIPs here", type="zip", accept_multiple_files=True)

# ---------- Helpers ----------
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

def blankify(df: pd.DataFrame) -> pd.DataFrame:
    """Make DataFrame display/export with blank cells (no NaNs or 'nan' literals)."""
    if df is None:
        return df
    df = df.fillna("")
    df = df.replace({r"^\s*(?i)nan\s*$": ""}, regex=True)
    for c in df.columns:
        if pd.api.types.is_string_dtype(df[c]):
            df[c] = df[c].map(lambda x: x.strip() if isinstance(x, str) else x)
    return df

# ---------- Main ----------
if not uploads:
    st.info("Upload your ZIPs to generate the Excel.")
else:
    # ------------ Collectors ------------
    ledger_names = set()
    legal_entity_names = set()

    # Ledger ↔ Legal Entity Identifier mapping
    ledger_to_idents = {}            # ledger -> {LE identifier}
    ident_to_le_name = {}            # LE identifier -> LE name

    # Business Units
    bu_rows = []                     # [{Name, PrimaryLedgerName, LegalEntityName}]

    # Cost Orgs (from CST_COST_ORGANIZATION)
    # Keep Name, LegalEntityIdentifier, and OrgInformation2 (join key to books)
    costorg_rows = []                # [{Name, LegalEntityIdentifier, JoinKey}]
    costorg_name_to_joinkeys = {}    # CO Name -> {JoinKey}

    # Cost Books (from CST_COST_ORG_BOOK): JoinKey(CostOrgCode) -> {CostBookCode}
    books_by_joinkey = {}

    # Inventory Orgs (from INV_ORGANIZATION_PARAMETER)
    # Store: IO Name, OrganizationCode, ProfitCenterBuName, BusinessUnitName, MfgPlantFlag
    invorg_rows = []  # [{InventoryOrg, OrganizationCode, ProfitCenterBU, ManagementBU, MfgFlag}]

    # CostOrg ↔ InventoryOrg relationship (from ORA_CST_COST_ORG_INV.csv)
    # OrganizationCode (IO) -> CostOrgCode (which equals JoinKey used by books)
    inv_code_to_costorg_code = {}

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
                st.warning(f"`GL_PRIMARY_LEDGER.csv` missing `ORA_GL_PRIMARY_LEDGER_CONFIG.Name`. Found: {list(df.columns)}")

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

        # Ledger ↔ LE identifier (Assign Legal Entities)
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

        # Backup map for identifier -> LE name (ObjectName in journal config)
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
                        "Name": str(r[bu_col]).strip(),
                        "PrimaryLedgerName": str(r[led_col]).strip(),
                        "LegalEntityName": str(r[le_col]).strip()
                    })
            else:
                st.warning(f"`FUN_BUSINESS_UNIT.csv` missing needed columns. Found: {list(df.columns)}")

        # Cost Orgs (MASTER) — Name, LegalEntityIdentifier, OrgInformation2 as JOIN KEY
        df = read_csv_from_zip(z, "CST_COST_ORGANIZATION.csv")
        if df is not None:
            name_col   = pick_col(df, ["Name"])
            ident_col  = pick_col(df, ["LegalEntityIdentifier"])
            join_col   = pick_col(df, ["OrgInformation2"])  # join to CST_COST_ORG_BOOK.CostOrgCode
            if name_col and ident_col and join_col:
                for _, r in df[[name_col, ident_col, join_col]].dropna(how="all").iterrows():
                    name  = str(r[name_col]).strip()
                    ident = str(r[ident_col]).strip()
                    joink = str(r[join_col]).strip()
                    costorg_rows.append({"Name": name, "LegalEntityIdentifier": ident, "JoinKey": joink})
                    if name and joink:
                        costorg_name_to_joinkeys.setdefault(name, set()).add(joink)
            else:
                st.warning(f"`CST_COST_ORGANIZATION.csv` missing needed columns (need Name, LegalEntityIdentifier, OrgInformation2). Found: {list(df.columns)}")

        # Cost Books — map JoinKey(CostOrgCode) -> {CostBookCode}
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
                st.warning(f"`CST_COST_ORG_BOOK.csv` missing needed columns. Found: {list(df.columns)}")

        # Inventory Orgs (Manage Inventory Organizations)
        df = read_csv_from_zip(z, "INV_ORGANIZATION_PARAMETER.csv")
        if df is not None:
            name_col  = pick_col(df, ["Name"])
            code_col  = pick_col(df, ["OrganizationCode"])
            pcbu_col  = pick_col(df, ["ProfitCenterBuName", "ProfitCenterBUName", "ProfitCenter Business Unit", "ProfitCenterBuName_Display"])
            mbu_col   = pick_col(df, ["BusinessUnitName", "ManagementBuName", "Management Business Unit", "Business Unit Name"])
            mfg_col   = pick_col(df, ["MfgPlantFlag", "ManufacturingPlantFlag", "Manufacturing Plant Flag"])
            # legal ident (used only for "hanging" grouping logic safety)
            # leident_col = pick_col(df, ["LegalEntityIdentifier"])

            if name_col and code_col:
                for _, r in df.dropna(how="all").iterrows():
                    invorg_rows.append({
                        "InventoryOrg": str(r.get(name_col, "")).strip(),
                        "OrganizationCode": str(r.get(code_col, "")).strip(),
                        "ProfitCenterBU": str(r.get(pcbu_col, "")).strip() if pcbu_col else "",
                        "ManagementBU": str(r.get(mbu_col, "")).strip() if mbu_col else "",
                        "MfgFlag": str(r.get(mfg_col, "")).strip() if mfg_col else ""
                    })
            else:
                st.warning(f"`INV_ORGANIZATION_PARAMETER.csv` missing Name and/or OrganizationCode. Found: {list(df.columns)}")

        # IO ↔ Cost Org mapping
        # Expected column names vary; try common patterns.
        df = read_csv_from_zip(z, "ORA_CST_COST_ORG_INV.csv")
        if df is None:
            # Some environments export as CST_COST_ORG_INV.csv
            df = read_csv_from_zip(z, "CST_COST_ORG_INV.csv")
        if df is not None:
            io_code_col  = pick_col(df, ["OrganizationCode", "InvOrgCode", "InventoryOrganizationCode"])
            co_code_col  = pick_col(df, ["CostOrgCode", "ORA_CST_ACCT_COST_ORG.CostOrgCode"])
            if io_code_col and co_code_col:
                for _, r in df[[io_code_col, co_code_col]].dropna(how="all").iterrows():
                    iocode = str(r[io_code_col]).strip()
                    cocode = str(r[co_code_col]).strip()
                    if iocode and cocode:
                        inv_code_to_costorg_code[iocode] = cocode
            else:
                st.warning(f"`*_CST_COST_ORG_INV.csv` missing OrganizationCode and/or CostOrgCode. Found: {list(df.columns)}")

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

    # Name-based cautious backfill (LE -> Ledgers)
    le_to_ledgers_namekey = {}
    for led, le_set in ledger_to_le_names.items():
        for le in le_set:
            le_to_ledgers_namekey.setdefault(le, set()).add(led)

    # ===================================================
    # Tab 1: Ledger – Legal Entity – Business Unit
    # ===================================================
    rows1, seen_triples, seen_ledgers_with_bu = [], set(), set()

    # 1) BU-driven rows (+ cautious backfill for missing Ledger/LE)
    for r in bu_rows:
        bu  = r["Name"]
        led = r["PrimaryLedgerName"] if r["PrimaryLedgerName"] in ledger_names else ""
        le  = r["LegalEntityName"]  if r["LegalEntityName"]  in legal_entity_names else ""

        if not led and le and le in le_to_ledgers_namekey and len(le_to_ledgers_namekey[le]) == 1:
            led = next(iter(le_to_ledgers_namekey[le]))
        if not le and led and led in ledger_to_le_names and len(ledger_to_le_names[led]) == 1:
            le = next(iter(ledger_to_le_names[led]))

        rows1.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": bu})
        seen_triples.add((led, le, bu))
        if led:
            seen_ledgers_with_bu.add(led)

    # 2) Ledger–LE pairs with no BU
    seen_pairs = {(a, b) for (a, b, _) in seen_triples}
    for led, le in sorted(known_pairs):
        if (led, le) not in seen_pairs:
            rows1.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": ""})

    # 3) Orphan ledgers in master list
    mapped_ledgers = set(ledger_to_le_names.keys())
    for led in sorted(ledger_names - mapped_ledgers - seen_ledgers_with_bu):
        rows1.append({"Ledger Name": led, "Legal Entity": "", "Business Unit": ""})

    # 4) True unassigned LEs
    le_names_in_pairs = {le for (_, le) in known_pairs}
    le_names_in_bu    = {r["LegalEntityName"] for r in bu_rows if r.get("LegalEntityName")}
    for le in sorted(legal_entity_names - le_names_in_pairs - le_names_in_bu):
        rows1.append({"Ledger Name": "", "Legal Entity": le, "Business Unit": ""})

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

    # ===================================================
    # Tab 2: Ledger – Legal Entity – Cost Organization – Cost Book – Inventory Org (+ BU columns & Mfg flag)
    # ===================================================
    # Helper: books list by Cost Org name (via set of join-keys)
    def books_for_costorg_name(co_name: str):
        co_name = (co_name or "").strip()
        if not co_name:
            return []
        keys = sorted(costorg_name_to_joinkeys.get(co_name, []))
        acc = set()
        for k in keys:
            acc |= books_by_joinkey.get(k, set())
        return sorted(acc)

    # Build quick lookups
    # (1) From costorg_rows: ident -> set(ledgers), ident -> LE name
    ident_to_ledgers = {k: v for k, v in ident_to_ledgers.items()}  # already built
    # (2) From invorg_rows + inv_code_to_costorg_code: org code -> cost books (via same join key as books)
    #     Also retain IO attributes (ProfitCenter BU, Management BU, MfgFlag).
    inv_records = {}  # InventoryOrgName -> dicts (but we’ll iterate by row instead)

    rows2 = []
    seen_pairs2 = set()

    # Base ledger–LE pairs from Tab 1 (to ensure alignment)
    base_pairs = {
        (r["Ledger Name"], r["Legal Entity"])
        for _, r in df1.iterrows()
        if str(r["Ledger Name"]).strip() and str(r["Legal Entity"]).strip()
    }

    # 1) Emit rows from Cost Orgs (identifier-driven → ledger/LE). Allow blank CO if IO exists later.
    for r in costorg_rows:
        co    = r.get("Name", "").strip()
        ident = r.get("LegalEntityIdentifier", "").strip()
        le    = ident_to_le_name.get(ident, "").strip()
        leds  = ident_to_ledgers.get(ident, set())
        cb    = "; ".join(books_for_costorg_name(co)) if co else ""

        if leds:
            for led in sorted(leds):
                rows2.append({
                    "Ledger Name": led, "Legal Entity": le,
                    "Cost Organization": co, "Cost Book": cb,
                    "Inventory Org": "", "Profit Center BU": "", "Management BU": "", "Manufacturing Plant": ""
                })
                seen_pairs2.add((led, le))
        else:
            rows2.append({
                "Ledger Name": "", "Legal Entity": le,
                "Cost Organization": co, "Cost Book": cb,
                "Inventory Org": "", "Profit Center BU": "", "Management BU": "", "Manufacturing Plant": ""
            })

    # 2) Add rows from Inventory Orgs (map IO → Cost Org Code → books; ledger/LE derived via base_pairs if present)
    # Build a helper to get ledger & LE for a given Cost Org JoinKey using costorg_rows + ident map
    joink_to_le_leds = {}  # JoinKey -> (LE name, set(ledgers))
    for r in costorg_rows:
        joink = r.get("JoinKey", "").strip()
        ident = r.get("LegalEntityIdentifier", "").strip()
        if not joink or not ident:
            continue
        le_name = ident_to_le_name.get(ident, "").strip()
        leds    = ident_to_ledgers.get(ident, set())
        if le_name or leds:
            joink_to_le_leds[joink] = (le_name, leds)

    for r in invorg_rows:
        io_name  = r["InventoryOrg"]
        iocode   = r["OrganizationCode"]
        pcbu     = r["ProfitCenterBU"]
        mbu      = r["ManagementBU"]
        mfgflag  = r["MfgFlag"]

        if not io_name:
            continue

        # Find its Cost Org code (JoinKey) and books
        co_code = inv_code_to_costorg_code.get(iocode, "")
        books   = sorted(books_by_joinkey.get(co_code, set()))
        cb      = "; ".join(books) if books else ""

        # Derive LE + Ledgers from the join key if possible
        le_name, leds = joink_to_le_leds.get(co_code, ("", set()))

        if leds:
            for led in sorted(leds):
                rows2.append({
                    "Ledger Name": led, "Legal Entity": le_name,
                    "Cost Organization": "", "Cost Book": cb,
                    "Inventory Org": io_name, "Profit Center BU": pcbu, "Management BU": mbu,
                    "Manufacturing Plant": ("Yes" if str(mfgflag).strip().lower() in ("y","yes","true","1") else "")
                })
                seen_pairs2.add((led, le_name))
        else:
            # Hanging IO (no clear ledger) — put LE if we know it; else blanks
            rows2.append({
                "Ledger Name": "", "Legal Entity": le_name,
                "Cost Organization": "", "Cost Book": cb,
                "Inventory Org": io_name, "Profit Center BU": pcbu, "Management BU": mbu,
                "Manufacturing Plant": ("Yes" if str(mfgflag).strip().lower() in ("y","yes","true","1") else "")
            })

    # 3) Ensure all Tab 1 ledger–LE pairs appear (only if there’s at least a CO or IO there? User asked:
    #    "tab 2 should only have a row if there is a cost org and or an inv org" → So DO NOT backfill pure empty pairs.)
    #    Therefore, we skip emitting empty pairs here.

    # 4) Orphan ledgers in masters but unseen (only if they have CO/IO? If none, skip.)
    #    We skip adding pure orphans to honor "only rows with CO and/or IO".

    # Collapse + sort
    df2 = pd.DataFrame(rows2)
    if df2.empty:
        df2 = pd.DataFrame(columns=[
            "Ledger Name","Legal Entity","Cost Organization","Cost Book",
            "Inventory Org","Profit Center BU","Management BU","Manufacturing Plant"
        ])
    else:
        # Remove rows where both Cost Org and Inventory Org are blank
        df2 = df2.loc[~((df2["Cost Organization"].fillna("") == "") & (df2["Inventory Org"].fillna("") == ""))]

    # Order & index
    df2["__LedgerEmpty"] = (df2["Ledger Name"].fillna("") == "").astype(int)
    sort_cols = ["__LedgerEmpty", "Ledger Name", "Legal Entity", "Cost Organization", "Cost Book", "Inventory Org"]
    df2 = (
        df2.sort_values(sort_cols, ascending=[True, True, True, True, True, True])
           .drop(columns="__LedgerEmpty")
           .reset_index(drop=True)
    )
    df2.insert(0, "Assignment", range(1, len(df2) + 1))

    # ------------ Clean blanks & Export ------------
    df1_clean = blankify(df1)
    df2_clean = blankify(df2)

    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        df1_clean.to_excel(writer, index=False, sheet_name="Ledger_LE_BU_Assignments")
        df2_clean.to_excel(writer, index=False, sheet_name="Ledger_LE_CostOrg_Books")

    st.success(f"Built {len(df1_clean)} BU rows and {len(df2_clean)} Inventory Org rows.")
    st.dataframe(df1_clean.head(25), use_container_width=True, height=280)
    st.dataframe(df2_clean.head(25), use_container_width=True, height=280)

    st.download_button(
        "⬇️ Download Excel (EnterpriseStructure.xlsx)",
        data=excel_buf.getvalue(),
        file_name="EnterpriseStructure.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )



 # ===================== DRAW.IO DIAGRAM BLOCK (IO min-gap + ledger group padding) =====================
if (
    "df1" in locals() and isinstance(df1, pd.DataFrame) and not df1.empty and
    "df2" in locals() and isinstance(df2, pd.DataFrame)
):
    import xml.etree.ElementTree as ET
    import zlib, base64, uuid

    def _make_drawio_xml(df_bu: pd.DataFrame, df_tab2: pd.DataFrame) -> str:
        # --- layout & spacing ---
        W, H           = 180, 48
        X_STEP         = 230                # spacing for BU/CO/Books columns
        IO_BASE_STEP   = 180                # desired base spacing for IOs
        IO_GAP         = 40                 # extra horizontal breathing room between IO boxes
        IO_STEP        = max(IO_BASE_STEP, W + IO_GAP)   # HARD minimum to prevent overlap

        LEDGER_PAD     = 320                # minimum horizontal gap between ledger "umbrellas"
        PAD_GROUP      = 60                 # minor pad inside a ledger cluster
        LEFT_PAD       = 260
        RIGHT_PAD      = 200

        Y_LEDGER   = 150
        Y_LE       = 310
        Y_BU       = 470
        Y_CO       = 630
        Y_CB       = 790
        Y_IO       = 1060  # lower so the merge point sits below the book row

        # --- styles ---
        S_LEDGER = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE6E6;strokeColor=#C86868;fontSize=12;"
        S_LE     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE2C2;strokeColor=#A66000;fontSize=12;"
        S_BU     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFF1B3;strokeColor=#B38F00;fontSize=12;"
        S_CO     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#E2F7E2;strokeColor=#3D8B3D;fontSize=12;"
        S_CB     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#7FBF7F;strokeColor=#2F7D2F;fontSize=12;"
        S_IO     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#D6EFFF;strokeColor=#2F71A8;fontSize=12;"
        S_IO_PLT = "rounded=1;whiteSpace=wrap;html=1;fillColor=#D6EFFF;strokeColor=#1F4D7A;strokeWidth=2;fontSize=12;"

        # Edges: child top-center → parent bottom-center (orthogonal elbows)
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
            # Track x-positions used by this ledger to compute its span
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

                    # Under each CO: books to LEFT, IOs centered UNDER
                    for c in cos:
                        base = co_x[c]

                        # Books
                        books = sorted(dict.fromkeys(cb_by_co.get((L, le, c), [])))
                        for i, bk in enumerate(books, start=1):
                            x_pos = base - i*X_STEP
                            cb_x[(L, le, c, bk)] = x_pos
                            ledger_x_used.append(x_pos)

                        # IOs (centered under CO), enforce min spacing
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

                # center ledger over its LEs
                xs_led = [le_x[(L, le)] for le in les]
                led_x[L] = int(sum(xs_led)/len(xs_led)) if xs_led else next_x
                ledger_x_used.append(led_x[L])

                # after finishing this ledger, push the cursor to at least max_x + LEDGER_PAD
                max_x_used = max(ledger_x_used) if ledger_x_used else next_x
                next_x = max(next_x + PAD_GROUP, max_x_used + LEDGER_PAD)

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

        def add_text(text, x, y):
            tid = uuid.uuid4().hex[:8]
            t = ET.SubElement(root, "mxCell", attrib={
                "id": tid, "value": text, "style": S_HDR, "vertex": "1", "parent": "1"})
            ET.SubElement(t, "mxGeometry", attrib={
                "x": str(int(x)), "y": str(int(y)), "width": "260", "height": "22", "as": "geometry"})
            return tid

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

                    # Books
                    for bk in sorted(set(cb_by_co.get((L, le, c), []))):
                        id_map[("CB", L, le, c, bk)] = add_vertex(bk, S_CB, cb_x[(L, le, c, bk)], Y_CB)

                    # IOs
                    for io in sorted(io_by_co.get((L, le, c), []), key=lambda k: k["Name"]):
                        label = f"🏭 {io['Name']}" if str(io["Mfg"]).lower() == "yes" else io["Name"]
                        style = S_IO_PLT if str(io["Mfg"]).lower() == "yes" else S_IO
                        id_map[("IO", L, le, c, io["Name"])] = add_vertex(label, style, io_x[(L, le, c, io["Name"])], Y_IO)

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


