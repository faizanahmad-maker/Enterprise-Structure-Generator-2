import io, zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Enterprise Structure Generator", page_icon="📊", layout="wide")
st.title("Enterprise Structure Generator — Excel with Cost Orgs (Two Tabs)")

st.markdown("""
Upload up to **5 Oracle export ZIPs** (any order):
- `Manage General Ledger` (Ledgers)
- `Manage Legal Entities` (Legal Entities)
- `Assign Legal Entities` (Ledger↔LE mapping)
- `Manage Business Units` (Business Units)
- `Manage Cost Organizations` (Cost Orgs)
""")

uploads = st.file_uploader("Drop your ZIPs here", type="zip", accept_multiple_files=True)

def read_csv_from_zip(zf, name):
    """Return a CSV as DataFrame if present in the zip; else None."""
    if name not in zf.namelist():
        return None
    with zf.open(name) as fh:
        return pd.read_csv(fh, dtype=str)

if not uploads:
    st.info("Upload your ZIPs to generate the Excel.")
else:
    # ------------ Collectors ------------
    # For BU tab (Tab 1)
    ledger_names = set()                 # GL_PRIMARY_LEDGER.csv :: ORA_GL_PRIMARY_LEDGER_CONFIG.Name
    legal_entity_names = set()           # XLE_ENTITY_PROFILE.csv :: Name
    ledger_to_idents = {}                # ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv :: GL_LEDGER.Name -> {LegalEntityIdentifier}
    ident_to_le_name = {}                # XLE_ENTITY_PROFILE / ORA_GL_JOURNAL_CONFIG_DETAIL
    bu_rows = []                         # FUN_BUSINESS_UNIT.csv :: Name, PrimaryLedgerName, LegalEntityName

    # For Cost Org tab (Tab 2)
    costorg_rows = []                    # CST_COST_ORGANIZATION.csv :: Name, LegalEntityIdentifier

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
            col = "ORA_GL_PRIMARY_LEDGER_CONFIG.Name"
            if col in df.columns:
                ledger_names |= set(df[col].dropna().map(str).str.strip())
            else:
                st.warning(f"`GL_PRIMARY_LEDGER.csv` missing `{col}`. Found: {list(df.columns)}")

        # Legal Entities (primary source for ident->name)
        df = read_csv_from_zip(z, "XLE_ENTITY_PROFILE.csv")
        if df is not None:
            need = {"Name", "LegalEntityIdentifier"}
            if need.issubset(df.columns):
                for _, r in df[list(need)].dropna(how="all").iterrows():
                    le_name = str(r["Name"]).strip()
                    le_ident = str(r["LegalEntityIdentifier"]).strip()
                    if le_name:
                        legal_entity_names.add(le_name)
                    if le_ident and le_name:
                        ident_to_le_name[le_ident] = le_name
            else:
                st.warning(f"`XLE_ENTITY_PROFILE.csv` missing {sorted(need - set(df.columns))}. Found: {list(df.columns)}")

        # Ledger ↔ LE identifier
        df = read_csv_from_zip(z, "ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv")
        if df is not None:
            need = {"GL_LEDGER.Name", "LegalEntityIdentifier"}
            if need.issubset(df.columns):
                for _, r in df[list(need)].dropna(how="all").iterrows():
                    led = str(r["GL_LEDGER.Name"]).strip()
                    ident = str(r["LegalEntityIdentifier"]).strip()
                    if led and ident:
                        ledger_to_idents.setdefault(led, set()).add(ident)
            else:
                st.warning(f"`ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv` missing {sorted(need - set(df.columns))}. Found: {list(df.columns)}")

        # Identifier ↔ LE name (backup source)
        df = read_csv_from_zip(z, "ORA_GL_JOURNAL_CONFIG_DETAIL.csv")
        if df is not None:
            need = {"LegalEntityIdentifier", "ObjectName"}
            if need.issubset(df.columns):
                for _, r in df[list(need)].dropna(how="all").iterrows():
                    ident = str(r["LegalEntityIdentifier"]).strip()
                    obj = str(r["ObjectName"]).strip()
                    if ident and obj and ident not in ident_to_le_name:
                        ident_to_le_name[ident] = obj
            # no warning if absent; it's merely a backup source

        # Business Units
        df = read_csv_from_zip(z, "FUN_BUSINESS_UNIT.csv")
        if df is not None:
            need = {"Name", "PrimaryLedgerName", "LegalEntityName"}
            if need.issubset(df.columns):
                for _, r in df[list(need)].dropna(how="all").iterrows():
                    bu_rows.append({
                        "Name": str(r["Name"]).strip(),
                        "PrimaryLedgerName": str(r["PrimaryLedgerName"]).strip(),
                        "LegalEntityName": str(r["LegalEntityName"]).strip()
                    })
            else:
                st.warning(f"`FUN_BUSINESS_UNIT.csv` missing {sorted(need - set(df.columns))}. Found: {list(df.columns)}")

        # Cost Orgs
        df = read_csv_from_zip(z, "CST_COST_ORGANIZATION.csv")
        if df is not None:
            need = {"Name", "LegalEntityIdentifier"}
            if need.issubset(df.columns):
                for _, r in df[list(need)].dropna(how="all").iterrows():
                    costorg_rows.append({
                        "Name": str(r["Name"]).strip(),
                        "LegalEntityIdentifier": str(r["LegalEntityIdentifier"]).strip()
                    })
            else:
                st.warning(f"`CST_COST_ORGANIZATION.csv` missing {sorted(need - set(df.columns))}. Found: {list(df.columns)}")

    # ------------ Build Ledger <-> LE maps ------------
    ledger_to_le_names = {}
    for led, idents in ledger_to_idents.items():
        for ident in idents:
            le_name = ident_to_le_name.get(ident, "").strip()
            if le_name:
                ledger_to_le_names.setdefault(led, set()).add(le_name)

    le_to_ledgers = {}
    for led, le_set in ledger_to_le_names.items():
        for le in le_set:
            le_to_ledgers.setdefault(le, set()).add(led)

    # ===================================================
    # Tab 1: Ledger – Legal Entity – Business Unit
    # ===================================================
    rows1 = []
    seen_triples = set()
    seen_ledgers_with_bu = set()
    seen_les_with_bu = set()

    # 1) BU-driven rows with smart back-fill
    for r in bu_rows:
        bu = r["Name"]
        led = r["PrimaryLedgerName"] if r["PrimaryLedgerName"] in ledger_names else ""
        le  = r["LegalEntityName"]  if r["LegalEntityName"]  in legal_entity_names else ""

        # back-fill ledger from LE if missing and unique
        if not led and le and le in le_to_ledgers and len(le_to_ledgers[le]) == 1:
            led = next(iter(le_to_ledgers[le]))
        # back-fill LE from ledger if missing and unique
        if not le and led and led in ledger_to_le_names and len(ledger_to_le_names[led]) == 1:
            le = next(iter(ledger_to_le_names[led]))

        rows1.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": bu})
        seen_triples.add((led, le, bu))
        if led: seen_ledgers_with_bu.add(led)
        if le:  seen_les_with_bu.add(le)

    # 2) Ledger–LE pairs with no BU
    seen_pairs = {(a, b) for (a, b, _) in seen_triples}
    for led, le_set in ledger_to_le_names.items():
        if not le_set:
            if led not in seen_ledgers_with_bu:
                rows1.append({"Ledger Name": led, "Legal Entity": "", "Business Unit": ""})
            continue
        for le in le_set:
            if (led, le) not in seen_pairs:
                rows1.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": ""})

    # 3) Orphan Ledgers in master list (no mapping & no BUs)
    for led in sorted(ledger_names - set(ledger_to_le_names.keys()) - seen_ledgers_with_bu):
        rows1.append({"Ledger Name": led, "Legal Entity": "", "Business Unit": ""})

    # 4) Orphan LEs (in master list) with no BU; back-fill ledger if uniquely known
    for le in sorted(legal_entity_names - seen_les_with_bu):
        led = next(iter(le_to_ledgers[le])) if le in le_to_ledgers and len(le_to_ledgers[le]) == 1 else ""
        rows1.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": ""})

    df1 = pd.DataFrame(rows1).drop_duplicates().reset_index(drop=True)
    # ordering: non-empty ledgers first, then by Ledger, LE, BU
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
    # Tab 2: Ledger – Legal Entity – Cost Organization
    # ===================================================
    rows2 = []
    seen_triples2 = set()
    seen_ledgers_with_co = set()
    seen_les_with_co = set()

    # Cost Org rows (LE-ident -> LE name -> (unique) Ledger)
    for r in costorg_rows:
        co = r["Name"]
        ident = r["LegalEntityIdentifier"]
        le = ident_to_le_name.get(ident, "")
        led = ""
        if le in le_to_ledgers and len(le_to_ledgers[le]) == 1:
            led = next(iter(le_to_ledgers[le]))
        rows2.append({"Ledger Name": led, "Legal Entity": le, "Cost Organization": co})
        seen_triples2.add((led, le, co))
        if led: seen_ledgers_with_co.add(led)
        if le:  seen_les_with_co.add(le)

    # Hanging LEs (no cost orgs); back-fill ledger if uniquely known
    for le in sorted(legal_entity_names - seen_les_with_co):
        led = next(iter(le_to_ledgers[le])) if le in le_to_ledgers and len(le_to_ledgers[le]) == 1 else ""
        rows2.append({"Ledger Name": led, "Legal Entity": le, "Cost Organization": ""})

    # Orphan Ledgers (present in ledgers list but no CO rows)
    for led in sorted(ledger_names - seen_ledgers_with_co):
        rows2.append({"Ledger Name": led, "Legal Entity": "", "Cost Organization": ""})

    df2 = pd.DataFrame(rows2).drop_duplicates().reset_index(drop=True)
    df2["__LedgerEmpty"] = (df2["Ledger Name"] == "").astype(int)
    df2 = (
        df2.sort_values(
            ["__LedgerEmpty", "Ledger Name", "Legal Entity", "Cost Organization"],
            ascending=[True, True, True, True]
        )
        .drop(columns="__LedgerEmpty")
        .reset_index(drop=True)
    )
    df2.insert(0, "Assignment", range(1, len(df2) + 1))

    # ------------ Excel Output ------------
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        df1.to_excel(writer, index=False, sheet_name="Ledger_LE_BU_Assignments")
        df2.to_excel(writer, index=False, sheet_name="Ledger_LE_CostOrg_Assignments")

    st.success(f"Built {len(df1)} BU rows and {len(df2)} Cost Org rows.")
    st.dataframe(df1.head(25), use_container_width=True, height=280)
    st.dataframe(df2.head(25), use_container_width=True, height=280)

    st.download_button(
        "⬇️ Download Excel (EnterpriseStructure.xlsx)",
        data=excel_buf.getvalue(),
        file_name="EnterpriseStructure.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
