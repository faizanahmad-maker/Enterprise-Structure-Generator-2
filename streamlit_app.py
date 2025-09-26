import io, zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Enterprise Structure Generator", page_icon="📊", layout="wide")
st.title("Enterprise Structure Generator — Excel with Cost Orgs")

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
    if name not in zf.namelist():
        return None
    with zf.open(name) as fh:
        return pd.read_csv(fh, dtype=str)

if not uploads:
    st.info("Upload your ZIPs to generate the Excel.")
else:
    # collectors for BU logic
    ledger_names = set()
    legal_entity_names = set()
    ledger_to_idents = {}
    ident_to_le_name = {}
    bu_rows = []

    # collectors for Cost Org logic
    costorg_rows = []  # Name, LegalEntityIdentifier

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

        # Legal Entities
        df = read_csv_from_zip(z, "XLE_ENTITY_PROFILE.csv")
        if df is not None and {"Name","LegalEntityIdentifier"}.issubset(df.columns):
            for _, r in df.iterrows():
                le_name = str(r["Name"]).strip()
                le_ident = str(r["LegalEntityIdentifier"]).strip()
                if le_name: legal_entity_names.add(le_name)
                if le_ident: ident_to_le_name[le_ident] = le_name

        # Ledger ↔ LE identifier
        df = read_csv_from_zip(z, "ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv")
        if df is not None and {"GL_LEDGER.Name","LegalEntityIdentifier"}.issubset(df.columns):
            for _, r in df.iterrows():
                led = str(r["GL_LEDGER.Name"]).strip()
                ident = str(r["LegalEntityIdentifier"]).strip()
                if led and ident:
                    ledger_to_idents.setdefault(led, set()).add(ident)

        # Business Units
        df = read_csv_from_zip(z, "FUN_BUSINESS_UNIT.csv")
        if df is not None and {"Name","PrimaryLedgerName","LegalEntityName"}.issubset(df.columns):
            for _, r in df.iterrows():
                bu_rows.append({
                    "Name": str(r["Name"]).strip(),
                    "PrimaryLedgerName": str(r["PrimaryLedgerName"]).strip(),
                    "LegalEntityName": str(r["LegalEntityName"]).strip()
                })

        # Cost Orgs
        df = read_csv_from_zip(z, "CST_COST_ORGANIZATION.csv")
        if df is not None and {"Name","LegalEntityIdentifier"}.issubset(df.columns):
            for _, r in df.iterrows():
                costorg_rows.append({
                    "Name": str(r["Name"]).strip(),
                    "LegalEntityIdentifier": str(r["LegalEntityIdentifier"]).strip()
                })

    # === Build Ledger→LE maps ===
    ledger_to_le_names = {}
    for led, idents in ledger_to_idents.items():
        for ident in idents:
            le_name = ident_to_le_name.get(ident,"").strip()
            if le_name:
                ledger_to_le_names.setdefault(led,set()).add(le_name)

    le_to_ledgers = {}
    for led, le_set in ledger_to_le_names.items():
        for le in le_set:
            le_to_ledgers.setdefault(le,set()).add(led)

    # === Tab 1: Ledger – LE – BU ===
    rows = []
    seen_triples, seen_ledgers_with_bu, seen_les_with_bu = set(), set(), set()

    # BU-driven rows
    for r in bu_rows:
        bu = r["Name"]
        led = r["PrimaryLedgerName"] if r["PrimaryLedgerName"] in ledger_names else ""
        le  = r["LegalEntityName"]  if r["LegalEntityName"]  in legal_entity_names else ""

        # back-fill ledger from LE if missing
        if not led and le and le in le_to_ledgers and len(le_to_ledgers[le])==1:
            led = next(iter(le_to_ledgers[le]))
        # back-fill LE from ledger if missing
        if not le and led and led in ledger_to_le_names and len(ledger_to_le_names[led])==1:
            le = next(iter(ledger_to_le_names[led]))

        rows.append({"Ledger Name": led,"Legal Entity": le,"Business Unit": bu})
        seen_triples.add((led,le,bu))
        if led: seen_ledgers_with_bu.add(led)
        if le:  seen_les_with_bu.add(le)

    # Ledger–LE pairs with no BU
    seen_pairs = {(a,b) for (a,b,_) in seen_triples}
    for led, le_set in ledger_to_le_names.items():
        for le in le_set:
            if (led,le) not in seen_pairs:
                rows.append({"Ledger Name": led,"Legal Entity": le,"Business Unit": ""})

    # Orphan ledgers
    for led in sorted(ledger_names - set(ledger_to_le_names.keys()) - seen_ledgers_with_bu):
        rows.append({"Ledger Name": led,"Legal Entity": "","Business Unit": ""})

    # Orphan LEs
    for le in sorted(legal_entity_names - seen_les_with_bu):
        led = next(iter(le_to_ledgers[le])) if le in le_to_ledgers and len(le_to_ledgers[le])==1 else ""
        rows.append({"Ledger Name": led,"Legal Entity": le,"Business Unit": ""})

    df1 = pd.DataFrame(rows).drop_duplicates().reset_index(drop=True)
    df1["__LedgerEmpty"] = (df1["Ledger Name"]=="").astype(int)
    df1 = df1.sort_values(["__LedgerEmpty","Ledger Name","Legal Entity","Business Unit"],
                          ascending=[True,True,True,True]).drop(columns="__LedgerEmpty").reset_index(drop=True)
    df1.insert(0,"Assignment",range(1,len(df1)+1))

    # === Tab 2: Ledger – LE – Cost Org ===
    rows2 = []
    seen_pairs2 = set()

    for r in costorg_rows:
        co = r["Name"]
        ident = r["LegalEntityIdentifier"]
        le = ident_to_le_name.get(ident,"")
        led = ""
        if le in le_to_ledgers and len(le_to_ledgers[le])==1:
            led = next(iter(le_to_ledgers[le]))
        rows2.append({"Ledger Name": led,"Legal Entity": le,"Cost Organization": co})
        seen_pairs2.add((led,le,co))

    # add hanging LEs
    for le in legal_entity_names:
        if not any(le==p[1] for p in seen_pairs2):
            led = next(iter(le_to_ledgers[le])) if le in le_to_ledgers and len(le_to_ledgers[le])==1 else ""
            rows2.append({"Ledger Name": led,"Legal Entity": le,"Cost Organization": ""})

    df2 = pd.DataFrame(rows2).drop_duplicates().reset_index(drop=True)
    df2["__LedgerEmpty"] = (df2["Ledger Name"]=="").astype(int)
    df2 = df2.sort_values(["__LedgerEmpty","Ledger Name","Legal Entity","Cost Organization"],
                          ascending=[True,True,True,True]).drop(columns="__LedgerEmpty").reset_index(drop=True)
    df2.insert(0,"Assignment",range(1,len(df2)+1))

    # === Excel Output ===
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        df1.to_excel(writer, index=False, sheet_name="Ledger_LE_BU_Assignments")
        df2.to_excel(writer, index=False, sheet_name="Ledger_LE_CostOrg_Assignments")

    st.success(f"Built {len(df1)} BU rows and {len(df2)} Cost Org rows.")
    st.dataframe(df1.head(20), use_container_width=True, height=250)
    st.dataframe(df
