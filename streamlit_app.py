import io, zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Enterprise Structure Generator 2", page_icon="🧭", layout="wide")
st.title("Enterprise Structure Generator 2 — Core + Cost Org lane")

uploads = st.file_uploader("Upload Oracle export ZIPs", type="zip", accept_multiple_files=True)

def read_csv_from_zip(zf, name):
    if name not in zf.namelist():
        return None
    with zf.open(name) as fh:
        return pd.read_csv(fh, dtype=str)

if uploads:
    # ---------------- Core collectors ----------------
    ledger_names = set()
    legal_entity_names = set()
    ledger_to_idents = {}
    ident_to_le_name = {}
    bu_rows = []

    # ---------------- Cost org collectors ----------------
    costorg_rows = []
    code_to_ledger = {}

    # ---------------- Scan zips ----------------
    for up in uploads:
        try:
            z = zipfile.ZipFile(up)
        except Exception as e:
            st.error(f"Could not open `{up.name}` as a ZIP: {e}")
            continue

        # Ledgers
        df = read_csv_from_zip(z, "GL_PRIMARY_LEDGER.csv")
        if df is not None and "ORA_GL_PRIMARY_LEDGER_CONFIG.Name" in df.columns:
            ledger_names |= set(df["ORA_GL_PRIMARY_LEDGER_CONFIG.Name"].dropna().str.strip())

        # Legal Entities
        df = read_csv_from_zip(z, "XLE_ENTITY_PROFILE.csv")
        if df is not None and "Name" in df.columns:
            legal_entity_names |= set(df["Name"].dropna().str.strip())

        # Ledger ↔ LE identifier
        df = read_csv_from_zip(z, "ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv")
        if df is not None and {"GL_LEDGER.Name", "LegalEntityIdentifier"} <= set(df.columns):
            for _, r in df[["GL_LEDGER.Name", "LegalEntityIdentifier"]].dropna().iterrows():
                ledger_to_idents.setdefault(r["GL_LEDGER.Name"].strip(), set()).add(r["LegalEntityIdentifier"].strip())

        # Identifier ↔ LE name
        df = read_csv_from_zip(z, "ORA_GL_JOURNAL_CONFIG_DETAIL.csv")
        if df is not None and {"LegalEntityIdentifier", "ObjectName"} <= set(df.columns):
            for _, r in df[["LegalEntityIdentifier", "ObjectName"]].dropna().iterrows():
                ident_to_le_name[r["LegalEntityIdentifier"].strip()] = r["ObjectName"].strip()

        # Business Units
        df = read_csv_from_zip(z, "FUN_BUSINESS_UNIT.csv")
        if df is not None and {"Name","PrimaryLedgerName","LegalEntityName"} <= set(df.columns):
            for _, r in df.iterrows():
                bu_rows.append({
                    "Name": str(r["Name"]).strip(),
                    "PrimaryLedgerName": str(r["PrimaryLedgerName"]).strip(),
                    "LegalEntityName": str(r["LegalEntityName"]).strip()
                })

        # Cost Org master
        df = read_csv_from_zip(z, "CST_COST_ORGANIZATION.csv")
        if df is not None and {"Name","LegalEntityIdentifier","OrgInformation2"} <= set(df.columns):
            for _, r in df.iterrows():
                costorg_rows.append({
                    "CostOrgName": str(r["Name"]).strip(),
                    "LE_Ident": str(r["LegalEntityIdentifier"]).strip(),
                    "CostOrgCode": str(r["OrgInformation2"]).strip()
                })

        # Cost Org → Ledger mapping
        df = read_csv_from_zip(z, "CST_COST_ORG_BOOK.csv")
        if df is not None and {"ORA_CST_ACCT_COST_ORG.CostOrgCode","Name"} <= set(df.columns):
            for _, r in df.iterrows():
                code = str(r["ORA_CST_ACCT_COST_ORG.CostOrgCode"]).strip()
                ledger = str(r["Name"]).strip()
                if code and ledger:
                    code_to_ledger.setdefault(code, set()).add(ledger)

    # ---------------- Build core sheet (Ledger–LE–BU) ----------------
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

    rows = []
    seen_triples = set()
    for r in bu_rows:
        bu, led, le = r["Name"], r["PrimaryLedgerName"], r["LegalEntityName"]
        rows.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": bu})
        seen_triples.add((led, le, bu))

    for led, le_set in ledger_to_le_names.items():
        for le in le_set:
            if (led, le, "") not in seen_triples:
                rows.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": ""})

    df_core = pd.DataFrame(rows).drop_duplicates().reset_index(drop=True)
    st.success(f"Sheet 1: {len(df_core)} rows (Ledger–LE–BU)")

    # ---------------- Build cost org sheet (Ledger–LE–CostOrg) ----------------
    out_rows = []
    for r in costorg_rows:
        le_name = ident_to_le_name.get(r["LE_Ident"], "").strip()
        cname, ccode = r["CostOrgName"], r["CostOrgCode"]
        ledgers = code_to_ledger.get(ccode, [])
        if ledgers:
            for L in ledgers:
                out_rows.append({"Ledger Name": L, "Legal Entity": le_name, "Business Unit": "", "Cost Organization": cname})
        else:
            out_rows.append({"Ledger Name": "", "Legal Entity": le_name, "Business Unit": "", "Cost Organization": cname})

    df_cost = pd.DataFrame(out_rows).drop_duplicates().reset_index(drop=True)
    st.success(f"Sheet 2: {len(df_cost)} rows (Ledger–LE–CostOrg, orphans included)")

    # ---------------- Excel export ----------------
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df_core.to_excel(writer, index=False, sheet_name="Core_Ledger_LE_BU")
        df_cost.to_excel(writer, index=False, sheet_name="Ledger_LE_CostOrg")
    st.download_button("⬇️ Download Excel", data=buf.getvalue(), file_name="EnterpriseStructure_v2.xlsx")

else:
    st.info("Upload your ZIPs to generate outputs.")
