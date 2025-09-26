import io, zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Enterprise Structure Generator", page_icon="📊", layout="wide")
st.title("Enterprise Structure Generator — Excel + Diagram (with Cost Orgs)")

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
    st.info("Upload your ZIPs to generate the Excel and diagram.")
else:
    # ---------------- Collectors ----------------
    ledger_names = set()               # GL_PRIMARY_LEDGER.csv :: ORA_GL_PRIMARY_LEDGER_CONFIG.Name
    legal_entity_names = set()         # XLE_ENTITY_PROFILE.csv :: Name
    ledger_to_le_ids = {}              # ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv :: GL_LEDGER.Name -> {LegalEntityIdentifier}
    ident_to_le_name = {}              # XLE_ENTITY_PROFILE/ORA_GL_JOURNAL_CONFIG_DETAIL :: ident -> LE Name
    bu_rows = []                       # FUN_BUSINESS_UNIT.csv :: Name, PrimaryLedgerName, LegalEntityName
    costorg_rows = []                  # CST_COST_ORGANIZATION.csv :: Name, LegalEntityIdentifier

    # ---------------- Scan uploads ----------------
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

        # Legal Entities (primary source: name + identifier)
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
                        ledger_to_le_ids.setdefault(led, set()).add(ident)
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

        # Business Units
        df = read_csv_from_zip(z, "FUN_BUSINESS_UNIT.csv")
        if df is not None:
            need = {"Name", "PrimaryLedgerName", "LegalEntityName"}
            if need.issubset(df.columns):
                for _, r in df[list(need)].dropna(how="all").iterrows():
                    bu_rows.append({
                        "Name": str(r["Name"]).strip(),
                        "PrimaryLedgerName": str(r["PrimaryLedgerName"]).strip(),
                        "LegalEntityName": str(r["LegalEntityName"]).strip(),
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
                        "LegalEntityIdentifier": str(r["LegalEntityIdentifier"]).strip(),
                    })
            else:
                st.warning(f"`CST_COST_ORGANIZATION.csv` missing {sorted(need - set(df.columns))}. Found: {list(df.columns)}")

    # --------------- Build maps (with duplicates handled) ---------------
    # 1) ledger -> {LE names}  (display purposes)
    ledger_to_le_names = {}
    for led, ident_set in ledger_to_le_ids.items():
        for ident in ident_set:
            le_name = ident_to_le_name.get(ident, "").strip()
            if le_name:
                ledger_to_le_names.setdefault(led, set()).add(le_name)

    # 2) Reverse: identifier -> {ledgers}
    ident_to_ledgers = {}
    for led, ident_set in ledger_to_le_ids.items():
        for ident in ident_set:
            ident_to_ledgers.setdefault(ident, set()).add(led)

    # 3) Reverse by name within a ledger (for BU disambiguation)
    #    We use (ledger, le_name) as the unique key for Tab 1.
    #    No global assumption about uniqueness of le_name across ledgers.
    # (No extra structure required beyond the BU rows themselves.)

    # =========================
    # Tab 1: Ledger – LE – BU
    # =========================
    rows1 = []
    seen_triples = set()
    seen_ledgers_with_bu = set()
    seen_les_with_bu_by_ledger = set()  # keys: (ledger, le_name)

    # 1) BU-driven rows; treat (ledger, le_name) as the unique LE key
    for r in bu_rows:
        bu = r["Name"]
        led = r["PrimaryLedgerName"] if r["PrimaryLedgerName"] in ledger_names else ""
        le  = r["LegalEntityName"]

        # only accept LE name if known in that ledger (prevents cross-ledger bleed)
        if led and le and le not in ledger_to_le_names.get(led, set()):
            # fallback: if ledger unknown but the le_name appears in exactly one ledger, assign it
            led_candidates = [L for L, names in ledger_to_le_names.items() if le in names]
            if not led and len(led_candidates) == 1:
                led = led_candidates[0]
            else:
                # keep as-is (may show as hanging later)
                pass

        rows1.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": bu})
        seen_triples.add((led, le, bu))
        if led: seen_ledgers_with_bu.add(led)
        if led and le: seen_les_with_bu_by_ledger.add((led, le))

    # 2) Ledger–LE pairs with no BU (emit per-ledger)
    seen_pairs = {(a, b) for (a, b, _) in seen_triples}
    for led, names in ledger_to_le_names.items():
        for le in names:
            if (led, le) not in seen_pairs:
                rows1.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": ""})

    # 3) Orphan ledgers present in master list (no mapping & no BUs)
    for led in sorted(ledger_names - set(ledger_to_le_names.keys()) - seen_ledgers_with_bu):
        rows1.append({"Ledger Name": led, "Legal Entity": "", "Business Unit": ""})

    # 4) Orphan LEs by name that appeared in master list but not via BU under any ledger
    #    We cannot place them under a specific ledger unless that (name) appears in exactly one ledger.
    for le in sorted(legal_entity_names):
        # If LE name appears under some ledger(s) in ledger_to_le_names, we handled it above.
        appearing_ledgers = [L for L, names in ledger_to_le_names.items() if le in names]
        if not appearing_ledgers:
            # no link to any ledger via mapping; emit hanging row
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

    # ===========================================
    # Tab 2: Ledger – LE – Cost Organization
    # ===========================================
    rows2 = []
    seen_tuples2 = set()
    seen_ledgers_with_co = set()
    seen_les_with_co_by_ledger = set()

    for r in costorg_rows:
        co = r["Name"]
        ident = r["LegalEntityIdentifier"]
        le = ident_to_le_name.get(ident, "")

        ledgers_for_ident = sorted(ident_to_ledgers.get(ident, set()))
        if not ledgers_for_ident:
            # Hanging (no ledger known for this identifier)
            rows2.append({"Ledger Name": "", "Legal Entity": le, "Cost Organization": co})
            seen_tuples2.add(("", le, co))
            continue

        # Emit one row per ledger if identifier is attached to multiple ledgers
        for led in ledgers_for_ident:
            rows2.append({"Ledger Name": led, "Legal Entity": le, "Cost Organization": co})
            seen_tuples2.add((led, le, co))
            seen_ledgers_with_co.add(led)
            seen_les_with_co_by_ledger.add((led, le))

    # Add hanging LEs (that appear in mapping) but have no cost org rows, per-ledger
    for led, names in ledger_to_le_names.items():
        for le in names:
            if (led, le) not in {(a, b) for (a, b, _) in seen_tuples2}:
                rows2.append({"Ledger Name": led, "Legal Entity": le, "Cost Organization": ""})

    # Orphan ledgers present in master list but no CO rows
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

    # ---------------- Excel Output ----------------
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

    # ===================== DRAW.IO DIAGRAM BLOCK =====================
    if not df1.empty or not df2.empty:
        import xml.etree.ElementTree as ET
        import zlib, base64, uuid

        def _make_drawio_xml(df_bu: pd.DataFrame, df_co: pd.DataFrame) -> str:
            # --- layout & spacing ---
            W, H       = 180, 48
            X_STEP     = 230
            PAD_GROUP  = 60
            LEFT_PAD   = 260
            RIGHT_PAD  = 160

            Y_LEDGER   = 150
            Y_LE       = 310
            Y_BU       = 470
            Y_CO       = 630

            # --- styles ---
            S_LEDGER = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE6E6;strokeColor=#C86868;fontSize=12;"
            S_LE     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE2C2;strokeColor=#A66000;fontSize=12;"
            S_BU     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFF1B3;strokeColor=#B38F00;fontSize=12;"
            S_CO     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#E2F7E2;strokeColor=#3D8B3D;fontSize=12;"

            S_EDGE = (
                "endArrow=block;rounded=1;edgeStyle=orthogonalEdgeStyle;orthogonal=1;"
                "jettySize=auto;strokeColor=#666666;exitX=0.5;exitY=0;entryX=0.5;entryY=1;"
            )

            # normalize inputs
            bu = df_bu[["Ledger Name", "Legal Entity", "Business Unit"]].copy()
            co = df_co[["Ledger Name", "Legal Entity", "Cost Organization"]].copy()
            for df in (bu, co):
                for c in df.columns:
                    df[c] = df[c].fillna("").map(str).str.strip()

            ledgers = sorted(set(bu["Ledger Name"]).union(set(co["Ledger Name"])) - {""})

            # maps
            le_map = {}   # {ledger: set(LEs)}
            for _, r in pd.concat([bu, co], ignore_index=True).iterrows():
                if r["Ledger Name"] and r["Legal Entity"]:
                    le_map.setdefault(r["Ledger Name"], set()).add(r["Legal Entity"])

            bu_map = {}   # {(ledger, le): set(BUs)}
            for _, r in bu.iterrows():
                if r["Ledger Name"] and r["Legal Entity"] and r["Business Unit"]:
                    bu_map.setdefault((r["Ledger Name"], r["Legal Entity"]), set()).add(r["Business Unit"])

            co_map = {}   # {(ledger, le): set(COs)}
            for _, r in co.iterrows():
                if r["Ledger Name"] and r["Legal Entity"] and r["Cost Organization"]:
                    co_map.setdefault((r["Ledger Name"], r["Legal Entity"]), set()).add(r["Cost Organization"])

            # x-coordinates; offset COs by half a step to reduce arrow overlap
            next_x = LEFT_PAD
            led_x, le_x, bu_x, co_x = {}, {}, {}, {}

            for L in ledgers:
                les = sorted(le_map.get(L, []))
                for le in les:
                    buses = sorted(bu_map.get((L, le), []
