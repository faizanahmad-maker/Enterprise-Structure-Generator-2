import io, zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Enterprise Structure Generator", page_icon="📊", layout="wide")
st.title("Enterprise Structure Generator — Excel + draw.io (with Inventory Orgs)")

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

if not uploads:
    st.info("Upload your ZIPs to generate the Excel & diagram.")
else:
    # ------------ Collectors ------------
    ledger_names = set()
    legal_entity_names = set()
    ledger_to_idents = {}            # ledger -> {LE identifier}
    ident_to_le_name = {}            # LE identifier -> LE name
    bu_rows = []                     # BU rows (for Tab 1)

    # Cost Orgs (MASTER)
    costorg_rows = []                # [{Name, LegalEntityIdentifier, JoinKey}]
    costorg_name_to_joinkeys = {}    # Name -> {JoinKey}
    books_by_joinkey = {}            # JoinKey -> {CostBookCode}

    # Inventory Orgs (MASTER) + relationships
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
    # Tab 1: Ledger – Legal Entity – Business Unit
    # ===================================================
    rows1, seen_triples, seen_ledgers_with_bu = [], set(), set()

    for r in bu_rows:
        bu  = r["Name"]
        led = r["PrimaryLedgerName"]
        le  = r["LegalEntityName"]
        rows1.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": bu})
        seen_triples.add((led, le, bu))
        if led:
            seen_ledgers_with_bu.add(led)

    seen_pairs = {(a, b) for (a, b, _) in seen_triples}
    for led, le in sorted(known_pairs):
        if (led, le) not in seen_pairs:
            rows1.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": ""})

    mapped_ledgers = set(ledger_to_le_names.keys())
    for led in sorted(ledger_names - mapped_ledgers - seen_ledgers_with_bu):
        rows1.append({"Ledger Name": led, "Legal Entity": "", "Business Unit": ""})

    df1 = pd.DataFrame(rows1).drop_duplicates().reset_index(drop=True)
    df1 = df1.fillna("")  # <-- NaN → blanks (requested)
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
    # Tab 2: Ledger – LE – Cost Org – Cost Book – Inv Org – ProfitCenter BU – Management BU – Mfg Plant
    #  - Emits only when there is an IO and/or a CO (driven by IOs)
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
        books   = "; ".join(sorted(books_by_joinkey.get(co_key, []))) if co_key else ""

        if not name and not co_name:
            continue  # skip rows with neither IO nor CO

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
    df2 = df2.fillna("")  # keep preview/export clean
    if not df2.empty:
        df2["__LedgerEmpty"] = (df2["Ledger Name"] == "").astype(int)
        df2["__COEmpty"]     = (df2["Cost Organization"] == "").astype(int)
        df2 = (
            df2.sort_values(
                ["__LedgerEmpty", "Ledger Name", "Legal Entity", "__COEmpty", "Cost Organization", "Inventory Org"],
                ascending=[True, True, True, True, True, True]
            )
            .drop(columns=["__LedgerEmpty", "__COEmpty"])
            .reset_index(drop=True)
        )
    df2.insert(0, "Assignment", range(1, len(df2) + 1))

    # ------------ Excel Output ------------
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        df1.to_excel(writer, index=False, sheet_name="Ledger_LE_BU_Assignments")
        df2.to_excel(writer, index=False, sheet_name="Ledger_LE_CostOrg_IOs")

    st.success(f"Built {len(df1)} BU rows and {len(df2)} Inventory Org rows.")
    st.dataframe(df1.head(25), use_container_width=True, height=280)
    st.dataframe(df2.head(25), use_container_width=True, height=320)

    st.download_button(
        "⬇️ Download Excel (EnterpriseStructure.xlsx)",
        data=excel_buf.getvalue(),
        file_name="EnterpriseStructure.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ===================== DRAW.IO DIAGRAM BLOCK (CO left=Books, right=IOs; centered edges) =====================
    if not df2.empty:
        import xml.etree.ElementTree as ET
        import zlib, base64, uuid

        def _make_drawio_xml(df_bu: pd.DataFrame, df_tab2: pd.DataFrame) -> str:
            # --- layout & spacing ---
            W, H       = 180, 48
            X_STEP     = 230            # spacing for BUs/COs/Books
            IO_STEP    = max(160, 160)  # minimum spacing for IOs so they never overlap
            PAD_GROUP  = 60
            LEFT_PAD   = 260
            RIGHT_PAD  = 200

            Y_LEDGER   = 150
            Y_LE       = 310
            Y_BU       = 470
            Y_CO       = 630
            Y_CB       = 790
            Y_IO       = 950

            # --- styles ---
            S_LEDGER = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE6E6;strokeColor=#C86868;fontSize=12;"
            S_LE     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE2C2;strokeColor=#A66000;fontSize=12;"
            S_BU     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFF1B3;strokeColor=#B38F00;fontSize=12;"
            S_CO     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#E2F7E2;strokeColor=#3D8B3D;fontSize=12;"
            S_CB     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#7FBF7F;strokeColor=#2F7D2F;fontSize=12;"
            S_IO     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#D6EFFF;strokeColor=#2F71A8;fontSize=12;"
            S_IO_PLT = "rounded=1;whiteSpace=wrap;html=1;fillColor=#D6EFFF;strokeColor=#1F4D7A;strokeWidth=2;fontSize=12;"

            # Edge: top-center (child) -> bottom-center (parent), orthogonal elbows
            S_EDGE   = ("endArrow=block;rounded=1;edgeStyle=orthogonalEdgeStyle;orthogonal=1;"
                        "jettySize=auto;strokeColor=#666666;exitX=0.5;exitY=0;entryX=0.5;entryY=1;")
            S_HDR    = "text;align=left;verticalAlign=middle;fontSize=13;fontStyle=1;"

            # --- normalize input ---
            df_bu = df_bu[["Ledger Name", "Legal Entity", "Business Unit"]].copy()
            df_bu = df_bu.fillna("").map(lambda x: x.strip() if isinstance(x,str) else x)
            df = df_tab2[["Ledger Name","Legal Entity","Cost Organization","Cost Book","Inventory Org","Manufacturing Plant"]].copy()
            for c in df.columns:
                df[c] = df[c].fillna("").map(str).str.strip()

            ledgers_all = sorted([x for x in set(df_bu["Ledger Name"]) | set(df["Ledger Name"]) if x])

            # --- maps ---
            le_map, bu_map, co_map = {}, {}, {}
            cb_map_by_co = {}   # (L,LE,C) -> [books...]
            io_map_by_co = {}   # (L,LE,C) -> [{"Name":..., "Mfg":...}, ...]

            # LE map
            for _, r in pd.concat([df_bu[["Ledger Name","Legal Entity"]],
                                   df[["Ledger Name","Legal Entity"]]]).drop_duplicates().iterrows():
                if r["Ledger Name"] and r["Legal Entity"]:
                    le_map.setdefault(r["Ledger Name"], set()).add(r["Legal Entity"])

            # BU under LE
            for _, r in df_bu.iterrows():
                L, E, B = r["Ledger Name"], r["Legal Entity"], r["Business Unit"]
                if L and E and B:
                    bu_map.setdefault((L,E), set()).add(B)

            # Cost Orgs under LE
            for _, r in df.iterrows():
                L, E, C = r["Ledger Name"], r["Legal Entity"], r["Cost Organization"]
                if L and E and C:
                    co_map.setdefault((L,E), set()).add(C)

            # Books & IOs grouped by CO (not by book for IOs — IOs live as siblings of Books)
            for _, r in df.iterrows():
                L, E, C, B = r["Ledger Name"], r["Legal Entity"], r["Cost Organization"], r["Cost Book"]
                IO, MFG = r["Inventory Org"], r["Manufacturing Plant"]
                if L and E and C and B:
                    for bk in [b.strip() for b in B.split(";") if b.strip()]:
                        cb_map_by_co.setdefault((L,E,C), []).append(bk)
                if L and E and C and IO:
                    io_map_by_co.setdefault((L,E,C), [])
                    rec = {"Name": IO, "Mfg": (MFG or "")}
                    # Dedup by name within the CO bucket
                    if all(x["Name"] != IO for x in io_map_by_co[(L,E,C)]):
                        io_map_by_co[(L,E,C)].append(rec)

            # --- x coordinates ---
            next_x = LEFT_PAD
            led_x, le_x, bu_x, co_x, cb_x, io_x = {}, {}, {}, {}, {}, {}

            for L in ledgers_all:
                les = sorted(le_map.get(L, []))
                # Allocate children columns under each LE (BUs first, then COs so BUs sit left)
                for le in les:
                    buses = sorted(bu_map.get((L, le), []))
                    cos   = sorted(co_map.get((L, le), []))
                    if not buses and not cos:
                        # solitary LE
                        le_x[(L, le)] = next_x; next_x += X_STEP
                    else:
                        # Allocate BUs (left of COs)
                        for b in buses:
                            if b not in bu_x:
                                bu_x[b] = next_x; next_x += X_STEP
                        # Allocate COs
                        for c in cos:
                            if c not in co_x:
                                co_x[c] = next_x; next_x += X_STEP

                        # Center LE above its children
                        xs = []
                        xs += [bu_x[b] for b in buses]
                        xs += [co_x[c] for c in cos]
                        if xs:
                            le_x[(L, le)] = int(sum(xs)/len(xs))
                        else:
                            le_x[(L, le)] = next_x; next_x += X_STEP

                    # Under each CO, place Books to the LEFT cluster, IOs to the RIGHT cluster
                    for c in sorted(co_map.get((L, le), [])):
                        base = co_x[c]

                        # Books left of CO
                        books = cb_map_by_co.get((L, le, c), [])
                        books = sorted(dict.fromkeys(books))  # dedup keep order
                        for i, bk in enumerate(books, start=1):
                            cb_x[(L, le, c, bk)] = base - i*X_STEP

                        # IOs right of CO (use IO_STEP for min spacing)
                        ios = io_map_by_co.get((L, le, c), [])
                        for j, io in enumerate(sorted(ios, key=lambda k: k["Name"])):
                            io_x[(L, le, c, io["Name"])] = base + (j+1)*IO_STEP

                if les:
                    xs = [le_x[(L, le)] for le in les]
                    led_x[L] = int(sum(xs)/len(xs))
                else:
                    led_x[L] = next_x; next_x += X_STEP

                next_x += PAD_GROUP

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
                        # Books (left)
                        for bk in sorted(set(cb_map_by_co.get((L, le, c), []))):
                            id_map[("CB", L, le, c, bk)] = add_vertex(bk, S_CB, cb_x[(L, le, c, bk)], Y_CB)
                        # IOs (right)
                        for io in sorted(io_map_by_co.get((L, le, c), []), key=lambda k: k["Name"]):
                            label = f"🏭 {io['Name']}" if str(io["Mfg"]).lower() == "yes" else io["Name"]
                            style = S_IO_PLT if str(io["Mfg"]).lower() == "yes" else S_IO
                            id_map[("IO", L, le, c, io["Name"])] = add_vertex(label, style, io_x[(L, le, c, io["Name"])], Y_IO)

            # --- edges (child→parent, center-to-center) ---
            for L in ledgers_all:
                for le in sorted(le_map.get(L, [])):
                    if ("E", L, le) in id_map: add_edge(id_map[("E", L, le)], id_map[("L", L)])
                    for b in sorted(bu_map.get((L, le), [])):
                        if ("B", L, le, b) in id_map: add_edge(id_map[("B", L, le, b)], id_map[("E", L, le)])
                    for c in sorted(co_map.get((L, le), [])):
                        if ("C", L, le, c) in id_map:
                            add_edge(id_map[("C", L, le, c)], id_map[("E", L, le)])
                            for bk in sorted(set(cb_map_by_co.get((L, le, c), []))):
                                if ("CB", L, le, c, bk) in id_map:
                                    add_edge(id_map[("CB", L, le, c, bk)], id_map[("C", L, le, c)])
                            for io in io_map_by_co.get((L, le, c), []):
                                k = ("IO", L, le, c, io["Name"])
                                if k in id_map:
                                    add_edge(id_map[k], id_map[("C", L, le, c)])

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
                        "x": str(x+36), "y": str(y+offset-4), "width": "220", "height": "20", "as": "geometry"})

                swatch("Ledger", "#FFE6E6", 36)
                swatch("Legal Entity", "#FFE2C2", 62)
                swatch("Business Unit (left of LE)", "#FFF1B3", 88)
                swatch("Cost Org (right of LE)", "#E2F7E2", 114)
                swatch("Cost Book (left of CO)", "#7FBF7F", 140)
                swatch("Inventory Org (right of CO)", "#D6EFFF", 166, stroke="#2F71A8")
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
