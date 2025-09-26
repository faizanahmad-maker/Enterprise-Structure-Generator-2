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
# ===================== DRAW.IO DIAGRAM BLOCK =====================
if "df1" in locals() and isinstance(df1, pd.DataFrame) and not df1.empty:
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
        Y_CO       = 630   # cost orgs row, below BUs

        # --- styles ---
        S_LEDGER = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE6E6;strokeColor=#C86868;fontSize=12;"
        S_LE     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE2C2;strokeColor=#A66000;fontSize=12;"
        S_BU     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFF1B3;strokeColor=#B38F00;fontSize=12;"
        S_CO     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#E2F7E2;strokeColor=#3D8B3D;fontSize=12;"

        S_EDGE = (
            "endArrow=block;rounded=1;edgeStyle=orthogonalEdgeStyle;orthogonal=1;"
            "jettySize=auto;strokeColor=#666666;exitX=0.5;exitY=0;entryX=0.5;entryY=1;"
        )

        # --- normalize input ---
        df_bu = df_bu[["Ledger Name", "Legal Entity", "Business Unit"]].copy()
        df_co = df_co[["Ledger Name", "Legal Entity", "Cost Organization"]].copy()
        for df in (df_bu, df_co):
            for c in df.columns:
                df[c] = df[c].fillna("").map(str).str.strip()

        ledgers = sorted([x for x in set(df_bu["Ledger Name"]) | set(df_co["Ledger Name"]) if x])
        le_map = {}  # {ledger: set(LEs)}
        for _, r in pd.concat([df_bu, df_co]).iterrows():
            if r["Ledger Name"] and r["Legal Entity"]:
                le_map.setdefault(r["Ledger Name"], set()).add(r["Legal Entity"])

        bu_map = {}  # {(ledger, le): set(BUs)}
        for _, r in df_bu.iterrows():
            if r["Ledger Name"] and r["Legal Entity"] and r["Business Unit"]:
                bu_map.setdefault((r["Ledger Name"], r["Legal Entity"]), set()).add(r["Business Unit"])

        co_map = {}  # {(ledger, le): set(COs)}
        for _, r in df_co.iterrows():
            if r["Ledger Name"] and r["Legal Entity"] and r["Cost Organization"]:
                co_map.setdefault((r["Ledger Name"], r["Legal Entity"]), set()).add(r["Cost Organization"])

        # --- x-coordinates ---
        next_x = LEFT_PAD
        led_x, le_x, bu_x, co_x = {}, {}, {}, {}

        for L in ledgers:
            les = sorted(le_map.get(L, []))
            for le in les:
                buses = sorted(bu_map.get((L, le), []))
                cos   = sorted(co_map.get((L, le), []))
                all_children = buses + cos if (buses or cos) else [le]

                for child in all_children:
                    if child not in bu_x and child not in co_x:
                        if child in buses:
                            bu_x[child] = next_x
                        else:
                            co_x[child] = next_x
                        next_x += X_STEP
                if buses or cos:
                    xs = []
                    if buses: xs += [bu_x[b] for b in buses]
                    if cos:   xs += [co_x[c] for c in cos]
                    le_x[(L, le)] = int(sum(xs)/len(xs))
                else:
                    le_x[(L, le)] = next_x
                    next_x += X_STEP
            if les:
                xs = [le_x[(L, le)] for le in les]
                led_x[L] = int(sum(xs)/len(xs))
            else:
                led_x[L] = next_x
                next_x += X_STEP
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

        # --- vertices ---
        id_map = {}
        for L in ledgers:
            id_map[("L", L)] = add_vertex(L, S_LEDGER, led_x[L], Y_LEDGER)
            for le in sorted(le_map.get(L, [])):
                id_map[("E", L, le)] = add_vertex(le, S_LE, le_x[(L, le)], Y_LE)
                for b in sorted(bu_map.get((L, le), [])):
                    id_map[("B", L, le, b)] = add_vertex(b, S_BU, bu_x[b], Y_BU)
                for c in sorted(co_map.get((L, le), [])):
                    id_map[("C", L, le, c)] = add_vertex(c, S_CO, co_x[c], Y_CO)

        # --- edges ---
        for L in ledgers:
            for le in sorted(le_map.get(L, [])):
                # LE → Ledger
                if ("E", L, le) in id_map and ("L", L) in id_map:
                    add_edge(id_map[("E", L, le)], id_map[("L", L)])
                # BU → LE
                for b in sorted(bu_map.get((L, le), [])):
                    if ("B", L, le, b) in id_map:
                        add_edge(id_map[("B", L, le, b)], id_map[("E", L, le)])
                # CO → LE
                for c in sorted(co_map.get((L, le), [])):
                    if ("C", L, le, c) in id_map:
                        add_edge(id_map[("C", L, le, c)], id_map[("E", L, le)])

        # Legend
        def add_legend(x=20, y=20):
            def swatch(lbl, color, ty, offset):
                box = ET.SubElement(root, "mxCell", attrib={
                    "id": uuid.uuid4().hex[:8], "value": "",
                    "style": f"rounded=1;fillColor={color};strokeColor=#666666;",
                    "vertex": "1", "parent": "1"})
                ET.SubElement(box, "mxGeometry", attrib={
                    "x": str(x+12), "y": str(y+offset), "width": "18", "height": "12", "as": "geometry"})
                txt = ET.SubElement(root, "mxCell", attrib={
                    "id": uuid.uuid4().hex[:8], "value": lbl,
                    "style": "text;align=left;verticalAlign=middle;fontSize=12;",
                    "vertex": "1", "parent": "1"})
                ET.SubElement(txt, "mxGeometry", attrib={
                    "x": str(x+36), "y": str(y+offset-4), "width": "130", "height": "20", "as": "geometry"})
            swatch("Ledger", "#FFE6E6", "L", 36)
            swatch("Legal Entity", "#FFE2C2", "E", 62)
            swatch("Business Unit", "#FFF1B3", "B", 88)
            swatch("Cost Org", "#E2F7E2", "C", 114)

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
