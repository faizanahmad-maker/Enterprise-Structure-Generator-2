import io, zipfile, base64, zlib, uuid
import pandas as pd
import streamlit as st
import xml.etree.ElementTree as ET

st.set_page_config(page_title="Enterprise Structure Generator 2", page_icon="🧭", layout="wide")
st.title("Enterprise Structure Generator 2 — Core + Cost Org (separate vertical, arrows upward)")

st.markdown("""
**Uploads (any order):**
- **Core**: `GL_PRIMARY_LEDGER.csv`, `XLE_ENTITY_PROFILE.csv`,
  `ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv`, `ORA_GL_JOURNAL_CONFIG_DETAIL.csv`, `FUN_BUSINESS_UNIT.csv`
- **Costing**: `CST_COST_ORGANIZATION.csv` (Name, LegalEntityIdentifier, OrgInformation2),
  `CST_COST_ORG_BOOK.csv` (**Name = Ledger**, `ORA_CST_ACCT_COST_ORG.CostOrgCode` = CostOrgKey)
""")

uploads = st.file_uploader("Drop Oracle export ZIPs", type="zip", accept_multiple_files=True)

# ----------------- helpers -----------------
def read_csv_from_zip(zf, name):
    if name not in zf.namelist():
        return None
    with zf.open(name) as fh:
        return pd.read_csv(fh, dtype=str)

def _drawio_url_from_xml(xml: str) -> str:
    raw = zlib.compress(xml.encode("utf-8"), level=9)[2:-4]  # raw DEFLATE
    b64 = base64.b64encode(raw).decode("ascii")
    return f"https://app.diagrams.net/?title=EnterpriseStructure.drawio#R{b64}"

# ----------------- ingestion -----------------
if not uploads:
    st.info("Upload your ZIPs to generate outputs.")
    st.stop()

# Core collectors
ledger_names = set()
legal_entity_names = set()
ledger_to_idents = {}
ident_to_le_name = {}
bu_rows = []

# Costing collectors
costorg_rows = []      # from CST_COST_ORGANIZATION.csv :: Name, LegalEntityIdentifier, OrgInformation2
code_to_ledger = {}    # from CST_COST_ORG_BOOK.csv    :: ORA_CST_ACCT_COST_ORG.CostOrgCode -> Name (Ledger)

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
    if df is not None and "Name" in df.columns:
        legal_entity_names |= set(df["Name"].dropna().map(str).str.strip())

    # Ledger ↔ LE identifier
    df = read_csv_from_zip(z, "ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv")
    if df is not None and {"GL_LEDGER.Name","LegalEntityIdentifier"} <= set(df.columns):
        for _, r in df[["GL_LEDGER.Name","LegalEntityIdentifier"]].dropna().iterrows():
            ledger_to_idents.setdefault(str(r["GL_LEDGER.Name"]).strip(), set()).add(str(r["LegalEntityIdentifier"]).strip())

    # Identifier ↔ LE name
    df = read_csv_from_zip(z, "ORA_GL_JOURNAL_CONFIG_DETAIL.csv")
    if df is not None and {"LegalEntityIdentifier","ObjectName"} <= set(df.columns):
        for _, r in df[["LegalEntityIdentifier","ObjectName"]].dropna().iterrows():
            ident_to_le_name[str(r["LegalEntityIdentifier"]).strip()] = str(r["ObjectName"]).strip()

    # Business Units
    df = read_csv_from_zip(z, "FUN_BUSINESS_UNIT.csv")
    if df is not None and {"Name","PrimaryLedgerName","LegalEntityName"} <= set(df.columns):
        t = df[["Name","PrimaryLedgerName","LegalEntityName"]].fillna("").astype(str).applymap(lambda x: x.strip())
        bu_rows += t.to_dict(orient="records")

    # Cost Org master
    df = read_csv_from_zip(z, "CST_COST_ORGANIZATION.csv")
    if df is not None and {"Name","LegalEntityIdentifier","OrgInformation2"} <= set(df.columns):
        t = df[["Name","LegalEntityIdentifier","OrgInformation2"]].fillna("").astype(str).applymap(lambda x: x.strip())
        t = t.rename(columns={"Name":"CostOrgName","LegalEntityIdentifier":"LE_Ident","OrgInformation2":"CostOrgCode"})
        costorg_rows += t.to_dict(orient="records")

    # Cost Org → Ledger (authoritative)
    df = read_csv_from_zip(z, "CST_COST_ORG_BOOK.csv")
    if df is not None and {"ORA_CST_ACCT_COST_ORG.CostOrgCode","Name"} <= set(df.columns):
        for _, r in df[["ORA_CST_ACCT_COST_ORG.CostOrgCode","Name"]].dropna().iterrows():
            code = str(r["ORA_CST_ACCT_COST_ORG.CostOrgCode"]).strip()
            ledger = str(r["Name"]).strip()  # in your export, 'Name' is the GL Ledger
            if code and ledger:
                code_to_ledger.setdefault(code, set()).add(ledger)

# ----------------- Sheet 1 (Ledger–LE–BU) with **no duplicate (L,LE) if BU exists** -----------------
# Build Ledger → LE (from identifiers)
ledger_to_le_names = {}
for led, idents in ledger_to_idents.items():
    for ident in idents:
        le_name = ident_to_le_name.get(ident, "").strip()
        if le_name:
            ledger_to_le_names.setdefault(led, set()).add(le_name)

# Collect BUs per (Ledger, LE)
bu_map = {}  # (L,E) -> set(BU)
for r in bu_rows:
    L = r["PrimaryLedgerName"].strip()
    E = r["LegalEntityName"].strip()
    B = r["Name"].strip()
    if L and E and B:
        bu_map.setdefault((L,E), set()).add(B)

rows_core = []
# For each ledger & LE we know about, either output all BUs OR a single L-LE without BU
for L, le_set in ledger_to_le_names.items():
    for E in sorted(le_set):
        buses = sorted(list(bu_map.get((L,E), set())))
        if buses:
            for B in buses:
                rows_core.append({"Ledger Name": L, "Legal Entity": E, "Business Unit": B})
        else:
            rows_core.append({"Ledger Name": L, "Legal Entity": E, "Business Unit": ""})

# Orphan Ledgers (no LE at all)
for L in sorted(ledger_names - set(ledger_to_le_names.keys())):
    rows_core.append({"Ledger Name": L, "Legal Entity": "", "Business Unit": ""})

# Orphan LEs (never appear under any ledger in the core files)
present_les = {r["Legal Entity"] for r in rows_core if r["Legal Entity"]}
for E in sorted(legal_entity_names - present_les):
    rows_core.append({"Ledger Name": "", "Legal Entity": E, "Business Unit": ""})

df_core = pd.DataFrame(rows_core).drop_duplicates().reset_index(drop=True)
# Order: non-empty ledger first, then LE, then BU
df_core["__LedgerEmpty"] = (df_core["Ledger Name"] == "").astype(int)
df_core = df_core.sort_values(
    ["__LedgerEmpty","Ledger Name","Legal Entity","Business Unit"],
    ascending=[True,True,True,True]
).drop(columns="__LedgerEmpty").reset_index(drop=True)

st.success(f"Sheet 1 (Core): {len(df_core)} rows (no duplicate L–LE when BUs exist)")
st.dataframe(df_core, use_container_width=True, height=360)

# ----------------- Sheet 2 (Ledger–LE–Cost Org; row-per-cost-org, include orphans) -----------------
rows_cost = []
for r in costorg_rows:
    cname = r["CostOrgName"] if r["CostOrgName"] else r["CostOrgCode"]
    le_name = ident_to_le_name.get(r["LE_Ident"], "").strip()
    ledgers = sorted(code_to_ledger.get(r["CostOrgCode"], []))
    if ledgers:
        for L in ledgers:
            rows_cost.append({"Ledger Name": L, "Legal Entity": le_name, "Business Unit": "", "Cost Organization": cname})
    else:
        rows_cost.append({"Ledger Name": "", "Legal Entity": le_name, "Business Unit": "", "Cost Organization": cname})

df_cost = pd.DataFrame(rows_cost, columns=["Ledger Name","Legal Entity","Business Unit","Cost Organization"]).drop_duplicates().reset_index(drop=True)
st.success(f"Sheet 2 (Cost Orgs): {len(df_cost)} rows")
st.dataframe(df_cost, use_container_width=True, height=320)

# ----------------- Excel export (openpyxl) -----------------
excel_buf = io.BytesIO()
with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
    df_core.to_excel(writer, index=False, sheet_name="Core_Ledger_LE_BU")
    df_cost.to_excel(writer, index=False, sheet_name="Ledger_LE_CostOrg")

st.download_button(
    "⬇️ Download Excel (EnterpriseStructure_v2.xlsx)",
    data=excel_buf.getvalue(),
    file_name="EnterpriseStructure_v2.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ================== DRAW.IO (Ledger top, LE below, BU left lane, Cost Org right lane; edges upward) ==================
def make_drawio_xml(df_core: pd.DataFrame, df_cost: pd.DataFrame) -> str:
    # --- layout constants ---
    LEFT_PAD   = 240
    RIGHT_PAD  = 200
    X_STEP     = 220
    PAD_GROUP  = 120
    W, H       = 180, 48

    Y_LEDGER   = 120
    Y_LE       = 280
    Y_BU       = 440
    Y_COST     = 520  # lower than BU (different vertical)

    # --- styles ---
    S_LEDGER = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE6E6;strokeColor=#C86868;fontSize=12;"
    S_LE     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE2C2;strokeColor=#A66000;fontSize=12;"
    S_BU     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFF1B3;strokeColor=#B38F00;fontSize=12;"
    S_COST   = "rounded=1;whiteSpace=wrap;html=1;fillColor=#DDEBFF;strokeColor=#3B82F6;fontSize=12;"

    # upward edges: top center (child) -> bottom center (parent)
    EDGE_UP_GRAY = (
        "endArrow=block;rounded=1;"
        "edgeStyle=orthogonalEdgeStyle;orthogonal=1;jettySize=auto;"
        "strokeColor=#666666;"
        "exitX=0.5;exitY=0;"   # from top center of child
        "entryX=0.5;entryY=1;" # to bottom center of parent
    )
    EDGE_UP_BLUE = EDGE_UP_GRAY.replace("#666666", "#3B82F6")

    # --- normalize inputs ---
    core = df_core[["Ledger Name","Legal Entity","Business Unit"]].copy().fillna("").astype(str).applymap(lambda x: x.strip())
    cost = df_cost[["Ledger Name","Legal Entity","Cost Organization"]].copy().fillna("").astype(str).applymap(lambda x: x.strip())

    # Build maps
    ledgers = sorted(list(set([x for x in core["Ledger Name"].unique() if x]) | set([x for x in cost["Ledger Name"].unique() if x])))

    led_to_les = {}
    for _, r in core.iterrows():
        L,E = r["Ledger Name"], r["Legal Entity"]
        if L and E:
            led_to_les.setdefault(L, set()).add(E)
    for _, r in cost.iterrows():
        L,E = r["Ledger Name"], r["Legal Entity"]
        if L and E:
            led_to_les.setdefault(L, set()).add(E)

    le_to_bu = {}
    for _, r in core.iterrows():
        L,E,B = r["Ledger Name"], r["Legal Entity"], r["Business Unit"]
        if L and E and B:
            le_to_bu.setdefault((L,E), set()).add(B)

    le_to_cost = {}
    for _, r in cost.iterrows():
        L,E,C = r["Ledger Name"], r["Legal Entity"], r["Cost Organization"]
        if L and E and C:
            le_to_cost.setdefault((L,E), set()).add(C)

    # --- positioning ---
    next_x = LEFT_PAD
    led_x, le_x, bu_x, cost_x = {}, {}, {}, {}

    for L in ledgers:
        les = sorted(list(led_to_les.get(L, set())))
        if not les:
            # ensure ledger shows even if no LE
            led_x[L] = next_x
            next_x += X_STEP + PAD_GROUP
            continue

        # 1) allocate BU left lane slots
        lane_left_positions = []
        for E in les:
            buses = sorted(list(le_to_bu.get((L,E), set())))
            if buses:
                for b in buses:
                    if b not in bu_x:
                        bu_x[b] = next_x
                        next_x += X_STEP
                lane_left_positions += [bu_x[b] for b in buses]
            else:
                # reserve left slot under this LE for symmetry
                lane_left_positions.append(next_x)
                next_x += X_STEP

        # gap between lanes
        next_x += 120

        # 2) allocate Cost Org right lane slots (lower row, own vertical)
        lane_right_positions = []
        for E in les:
            costs = sorted(list(le_to_cost.get((L,E), set())))
            if costs:
                for c in costs:
                    if c not in cost_x:
                        cost_x[c] = next_x
                        next_x += X_STEP
                lane_right_positions += [cost_x[c] for c in costs]
            else:
                lane_right_positions.append(next_x)
                next_x += X_STEP

        # center each LE over midpoint of its local left/right children
        # (if only one side has children, center over that side)
        # compute per-LE centers
        # collect per-LE xs
        tmp_centers = {}
        idx_left = 0
        idx_right = 0
        for E in les:
            buses = sorted(list(le_to_bu.get((L,E), set())))
            costs = sorted(list(le_to_cost.get((L,E), set())))
            xs = []
            if buses:
                xs += [bu_x[b] for b in buses]
            else:
                xs += [lane_left_positions[idx_left]]
            idx_left += (len(buses) or 1)
            if costs:
                xs += [cost_x[c] for c in costs]
            else:
                xs += [lane_right_positions[idx_right]]
            idx_right += (len(costs) or 1)
            tmp_centers[E] = int(sum(xs)/len(xs))

        # place LEs and Ledger
        for E in les:
            le_x[(L,E)] = tmp_centers[E]
        led_x[L] = int(sum(le_x[(L,E)] for E in les)/len(les))

        # gap before next ledger block
        next_x += PAD_GROUP

    # right pad space for any orphans
    next_x += RIGHT_PAD

    # --- XML skeleton ---
    mxfile  = ET.Element("mxfile", attrib={"host": "app.diagrams.net"})
    diagram = ET.SubElement(mxfile, "diagram", attrib={"id": str(uuid.uuid4()), "name": "Enterprise Structure"})
    model   = ET.SubElement(diagram, "mxGraphModel", attrib={
        "dx": "1284", "dy": "682", "grid": "1", "gridSize": "10",
        "page": "1", "pageWidth": "2400", "pageHeight": "1400",
        "background": "#ffffff"
    })
    root    = ET.SubElement(model, "root")
    ET.SubElement(root, "mxCell", attrib={"id": "0"})
    ET.SubElement(root, "mxCell", attrib={"id": "1", "parent": "0"})

    def add_vertex(label, style, x, y, w=W, h=H):
        vid = uuid.uuid4().hex[:8]
        c = ET.SubElement(root, "mxCell", attrib={"id": vid, "value": label, "style": style, "vertex": "1", "parent": "1"})
        ET.SubElement(c, "mxGeometry", attrib={"x": str(int(x)), "y": str(int(y)), "width": str(w), "height": str(h), "as": "geometry"})
        return vid

    def add_edge(src, tgt, style):
        eid = uuid.uuid4().hex[:8]
        c = ET.SubElement(root, "mxCell", attrib={
            "id": eid, "value": "", "style": style, "edge": "1", "parent": "1",
            "source": src, "target": tgt
        })
        ET.SubElement(c, "mxGeometry", attrib={"relative": "1", "as": "geometry"})

    id_map = {}

    # Ledger nodes
    for L in ledgers:
        id_map[("L", L)] = add_vertex(L, S_LEDGER, led_x[L], Y_LEDGER)

    # LE nodes
    for L in ledgers:
        for E in sorted(list(led_to_les.get(L, set()))):
            id_map[("E", L, E)] = add_vertex(E, S_LE, le_x[(L,E)], Y_LE)

    # BU nodes (left lane)
    for (L,E), buses in le_to_bu.items():
        for b in sorted(list(buses)):
            id_map[("B", L, E, b)] = add_vertex(f"{b} BU", S_BU, bu_x[b], Y_BU)

    # Cost Org nodes (right lane; lower)
    cost_nodes_drawn = set()
    for (L,E), costs in le_to_cost.items():
        for c in sorted(list(costs)):
            if c not in cost_nodes_drawn:
                id_map[("C", L, E, c)] = add_vertex(c, S_COST, cost_x[c], Y_COST)
                cost_nodes_drawn.add(c)

    # Edges upward with top/bottom anchoring: Cost→BU (blue), BU→LE (gray), LE→Ledger (gray)
    for (L,E), costs in le_to_cost.items():
        for c in sorted(list(costs)):
            # connect Cost Org -> nearest BU under same LE if one exists, else directly to LE
            buses = sorted(list(le_to_bu.get((L,E), set())))
            src = id_map.get(("C", L, E, c))
            if buses:
                for b in buses:
                    tgt = id_map.get(("B", L, E, b))
                    if src and tgt:
                        add_edge(src, tgt, EDGE_UP_BLUE)
            else:
                tgt = id_map.get(("E", L, E))
                if src and tgt:
                    add_edge(src, tgt, EDGE_UP_BLUE)

    for (L,E), buses in le_to_bu.items():
        for b in sorted(list(buses)):
            src = id_map.get(("B", L, E, b))
            tgt = id_map.get(("E", L, E))
            if src and tgt:
                add_edge(src, tgt, EDGE_UP_GRAY)

    for L, les in led_to_les.items():
        for E in les:
            src = id_map.get(("E", L, E))
            tgt = id_map.get(("L", L))
            if src and tgt:
                add_edge(src, tgt, EDGE_UP_GRAY)

    # Legend
    def add_legend(x=20, y=20):
        def swatch(lbl, color, gy):
            add_vertex("", f"rounded=1;fillColor={color};strokeColor=#666666;", x+12, y+gy, 18, 12)
            add_vertex(lbl, "text;align=left;verticalAlign=middle;fontSize=12;", x+36, y+gy-4, 280, 20)
        add_vertex("", "rounded=1;fillColor=#FFFFFF;strokeColor=#CBD5E1;", x, y, 320, 160)
        swatch("Ledger", "#FFE6E6", 36)
        swatch("Legal Entity", "#FFE2C2", 62)
        swatch("Business Unit (left lane)", "#FFF1B3", 88)
        swatch("Cost Organization (right lane, lower)", "#DDEBFF", 114)

    add_legend()
    return ET.tostring(mxfile, encoding="utf-8", method="xml").decode("utf-8")

# Build & offer the diagram
xml = make_drawio_xml(df_core, df_cost)
st.download_button(
    "⬇️ Download diagram (.drawio)",
    data=xml.encode("utf-8"),
    file_name="EnterpriseStructure.drawio",
    mime="application/xml"
)
st.markdown(f"[🔗 Open in draw.io (preview)]({_drawio_url_from_xml(xml)})")
st.caption("Arrows go upward: Cost Org → BU → Legal Entity → Ledger. BU lane left; Cost Org lane right & lower. No backdrop boxes.")
