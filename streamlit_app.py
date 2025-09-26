import io, zipfile
import base64, zlib, uuid
import pandas as pd
import streamlit as st
import xml.etree.ElementTree as ET

st.set_page_config(page_title="Enterprise Structure Generator 2", page_icon="🧭", layout="wide")
st.title("Enterprise Structure Generator 2 — Core + Cost Org lane (siloed diagram)")

st.markdown("""
**Uploads (any order):**
- **Core**: `GL_PRIMARY_LEDGER.csv`, `XLE_ENTITY_PROFILE.csv`,
  `ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv`, `ORA_GL_JOURNAL_CONFIG_DETAIL.csv`, `FUN_BUSINESS_UNIT.csv`
- **Costing**: `CST_COST_ORGANIZATION.csv` (Name, LegalEntityIdentifier, OrgInformation2),
  `CST_COST_ORG_BOOK.csv` (**Name = Ledger**, `ORA_CST_ACCT_COST_ORG.CostOrgCode` = CostOrgKey)
""")

uploads = st.file_uploader("Drop Oracle export ZIPs", type="zip", accept_multiple_files=True)

# ---------- helpers ----------
def read_csv_from_zip(zf, name):
    if name not in zf.namelist():
        return None
    with zf.open(name) as fh:
        return pd.read_csv(fh, dtype=str)

def _drawio_url_from_xml(xml: str) -> str:
    raw = zlib.compress(xml.encode("utf-8"), level=9)[2:-4]
    b64 = base64.b64encode(raw).decode("ascii")
    return f"https://app.diagrams.net/?title=EnterpriseStructure.drawio#R{b64}"

# ---------- ingestion ----------
if not uploads:
    st.info("Upload your ZIPs to generate outputs.")
    st.stop()

# Core collectors
ledger_names = set()
legal_entity_names = set()
ledger_to_idents = {}
ident_to_le_name = {}
bu_rows = []

# Cost collectors
costorg_rows = []    # {CostOrgName, LE_Ident, CostOrgCode}
code_to_ledger = {}  # CostOrgCode -> set(LedgerName)

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

    # Ledger ↔ LE Identifier
    df = read_csv_from_zip(z, "ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv")
    if df is not None and {"GL_LEDGER.Name", "LegalEntityIdentifier"} <= set(df.columns):
        for _, r in df[["GL_LEDGER.Name", "LegalEntityIdentifier"]].dropna().iterrows():
            ledger_to_idents.setdefault(str(r["GL_LEDGER.Name"]).strip(), set()).add(str(r["LegalEntityIdentifier"]).strip())

    # Identifier ↔ LE Name
    df = read_csv_from_zip(z, "ORA_GL_JOURNAL_CONFIG_DETAIL.csv")
    if df is not None and {"LegalEntityIdentifier", "ObjectName"} <= set(df.columns):
        for _, r in df[["LegalEntityIdentifier", "ObjectName"]].dropna().iterrows():
            ident_to_le_name[str(r["LegalEntityIdentifier"]).strip()] = str(r["ObjectName"]).strip()

    # BUs
    df = read_csv_from_zip(z, "FUN_BUSINESS_UNIT.csv")
    if df is not None and {"Name", "PrimaryLedgerName", "LegalEntityName"} <= set(df.columns):
        t = df[["Name", "PrimaryLedgerName", "LegalEntityName"]].fillna("").astype(str).applymap(lambda x: x.strip())
        bu_rows += t.to_dict(orient="records")

    # Cost Org master
    df = read_csv_from_zip(z, "CST_COST_ORGANIZATION.csv")
    if df is not None and {"Name", "LegalEntityIdentifier", "OrgInformation2"} <= set(df.columns):
        t = df[["Name", "LegalEntityIdentifier", "OrgInformation2"]].fillna("").astype(str).applymap(lambda x: x.strip())
        t = t.rename(columns={"Name": "CostOrgName", "LegalEntityIdentifier":"LE_Ident", "OrgInformation2":"CostOrgCode"})
        costorg_rows += t.to_dict(orient="records")

    # Cost Org → Ledger (authoritative)
    df = read_csv_from_zip(z, "CST_COST_ORG_BOOK.csv")
    if df is not None and {"ORA_CST_ACCT_COST_ORG.CostOrgCode", "Name"} <= set(df.columns):
        for _, r in df[["ORA_CST_ACCT_COST_ORG.CostOrgCode", "Name"]].dropna().iterrows():
            code = str(r["ORA_CST_ACCT_COST_ORG.CostOrgCode"]).strip()
            ledg = str(r["Name"]).strip()
            if code and ledg:
                code_to_ledger.setdefault(code, set()).add(ledg)

# ---------- Sheet 1 (ordered like v1) ----------
ledger_to_le_names = {}
for led, idents in ledger_to_idents.items():
    for ident in idents:
        le_name = ident_to_le_name.get(ident, "").strip()
        if le_name:
            ledger_to_le_names.setdefault(led, set()).add(le_name)

rows_core = []
seen_triples = set()

# BU rows
for r in bu_rows:
    led, le, bu = r["PrimaryLedgerName"], r["LegalEntityName"], r["Name"]
    rows_core.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": bu})
    seen_triples.add((led, le, bu))

# Ledger–LE pairs lacking BU
for led, les in ledger_to_le_names.items():
    for le in les:
        if (led, le, "") not in seen_triples:
            rows_core.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": ""})

# Orphan ledgers
for led in sorted(ledger_names - set(ledger_to_le_names.keys())):
    rows_core.append({"Ledger Name": led, "Legal Entity": "", "Business Unit": ""})

# Orphan LEs
present_les = {r["Legal Entity"] for r in rows_core if r["Legal Entity"]}
for le in sorted(legal_entity_names - present_les):
    rows_core.append({"Ledger Name": "", "Legal Entity": le, "Business Unit": ""})

df_core = pd.DataFrame(rows_core).drop_duplicates().reset_index(drop=True)
df_core["__LedgerEmpty"] = (df_core["Ledger Name"] == "").astype(int)
df_core = df_core.sort_values(
    ["__LedgerEmpty", "Ledger Name", "Legal Entity", "Business Unit"],
    ascending=[True, True, True, True]
).drop(columns="__LedgerEmpty").reset_index(drop=True)

st.success(f"Sheet 1 (Core): {len(df_core)} rows")
st.dataframe(df_core, use_container_width=True, height=360)

# ---------- Sheet 2 (Ledger–LE–Cost Org) ----------
out_rows = []
for r in costorg_rows:
    cname = r["CostOrgName"] if r["CostOrgName"] else r["CostOrgCode"]
    le_name = ident_to_le_name.get(r["LE_Ident"], "").strip()
    ledgers = sorted(code_to_ledger.get(r["CostOrgCode"], []))
    if ledgers:
        for L in ledgers:
            out_rows.append({"Ledger Name": L, "Legal Entity": le_name, "Business Unit": "", "Cost Organization": cname})
    else:
        out_rows.append({"Ledger Name": "", "Legal Entity": le_name, "Business Unit": "", "Cost Organization": cname})

df_cost = pd.DataFrame(out_rows, columns=["Ledger Name","Legal Entity","Business Unit","Cost Organization"]).drop_duplicates().reset_index(drop=True)
st.success(f"Sheet 2 (Cost Orgs): {len(df_cost)} rows (row-per-cost-org; orphans included)")
st.dataframe(df_cost, use_container_width=True, height=320)

# ---------- Excel export (openpyxl) ----------
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

# ======================= SILOED DRAW.IO (Ledger boxes; BU left; Cost Org right/lower; arrows down) =======================
def make_siloed_drawio(df_core: pd.DataFrame, df_cost: pd.DataFrame) -> str:
    # Layout constants
    LEFT_PAD, RIGHT_PAD = 80, 80
    SILO_PAD = 120        # gap between ledger silos
    W, H = 180, 48
    X_STEP = 220
    Y_LEDGER, Y_LE, Y_BU, Y_COST = 120, 260, 400, 470  # Cost Orgs lower than BU

    # Styles
    S_LEDGER = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE6E6;strokeColor=#C86868;fontSize=12;"
    S_LE     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE2C2;strokeColor=#A66000;fontSize=12;"
    S_BU     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFF1B3;strokeColor=#B38F00;fontSize=12;"
    S_COST   = "rounded=1;whiteSpace=wrap;html=1;fillColor=#DDEBFF;strokeColor=#3B82F6;fontSize=12;"
    S_EDGE   = "endArrow=block;rounded=1;edgeStyle=orthogonalEdgeStyle;orthogonal=1;jettySize=auto;strokeColor=#666666;"
    S_EDGE_COST = "endArrow=block;rounded=1;edgeStyle=orthogonalEdgeStyle;orthogonal=1;jettySize=auto;strokeColor=#3B82F6;"
    S_SILO   = "rounded=1;whiteSpace=wrap;fillColor=#FFFFFF;strokeColor=#E5E7EB;dashed=1;"  # container per ledger

    # Normalize inputs
    core = df_core[["Ledger Name","Legal Entity","Business Unit"]].fillna("").astype(str).applymap(lambda x: x.strip())
    cost = df_cost[["Ledger Name","Legal Entity","Cost Organization"]].fillna("").astype(str).applymap(lambda x: x.strip())

    # Build per-ledger structures (include ledgers that appear only in cost)
    ledgers = sorted(list(set([x for x in core["Ledger Name"].unique() if x]) | set([x for x in cost["Ledger Name"].unique() if x])))

    # Map L->LEs
    led_to_les = {}
    for _, r in core.iterrows():
        if r["Ledger Name"] and r["Legal Entity"]:
            led_to_les.setdefault(r["Ledger Name"], set()).add(r["Legal Entity"])
    for _, r in cost.iterrows():
        if r["Ledger Name"] and r["Legal Entity"]:
            led_to_les.setdefault(r["Ledger Name"], set()).add(r["Legal Entity"])

    # BU map and CostOrg map (scoped by L,E)
    le_to_bu = {}
    for _, r in core.iterrows():
        L, E, B = r["Ledger Name"], r["Legal Entity"], r["Business Unit"]
        if L and E and B:
            le_to_bu.setdefault((L,E), set()).add(B)

    le_to_cost = {}
    for _, r in cost.iterrows():
        L, E, C = r["Ledger Name"], r["Legal Entity"], r["Cost Organization"]
        if L and E and C:
            le_to_cost.setdefault((L,E), set()).add(C)

    # XML skeleton
    mxfile  = ET.Element("mxfile", attrib={"host": "app.diagrams.net"})
    diagram = ET.SubElement(mxfile, "diagram", attrib={"id": str(uuid.uuid4()), "name": "Enterprise Structure"})
    model   = ET.SubElement(diagram, "mxGraphModel", attrib={
        "dx": "1284", "dy": "682", "grid": "1", "gridSize": "10",
        "page": "1", "pageWidth": "2200", "pageHeight": "1200",
        "background": "#ffffff"
    })
    root    = ET.SubElement(model, "root")
    ET.SubElement(root, "mxCell", attrib={"id": "0"})
    ET.SubElement(root, "mxCell", attrib={"id": "1", "parent": "0"})

    def add_vertex(label, style, x, y, w=W, h=H, parent="1"):
        vid = uuid.uuid4().hex[:8]
        c = ET.SubElement(root, "mxCell", attrib={"id": vid, "value": label, "style": style, "vertex": "1", "parent": parent})
        ET.SubElement(c, "mxGeometry", attrib={"x": str(int(x)), "y": str(int(y)), "width": str(w), "height": str(h), "as": "geometry"})
        return vid

    def add_edge(src, tgt, style=S_EDGE, parent="1"):
        eid = uuid.uuid4().hex[:8]
        c = ET.SubElement(root, "mxCell", attrib={"id": eid, "value": "", "style": style, "edge": "1", "parent": parent, "source": src, "target": tgt})
        ET.SubElement(c, "mxGeometry", attrib={"relative": "1", "as": "geometry"})

    # Global cursor for silos
    cursor_x = LEFT_PAD

    for L in ledgers:
        LEs = sorted(list(led_to_les.get(L, set())))
        if not LEs:
            # create a placeholder LE so the ledger still shows
            LEs = [""]

        # compute how many slots we need left/right to size the silo
        max_left = 1
        max_right = 1
        for E in LEs:
            max_left  = max(max_left,  len(le_to_bu.get((L,E), set())) or 1)
            max_right = max(max_right, len(le_to_cost.get((L,E), set())) or 1)

        # Silo width = (left slots + gap + right slots) * X_STEP
        GAP = 2  # virtual columns of separation
        cols = max_left + GAP + max_right
        silo_w = max(cols * X_STEP, W + 2*X_STEP)
        silo_h = 520
        silo_id = add_vertex("", S_SILO, cursor_x, 80, silo_w, silo_h)  # container

        # Ledger at top-center of silo
        ledger_x = cursor_x + silo_w/2 - W/2
        v_ledger = add_vertex(L, S_LEDGER, ledger_x, Y_LEDGER-60, parent=silo_id)

        # Per-LE placement inside silo
        # We'll stack LEs horizontally centered (average of left/right lanes)
        # Determine per-LE centers equally spaced across silo width
        if len(LEs) == 1:
            centers = [cursor_x + silo_w/2 - W/2]
        else:
            span = silo_w - 2*X_STEP
            step = span / (len(LEs)-1)
            centers = [cursor_x + X_STEP + i*step - W/2 for i in range(len(LEs))]

        for E, le_center_x in zip(LEs, centers):
            v_le = add_vertex(E, S_LE, le_center_x, Y_LE, parent=silo_id)
            # Arrow: Ledger → LE
            add_edge(v_ledger, v_le, style=S_EDGE, parent=silo_id)

            # Left lane (BU) under this LE
            buses = sorted(list(le_to_bu.get((L,E), set())))
            if not buses:
                # reserve one slot to balance
                buses = []
            left_start = le_center_x - (max_left*X_STEP)/2
            cur_x = left_start
            for b in (buses or [""]):
                if b:
                    v_bu = add_vertex(f"{b} BU", S_BU, cur_x, Y_BU, parent=silo_id)
                    # Arrow: LE → BU
                    add_edge(v_le, v_bu, style=S_EDGE, parent=silo_id)
                cur_x += X_STEP

            # Right lane (Cost Orgs) under-and-right
            costs = sorted(list(le_to_cost.get((L,E), set())))
            right_start = le_center_x + (GAP*X_STEP)/2  # slight gap to the right
            cur_x = right_start
            if not costs:
                costs = []
            for c in (costs or [""]):
                if c:
                    v_cost = add_vertex(c, S_COST, cur_x, Y_COST, parent=silo_id)
                    # Arrow: LE → Cost Org (blue)
                    add_edge(v_le, v_cost, style=S_EDGE_COST, parent=silo_id)
                cur_x += X_STEP

        # advance to next silo
        cursor_x += silo_w + SILO_PAD

    # Legend
    def add_legend(x=20, y=20):
        def swatch(lbl, color, gy):
            add_vertex("", f"rounded=1;fillColor={color};strokeColor=#666666;", x+12, y+gy, 18, 12)
            add_vertex(lbl, "text;align=left;verticalAlign=middle;fontSize=12;", x+36, y+gy-4, 260, 20)
        add_vertex("", "rounded=1;fillColor=#FFFFFF;strokeColor=#CBD5E1;", x, y, 300, 160)
        swatch("Ledger", "#FFE6E6", 36)
        swatch("Legal Entity", "#FFE2C2", 62)
        swatch("Business Unit (left lane)", "#FFF1B3", 88)
        swatch("Cost Organization (right, lower)", "#DDEBFF", 114)

    add_legend()
    return ET.tostring(mxfile, encoding="utf-8", method="xml").decode("utf-8")

# Build diagram data: only render when there is something for either sheet
if not (df_core.empty and df_cost.empty):
    xml = make_siloed_drawio(df_core, df_cost)
    st.download_button(
        "⬇️ Download diagram (.drawio)",
        data=xml.encode("utf-8"),
        file_name="EnterpriseStructure.drawio",
        mime="application/xml"
    )
    st.markdown(f"[🔗 Open in draw.io (preview)]({_drawio_url_from_xml(xml)})")
    st.caption("Each ledger is a separate silo. Arrows flow down: Ledger → LE → BU / Cost Org. BU lane left; Cost Org lane right & lower.")
