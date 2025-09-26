import io, zipfile
import base64, zlib, uuid
import pandas as pd
import streamlit as st
import xml.etree.ElementTree as ET

st.set_page_config(page_title="Enterprise Structure Generator 2", page_icon="🧭", layout="wide")
st.title("Enterprise Structure Generator 2 — Core + Cost Org lane")

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
    raw = zlib.compress(xml.encode("utf-8"), level=9)[2:-4]
    b64 = base64.b64encode(raw).decode("ascii")
    return f"https://app.diagrams.net/?title=EnterpriseStructure.drawio#R{b64}"

# ----------------- app body -----------------
if not uploads:
    st.info("Upload your ZIPs to generate outputs.")
    st.stop()

# ---- collectors (core) ----
ledger_names = set()
legal_entity_names = set()
ledger_to_idents = {}
ident_to_le_name = {}
bu_rows = []

# ---- collectors (cost) ----
costorg_rows = []      # from CST_COST_ORGANIZATION.csv :: Name, LegalEntityIdentifier, OrgInformation2
code_to_ledger = {}    # from CST_COST_ORG_BOOK.csv :: ORA_CST_ACCT_COST_ORG.CostOrgCode -> Name (Ledger)

# ---- scan every zip ----
for up in uploads:
    try:
        z = zipfile.ZipFile(up)
    except Exception as e:
        st.error(f"Could not open `{up.name}` as ZIP: {e}")
        continue

    # Ledgers
    df = read_csv_from_zip(z, "GL_PRIMARY_LEDGER.csv")
    if df is not None:
        col = "ORA_GL_PRIMARY_LEDGER_CONFIG.Name"
        if col in df.columns:
            ledger_names |= set(df[col].dropna().map(str).str.strip())
        else:
            st.warning(f"`GL_PRIMARY_LEDGER.csv` missing `{col}`. Found: {list(df.columns)}")

    # Legal Entities
    df = read_csv_from_zip(z, "XLE_ENTITY_PROFILE.csv")
    if df is not None:
        if "Name" in df.columns:
            legal_entity_names |= set(df["Name"].dropna().map(str).str.strip())
        else:
            st.warning(f"`XLE_ENTITY_PROFILE.csv` missing `Name`. Found: {list(df.columns)}")

    # Ledger ↔ LE Identifier
    df = read_csv_from_zip(z, "ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv")
    if df is not None:
        need = ["GL_LEDGER.Name", "LegalEntityIdentifier"]
        if all(c in df.columns for c in need):
            for _, r in df[need].dropna().iterrows():
                led = str(r["GL_LEDGER.Name"]).strip()
                ident = str(r["LegalEntityIdentifier"]).strip()
                if led and ident:
                    ledger_to_idents.setdefault(led, set()).add(ident)
        else:
            st.warning(f"`ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv` missing required columns. Found: {list(df.columns)}")

    # Identifier ↔ LE Name
    df = read_csv_from_zip(z, "ORA_GL_JOURNAL_CONFIG_DETAIL.csv")
    if df is not None:
        need = ["LegalEntityIdentifier", "ObjectName"]
        if all(c in df.columns for c in need):
            for _, r in df[need].dropna().iterrows():
                ident = str(r["LegalEntityIdentifier"]).strip()
                name  = str(r["ObjectName"]).strip()
                if ident:
                    ident_to_le_name[ident] = name
        else:
            st.warning(f"`ORA_GL_JOURNAL_CONFIG_DETAIL.csv` missing required columns. Found: {list(df.columns)}")

    # Business Units
    df = read_csv_from_zip(z, "FUN_BUSINESS_UNIT.csv")
    if df is not None:
        need = ["Name", "PrimaryLedgerName", "LegalEntityName"]
        if all(c in df.columns for c in need):
            t = df[need].fillna("").astype(str).applymap(lambda x: x.strip())
            bu_rows += t.to_dict(orient="records")
        else:
            st.warning(f"`FUN_BUSINESS_UNIT.csv` missing required columns. Found: {list(df.columns)}")

    # Cost Org master
    df = read_csv_from_zip(z, "CST_COST_ORGANIZATION.csv")
    if df is not None:
        need = ["Name", "LegalEntityIdentifier", "OrgInformation2"]
        if all(c in df.columns for c in need):
            t = df[need].fillna("").astype(str).applymap(lambda x: x.strip())
            t = t.rename(columns={
                "Name": "CostOrgName",
                "LegalEntityIdentifier": "LE_Ident",
                "OrgInformation2": "CostOrgCode"
            })
            costorg_rows += t.to_dict(orient="records")
        else:
            st.warning(f"`CST_COST_ORGANIZATION.csv` missing required columns. Found: {list(df.columns)}")

    # Cost Org → Ledger mapping (**authoritative**)
    df = read_csv_from_zip(z, "CST_COST_ORG_BOOK.csv")
    if df is not None:
        code_col = "ORA_CST_ACCT_COST_ORG.CostOrgCode"
        ledger_col = "Name"  # in your export, 'Name' holds the Ledger
        if code_col in df.columns and ledger_col in df.columns:
            for _, r in df[[code_col, ledger_col]].dropna().iterrows():
                ccode  = str(r[code_col]).strip()
                ledger = str(r[ledger_col]).strip()
                if ccode and ledger:
                    code_to_ledger.setdefault(ccode, set()).add(ledger)
        else:
            st.warning("`CST_COST_ORG_BOOK.csv` present but missing `ORA_CST_ACCT_COST_ORG.CostOrgCode` or `Name` columns.")

# ---------- build Sheet 1 (Ledger–LE–BU) ----------
ledger_to_le_names = {}
for led, idents in ledger_to_idents.items():
    for ident in idents:
        le_name = ident_to_le_name.get(ident, "").strip()
        if le_name:
            ledger_to_le_names.setdefault(led, set()).add(le_name)

rows_core = []
seen = set()

# BU-driven rows
for r in bu_rows:
    led = r["PrimaryLedgerName"]
    le  = r["LegalEntityName"]
    bu  = r["Name"]
    rows_core.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": bu})
    seen.add((led, le, bu))

# Ledger–LE pairs without BU
for led, les in ledger_to_le_names.items():
    for le in les:
        if (led, le, "") not in seen:
            rows_core.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": ""})

df_core = pd.DataFrame(rows_core).drop_duplicates().reset_index(drop=True)
st.success(f"Sheet 1 (Core): {len(df_core)} rows")
st.dataframe(df_core, use_container_width=True, height=360)

# ---------- build Sheet 2 (Ledger–LE–CostOrg) ----------
out_rows = []
for r in costorg_rows:
    cname = r["CostOrgName"]
    ccode = r["CostOrgCode"]
    le_name = ident_to_le_name.get(r["LE_Ident"], "").strip()
    ledgers = sorted(code_to_ledger.get(ccode, []))
    if ledgers:
        for L in ledgers:
            out_rows.append({
                "Ledger Name": L,
                "Legal Entity": le_name,
                "Business Unit": "",
                "Cost Organization": cname if cname else ccode
            })
    else:
        # orphan (no ledger mapping) — keep row with blanks
        out_rows.append({
            "Ledger Name": "",
            "Legal Entity": le_name,
            "Business Unit": "",
            "Cost Organization": cname if cname else ccode
        })

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

# ===================== DRAW.IO DIAGRAM (BU left, Cost Org right, blue) =====================
# Build a combined view so we don't crash when a (Ledger, LE) exists only in cost data
def make_drawio_xml(df_core: pd.DataFrame, df_cost: pd.DataFrame) -> str:
    # layout
    LEFT_PAD, RIGHT_PAD = 260, 160
    W, H, X_STEP, PAD_GROUP = 180, 48, 230, 60
    Y_LEDGER, Y_LE, Y_BU, Y_COST = 170, 330, 490, 560
    BUS_Y = 250

    # styles
    S_LEDGER = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE6E6;strokeColor=#C86868;fontSize=12;"
    S_LE     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE2C2;strokeColor=#A66000;fontSize=12;"
    S_BU     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFF1B3;strokeColor=#B38F00;fontSize=12;"
    S_COST   = "rounded=1;whiteSpace=wrap;html=1;fillColor=#DDEBFF;strokeColor=#3B82F6;fontSize=12;"

    S_EDGE = "endArrow=block;rounded=1;edgeStyle=orthogonalEdgeStyle;orthogonal=1;jettySize=auto;strokeColor=#666666;exitX=0.5;exitY=0;entryX=0.5;entryY=1;"
    S_EDGE_LEDGER = "endArrow=block;rounded=1;edgeStyle=orthogonalEdgeStyle;orthogonal=1;jettySize=auto;strokeColor=#444444;exitX=0.5;exitY=0;entryX=0.5;entryY=1;"
    S_EDGE_COST = "endArrow=block;rounded=1;edgeStyle=orthogonalEdgeStyle;orthogonal=1;jettySize=auto;strokeColor=#3B82F6;exitX=0.5;exitY=0;entryX=0.5;entryY=1;"

    # normalize
    core = df_core[["Ledger Name","Legal Entity","Business Unit"]].copy()
    core = core.fillna("").astype(str).applymap(lambda x: x.strip())
    cost = df_cost[["Ledger Name","Legal Entity","Cost Organization"]].copy() if not df_cost.empty else pd.DataFrame(columns=["Ledger Name","Legal Entity","Cost Organization"])
    cost = cost.fillna("").astype(str).applymap(lambda x: x.strip())

    # ledger list should include both sources
    ledgers = sorted(list(set([x for x in core["Ledger Name"].unique() if x]) | set([x for x in cost["Ledger Name"].unique() if x])))

    # map ledger -> set of LEs (from both sources)
    led_to_les = {}
    for _, r in core.iterrows():
        L, E = r["Ledger Name"], r["Legal Entity"]
        if L and E:
            led_to_les.setdefault(L, set()).add(E)
    for _, r in cost.iterrows():
        L, E = r["Ledger Name"], r["Legal Entity"]
        if L and E:
            led_to_les.setdefault(L, set()).add(E)

    # children collections
    le_to_bus = {}
    for _, r in core.iterrows():
        L, E, B = r["Ledger Name"], r["Legal Entity"], r["Business Unit"]
        if L and E and B:
            le_to_bus.setdefault((L, E), set()).add(B)

    le_to_cost = {}
    orphan_cost = []
    for _, r in cost.iterrows():
        L, E, C = r["Ledger Name"], r["Legal Entity"], r["Cost Organization"]
        if C and L and E:
            le_to_cost.setdefault((L, E), set()).add(C)
        elif C:
            orphan_cost.append(C)

    # position maps
    led_x, le_x, bu_x, cost_x = {}, {}, {}, {}

    next_x_left = LEFT_PAD
    next_x_right = LEFT_PAD
    SEP = 180  # gap between lanes

    for L in ledgers:
        les = sorted(list(led_to_les.get(L, set())))
        lane_left_xs, lane_right_xs = [], []

        # BU lane (left)
        for E in les:
            buses = sorted(list(le_to_bus.get((L, E), set())))
            if buses:
                for b in buses:
                    if b not in bu_x:
                        bu_x[b] = next_x_left
                        next_x_left += X_STEP
                lane_left_xs += [bu_x[b] for b in buses]
            else:
                lane_left_xs.append(next_x_left)
                next_x_left += X_STEP

        # Cost lane (right) starts after left lane + SEP for this ledger cluster
        start_right = max(next_x_left + SEP, LEFT_PAD + SEP)
        if next_x_right < start_right:
            next_x_right = start_right

        for E in les:
            costs = sorted(list(le_to_cost.get((L, E), set())))
            if costs:
                for c in costs:
                    if c not in cost_x:
                        cost_x[c] = next_x_right
                        next_x_right += X_STEP
                lane_right_xs += [cost_x[c] for c in costs]
            else:
                lane_right_xs.append(next_x_right)
                next_x_right += X_STEP

        all_xs = (lane_left_xs or []) + (lane_right_xs or [])
        if all_xs:
            le_center = int(sum(all_xs) / len(all_xs))
        else:
            le_center = next_x_left
            next_x_left += X_STEP
        for E in les:
            le_x[(L, E)] = le_center

        xs_this_ledger = [le_x[(L, e)] for e in les] if les else [le_center]
        led_x[L] = int(sum(xs_this_ledger) / len(xs_this_ledger))

        next_x_left += PAD_GROUP
        next_x_right += PAD_GROUP

    # orphan Cost Orgs on the far right
    next_x_right += RIGHT_PAD
    for c in orphan_cost:
        if c not in cost_x:
            cost_x[c] = next_x_right
            next_x_right += X_STEP

    # XML
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

    def add_vertex(label, style, x, y, w=W, h=H):
        vid = uuid.uuid4().hex[:8]
        c = ET.SubElement(root, "mxCell", attrib={"id": vid, "value": label, "style": style, "vertex": "1", "parent": "1"})
        ET.SubElement(c, "mxGeometry", attrib={"x": str(int(x)), "y": str(int(y)), "width": str(w), "height": str(h), "as": "geometry"})
        return vid

    def add_edge(src, tgt, style=S_EDGE, points=None):
        eid = uuid.uuid4().hex[:8]
        c = ET.SubElement(root, "mxCell", attrib={"id": eid, "value": "", "style": style, "edge": "1", "parent": "1", "source": src, "target": tgt})
        g = ET.SubElement(c, "mxGeometry", attrib={"relative": "1", "as": "geometry"})
        if points:
            arr = ET.SubElement(g, "Array", attrib={"as": "points"})
            for px, py in points:
                ET.SubElement(arr, "mxPoint", attrib={"x": str(int(px)), "y": str(int(py))})

    def add_bus_edge(src_id, src_center_x, tgt_id, tgt_center_x):
        add_edge(src_id, tgt_id, style=S_EDGE_LEDGER, points=[(src_center_x, BUS_Y), (tgt_center_x, BUS_Y)])

    id_map = {}

    # Ledgers
    for L in ledgers:
        id_map[("L", L)] = add_vertex(L, S_LEDGER, led_x[L], Y_LEDGER)

    # LEs
    for L in ledgers:
        for E in sorted(list(led_to_les.get(L, set()))):
            id_map[("E", L, E)] = add_vertex(E, S_LE, le_x[(L, E)], Y_LE)

    # BUs (left)
    for (L, E), buses in le_to_bus.items():
        for b in sorted(list(buses)):
            id_map[("B", b)] = add_vertex(b, S_BU, bu_x.setdefault(b, LEFT_PAD), Y_BU)

    # Cost Orgs (right, blue)
    for (L, E), costs in le_to_cost.items():
        for c in sorted(list(costs)):
            id_map[("C", c)] = add_vertex(c, S_COST, cost_x[c], Y_COST)
    for c in orphan_cost:
        id_map[("C", c)] = add_vertex(c, S_COST, cost_x[c], Y_COST)

    drawn = set()
    # BU → LE
    for (L, E), buses in le_to_bus.items():
        for b in sorted(list(buses)):
            if (("B", b) in id_map) and (("E", L, E) in id_map):
                k = ("B2E", b, L, E)
                if k not in drawn:
                    add_edge(id_map[("B", b)], id_map[("E", L, E)], style=S_EDGE)
                    drawn.add(k)

    # Cost Org → LE (blue)
    for (L, E), costs in le_to_cost.items():
        for c in sorted(list(costs)):
            if (("C", c) in id_map) and (("E", L, E) in id_map):
                k = ("C2E", c, L, E)
                if k not in drawn:
                    add_edge(id_map[("C", c)], id_map[("E", L, E)], style=S_EDGE_COST)
                    drawn.add(k)

    # LE → Ledger via bus
    for L, les in led_to_les.items():
        for E in les:
            if (("E", L, E) in id_map) and (("L", L) in id_map):
                k = ("E2L", L, E)
                if k not in drawn:
                    src_x_center = le_x[(L, E)] + W/2
                    tgt_x_center = led_x[L] + W/2
                    add_bus_edge(id_map[("E", L, E)], src_x_center, id_map[("L", L)], tgt_x_center)
                    drawn.add(k)

    # Legend
    def add_legend(x=20, y=20):
        def swatch(lbl, color, gy):
            add_vertex("", f"rounded=1;fillColor={color};strokeColor=#666666;", x+12, y+gy, 18, 12)
            add_vertex(lbl, "text;align=left;verticalAlign=middle;fontSize=12;", x+36, y+gy-4, 220, 20)
        add_vertex("", "rounded=1;fillColor=#FFFFFF;strokeColor=#CBD5E1;", x, y, 260, 150)
        swatch("Ledger", "#FFE6E6", 36)
        swatch("Legal Entity", "#FFE2C2", 62)
        swatch("Business Unit (left lane)", "#FFF1B3", 88)
        swatch("Cost Organization (right lane)", "#DDEBFF", 114)

    add_legend()
    return ET.tostring(mxfile, encoding="utf-8", method="xml").decode("utf-8")

# Only render diagram if we have at least something to show
if not df_core.empty or not df_cost.empty:
    xml = make_drawio_xml(df_core, df_cost)
    st.download_button(
        "⬇️ Download diagram (.drawio)",
        data=xml.encode("utf-8"),
        file_name="EnterpriseStructure.drawio",
        mime="application/xml"
    )
    st.markdown(f"[🔗 Open in draw.io (preview)]({_drawio_url_from_xml(xml)})")
    st.caption("Left lane = BUs • Right lane = Cost Orgs (blue). LE bridges both; Ledger sits above via the bus.")
