import io
import zlib
import base64
import uuid
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
import xml.etree.ElementTree as ET

st.set_page_config(page_title="Enterprise Structure + Cost Orgs", layout="wide")

# =========================
# Utilities
# =========================

def _deflate_base64(xml_text: str) -> str:
    """Deflate (zlib) + base64 encode, with no zlib header (raw)."""
    compressor = zlib.compressobj(level=9, wbits=-15)
    compressed = compressor.compress(xml_text.encode("utf-8")) + compressor.flush()
    return base64.b64encode(compressed).decode("ascii")

def _new_id() -> str:
    return str(uuid.uuid4()).replace("-", "")[:12]

# =========================
# Graph Builder for draw.io
# =========================

class GraphBuilder:
    """
    Minimal mxGraph builder that produces a .drawio (mxfile) with one diagram.
    Handles nodes as mxCell vertices and edges with styles.
    Coordinates are absolute; we keep parent as the root layer.
    """
    def __init__(self, name: str = "ES Diagram"):
        self.name = name
        self.root_id = "0"
        self.layer_id = "1"
        self.cells = []  # list of dicts representing mxCell attrs
        self.node_ids = set()

        # create root + layer cells
        self.cells.append({
            "id": self.root_id,
        })
        self.cells.append({
            "id": self.layer_id,
            "parent": self.root_id
        })

    def add_node(self, label: str, x: int, y: int, w: int, h: int, style: str = "") -> str:
        cid = _new_id()
        self.node_ids.add(cid)
        self.cells.append({
            "id": cid,
            "value": label,
            "style": f"whiteSpace=wrap;html=1;align=center;{style}",
            "vertex": "1",
            "parent": self.layer_id,
            "geometry": {"x": x, "y": y, "width": w, "height": h}
        })
        return cid

    def add_edge(self, src: str, dst: str, style: str = "") -> str:
        eid = _new_id()
        self.cells.append({
            "id": eid,
            "edge": "1",
            "parent": self.layer_id,
            "style": style,
            "source": src,
            "target": dst
        })
        return eid

    def to_mxgraphmodel_xml(self) -> str:
        """
        Build <mxGraphModel><root>...</root></mxGraphModel> XML string.
        """
        mxGraphModel = ET.Element("mxGraphModel")
        root = ET.SubElement(mxGraphModel, "root")

        # write cells
        for c in self.cells:
            mxCell = ET.SubElement(root, "mxCell", {k: v for k, v in c.items()
                                                    if k not in ("geometry",)})
            if "geometry" in c:
                g = c["geometry"]
                ET.SubElement(mxCell, "mxGeometry", {
                    "x": str(g["x"]), "y": str(g["y"]),
                    "width": str(g["width"]), "height": str(g["height"]),
                    "as": "geometry"
                })
        return ET.tostring(mxGraphModel, encoding="utf-8", xml_declaration=False).decode("utf-8")

    def to_drawio_xml(self) -> str:
        """
        Wrap mxGraphModel in <mxfile><diagram>compressed</diagram></mxfile>.
        """
        inner = self.to_mxgraphmodel_xml()
        enc = _deflate_base64(inner)
        mxfile = ET.Element("mxfile", {"host": "app.diagrams.net"})
        diagram = ET.SubElement(mxfile, "diagram", {"name": self.name, "id": _new_id()})
        diagram.text = enc
        return ET.tostring(mxfile, encoding="utf-8", xml_declaration=True).decode("utf-8")

# =========================
# Core data builders
# =========================

REQUIRED_CO_COLS = {"Name", "LegalEntityIdentifier"}
REQUIRED_LE_COLS = {"Identifier", "Name"}

def load_cost_orgs(cost_org_csv: io.BytesIO, legal_entities_csv: io.BytesIO) -> pd.DataFrame:
    df_co = pd.read_csv(cost_org_csv, dtype=str).rename(columns=lambda c: c.strip())
    df_le = pd.read_csv(legal_entities_csv, dtype=str).rename(columns=lambda c: c.strip())

    missing_co = REQUIRED_CO_COLS - set(df_co.columns)
    missing_le = REQUIRED_LE_COLS - set(df_le.columns)
    if missing_co:
        raise ValueError(f"CST_COST_ORGANIZATION.csv missing columns: {missing_co}")
    if missing_le:
        raise ValueError(f"LEGAL_ENTITIES.csv missing columns: {missing_le}")

    df_le = df_le.rename(columns={"Name": "LegalEntityName",
                                  "Identifier": "LegalEntityIdentifier"})

    # Try to locate ledger column if present
    ledger_guess = None
    for c in df_le.columns:
        lc = c.lower()
        if lc in ("ledger", "ledgername", "primaryledger", "ledger_name"):
            ledger_guess = c
            break
    if ledger_guess:
        df_le = df_le.rename(columns={ledger_guess: "Ledger"})
    else:
        df_le["Ledger"] = ""

    df = (df_co.rename(columns={"Name": "CostOrganization"})
                .merge(df_le[["LegalEntityIdentifier", "LegalEntityName", "Ledger"]],
                       on="LegalEntityIdentifier", how="left"))

    return df[["Ledger", "LegalEntityIdentifier", "LegalEntityName", "CostOrganization"]]

def build_cost_org_tab(xlsx_bytes_or_none, df_cost: pd.DataFrame,
                       out_sheet_name: str = "ES – Ledger–LE–CostOrg") -> bytes:
    """
    Writes/overwrites the cost-org mapping sheet into supplied workbook (if provided),
    else builds a new workbook with only that sheet. Returns bytes of the workbook.
    """
    df = df_cost.copy()
    df["__WARN_UnmappedLE"] = df["LegalEntityName"].isna() | (df["LegalEntityName"] == "")
    df = df[["Ledger", "LegalEntityName", "CostOrganization", "LegalEntityIdentifier", "__WARN_UnmappedLE"]]
    df = df.sort_values(["Ledger", "LegalEntityName", "CostOrganization"],
                        na_position="last").reset_index(drop=True)

    output = io.BytesIO()
    if xlsx_bytes_or_none:
        # load and replace sheet
        wb = load_workbook(filename=io.BytesIO(xlsx_bytes_or_none))
        if out_sheet_name in wb.sheetnames:
            ws = wb[out_sheet_name]
            wb.remove(ws)
        ws = wb.create_sheet(out_sheet_name)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = out_sheet_name

    # write headers
    headers = list(df.columns)
    ws.append(headers)
    for _, row in df.iterrows():
        ws.append(list(row.values))

    wb.save(output)
    output.seek(0)
    return output.read()

# =========================
# Diagram wiring
# =========================

# Layout constants (tweak if your canvases are wider)
X_LEDGER = 40
X_LE     = 320
X_BU     = 650
X_CO     = 950

Y_START_LEDGER = 40
Y_START_LE     = 40
Y_START_BU     = 40
Y_START_CO     = 220   # lower baseline so Cost Orgs sit below BUs

Y_STEP = 90
NODE_W, NODE_H = 170, 48

STYLE_NODE_LEDGER = "rounded=1;fillColor=#F5F5F5;strokeColor=#666666;fontStyle=1;"
STYLE_NODE_LE     = "rounded=1;fillColor=#FFFFFF;strokeColor=#222222;"
STYLE_NODE_BU     = "rounded=1;fillColor=#FFFFFF;strokeColor=#888888;"
STYLE_NODE_CO     = "rounded=1;fillColor=#E9F2FF;strokeColor=#1F75FE;"

STYLE_EDGE_BASE = "endArrow=block;endFill=1;rounded=1;jettySize=auto;orthogonalLoop=1;edgeStyle=orthogonalEdgeStyle;curved=1;jumpStyle=arc;jumpSize=10;"
STYLE_LEDGER_TO_LE = STYLE_EDGE_BASE + "strokeColor=#666666;strokeWidth=2;"
STYLE_LE_TO_BU     = STYLE_EDGE_BASE + "strokeColor=#FFD400;strokeWidth=2;"   # yellow
STYLE_LE_TO_CO     = STYLE_EDGE_BASE + "strokeColor=#1F75FE;strokeWidth=2;"   # blue

def build_diagram(df_ledger_le_bu: pd.DataFrame,
                  df_cost: pd.DataFrame) -> str:
    """
    df_ledger_le_bu must have columns: Ledger, LegalEntityName, BusinessUnitName
    df_cost must have columns: Ledger, LegalEntityName, CostOrganization
    Returns .drawio XML string.
    """
    g = GraphBuilder(name="Enterprise Structure (with Cost Orgs)")

    # --- Place Ledgers
    ledger_node_ids = {}
    y = Y_START_LEDGER
    for ledger in sorted(set(df_ledger_le_bu["Ledger"].dropna()) | set(df_cost["Ledger"].dropna())):
        label = ledger if pd.notna(ledger) and ledger != "" else "(No Ledger)"
        nid = g.add_node(label=label, x=X_LEDGER, y=y, w=NODE_W, h=NODE_H, style=STYLE_NODE_LEDGER)
        ledger_node_ids[label] = nid
        y += Y_STEP

    # --- Place LEs (group by ledger)
    le_node_ids = {}
    # Ensure LE rows exist from either frame
    df_le_all = (pd.concat([
                    df_ledger_le_bu[["Ledger", "LegalEntityName"]],
                    df_cost[["Ledger", "LegalEntityName"]]
                 ], ignore_index=True)
                 .drop_duplicates()
                 .sort_values(["Ledger", "LegalEntityName"], na_position="last"))

    last_ledger = None
    y_le = Y_START_LE
    for _, r in df_le_all.iterrows():
        ledger = r.get("Ledger") if pd.notna(r.get("Ledger")) else "(No Ledger)"
        le     = r.get("LegalEntityName") if pd.notna(r.get("LegalEntityName")) else "(Unnamed LE)"
        if ledger != last_ledger:
            y_le = Y_START_LE
            last_ledger = ledger
        nid = g.add_node(label=le, x=X_LE, y=y_le, w=NODE_W, h=NODE_H, style=STYLE_NODE_LE)
        le_node_ids[(ledger, le)] = nid
        # edge Ledger -> LE
        g.add_edge(src=ledger_node_ids.get(ledger, list(ledger_node_ids.values())[0]),
                   dst=nid, style=STYLE_LEDGER_TO_LE)
        y_le += Y_STEP

    # --- Place BUs (group by LE)
    bu_node_ids = {}
    if not df_ledger_le_bu.empty:
        for (ledger, le), grp in df_ledger_le_bu.groupby(["Ledger", "LegalEntityName"], dropna=False):
            y_bu = Y_START_BU
            for bu in sorted(set(grp["BusinessUnitName"].dropna())):
                bu_label = bu if bu else "(Unnamed BU)"
                nid = g.add_node(label=bu_label, x=X_BU, y=y_bu, w=NODE_W, h=NODE_H, style=STYLE_NODE_BU)
                bu_node_ids[(ledger or "(No Ledger)", le or "(Unnamed LE)", bu_label)] = nid
                # edge LE -> BU (yellow)
                g.add_edge(src=le_node_ids[(ledger or "(No Ledger)", le or "(Unnamed LE)")],
                           dst=nid, style=STYLE_LE_TO_BU)
                y_bu += Y_STEP

    # --- Place Cost Orgs (lower layer)
    co_node_ids = {}
    if not df_cost.empty:
        for (ledger, le), grp in df_cost.groupby(["Ledger", "LegalEntityName"], dropna=False):
            y_co = Y_START_CO
            for co in sorted(set(grp["CostOrganization"].dropna())):
                co_label = co if co else "(Unnamed Cost Org)"
                nid = g.add_node(label=co_label, x=X_CO, y=y_co, w=NODE_W, h=NODE_H, style=STYLE_NODE_CO)
                co_node_ids[(ledger or "(No Ledger)", le or "(Unnamed LE)", co_label)] = nid
                # edge LE -> Cost Org (blue)
                g.add_edge(src=le_node_ids[(ledger or "(No Ledger)", le or "(Unnamed LE)")],
                           dst=nid, style=STYLE_LE_TO_CO)
                y_co += Y_STEP

    return g.to_drawio_xml()

# =========================
# Streamlit UI
# =========================

st.title("Enterprise Structure Generator — Increment: Cost Organizations")
st.caption("Adds a second tab (Ledger–LE–CostOrg) and a diagram layer with LE→Cost Org (blue) edges. BU remains yellow, Cost Orgs sit lower vertically.")

with st.sidebar:
    st.header("Inputs")
    xlsx_file = st.file_uploader("(Optional) Existing ES Workbook (.xlsx)", type=["xlsx"])
    le_bu_csv = st.file_uploader("Ledger–LE–BU CSV (columns: Ledger, LegalEntityName, BusinessUnitName)", type=["csv"])
    cost_org_csv = st.file_uploader("CST_COST_ORGANIZATION.csv", type=["csv"])
    legal_entities_csv = st.file_uploader("MANAGE_LEGAL_ENTITIES.csv", type=["csv"])

    run = st.button("Build Tab + Diagram")

# Hints
st.markdown("""
**Expected columns**  
- *CST_COST_ORGANIZATION.csv*: `Name`, `LegalEntityIdentifier`  
- *MANAGE_LEGAL_ENTITIES.csv*: `Identifier`, `Name`, and ideally a `Ledger` column (or `LedgerName` / `PrimaryLedger`)  
- *Ledger–LE–BU CSV*: `Ledger`, `LegalEntityName`, `BusinessUnitName`
""")

if run:
    # Validate mandatory files
    if not cost_org_csv or not legal_entities_csv:
        st.error("Please upload both **CST_COST_ORGANIZATION.csv** and **MANAGE_LEGAL_ENTITIES.csv**.")
        st.stop()
    if not le_bu_csv:
        st.warning("No Ledger–LE–BU CSV provided. Diagram will include Ledgers and LEs and skip BUs.")

    try:
        # Load frames
        df_cost = load_cost_orgs(cost_org_csv, legal_entities_csv)

        if le_bu_csv:
            df_le_bu = pd.read_csv(le_bu_csv, dtype=str).rename(columns=lambda c: c.strip())
            need = {"Ledger", "LegalEntityName", "BusinessUnitName"}
            missing = need - set(df_le_bu.columns)
            if missing:
                st.error(f"Ledger–LE–BU CSV missing columns: {missing}")
                st.stop()
        else:
            # build minimal frame with whatever ledgers/LEs exist from cost-org file
            df_le_bu = (df_cost[["Ledger", "LegalEntityName"]].drop_duplicates())
            df_le_bu["BusinessUnitName"] = pd.NA

        # Build/Update workbook tab
        xlsx_in_bytes = xlsx_file.read() if xlsx_file is not None else None
        wb_bytes = build_cost_org_tab(xlsx_in_bytes, df_cost, out_sheet_name="ES – Ledger–LE–CostOrg")

        st.success("Workbook tab created.")
        st.download_button("⬇️ Download updated workbook", data=wb_bytes,
                           file_name="enterprise_structure.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # Build diagram
        drawio_xml = build_diagram(df_ledger_le_bu=df_le_bu, df_cost=df_cost)
        st.success("Diagram generated.")
        st.download_button("⬇️ Download .drawio diagram", data=drawio_xml.encode("utf-8"),
                           file_name="enterprise_structure.drawio", mime="application/xml")

        # Quick previews
        with st.expander("Preview: ES – Ledger–LE–CostOrg (first 50 rows)"):
            st.dataframe(df_cost.head(50))

        with st.expander("Preview: Diagram XML (first 60 lines)"):
            snippet = "\n".join(drawio_xml.splitlines()[:60])
            st.code(snippet, language="xml")

    except Exception as e:
        st.exception(e)
