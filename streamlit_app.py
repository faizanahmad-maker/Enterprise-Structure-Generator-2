import io
import zlib
import base64
import uuid
import re
import zipfile
from typing import Dict, Tuple, Optional

import pandas as pd
import streamlit as st
from openpyxl import Workbook
import xml.etree.ElementTree as ET


st.set_page_config(page_title="ES Generator — ZIP Input (Cost Orgs)", layout="wide")


# =========================
# Small utilities
# =========================

def _deflate_base64(xml_text: str) -> str:
    compressor = zlib.compressobj(level=9, wbits=-15)  # raw DEFLATE
    compressed = compressor.compress(xml_text.encode("utf-8")) + compressor.flush()
    return base64.b64encode(compressed).decode("ascii")

def _new_id() -> str:
    return str(uuid.uuid4()).replace("-", "")[:12]

def _norm(s: Optional[str]) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip().lower()

def _rename_like(df: pd.DataFrame, want: str) -> str:
    """Return the actual column name in df that matches logical name `want` by normalization."""
    targets = { _norm(c): c for c in df.columns }
    if want in targets:                     # exact normalized match
        return targets[want]
    # common aliases
    aliases = {
        "ledger": {"ledger", "ledgername", "primaryledger", "ledger_name"},
        "legalentityname": {"legalentityname", "legal entity name", "le name", "name (legal entity)"},
        "businessunitname": {"businessunitname", "business unit name", "bu name"},
        "identifier": {"identifier", "legalentityidentifier", "le identifier", "id"},
        "name": {"name"}
    }
    for alias in aliases.get(want, set()):
        if alias in targets:
            return targets[alias]
    # loose contains
    for k,v in targets.items():
        if want in k:
            return v
    raise KeyError(f"Could not locate a column like '{want}' in {list(df.columns)}")


# =========================
# Draw.io (mxGraph) builder
# =========================

class GraphBuilder:
    def __init__(self, name: str = "ES Diagram"):
        self.name = name
        self.root_id = "0"
        self.layer_id = "1"
        self.cells = []
        self.cells.append({"id": self.root_id})
        self.cells.append({"id": self.layer_id, "parent": self.root_id})

    def add_node(self, label: str, x: int, y: int, w: int, h: int, style: str = "") -> str:
        cid = _new_id()
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

    def _mx(self) -> str:
        mx = ET.Element("mxGraphModel")
        root = ET.SubElement(mx, "root")
        for c in self.cells:
            attrs = {k: v for k, v in c.items() if k not in ("geometry",)}
            cell = ET.SubElement(root, "mxCell", attrs)
            if "geometry" in c:
                g = c["geometry"]
                ET.SubElement(cell, "mxGeometry", {
                    "x": str(g["x"]), "y": str(g["y"]),
                    "width": str(g["width"]), "height": str(g["height"]),
                    "as": "geometry"
                })
        return ET.tostring(mx, encoding="utf-8").decode("utf-8")

    def to_drawio_xml(self) -> str:
        inner = self._mx()
        enc = _deflate_base64(inner)
        mxfile = ET.Element("mxfile", {"host": "app.diagrams.net"})
        diagram = ET.SubElement(mxfile, "diagram", {"name": self.name, "id": _new_id()})
        diagram.text = enc
        return ET.tostring(mxfile, encoding="utf-8", xml_declaration=True).decode("utf-8")


# =========================
# ZIP ingestion
# =========================

def extract_csvs_from_zip(zip_bytes: bytes) -> Dict[str, pd.DataFrame]:
    """
    Read all CSVs from the uploaded Oracle export ZIP.
    Return dict: {lower_filename: DataFrame}
    """
    out = {}
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
        for name in z.namelist():
            if name.lower().endswith(".csv"):
                with z.open(name) as f:
                    df = pd.read_csv(f, dtype=str).rename(columns=lambda c: c.strip())
                    out[name.lower()] = df
    if not out:
        raise ValueError("No CSV files found inside the ZIP.")
    return out

def select_frames_from_zip(csvs: Dict[str, pd.DataFrame]) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Return (df_le_bu, df_le, df_co)
    - df_le_bu columns: Ledger, LegalEntityName, BusinessUnitName  (may be empty if not found)
    - df_le columns: Identifier, Name, [Ledger optional]
    - df_co columns: CostOrganization, LegalEntityIdentifier
    """
    # Cost Orgs (required)
    co_key = next((k for k in csvs if "cst_cost_organization" in k or "cost_organization" in k), None)
    if not co_key:
        raise ValueError("ZIP must include CST_COST_ORGANIZATION.csv (Cost Organizations).")
    df_co_raw = csvs[co_key].copy()

    # Legal Entities (required)
    le_key = next((k for k in csvs if "legal_entities" in k or "legal entity" in k), None)
    if not le_key:
        raise ValueError("ZIP must include MANAGE_LEGAL_ENTITIES.csv (Legal Entities).")
    df_le_raw = csvs[le_key].copy()

    # Ledger–LE–BU (optional; many orgs have a curated export; otherwise we skip BUs)
    llb_key = next(
        (k for k in csvs if "ledger" in k and "legal" in k and "business" in k),
        None
    )
    df_le_bu_raw = csvs[llb_key].copy() if llb_key else pd.DataFrame()

    # --- Standardize CO
    name_col = _rename_like(df_co_raw, "name")
    leid_col = _rename_like(df_co_raw, "legalentityidentifier")
    df_co = df_co_raw.rename(columns={name_col: "CostOrganization", leid_col: "LegalEntityIdentifier"})[
        ["CostOrganization", "LegalEntityIdentifier"]
    ].astype(str)

    # --- Standardize LE
    id_col = _rename_like(df_le_raw, "identifier")
    nm_col = _rename_like(df_le_raw, "name")
    df_le = df_le_raw.rename(columns={id_col: "Identifier", nm_col: "Name"}).astype(str)

    # Optional ledger col
    ledger_col = None
    for c in df_le.columns:
        if _norm(c) in {"ledger", "ledgername", "primaryledger", "ledger_name"}:
            ledger_col = c
            break
    if ledger_col and ledger_col != "Ledger":
        df_le = df_le.rename(columns={ledger_col: "Ledger"})
    if "Ledger" not in df_le.columns:
        df_le["Ledger"] = ""

    # --- Standardize Ledger–LE–BU (if present)
    if not df_le_bu_raw.empty:
        L = _rename_like(df_le_bu_raw, "ledger")
        LE = _rename_like(df_le_bu_raw, "legalentityname")
        BU = _rename_like(df_le_bu_raw, "businessunitname")
        df_le_bu = df_le_bu_raw.rename(columns={L: "Ledger", LE: "LegalEntityName", BU: "BusinessUnitName"})[
            ["Ledger", "LegalEntityName", "BusinessUnitName"]
        ].astype(str)
    else:
        # minimal frame (no BUs)
        df_le_bu = (
            df_le.rename(columns={"Name": "LegalEntityName"})[["Ledger", "LegalEntityName"]]
            .drop_duplicates()
        )
        df_le_bu["BusinessUnitName"] = pd.NA

    return df_le_bu, df_le[["Identifier", "Name", "Ledger"]], df_co


# =========================
# Cost Org tab builder
# =========================

def build_costorg_tab(df_le_bu: pd.DataFrame, df_le: pd.DataFrame, df_co: pd.DataFrame) -> pd.DataFrame:
    le_lookup = df_le.rename(columns={"Identifier": "LegalEntityIdentifier",
                                      "Name": "LegalEntityName"})[
        ["LegalEntityIdentifier", "LegalEntityName", "Ledger"]
    ]
    out = (df_co.merge(le_lookup, on="LegalEntityIdentifier", how="left"))
    out["__WARN_UnmappedLE"] = out["LegalEntityName"].isna() | (out["LegalEntityName"] == "")
    out = out[["Ledger", "LegalEntityName", "CostOrganization", "LegalEntityIdentifier", "__WARN_UnmappedLE"]]
    out = out.sort_values(["Ledger", "LegalEntityName", "CostOrganization"],
                          na_position="last").reset_index(drop=True)
    return out


def dataframe_to_xlsx_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(list(df.columns))
    for _, r in df.iterrows():
        ws.append(list(r.values))
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.read()


# =========================
# Diagram builder
# =========================

# Layout (tweak to taste)
X_LEDGER = 40
X_LE     = 320
X_BU     = 650
X_CO     = 950

Y_START_LEDGER = 40
Y_START_LE     = 40
Y_START_BU     = 40
Y_START_CO     = 220   # Cost Orgs lower than BU
Y_STEP = 90
NODE_W, NODE_H = 170, 48

STYLE_NODE_LEDGER = "rounded=1;fillColor=#F5F5F5;strokeColor=#666666;fontStyle=1;"
STYLE_NODE_LE     = "rounded=1;fillColor=#FFFFFF;strokeColor=#222222;"
STYLE_NODE_BU     = "rounded=1;fillColor=#FFFFFF;strokeColor=#888888;"
STYLE_NODE_CO     = "rounded=1;fillColor=#E9F2FF;strokeColor=#1F75FE;"

STYLE_EDGE_BASE = ("endArrow=block;endFill=1;rounded=1;jettySize=auto;"
                   "orthogonalLoop=1;edgeStyle=orthogonalEdgeStyle;curved=1;"
                   "jumpStyle=arc;jumpSize=10;")
STYLE_LEDGER_TO_LE = STYLE_EDGE_BASE + "strokeColor=#666666;strokeWidth=2;"
STYLE_LE_TO_BU     = STYLE_EDGE_BASE + "strokeColor=#FFD400;strokeWidth=2;"  # yellow
STYLE_LE_TO_CO     = STYLE_EDGE_BASE + "strokeColor=#1F75FE;strokeWidth=2;"  # blue


def build_drawio(df_le_bu: pd.DataFrame, df_cost_tab: pd.DataFrame) -> str:
    g = GraphBuilder(name="Enterprise Structure (+ Cost Orgs)")

    ledgers = sorted(set(df_le_bu["Ledger"].dropna()) | set(df_cost_tab["Ledger"].dropna()))
    if not ledgers:
        ledgers = ["(No Ledger)"]

    # Ledgers
    ledger_ids, y = {}, Y_START_LEDGER
    for L in ledgers:
        lbl = L if L else "(No Ledger)"
        nid = g.add_node(lbl, X_LEDGER, y, NODE_W, NODE_H, STYLE_NODE_LEDGER)
        ledger_ids[lbl] = nid
        y += Y_STEP

    # LEs (from both frames)
    df_le_all = (pd.concat([df_le_bu[["Ledger", "LegalEntityName"]],
                            df_cost_tab[["Ledger", "LegalEntityName"]]],
                           ignore_index=True)
                 .drop_duplicates()
                 .sort_values(["Ledger", "LegalEntityName"], na_position="last"))

    le_ids, last_L, y_le = {}, None, Y_START_LE
    for _, r in df_le_all.iterrows():
        L = r["Ledger"] if pd.notna(r["Ledger"]) else ""
        LE = r["LegalEntityName"] if pd.notna(r["LegalEntityName"]) else "(Unnamed LE)"
        if L != last_L:
            y_le = Y_START_LE
            last_L = L
        nid = g.add_node(LE, X_LE, y_le, NODE_W, NODE_H, STYLE_NODE_LE)
        le_ids[(L or "", LE)] = nid
        g.add_edge(ledger_ids.get(L or "", list(ledger_ids.values())[0]), nid, STYLE_LEDGER_TO_LE)
        y_le += Y_STEP

    # BUs (if we have them)
    if df_le_bu["BusinessUnitName"].notna().any():
        for (L, LE), grp in df_le_bu.groupby(["Ledger", "LegalEntityName"], dropna=False):
            y_bu = Y_START_BU
            for bu in sorted(set(grp["BusinessUnitName"].dropna())):
                bl = bu if bu else "(Unnamed BU)"
                nid = g.add_node(bl, X_BU, y_bu, NODE_W, NODE_H, STYLE_NODE_BU)
                g.add_edge(le_ids[(L or "", LE or "(Unnamed LE)")], nid, STYLE_LE_TO_BU)
                y_bu += Y_STEP

    # Cost Orgs (lower)
    for (L, LE), grp in df_cost_tab.groupby(["Ledger", "LegalEntityName"], dropna=False):
        y_co = Y_START_CO
        for co in sorted(set(grp["CostOrganization"].dropna())):
            colbl = co if co else "(Unnamed Cost Org)"
            nid = g.add_node(colbl, X_CO, y_co, NODE_W, NODE_H, STYLE_NODE_CO)
            g.add_edge(le_ids[(L or "", LE or "(Unnamed LE)")], nid, STYLE_LE_TO_CO)
            y_co += Y_STEP

    return g.to_drawio_xml()


# =========================
# UI
# =========================

st.title("Enterprise Structure Generator — ZIP Input with Cost Orgs")
st.caption(
    "Upload your Oracle export **ZIP**. I’ll read `CST_COST_ORGANIZATION.csv` and `MANAGE_LEGAL_ENTITIES.csv`, "
    "build a new **ES – Ledger–LE–CostOrg** sheet, and produce a .drawio with **yellow LE→BU** and **blue LE→Cost Org** "
    "(with line-jump bridges). Cost Orgs sit on a lower vertical layer than BUs."
)

with st.sidebar:
    zip_file = st.file_uploader("Oracle Export (.zip)", type=["zip"])
    run = st.button("Build Tab + Diagram")

st.markdown("""
**What I look for inside the ZIP (filenames can vary):**
- `CST_COST_ORGANIZATION*.csv` – must include columns **Name**, **LegalEntityIdentifier**  
- `MANAGE_LEGAL_ENTITIES*.csv` – must include columns **Identifier**, **Name** (Ledger optional)  
- *(Optional)* a curated **Ledger–LE–BU** CSV containing **Ledger**, **LegalEntityName**, **BusinessUnitName**  
""")

if run:
    if not zip_file:
        st.error("Upload a ZIP file first.")
        st.stop()
    try:
        csvs = extract_csvs_from_zip(zip_file.read())
        df_le_bu, df_le, df_co = select_frames_from_zip(csvs)
        df_cost_tab = build_costorg_tab(df_le_bu, df_le, df_co)

        # Output workbook containing only the new tab (by design for this ZIP flow)
        xlsx_bytes = dataframe_to_xlsx_bytes(df_cost_tab, "ES – Ledger–LE–CostOrg")

        # Diagram
        drawio_xml = build_drawio(df_le_bu, df_cost_tab)

        st.success("✅ Built the Cost Org tab and diagram from the ZIP.")

        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "⬇️ Download ES – Ledger–LE–CostOrg.xlsx",
                data=xlsx_bytes,
                file_name="enterprise_structure.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with c2:
            st.download_button(
                "⬇️ Download enterprise_structure.drawio",
                data=drawio_xml.encode("utf-8"),
                file_name="enterprise_structure.drawio",
                mime="application/xml",
            )

        with st.expander("Preview: ES – Ledger–LE–CostOrg (first 50 rows)"):
            st.dataframe(df_cost_tab.head(50))

        unmapped = int(df_cost_tab["__WARN_UnmappedLE"].sum())
        st.info(f"Unmapped Cost Orgs (no LE name found): **{unmapped}**")

        with st.expander("Diagram XML (first ~60 lines)"):
            st.code("\n".join(drawio_xml.splitlines()[:60]), language="xml")

    except Exception as e:
        st.exception(e)

