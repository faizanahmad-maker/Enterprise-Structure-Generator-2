import io
import zlib
import base64
import uuid
import re
import zipfile
from typing import List, Optional, Tuple

import pandas as pd
import streamlit as st
from openpyxl import Workbook
import xml.etree.ElementTree as ET


st.set_page_config(page_title="ES Generator — Multi-ZIP (Cost Orgs)", layout="wide")


# =========================
# Utilities
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
    if want in targets:
        return targets[want]
    aliases = {
        "ledger": {"ledger", "ledgername", "primaryledger", "ledger_name"},
        "legalentityname": {"legalentityname", "legal entity name", "le name", "name (legal entity)"},
        "businessunitname": {"businessunitname", "business unit name", "bu name"},
        "identifier": {"identifier", "legalentityidentifier", "le identifier", "id"},
        "name": {"name"},
    }
    for alias in aliases.get(want, set()):
        if alias in targets:
            return targets[alias]
    # loose contains
    for k, v in targets.items():
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
            "id": cid, "value": label,
            "style": f"whiteSpace=wrap;html=1;align=center;{style}",
            "vertex": "1", "parent": self.layer_id,
            "geometry": {"x": x, "y": y, "width": w, "height": h}
        })
        return cid

    def add_edge(self, src: str, dst: str, style: str = "") -> str:
        eid = _new_id()
        self.cells.append({
            "id": eid, "edge": "1", "parent": self.layer_id,
            "style": style, "source": src, "target": dst
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
# Multi-ZIP ingestion & standardization
# =========================

def _iter_zip_csvs(file_bytes: bytes):
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        for name in z.namelist():
            if name.lower().endswith(".csv"):
                with z.open(name) as f:
                    df = pd.read_csv(f, dtype=str).rename(columns=lambda c: c.strip())
                    yield name.lower(), df

def _standardize_co(df: pd.DataFrame) -> pd.DataFrame:
    name_col = _rename_like(df, "name")
    leid_col = _rename_like(df, "legalentityidentifier")
    out = df.rename(columns={name_col: "CostOrganization", leid_col: "LegalEntityIdentifier"})[
        ["CostOrganization", "LegalEntityIdentifier"]
    ].astype(str)
    return out

def _standardize_le(df: pd.DataFrame) -> pd.DataFrame:
    id_col = _rename_like(df, "identifier")
    nm_col = _rename_like(df, "name")
    out = df.rename(columns={id_col: "Identifier", nm_col: "Name"}).astype(str)
    # try to normalize ledger name if present
    ledger_col = None
    for c in out.columns:
        if _norm(c) in {"ledger", "ledgername", "primaryledger", "ledger_name"}:
            ledger_col = c
            break
    if ledger_col and ledger_col != "Ledger":
        out = out.rename(columns={ledger_col: "Ledger"})
    if "Ledger" not in out.columns:
        out["Ledger"] = ""
    return out[["Identifier", "Name", "Ledger"]]

def _standardize_llb(df: pd.DataFrame) -> pd.DataFrame:
    L = _rename_like(df, "ledger")
    LE = _rename_like(df, "legalentityname")
    BU = _rename_like(df, "businessunitname")
    out = df.rename(columns={L: "Ledger", LE: "LegalEntityName", BU: "BusinessUnitName"})[
        ["Ledger", "LegalEntityName", "BusinessUnitName"]
    ].astype(str)
    return out

def load_from_multi_zips(files: List[io.BytesIO]) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Scan all uploaded zips and merge:
      - Cost Orgs (required)
      - Legal Entities (required)
      - Ledger–LE–BU (optional)
    Returns (df_le_bu, df_le, df_co)
    """
    co_parts, le_parts, llb_parts = [], [], []

    for f in files:
        for fname, df in _iter_zip_csvs(f):
            lf = fname.lower()
            if "cst_cost_organization" in lf or "cost_organization" in lf:
                try:
                    co_parts.append(_standardize_co(df))
                except Exception:
                    pass
            elif "legal_entities" in lf or "legal entity" in lf:
                try:
                    le_parts.append(_standardize_le(df))
                except Exception:
                    pass
            elif ("ledger" in lf and "legal" in lf and "business" in lf) or ("ledger-le-bu" in lf):
                try:
                    llb_parts.append(_standardize_llb(df))
                except Exception:
                    pass

    if not co_parts:
        raise ValueError("None of the ZIPs contained CST_COST_ORGANIZATION*.csv (Cost Organizations).")
    if not le_parts:
        raise ValueError("None of the ZIPs contained MANAGE_LEGAL_ENTITIES*.csv (Legal Entities).")

    df_co = (pd.concat(co_parts, ignore_index=True).drop_duplicates().reset_index(drop=True))
    df_le = (pd.concat(le_parts, ignore_index=True).drop_duplicates().reset_index(drop=True))

    if llb_parts:
        df_le_bu = (pd.concat(llb_parts, ignore_index=True)
                    .dropna(subset=["LegalEntityName", "Ledger"], how="all")
                    .drop_duplicates().reset_index(drop=True))
    else:
        # minimal when BU data missing
        df_le_bu = df_le.rename(columns={"Name": "LegalEntityName"})[["Ledger", "LegalEntityName"]].drop_duplicates()
        df_le_bu["BusinessUnitName"] = pd.NA

    return df_le_bu, df_le, df_co


# =========================
# Cost Org tab & XLSX builder
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

# Layout (disciplined columns; CO lower)
X_LEDGER = 40
X_LE     = 320
X_BU     = 650
X_CO     = 950

Y_START_LEDGER = 40
Y_START_LE     = 40
Y_START_BU     = 40
Y_START_CO     = 220
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

    # BUs (if present)
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

st.title("Enterprise Structure Generator — Multi-ZIP Input with Cost Orgs")
st.caption(
    "Drop **one or more Oracle export ZIPs**. I’ll merge `CST_COST_ORGANIZATION*.csv` and `MANAGE_LEGAL_ENTITIES*.csv` "
    "across all files, build **ES – Ledger–LE–CostOrg**, and generate a .drawio with **yellow LE→BU** and **blue LE→Cost Org** "
    "(line-jump bridges). Cost Orgs render on a lower vertical layer than BUs."
)

with st.sidebar:
    zip_files = st.file_uploader("Oracle Export ZIPs", type=["zip"], accept_multiple_files=True)
    run = st.button("Build Tab + Diagram")

st.markdown("""
**Inside each ZIP, I try to find (names can vary):**
- `CST_COST_ORGANIZATION*.csv` → columns **Name**, **LegalEntityIdentifier** *(required)*  
- `MANAGE_LEGAL_ENTITIES*.csv` → columns **Identifier**, **Name** *(required; Ledger optional)*  
- `*Ledger*Legal*Business*.csv` or `*ledger-le-bu*.csv` → **Ledger**, **LegalEntityName**, **BusinessUnitName** *(optional)*
""")

if run:
    if not zip_files:
        st.error("Upload at least one ZIP.")
        st.stop()
    try:
        # Read all zips into memory bytes list (Streamlit UploadedFile is file-like)
        zip_bytes_list = [zf.read() for zf in zip_files]

        df_le_bu, df_le, df_co = load_from_multi_zips(zip_bytes_list)
        df_cost_tab = build_costorg_tab(df_le_bu, df_le, df_co)

        # Output workbook containing the new tab
        xlsx_bytes = dataframe_to_xlsx_bytes(df_cost_tab, "ES – Ledger–LE–CostOrg")

        # Diagram
        drawio_xml = build_drawio(df_le_bu, df_cost_tab)

        st.success("✅ Built the Cost Org tab and diagram from multiple ZIPs.")

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
        st.info(f"Unmapped Cost Orgs (no LE name found after merges): **{unmapped}**")

        with st.expander("Diagram XML (first ~60 lines)"):
            st.code("\n".join(drawio_xml.splitlines()[:60]), language="xml")

    except Exception as e:
        st.exception(e)
