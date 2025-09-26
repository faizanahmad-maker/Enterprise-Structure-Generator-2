import io
import zlib
import base64
import uuid
from pathlib import Path
import re

import pandas as pd
import streamlit as st
from openpyxl import load_workbook, Workbook
import xml.etree.ElementTree as ET

st.set_page_config(page_title="ES Generator (OG Input) + Cost Orgs", layout="wide")

# =========================
# Helpers
# =========================

def _deflate_base64(xml_text: str) -> str:
    compressor = zlib.compressobj(level=9, wbits=-15)
    compressed = compressor.compress(xml_text.encode("utf-8")) + compressor.flush()
    return base64.b64encode(compressed).decode("ascii")

def _new_id() -> str:
    return str(uuid.uuid4()).replace("-", "")[:12]

def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip().lower()

def _best_sheet(sheets: dict, candidates: list[str]):
    """
    Pick a sheet by fuzzy match on name.
    - sheets: dict[str, DataFrame]
    - candidates: list of candidate names/regex (lowercased)
    """
    keys = {k: v for k, v in sheets.items()}
    lower = {k.lower(): k for k in keys.keys()}
    # Exact-lower match first
    for c in candidates:
        if c in lower:
            return keys[lower[c]]
    # Loose contains
    for k in keys.keys():
        lk = k.lower()
        for c in candidates:
            if c in lk:
                return keys[k]
    return None

# =========================
# Minimal draw.io builder
# =========================

class GraphBuilder:
    def __init__(self, name: str = "ES Diagram"):
        self.name = name
        self.root_id = "0"
        self.layer_id = "1"
        self.cells = []
        self.cells.append({"id": self.root_id})
        self.cells.append({"id": self.layer_id, "parent": self.root_id})

    def add_node(self, label, x, y, w, h, style=""):
        cid = _new_id()
        self.cells.append({
            "id": cid, "value": label,
            "style": f"whiteSpace=wrap;html=1;align=center;{style}",
            "vertex": "1", "parent": self.layer_id,
            "geometry": {"x": x, "y": y, "width": w, "height": h}
        })
        return cid

    def add_edge(self, src, dst, style=""):
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
# Core logic
# =========================

RE_CO = {"name", "legalentityidentifier"}  # column set after lowercase/strip
RE_LE = {"identifier", "name"}             # ledger optional

def read_es_workbook(xlsx_bytes: bytes) -> dict[str, pd.DataFrame]:
    sheets = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=None, dtype=str)
    # normalize headers
    for k, df in sheets.items():
        df.columns = [c.strip() for c in df.columns]
    return sheets

def get_frames_from_og_wb(sheets: dict[str, pd.DataFrame]):
    """
    Extract:
      - df_le_bu: columns Ledger, LegalEntityName, BusinessUnitName
      - df_le:    columns Identifier, Name, [Ledger?]
      - df_co:    columns Name, LegalEntityIdentifier
    """
    # 1) Ledger–LE–BU (your OG tab)
    df_le_bu = _best_sheet(
        sheets,
        ["es – ledger–le–bu", "es- ledger-le-bu", "es - ledger-le-bu", "ledger–le–bu", "ledger-le-bu"]
    )
    if df_le_bu is None:
        # try a generic finder by required columns
        for name, df in sheets.items():
            cols = {_norm(c) for c in df.columns}
            if {"ledger", "legalentityname", "businessunitname"}.issubset(cols):
                df_le_bu = df
                break
    if df_le_bu is None:
        raise ValueError("Could not find a sheet with columns: Ledger, LegalEntityName, BusinessUnitName.")

    # 2) Legal Entities dump
    df_le = _best_sheet(
        sheets,
        ["manage legal entities", "legal entities", "legalentities", "le", "man legal entities"]
    )
    if df_le is None:
        # fallback: look for Identifier + Name
        for df in sheets.values():
            cols = {_norm(c) for c in df.columns}
            if {"identifier", "name"}.issubset(cols):
                df_le = df
                break
    if df_le is None:
        raise ValueError("Could not find a Legal Entities sheet (needs columns: Identifier, Name, [Ledger optional]).")

    # 3) Cost Organizations dump
    df_co = _best_sheet(
        sheets,
        ["cst_cost_organization", "cost organization", "manage cost organizations", "cost org"]
    )
    if df_co is None:
        # fallback by columns
        for df in sheets.values():
            cols = {_norm(c) for c in df.columns}
            if {"name", "legalentityidentifier"}.issubset(cols):
                df_co = df
                break
    if df_co is None:
        raise ValueError("Could not find CST_COST_ORGANIZATION sheet (needs columns: Name, LegalEntityIdentifier).")

    # --- standardize column names
    # le-bu
    df_le_bu = df_le_bu.rename(columns={
        next(c for c in df_le_bu.columns if _norm(c) == "ledger"): "Ledger",
        next(c for c in df_le_bu.columns if _norm(c) == "legalentityname"): "LegalEntityName",
        next(c for c in df_le_bu.columns if _norm(c) == "businessunitname"): "BusinessUnitName",
    })[["Ledger", "LegalEntityName", "BusinessUnitName"]].astype(str)

    # legal entities
    le_map = {"Identifier": None, "Name": None}
    for c in df_le.columns:
        if _norm(c) == "identifier": le_map["Identifier"] = c
        if _norm(c) == "name":       le_map["Name"] = c
    if not all(le_map.values()):
        raise ValueError("Legal Entities sheet is missing Identifier/Name columns.")
    df_le_std = df_le.rename(columns={le_map["Identifier"]: "Identifier", le_map["Name"]: "Name"}).astype(str)

    # try to find ledger in LE sheet; else derive from le-bu
    ledger_col = None
    for c in df_le_std.columns:
        if _norm(c) in {"ledger", "ledgername", "primaryledger", "ledger_name"}:
            ledger_col = c
            break
    if ledger_col and ledger_col != "Ledger":
        df_le_std = df_le_std.rename(columns={ledger_col: "Ledger"})
    if "Ledger" not in df_le_std.columns:
        # derive from LE name as seen in ES – Ledger–LE–BU
        le_to_ledger = (df_le_bu.dropna(subset=["LegalEntityName", "Ledger"])
                        .groupby("LegalEntityName")["Ledger"]
                        .agg(lambda s: s.mode().iat[0] if not s.mode().empty else s.iloc[0]))
        df_le_std["Ledger"] = df_le_std["Name"].map(le_to_ledger).fillna("")

    # cost orgs
    co_map = {"Name": None, "LegalEntityIdentifier": None}
    for c in df_co.columns:
        if _norm(c) == "name": co_map["Name"] = c
        if _norm(c) == "legalentityidentifier": co_map["LegalEntityIdentifier"] = c
    if not all(co_map.values()):
        raise ValueError("CST_COST_ORGANIZATION sheet must have Name and LegalEntityIdentifier.")
    df_co_std = df_co.rename(columns={
        co_map["Name"]: "CostOrganization",
        co_map["LegalEntityIdentifier"]: "LegalEntityIdentifier"
    })[["CostOrganization", "LegalEntityIdentifier"]].astype(str)

    return df_le_bu, df_le_std[["Identifier", "Name", "Ledger"]], df_co_std

def build_costorg_tab(df_le_bu: pd.DataFrame, df_le: pd.DataFrame, df_co: pd.DataFrame) -> pd.DataFrame:
    le_lookup = df_le.rename(columns={"Identifier": "LegalEntityIdentifier",
                                      "Name": "LegalEntityName"})[["LegalEntityIdentifier", "LegalEntityName", "Ledger"]]
    out = (df_co.merge(le_lookup, on="LegalEntityIdentifier", how="left")
                .rename(columns={"CostOrganization": "CostOrganization"}))
    out["__WARN_UnmappedLE"] = out["LegalEntityName"].isna() | (out["LegalEntityName"] == "")
    out = out[["Ledger", "LegalEntityName", "CostOrganization", "LegalEntityIdentifier", "__WARN_UnmappedLE"]]
    out = out.sort_values(["Ledger", "LegalEntityName", "CostOrganization"], na_position="last").reset_index(drop=True)
    return out

def write_sheet_into_wb(xlsx_bytes: bytes, df: pd.DataFrame, sheet_name: str) -> bytes:
    bio = io.BytesIO(xlsx_bytes)
    try:
        wb = load_workbook(bio)
    except Exception:
        wb = Workbook()
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        wb.remove(ws)
    ws = wb.create_sheet(sheet_name) if wb.worksheets else wb.active
    if wb.worksheets and ws.title != sheet_name:
        ws.title = sheet_name

    ws.append(list(df.columns))
    for _, r in df.iterrows():
        ws.append(list(r.values))
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()

# ---------------- Diagram ----------------

X_LEDGER = 40
X_LE     = 320
X_BU     = 650
X_CO     = 950

Y_START_LEDGER = 40
Y_START_LE     = 40
Y_START_BU     = 40
Y_START_CO     = 220  # cost orgs lower than BU
Y_STEP = 90
NODE_W, NODE_H = 170, 48

STYLE_NODE_LEDGER = "rounded=1;fillColor=#F5F5F5;strokeColor=#666666;fontStyle=1;"
STYLE_NODE_LE     = "rounded=1;fillColor=#FFFFFF;strokeColor=#222222;"
STYLE_NODE_BU     = "rounded=1;fillColor=#FFFFFF;strokeColor=#888888;"
STYLE_NODE_CO     = "rounded=1;fillColor=#E9F2FF;strokeColor=#1F75FE;"

STYLE_EDGE_BASE = "endArrow=block;endFill=1;rounded=1;jettySize=auto;orthogonalLoop=1;edgeStyle=orthogonalEdgeStyle;curved=1;jumpStyle=arc;jumpSize=10;"
STYLE_LEDGER_TO_LE = STYLE_EDGE_BASE + "strokeColor=#666666;strokeWidth=2;"
STYLE_LE_TO_BU     = STYLE_EDGE_BASE + "strokeColor=#FFD400;strokeWidth=2;"  # yellow
STYLE_LE_TO_CO     = STYLE_EDGE_BASE + "strokeColor=#1F75FE;strokeWidth=2;"  # blue

def build_drawio(df_le_bu: pd.DataFrame, df_cost: pd.DataFrame) -> str:
    g = GraphBuilder(name="Enterprise Structure (+ Cost Orgs)")

    # Ledgers
    ledgers = sorted(set(df_le_bu["Ledger"].dropna()) | set(df_cost["Ledger"].dropna()))
    ledger_ids, y = {}, Y_START_LEDGER
    for L in ledgers:
        lbl = L if L else "(No Ledger)"
        nid = g.add_node(lbl, X_LEDGER, y, NODE_W, NODE_H, STYLE_NODE_LEDGER)
        ledger_ids[lbl] = nid
        y += Y_STEP

    # LEs (from both frames)
    df_le_all = (pd.concat([df_le_bu[["Ledger", "LegalEntityName"]],
                            df_cost[["Ledger", "LegalEntityName"]]], ignore_index=True)
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

    # BUs
    bu_ids = {}
    if not df_le_bu.empty:
        for (L, LE), grp in df_le_bu.groupby(["Ledger", "LegalEntityName"], dropna=False):
            y_bu = Y_START_BU
            for bu in sorted(set(grp["BusinessUnitName"].dropna())):
                bl = bu if bu else "(Unnamed BU)"
                nid = g.add_node(bl, X_BU, y_bu, NODE_W, NODE_H, STYLE_NODE_BU)
                bu_ids[(L or "", LE or "", bl)] = nid
                g.add_edge(le_ids[(L or "", LE or "(Unnamed LE)")], nid, STYLE_LE_TO_BU)
                y_bu += Y_STEP

    # Cost Orgs (lower vertical)
    if not df_cost.empty:
        for (L, LE), grp in df_cost.groupby(["Ledger", "LegalEntityName"], dropna=False):
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

st.title("Enterprise Structure Generator — OG Input with Cost Orgs")
st.caption("Upload your original ES workbook (.xlsx). I’ll add **ES – Ledger–LE–CostOrg** and generate a .drawio with yellow LE→BU and blue LE→Cost Org. Cost Orgs sit lower than BUs.")

with st.sidebar:
    xlsx = st.file_uploader("ES Workbook (.xlsx)", type=["xlsx"])
    run = st.button("Build New Tab + Diagram")

st.markdown("**Expected to be present in your workbook (names can vary):**")
st.markdown("- An OG tab with columns **Ledger, LegalEntityName, BusinessUnitName** (e.g., *ES – Ledger–LE–BU*)\n- A **Legal Entities** tab with **Identifier, Name** (Ledger optional)\n- A **Cost Organizations** tab with **Name, LegalEntityIdentifier** (e.g., *CST_COST_ORGANIZATION*)")

if run:
    if not xlsx:
        st.error("Upload your ES workbook first.")
        st.stop()
    try:
        wb_bytes = xlsx.read()
        sheets = read_es_workbook(wb_bytes)
        df_le_bu, df_le, df_co = get_frames_from_og_wb(sheets)

        df_cost_tab = build_costorg_tab(df_le_bu, df_le, df_co)
        new_wb_bytes = write_sheet_into_wb(wb_bytes, df_cost_tab, "ES – Ledger–LE–CostOrg")

        # Diagram
        drawio_xml = build_drawio(df_le_bu, df_cost_tab)

        st.success("✅ Built new tab and diagram.")

        col1, col2 = st.columns(2)
        with col1:
            st.download_button("⬇️ Download updated workbook",
                               data=new_wb_bytes,
                               file_name="enterprise_structure.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col2:
            st.download_button("⬇️ Download .drawio diagram",
                               data=drawio_xml.encode("utf-8"),
                               file_name="enterprise_structure.drawio",
                               mime="application/xml")

        with st.expander("Preview: ES – Ledger–LE–CostOrg (first 50 rows)"):
            st.dataframe(df_cost_tab.head(50))

        with st.expander("Diagram XML (first ~60 lines)"):
            st.code("\n".join(drawio_xml.splitlines()[:60]), language="xml")

        # Quick sanity counts
        unmapped = int(df_cost_tab["__WARN_UnmappedLE"].sum())
        st.info(f"Unmapped Cost Orgs (no LE name found): **{unmapped}**")

    except Exception as e:
        st.exception(e)
