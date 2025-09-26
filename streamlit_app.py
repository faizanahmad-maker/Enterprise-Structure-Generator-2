import io
import zlib
import base64
import uuid
import zipfile
from typing import List, Optional, Tuple

import pandas as pd
import streamlit as st
from openpyxl import Workbook
import xml.etree.ElementTree as ET


st.set_page_config(page_title="ES Generator — Multi-ZIP (OG filenames) + Cost Orgs", layout="wide")


# =========================
# Small helpers
# =========================

def _new_id() -> str:
    return str(uuid.uuid4()).replace("-", "")[:12]

def _deflate_base64(xml_text: str) -> str:
    comp = zlib.compressobj(level=9, wbits=-15)  # raw DEFLATE
    data = comp.compress(xml_text.encode("utf-8")) + comp.flush()
    return base64.b64encode(data).decode("ascii")

def _read_csv_from_zip(z: zipfile.ZipFile, name: str) -> Optional[pd.DataFrame]:
    if name not in z.namelist():
        return None
    with z.open(name) as fh:
        return pd.read_csv(fh, dtype=str)


# =========================
# Draw.io (mxGraph) builder
# =========================

class GraphBuilder:
    def __init__(self, name="ES Diagram"):
        self.name = name
        self.cells = [{"id": "0"}, {"id": "1", "parent": "0"}]

    def add_node(self, label, x, y, w, h, style=""):
        cid = _new_id()
        self.cells.append({
            "id": cid, "value": label, "vertex": "1", "parent": "1",
            "style": f"whiteSpace=wrap;html=1;align=center;{style}",
            "geometry": {"x": x, "y": y, "width": w, "height": h}
        })
        return cid

    def add_edge(self, src, dst, style=""):
        eid = _new_id()
        self.cells.append({
            "id": eid, "edge": "1", "parent": "1",
            "style": style, "source": src, "target": dst
        })
        return eid

    def _mx(self) -> str:
        mx = ET.Element("mxGraphModel")
        root = ET.SubElement(mx, "root")
        for c in self.cells:
            attrs = {k: v for k, v in c.items() if k != "geometry"}
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
# OG filenames & columns
# =========================

FN_LEDGER_LIST          = "GL_PRIMARY_LEDGER.csv"                   # optional catalog
COL_LEDGER_LIST_NAME    = "ORA_GL_PRIMARY_LEDGER_CONFIG.Name"

FN_LE_PROFILE           = "XLE_ENTITY_PROFILE.csv"                  # optional catalog
COL_LE_PROFILE_NAME     = "Name"

FN_IDENT_TO_LEDGER      = "ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv"    # required
COL_IDENT_LEDGER_NAME   = "GL_LEDGER.Name"
COL_IDENT_IDENTIFIER    = "LegalEntityIdentifier"

FN_IDENT_TO_LENAME      = "ORA_GL_JOURNAL_CONFIG_DETAIL.csv"        # required
COL_JCFG_IDENTIFIER     = "LegalEntityIdentifier"
COL_JCFG_LENAME         = "ObjectName"

FN_BUSINESS_UNITS       = "FUN_BUSINESS_UNIT.csv"                    # optional
COL_BU_NAME             = "Name"
COL_BU_LEDGER           = "PrimaryLedgerName"
COL_BU_LENAME           = "LegalEntityName"

FN_COST_ORGS            = "CST_COST_ORGANIZATION.csv"               # required
COL_CO_NAME             = "Name"
COL_CO_IDENTIFIER       = "LegalEntityIdentifier"


# =========================
# Load from multiple ZIPs
# =========================

def load_multi_zips(files: List[bytes]) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, set, set]:
    """
    Returns:
      df_le_bu (Ledger, LegalEntityName, BusinessUnitName)  [BU optional]
      df_ident_ledger (LegalEntityIdentifier, Ledger)
      df_ident_lename (LegalEntityIdentifier, LegalEntityName)
      ledger_names_set, le_names_set (catalogs; optional)
    """
    bu_parts, ident_ledger_parts, ident_lename_parts = [], [], []
    ledger_names, le_names = set(), set()

    for blob in files:
        with zipfile.ZipFile(io.BytesIO(blob)) as z:
            # ledger catalog (optional)
            df = _read_csv_from_zip(z, FN_LEDGER_LIST)
            if df is not None and COL_LEDGER_LIST_NAME in df.columns:
                ledger_names |= set(df[COL_LEDGER_LIST_NAME].dropna().map(str).str.strip())

            # LE catalog (optional)
            df = _read_csv_from_zip(z, FN_LE_PROFILE)
            if df is not None and COL_LE_PROFILE_NAME in df.columns:
                le_names |= set(df[COL_LE_PROFILE_NAME].dropna().map(str).str.strip())

            # Identifier → Ledger (required)
            df = _read_csv_from_zip(z, FN_IDENT_TO_LEDGER)
            if df is not None:
                need = {COL_IDENT_LEDGER_NAME, COL_IDENT_IDENTIFIER}
                if need.issubset(set(df.columns)):
                    ident_ledger_parts.append(
                        df[list(need)].rename(columns={
                            COL_IDENT_LEDGER_NAME: "Ledger",
                            COL_IDENT_IDENTIFIER: "LegalEntityIdentifier"
                        })
                    )

            # Identifier → LE Name (required)
            df = _read_csv_from_zip(z, FN_IDENT_TO_LENAME)
            if df is not None:
                need = {COL_JCFG_IDENTIFIER, COL_JCFG_LENAME}
                if need.issubset(set(df.columns)):
                    ident_lename_parts.append(
                        df[list(need)].rename(columns={
                            COL_JCFG_IDENTIFIER: "LegalEntityIdentifier",
                            COL_JCFG_LENAME: "LegalEntityName"
                        })
                    )

            # Business Units (optional)
            df = _read_csv_from_zip(z, FN_BUSINESS_UNITS)
            if df is not None:
                need = {COL_BU_NAME, COL_BU_LEDGER, COL_BU_LENAME}
                if need.issubset(set(df.columns)):
                    tmp = df[[COL_BU_NAME, COL_BU_LEDGER, COL_BU_LENAME]].copy()
                    tmp.columns = ["BusinessUnitName", "Ledger", "LegalEntityName"]
                    bu_parts.append(tmp)

    if not ident_ledger_parts:
        raise ValueError(f"None of the ZIPs had a valid `{FN_IDENT_TO_LEDGER}` "
                         f"with `{COL_IDENT_LEDGER_NAME}`, `{COL_IDENT_IDENTIFIER}`.")
    if not ident_lename_parts:
        raise ValueError(f"None of the ZIPs had a valid `{FN_IDENT_TO_LENAME}` "
                         f"with `{COL_JCFG_IDENTIFIER}`, `{COL_JCFG_LENAME}`.")

    df_ident_ledger = (pd.concat(ident_ledger_parts, ignore_index=True)
                       .dropna(subset=["LegalEntityIdentifier"])
                       .drop_duplicates().reset_index(drop=True))
    df_ident_lename = (pd.concat(ident_lename_parts, ignore_index=True)
                       .dropna(subset=["LegalEntityIdentifier"])
                       .drop_duplicates().reset_index(drop=True))

    if bu_parts:
        df_le_bu = (pd.concat(bu_parts, ignore_index=True)
                    .drop_duplicates().reset_index(drop=True))
    else:
        # minimal frame: all LE names + ledgers, BU column empty
        tmp = (df_ident_lename.merge(df_ident_ledger, on="LegalEntityIdentifier", how="outer")
               [["Ledger", "LegalEntityName"]].drop_duplicates())
        tmp["BusinessUnitName"] = pd.NA
        df_le_bu = tmp

    return df_le_bu, df_ident_ledger, df_ident_lename, ledger_names, le_names


def load_cost_orgs_from_zips(files: List[bytes]) -> pd.DataFrame:
    parts = []
    for blob in files:
        with zipfile.ZipFile(io.BytesIO(blob)) as z:
            df = _read_csv_from_zip(z, FN_COST_ORGS)
            if df is not None:
                need = {COL_CO_NAME, COL_CO_IDENTIFIER}
                if need.issubset(set(df.columns)):
                    tmp = df[[COL_CO_NAME, COL_CO_IDENTIFIER]].copy()
                    tmp.columns = ["CostOrganization", "LegalEntityIdentifier"]
                    parts.append(tmp)
    if not parts:
        raise ValueError(f"None of the ZIPs had `{FN_COST_ORGS}` with `{COL_CO_NAME}`, `{COL_CO_IDENTIFIER}`.")
    return (pd.concat(parts, ignore_index=True)
            .dropna(subset=["LegalEntityIdentifier"])
            .drop_duplicates().reset_index(drop=True))


# =========================
# Build ES – Ledger–LE–CostOrg
# =========================

def build_costorg_tab(
    df_cost: pd.DataFrame,
    df_ident_lename: pd.DataFrame,
    df_ident_ledger: pd.DataFrame
) -> pd.DataFrame:
    out = (df_cost
           .merge(df_ident_lename, on="LegalEntityIdentifier", how="left")
           .merge(df_ident_ledger, on="LegalEntityIdentifier", how="left"))
    out["Ledger"] = out["Ledger"].fillna("")
    out["LegalEntityName"] = out["LegalEntityName"].fillna("")
    out["__WARN_UnmappedLE"] = (out["LegalEntityName"] == "")
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
# Diagram (with robust normalization)
# =========================

def build_drawio(df_le_bu: pd.DataFrame, df_cost_tab: pd.DataFrame) -> str:
    # ---- normalize to safe strings ----
    def norm(v) -> str:
        if pd.isna(v):
            return ""
        s = str(v).strip()
        return "" if s.lower() in ("nan", "none", "null") else s

    for df, cols in [
        (df_le_bu, ["Ledger", "LegalEntityName", "BusinessUnitName"]),
        (df_cost_tab, ["Ledger", "LegalEntityName", "CostOrganization"])
    ]:
        for c in cols:
            if c in df.columns:
                df[c] = df[c].map(norm)

    # ---- graph styles & layout ----
    X_LEDGER, X_LE, X_BU, X_CO = 40, 320, 650, 950
    Y_LEDGER, Y_LE, Y_BU, Y_CO = 40, 40, 40, 220
    Y_STEP, W, H = 90, 170, 48

    S_NODE_LEDGER = "rounded=1;fillColor=#F5F5F5;strokeColor=#666666;fontStyle=1;"
    S_NODE_LE     = "rounded=1;fillColor=#FFFFFF;strokeColor=#222222;"
    S_NODE_BU     = "rounded=1;fillColor=#FFFFFF;strokeColor=#888888;"
    S_NODE_CO     = "rounded=1;fillColor=#E9F2FF;strokeColor=#1F75FE;"

    S_EDGE_BASE   = ("endArrow=block;endFill=1;rounded=1;jettySize=auto;"
                     "orthogonalLoop=1;edgeStyle=orthogonalEdgeStyle;curved=1;"
                     "jumpStyle=arc;jumpSize=10;")
    S_LEDGER_LE   = S_EDGE_BASE + "strokeColor=#666666;strokeWidth=2;"
    S_LE_BU       = S_EDGE_BASE + "strokeColor=#FFD400;strokeWidth=2;"  # yellow
    S_LE_CO       = S_EDGE_BASE + "strokeColor=#1F75FE;strokeWidth=2;"  # blue

    g = GraphBuilder(name="Enterprise Structure (+ Cost Orgs)")

    # Ledgers present anywhere
    ledgers = sorted({norm(x) for x in pd.concat([df_le_bu["Ledger"], df_cost_tab["Ledger"]], ignore_index=True)})
    if not ledgers:
        ledgers = [""]

    # ledger nodes
    ledger_ids, y = {}, Y_LEDGER
    for L in ledgers:
        lbl = L if L else "(No Ledger)"
        ledger_ids[L] = g.add_node(lbl, X_LEDGER, y, W, H, S_NODE_LEDGER)
        y += Y_STEP

    # unified LE list
    df_le_all = (pd.concat([
                    df_le_bu[["Ledger", "LegalEntityName"]],
                    df_cost_tab[["Ledger", "LegalEntityName"]]
                ], ignore_index=True)
                .dropna(how="all")
                .drop_duplicates()
                .sort_values(["Ledger", "LegalEntityName"], na_position="last"))

    le_ids, last_L, y_le = {}, None, Y_LE
    for _, r in df_le_all.iterrows():
        L = norm(r["Ledger"])
        E = norm(r["LegalEntityName"])
        if L != last_L:
            y_le = Y_LE
            last_L = L
        E_label = E if E else "(Unnamed LE)"
        nid = g.add_node(E_label, X_LE, y_le, W, H, S_NODE_LE)
        le_ids[(L, E_label)] = nid
        g.add_edge(ledger_ids.get(L, ledger_ids.get("", next(iter(ledger_ids.values())))), nid, S_LEDGER_LE)
        y_le += Y_STEP

    # BUs (only if parent LE exists)
    if "BusinessUnitName" in df_le_bu.columns and df_le_bu["BusinessUnitName"].notna().any():
        for (L_raw, E_raw), grp in df_le_bu.groupby(["Ledger", "LegalEntityName"], dropna=False):
            L, E = norm(L_raw), norm(E_raw)
            parent_key = (L, E if E else "(Unnamed LE)")
            if parent_key not in le_ids:
                continue
            y_bu = Y_BU
            for bu in sorted({norm(x) for x in grp["BusinessUnitName"] if norm(x)}):
                nid = g.add_node(bu, X_BU, y_bu, W, H, S_NODE_BU)
                g.add_edge(le_ids[parent_key], nid, S_LE_BU)
                y_bu += Y_STEP

    # Cost Orgs (lower; only if parent LE exists)
    for (L_raw, E_raw), grp in df_cost_tab.groupby(["Ledger", "LegalEntityName"], dropna=False):
        L, E = norm(L_raw), norm(E_raw)
        parent_key = (L, E if E else "(Unnamed LE)")
        if parent_key not in le_ids:
            continue
        y_co = Y_CO
        for co in sorted({norm(x) for x in grp["CostOrganization"] if norm(x)}):
            nid = g.add_node(co, X_CO, y_co, W, H, S_NODE_CO)
            g.add_edge(le_ids[parent_key], nid, S_LE_CO)
            y_co += Y_STEP

    return g.to_drawio_xml()


# =========================
# UI
# =========================

st.title("Enterprise Structure Generator — Multi-ZIP (OG filenames) with Cost Orgs")
st.caption(
    "Drop one or more Oracle export ZIPs. I’ll use **CST_COST_ORGANIZATION.csv**, "
    "**ORA_GL_JOURNAL_CONFIG_DETAIL.csv**, **ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv** "
    "to build **ES – Ledger–LE–CostOrg**, and (optionally) **FUN_BUSINESS_UNIT.csv** for BUs. "
    "Diagram: **yellow** LE→BU, **blue** LE→Cost Org (with bridge jumps). Cost Orgs sit lower than BUs."
)

with st.sidebar:
    zip_files = st.file_uploader("Oracle Export ZIPs", type=["zip"], accept_multiple_files=True)
    run = st.button("Build Tab + Diagram")

st.markdown("""
**I look for these exact OG files inside your ZIPs (case-sensitive names are fine):**
- `CST_COST_ORGANIZATION.csv` → **Name**, **LegalEntityIdentifier** *(required)*
- `ORA_GL_JOURNAL_CONFIG_DETAIL.csv` → **LegalEntityIdentifier**, **ObjectName** *(required for LE Names)*
- `ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv` → **GL_LEDGER.Name**, **LegalEntityIdentifier** *(required for Ledgers)*
- `FUN_BUSINESS_UNIT.csv` → **Name**, **PrimaryLedgerName**, **LegalEntityName** *(optional for BU layer)*
- `GL_PRIMARY_LEDGER.csv` / `XLE_ENTITY_PROFILE.csv` *(optional catalogs)*
""")

if run:
    if not zip_files:
        st.error("Upload at least one ZIP.")
        st.stop()

    try:
        blobs = [f.read() for f in zip_files]

        # Load core data
        df_le_bu, df_ident_ledger, df_ident_lename, _, _ = load_multi_zips(blobs)
        df_cost = load_cost_orgs_from_zips(blobs)

        # Build target tab
        df_cost_tab = build_costorg_tab(df_cost, df_ident_lename, df_ident_ledger)

        # Excel (single sheet for this increment)
        xlsx_bytes = dataframe_to_xlsx_bytes(df_cost_tab, "ES – Ledger–LE–CostOrg")

        # Diagram
        drawio_xml = build_drawio(df_le_bu, df_cost_tab)

        st.success("✅ Built **ES – Ledger–LE–CostOrg** and the diagram from OG-named CSVs.")

        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "⬇️ Download Excel (ES – Ledger–LE–CostOrg.xlsx)",
                data=xlsx_bytes,
                file_name="enterprise_structure.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with c2:
            st.download_button(
                "⬇️ Download Diagram (.drawio)",
                data=drawio_xml.encode("utf-8"),
                file_name="enterprise_structure.drawio",
                mime="application/xml"
            )

        with st.expander("Preview: ES – Ledger–LE–CostOrg (first 50 rows)"):
            st.dataframe(df_cost_tab.head(50))

        unmapped = int(df_cost_tab["__WARN_UnmappedLE"].sum())
        st.info(f"Unmapped Cost Orgs (no LE name found): **{unmapped}**")

        with st.expander("Diagram XML (first ~60 lines)"):
            st.code("\n".join(drawio_xml.splitlines()[:60]), language="xml")

    except Exception as e:
        st.exception(e)
