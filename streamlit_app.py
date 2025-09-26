import io, zipfile, uuid, zlib, base64
from typing import List, Optional, Tuple

import pandas as pd
import streamlit as st
from openpyxl import Workbook
import xml.etree.ElementTree as ET

st.set_page_config(page_title="ES Generator — Two Tabs + Diagram", layout="wide")
st.title("Enterprise Structure Generator — Two Tabs + Diagram")

# -----------------------------
# OG filenames / columns
# -----------------------------
FN_LEDGER_LIST        = "GL_PRIMARY_LEDGER.csv"                   # optional catalog
COL_LEDGER_LIST_NAME  = "ORA_GL_PRIMARY_LEDGER_CONFIG.Name"

FN_LE_PROFILE         = "XLE_ENTITY_PROFILE.csv"                  # optional catalog
COL_LE_PROFILE_NAME   = "Name"

FN_IDENT_TO_LEDGER    = "ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv"    # required for ledger from identifier
COL_IDENT_LEDGER_NAME = "GL_LEDGER.Name"
COL_IDENT_IDENTIFIER  = "LegalEntityIdentifier"

FN_IDENT_TO_LENAME    = "ORA_GL_JOURNAL_CONFIG_DETAIL.csv"        # required for LE name from identifier
COL_JCFG_IDENTIFIER   = "LegalEntityIdentifier"
COL_JCFG_LENAME       = "ObjectName"

FN_BUSINESS_UNITS     = "FUN_BUSINESS_UNIT.csv"                    # optional for BU layer
COL_BU_NAME           = "Name"
COL_BU_LEDGER         = "PrimaryLedgerName"
COL_BU_LENAME         = "LegalEntityName"

FN_COST_ORGS          = "CST_COST_ORGANIZATION.csv"               # required for cost orgs
COL_CO_NAME           = "Name"
COL_CO_IDENTIFIER     = "LegalEntityIdentifier"

# -----------------------------
# Helpers
# -----------------------------
def _read_csv_from_zip(z: zipfile.ZipFile, name: str) -> Optional[pd.DataFrame]:
    if name not in z.namelist():
        return None
    with z.open(name) as fh:
        return pd.read_csv(fh, dtype=str)

def _norm(s) -> str:
    if pd.isna(s): return ""
    x = str(s).strip()
    return "" if x.lower() in ("nan", "none", "null") else x

# -----------------------------
# Build ES – Ledger–LE–BU (reuse your OG logic)
# -----------------------------
def build_ledger_le_bu_from_zips(zip_bytes_list: List[bytes]) -> pd.DataFrame:
    ledger_names = set()
    legal_entity_names = set()
    ledger_to_idents = {}   # ledger -> {ident}
    ident_to_le_name = {}   # ident -> LE name
    bu_rows = []            # Name, PrimaryLedgerName, LegalEntityName

    for blob in zip_bytes_list:
        with zipfile.ZipFile(io.BytesIO(blob)) as z:
            # Ledgers
            df = _read_csv_from_zip(z, FN_LEDGER_LIST)
            if df is not None and COL_LEDGER_LIST_NAME in df.columns:
                ledger_names |= set(df[COL_LEDGER_LIST_NAME].dropna().map(str).str.strip())

            # LE list
            df = _read_csv_from_zip(z, FN_LE_PROFILE)
            if df is not None and COL_LE_PROFILE_NAME in df.columns:
                legal_entity_names |= set(df[COL_LE_PROFILE_NAME].dropna().map(str).str.strip())

            # Identifier ↔ Ledger
            df = _read_csv_from_zip(z, FN_IDENT_TO_LEDGER)
            if df is not None and {COL_IDENT_LEDGER_NAME, COL_IDENT_IDENTIFIER}.issubset(df.columns):
                for _, r in df[[COL_IDENT_LEDGER_NAME, COL_IDENT_IDENTIFIER]].dropna().iterrows():
                    led  = _norm(r[COL_IDENT_LEDGER_NAME])
                    ident = _norm(r[COL_IDENT_IDENTIFIER])
                    if led and ident:
                        ledger_to_idents.setdefault(led, set()).add(ident)

            # Identifier ↔ LE name
            df = _read_csv_from_zip(z, FN_IDENT_TO_LENAME)
            if df is not None and {COL_JCFG_IDENTIFIER, COL_JCFG_LENAME}.issubset(df.columns):
                for _, r in df[[COL_JCFG_IDENTIFIER, COL_JCFG_LENAME]].dropna().iterrows():
                    ident = _norm(r[COL_JCFG_IDENTIFIER])
                    name  = _norm(r[COL_JCFG_LENAME])
                    if ident:
                        ident_to_le_name[ident] = name

            # Business Units
            df = _read_csv_from_zip(z, FN_BUSINESS_UNITS)
            if df is not None and {COL_BU_NAME, COL_BU_LEDGER, COL_BU_LENAME}.issubset(df.columns):
                for _, r in df[[COL_BU_NAME, COL_BU_LEDGER, COL_BU_LENAME]].iterrows():
                    bu = _norm(r[COL_BU_NAME])
                    led = _norm(r[COL_BU_LEDGER])
                    le  = _norm(r[COL_BU_LENAME])
                    bu_rows.append({"Name": bu, "PrimaryLedgerName": led, "LegalEntityName": le})

    # Map ledger -> LE names via identifiers
    ledger_to_le_names = {}
    for led, idents in ledger_to_idents.items():
        for ident in idents:
            le_name = ident_to_le_name.get(ident, "").strip()
            if le_name:
                ledger_to_le_names.setdefault(led, set()).add(le_name)

    le_to_ledgers = {}
    for led, les in ledger_to_le_names.items():
        for le in les:
            le_to_ledgers.setdefault(le, set()).add(led)

    # Build final rows (OG approach)
    rows = []
    seen_triples = set()
    seen_ledgers_with_bu = set()
    seen_les_with_bu = set()

    # 1) BU-driven rows with smart back-fill
    for r in bu_rows:
        bu = r["Name"]
        led = r["PrimaryLedgerName"] if r["PrimaryLedgerName"] in ledger_names else ""
        le  = r["LegalEntityName"]  if r["LegalEntityName"]  in legal_entity_names else ""

        if not led and le and le in le_to_ledgers and len(le_to_ledgers[le]) == 1:
            led = next(iter(le_to_ledgers[le]))
        if not le and led and led in ledger_to_le_names and len(ledger_to_le_names[led]) == 1:
            le = next(iter(ledger_to_le_names[led]))

        rows.append({"Ledger": led, "LegalEntityName": le, "BusinessUnitName": bu})
        seen_triples.add((led, le, bu))
        if led: seen_ledgers_with_bu.add(led)
        if le:  seen_les_with_bu.add(le)

    # 2) Ledger–LE pairs with no BU
    seen_pairs = {(a, b) for (a, b, _) in seen_triples}
    for led, le_set in ledger_to_le_names.items():
        if not le_set:
            if led not in seen_ledgers_with_bu:
                rows.append({"Ledger": led, "LegalEntityName": "", "BusinessUnitName": ""})
            continue
        for le in le_set:
            if (led, le) not in seen_pairs:
                rows.append({"Ledger": led, "LegalEntityName": le, "BusinessUnitName": ""})

    # 3) Orphan ledgers with no mapping & no BU
    for led in sorted(ledger_names - set(ledger_to_le_names.keys()) - seen_ledgers_with_bu):
        rows.append({"Ledger": led, "LegalEntityName": "", "BusinessUnitName": ""})

    # 4) Orphan LEs with no BU; back-fill ledger if uniquely known
    for le in sorted(legal_entity_names - seen_les_with_bu):
        if le in le_to_ledgers and len(le_to_ledgers[le]) == 1:
            led = next(iter(le_to_ledgers[le]))
        else:
            led = ""
        rows.append({"Ledger": led, "LegalEntityName": le, "BusinessUnitName": ""})

    df = pd.DataFrame(rows).drop_duplicates().reset_index(drop=True)
    # sort
    df["__empty"] = (df["Ledger"] == "").astype(int)
    df = df.sort_values(["__empty","Ledger","LegalEntityName","BusinessUnitName"]).drop(columns="__empty").reset_index(drop=True)
    return df

# -----------------------------
# Build ES – Ledger–LE–CostOrg
# -----------------------------
def build_ledger_le_costorg_from_zips(zip_bytes_list: List[bytes]) -> pd.DataFrame:
    # Collect required frames
    parts_cost = []
    parts_ident_ledger = []
    parts_ident_lename = []

    for blob in zip_bytes_list:
        with zipfile.ZipFile(io.BytesIO(blob)) as z:
            df = _read_csv_from_zip(z, FN_COST_ORGS)
            if df is not None and {COL_CO_NAME, COL_CO_IDENTIFIER}.issubset(df.columns):
                tmp = df[[COL_CO_NAME, COL_CO_IDENTIFIER]].copy()
                tmp.columns = ["CostOrganization","LegalEntityIdentifier"]
                parts_cost.append(tmp)

            df = _read_csv_from_zip(z, FN_IDENT_TO_LEDGER)
            if df is not None and {COL_IDENT_LEDGER_NAME, COL_IDENT_IDENTIFIER}.issubset(df.columns):
                tmp = df[[COL_IDENT_LEDGER_NAME, COL_IDENT_IDENTIFIER]].copy()
                tmp.columns = ["Ledger","LegalEntityIdentifier"]
                parts_ident_ledger.append(tmp)

            df = _read_csv_from_zip(z, FN_IDENT_TO_LENAME)
            if df is not None and {COL_JCFG_IDENTIFIER, COL_JCFG_LENAME}.issubset(df.columns):
                tmp = df[[COL_JCFG_IDENTIFIER, COL_JCFG_LENAME]].copy()
                tmp.columns = ["LegalEntityIdentifier","LegalEntityName"]
                parts_ident_lename.append(tmp)

    if not parts_cost:
        raise ValueError(f"Missing `{FN_COST_ORGS}` with `{COL_CO_NAME}`, `{COL_CO_IDENTIFIER}` in your ZIPs.")
    if not parts_ident_ledger:
        raise ValueError(f"Missing `{FN_IDENT_TO_LEDGER}` with `{COL_IDENT_LEDGER_NAME}`, `{COL_IDENT_IDENTIFIER}`.")
    if not parts_ident_lename:
        raise ValueError(f"Missing `{FN_IDENT_TO_LENAME}` with `{COL_JCFG_IDENTIFIER}`, `{COL_JCFG_LENAME}`.")

    df_cost = pd.concat(parts_cost, ignore_index=True).drop_duplicates()
    df_il   = pd.concat(parts_ident_ledger, ignore_index=True).drop_duplicates()
    df_in   = pd.concat(parts_ident_lename, ignore_index=True).drop_duplicates()

    out = (df_cost.merge(df_in, on="LegalEntityIdentifier", how="left")
                  .merge(df_il, on="LegalEntityIdentifier", how="left"))
    # Only the 3 columns you asked for
    out = out[["Ledger","LegalEntityName","CostOrganization"]].fillna("").sort_values(
        ["Ledger","LegalEntityName","CostOrganization"], na_position="last"
    ).reset_index(drop=True)
    return out

# -----------------------------
# Workbook writer (two tabs)
# -----------------------------
def two_tab_workbook_bytes(df_bu: pd.DataFrame, df_co: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "ES – Ledger–LE–BU"
    ws1.append(["Ledger","LegalEntityName","BusinessUnitName"])
    for _, r in df_bu.iterrows():
        ws1.append([_norm(r["Ledger"]), _norm(r["LegalEntityName"]), _norm(r["BusinessUnitName"])])

    ws2 = wb.create_sheet("ES – Ledger–LE–CostOrg")
    ws2.append(["Ledger","LegalEntityName","CostOrganization"])
    for _, r in df_co.iterrows():
        ws2.append([_norm(r["Ledger"]), _norm(r["LegalEntityName"]), _norm(r["CostOrganization"])])

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()

# -----------------------------
# Diagram (yellow BU, blue Cost Org)
# -----------------------------
def drawio_from_frames(df_bu: pd.DataFrame, df_co: pd.DataFrame) -> str:
    # normalize
    for df, cols in [(df_bu,["Ledger","LegalEntityName","BusinessUnitName"]),
                     (df_co,["Ledger","LegalEntityName","CostOrganization"])]:
        for c in cols: df[c] = df[c].map(_norm)

    X_LEDGER, X_LE, X_BU, X_CO = 40, 320, 650, 950
    Y_LEDGER, Y_LE, Y_BU, Y_CO = 40, 40, 40, 220
    STEP, W, H = 90, 170, 48

    S_NODE_LEDGER = "rounded=1;fillColor=#F5F5F5;strokeColor=#666666;fontStyle=1;"
    S_NODE_LE     = "rounded=1;fillColor=#FFFFFF;strokeColor=#222222;"
    S_NODE_BU     = "rounded=1;fillColor=#FFFFFF;strokeColor=#888888;"
    S_NODE_CO     = "rounded=1;fillColor=#E9F2FF;strokeColor=#1F75FE;"

    S_BASE = ("endArrow=block;endFill=1;rounded=1;jettySize=auto;orthogonalLoop=1;"
              "edgeStyle=orthogonalEdgeStyle;curved=1;jumpStyle=arc;jumpSize=10;")
    S_LEDGER_LE = S_BASE + "strokeColor=#666666;strokeWidth=2;"
    S_LE_BU     = S_BASE + "strokeColor=#FFD400;strokeWidth=2;"   # yellow
    S_LE_CO     = S_BASE + "strokeColor=#1F75FE;strokeWidth=2;"   # blue

    class G:
        def __init__(self, name="Enterprise Structure"):
            self.cells=[{"id":"0"},{"id":"1","parent":"0"}]; self.name=name
        def add_node(self,label,x,y,w,h,sty):
            i=uuid.uuid4().hex[:10]; self.cells.append({"id":i,"value":label,"vertex":"1","parent":"1",
                "style":f"whiteSpace=wrap;html=1;align=center;{sty}","geometry":{"x":x,"y":y,"width":w,"height":h}}); return i
        def add_edge(self,s,t,sty):
            i=uuid.uuid4().hex[:10]; self.cells.append({"id":i,"edge":"1","parent":"1","style":sty,"source":s,"target":t}); return i
        def xml(self):
            mx=ET.Element("mxGraphModel"); root=ET.SubElement(mx,"root")
            for c in self.cells:
                a={k:v for k,v in c.items() if k!="geometry"}; cell=ET.SubElement(root,"mxCell",a)
                if "geometry" in c:
                    g=c["geometry"]; ET.SubElement(cell,"mxGeometry",{"x":str(g["x"]),"y":str(g["y"]),
                        "width":str(g["width"]),"height":str(g["height"]),"as":"geometry"})
            enc=_deflate_base64(ET.tostring(mx,encoding="utf-8").decode("utf-8"))
            mxfile=ET.Element("mxfile",{"host":"app.diagrams.net"})
            diagram=ET.SubElement(mxfile,"diagram",{"name":self.name,"id":uuid.uuid4().hex[:12]})
            diagram.text=enc
            return ET.tostring(mxfile,encoding="utf-8",xml_declaration=True).decode("utf-8")

    g = G("Enterprise Structure (+ Cost Orgs)")

    ledgers = sorted(set(df_bu["Ledger"]) | set(df_co["Ledger"]))
    if not ledgers: ledgers = [""]

    ledger_ids = {}; y = Y_LEDGER
    for L in ledgers:
        lbl = L if L else "(No Ledger)"
        ledger_ids[L] = g.add_node(lbl, X_LEDGER, y, W, H, S_NODE_LEDGER); y += STEP

    # LE nodes
    df_le = (pd.concat([df_bu[["Ledger","LegalEntityName"]],
                        df_co[["Ledger","LegalEntityName"]]], ignore_index=True)
             .drop_duplicates().sort_values(["Ledger","LegalEntityName"]))
    le_ids = {}; last=None; y_le=Y_LE
    for _, r in df_le.iterrows():
        L, E = r["Ledger"], r["LegalEntityName"]; label = E if E else "(Unnamed LE)"
        if L != last: y_le = Y_LE; last = L
        nid = g.add_node(label, X_LE, y_le, W, H, S_NODE_LE)
        le_ids[(L,label)] = nid
        g.add_edge(ledger_ids.get(L, ledger_ids.get("", next(iter(ledger_ids.values())))), nid, S_LEDGER_LE)
        y_le += STEP

    # BUs
    if df_bu["BusinessUnitName"].notna().any():
        for (L,E), grp in df_bu.groupby(["Ledger","LegalEntityName"], dropna=False):
            parent = (L, E if E else "(Unnamed LE)")
            if parent not in le_ids: continue
            y_bu = Y_BU
            for bu in sorted({b for b in grp["BusinessUnitName"].map(_norm) if b}):
                bn = g.add_node(bu, X_BU, y_bu, W, H, S_NODE_BU)
                g.add_edge(le_ids[parent], bn, S_LE_BU)
                y_bu += STEP

    # Cost Orgs (lower)
    for (L,E), grp in df_co.groupby(["Ledger","LegalEntityName"], dropna=False):
        parent = (L, E if E else "(Unnamed LE)")
        if parent not in le_ids: continue
        y_co = Y_CO
        for co in sorted({c for c in grp["CostOrganization"].map(_norm) if c}):
            cn = g.add_node(co, X_CO, y_co, W, H, S_NODE_CO)
            g.add_edge(le_ids[parent], cn, S_LE_CO)
            y_co += STEP

    return g.xml()

def drawio_link_from_xml(xml: str) -> str:
    # URL payload = raw DEFLATE (no zlib header/footer) + base64
    raw = zlib.compress(xml.encode("utf-8"), level=9)[2:-4]
    b64 = base64.b64encode(raw).decode("ascii")
    return f"https://app.diagrams.net/?title=EnterpriseStructure.drawio#R{b64}"

# -----------------------------
# UI
# -----------------------------
with st.sidebar:
    zips = st.file_uploader("Oracle Export ZIPs", type=["zip"], accept_multiple_files=True)
    run = st.button("Build")

st.caption("I’ll output **one Excel** with exactly two tabs and a **diagram link**. "
           "Files I expect across your ZIPs: "
           "`GL_PRIMARY_LEDGER.csv`, `XLE_ENTITY_PROFILE.csv`, "
           "`ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv`, `ORA_GL_JOURNAL_CONFIG_DETAIL.csv`, "
           "`FUN_BUSINESS_UNIT.csv` (optional), `CST_COST_ORGANIZATION.csv`.")

if run:
    if not zips:
        st.error("Upload at least one ZIP.")
        st.stop()
    try:
        blobs = [f.read() for f in zips]

        # 1) Build both tabs
        df_bu = build_ledger_le_bu_from_zips(blobs)[["Ledger","LegalEntityName","BusinessUnitName"]]
        df_co = build_ledger_le_costorg_from_zips(blobs)[["Ledger","LegalEntityName","CostOrganization"]]

        # 2) Workbook (exactly two tabs)
        xlsx_bytes = two_tab_workbook_bytes(df_bu, df_co)

        # 3) Diagram + link
        drawio_xml = drawio_from_frames(df_bu, df_co)
        link = drawio_link_from_xml(drawio_xml)

        # Downloads + link
        c1, c2, c3 = st.columns([1,1,2])
        with c1:
            st.download_button("⬇️ Download Excel (2 tabs)",
                               data=xlsx_bytes,
                               file_name="EnterpriseStructure.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with c2:
            st.download_button("⬇️ Download Diagram (.drawio)",
                               data=drawio_xml.encode("utf-8"),
                               file_name="EnterpriseStructure.drawio",
                               mime="application/xml")
        with c3:
            st.success("🔗 Open in draw.io:")
            st.markdown(f"[**Open Diagram in diagrams.net**]({link})")

        # Optional previews (collapsed by default)
        with st.expander("Preview: ES – Ledger–LE–BU (first 30)"):
            st.dataframe(df_bu.head(30))
        with st.expander("Preview: ES – Ledger–LE–CostOrg (first 30)"):
            st.dataframe(df_co.head(30))

    except Exception as e:
        st.exception(e)
