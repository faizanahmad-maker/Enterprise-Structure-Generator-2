import io, zipfile, uuid, base64, zlib, xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Enterprise Structure Generator", page_icon="📊", layout="wide")
st.title("Enterprise Structure Generator — Two Tabs + Diagram")

st.markdown("""
I’ll output one Excel with exactly **two tabs** and a **diagram link**.  
Files I expect across your ZIPs:  
- `GL_PRIMARY_LEDGER.csv`  
- `XLE_ENTITY_PROFILE.csv`  
- `ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv`  
- `ORA_GL_JOURNAL_CONFIG_DETAIL.csv`  
- `FUN_BUSINESS_UNIT.csv` (optional)  
- `CST_COST_ORGANIZATION.csv`  
""")

uploads = st.file_uploader("Drop your Oracle Export ZIPs", type="zip", accept_multiple_files=True)

# ---------------- HELPERS ----------------
def _norm(x):
    return str(x).strip() if pd.notna(x) else ""

def _deflate_base64(text: str) -> str:
    """Raw DEFLATE + base64 (what draw.io expects)."""
    comp = zlib.compressobj(level=9, wbits=-15)
    raw = comp.compress(text.encode("utf-8")) + comp.flush()
    return base64.b64encode(raw).decode("ascii")

def read_csv_from_zip(zf, fname):
    if fname not in zf.namelist():
        return None
    with zf.open(fname) as fh:
        return pd.read_csv(fh, dtype=str)

# ---------------- MAIN ----------------
if not uploads:
    st.info("Upload your ZIPs to generate the workbook + diagram.")
else:
    ledger_names = set()
    legal_entity_names = set()
    ledger_to_idents, ident_to_le_name = {}, {}
    bu_rows, co_rows = [], []

    for up in uploads:
        try:
            z = zipfile.ZipFile(up)
        except Exception as e:
            st.error(f"Could not open `{up.name}`: {e}")
            continue

        # Ledgers
        df = read_csv_from_zip(z, "GL_PRIMARY_LEDGER.csv")
        if df is not None and "ORA_GL_PRIMARY_LEDGER_CONFIG.Name" in df.columns:
            ledger_names |= set(df["ORA_GL_PRIMARY_LEDGER_CONFIG.Name"].dropna().map(str).str.strip())

        # Legal Entities (catalog)
        df = read_csv_from_zip(z, "XLE_ENTITY_PROFILE.csv")
        if df is not None and "Name" in df.columns:
            legal_entity_names |= set(df["Name"].dropna().map(str).str.strip())

        # Ledger ↔ LE Identifier
        df = read_csv_from_zip(z, "ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv")
        if df is not None and {"GL_LEDGER.Name","LegalEntityIdentifier"}.issubset(df.columns):
            for _, r in df[["GL_LEDGER.Name","LegalEntityIdentifier"]].dropna().iterrows():
                led, ident = str(r["GL_LEDGER.Name"]).strip(), str(r["LegalEntityIdentifier"]).strip()
                if led and ident:
                    ledger_to_idents.setdefault(led,set()).add(ident)

        # Identifier ↔ LE Name
        df = read_csv_from_zip(z, "ORA_GL_JOURNAL_CONFIG_DETAIL.csv")
        if df is not None and {"LegalEntityIdentifier","ObjectName"}.issubset(df.columns):
            for _, r in df[["LegalEntityIdentifier","ObjectName"]].dropna().iterrows():
                ident, obj = str(r["LegalEntityIdentifier"]).strip(), str(r["ObjectName"]).strip()
                if ident:
                    ident_to_le_name[ident] = obj

        # Business Units
        df = read_csv_from_zip(z, "FUN_BUSINESS_UNIT.csv")
        if df is not None and {"Name","PrimaryLedgerName","LegalEntityName"}.issubset(df.columns):
            for _, r in df[["Name","PrimaryLedgerName","LegalEntityName"]].iterrows():
                bu_rows.append({
                    "BusinessUnitName": _norm(r["Name"]),
                    "Ledger": _norm(r["PrimaryLedgerName"]),
                    "LegalEntityName": _norm(r["LegalEntityName"])
                })

        # Cost Orgs
        df = read_csv_from_zip(z, "CST_COST_ORGANIZATION.csv")
        if df is not None and {"Name","LegalEntityIdentifier"}.issubset(df.columns):
            for _, r in df[["Name","LegalEntityIdentifier"]].iterrows():
                co_rows.append({
                    "CostOrganization": _norm(r["Name"]),
                    "LegalEntityIdentifier": _norm(r["LegalEntityIdentifier"])
                })

    # Mappings
    ledger_to_le_names = {}
    for led, ids in ledger_to_idents.items():
        for ident in ids:
            le_name = ident_to_le_name.get(ident,"").strip()
            if le_name:
                ledger_to_le_names.setdefault(led,set()).add(le_name)

    # -------- TAB 1: Ledger–LE–BU --------
    df_bu = pd.DataFrame(bu_rows).drop_duplicates()
    df_bu = df_bu[["Ledger","LegalEntityName","BusinessUnitName"]]

    # -------- TAB 2: Ledger–LE–CostOrg --------
    co_enriched = []
    for r in co_rows:
        ident = r["LegalEntityIdentifier"]
        le_name = ident_to_le_name.get(ident,"")
        ledger = ""
        for led, ids in ledger_to_idents.items():
            if ident in ids:
                ledger = led
                break
        co_enriched.append({
            "Ledger": ledger,
            "LegalEntityName": le_name,
            "CostOrganization": r["CostOrganization"],
            "LegalEntityIdentifier": ident
        })
    df_co = pd.DataFrame(co_enriched).drop_duplicates()

    # -------- Excel Output --------
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        df_bu.to_excel(writer, index=False, sheet_name="ES_Ledger-LE-BU")
        df_co.to_excel(writer, index=False, sheet_name="ES_Ledger-LE-CostOrg")

    st.download_button("⬇️ Download Excel (2 tabs)",
        data=excel_buf.getvalue(),
        file_name="EnterpriseStructure.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # -------- Diagram --------
    def drawio_from_frames(df_bu, df_co):
        for df, cols in [(df_bu,["Ledger","LegalEntityName","BusinessUnitName"]),
                         (df_co,["Ledger","LegalEntityName","CostOrganization"])]:
            for c in cols: df[c] = df[c].map(_norm)

        X_LEDGER, X_LE, X_BU, X_CO = 40, 320, 650, 950
        Y_LEDGER, Y_LE, Y_BU, Y_CO = 40, 40, 40, 220
        STEP, W, H = 90, 170, 48

        S_NODE_LEDGER="rounded=1;fillColor=#F5F5F5;strokeColor=#666666;fontStyle=1;"
        S_NODE_LE="rounded=1;fillColor=#FFFFFF;strokeColor=#222222;"
        S_NODE_BU="rounded=1;fillColor=#FFFFFF;strokeColor=#888888;"
        S_NODE_CO="rounded=1;fillColor=#E9F2FF;strokeColor=#1F75FE;"

        S_BASE=("endArrow=block;endFill=1;rounded=1;jettySize=auto;orthogonalLoop=1;"
                "edgeStyle=orthogonalEdgeStyle;orthogonal=1;curved=1;jumpStyle=arc;jumpSize=10;")
        S_LEDGER_LE=S_BASE+"strokeColor=#666666;strokeWidth=2;"
        S_LE_BU=S_BASE+"strokeColor=#FFD400;strokeWidth=2;"
        S_LE_CO=S_BASE+"strokeColor=#1F75FE;strokeWidth=2;"

        class G:
            def __init__(self,name="Enterprise Structure (+ Cost Orgs)"):
                self.name=name; self.root=[{"id":"0"},{"id":"1","parent":"0"}]
            def add_node(self,label,x,y,w,h,style):
                i=uuid.uuid4().hex[:8]
                self.root.append({"id":i,"value":label,"vertex":"1","parent":"1","style":f"whiteSpace=wrap;html=1;align=center;{style}",
                                  "geometry":{"x":x,"y":y,"width":w,"height":h}})
                return i
            def add_edge(self,s,t,style):
                i=uuid.uuid4().hex[:8]
                self.root.append({"id":i,"edge":"1","parent":"1","style":style,"source":s,"target":t,
                                  "geometry":{"relative":"1"}})
            def xml(self):
                mx=ET.Element("mxGraphModel"); root=ET.SubElement(mx,"root")
                for c in self.root:
                    attrs={k:v for k,v in c.items() if k!="geometry"}
                    cell=ET.SubElement(root,"mxCell",attrs)
                    if "geometry" in c:
                        g=ET.SubElement(cell,"mxGeometry",{**{k:str(v) for k,v in c["geometry"].items()}, "as":"geometry"})
                enc=_deflate_base64(ET.tostring(mx,encoding="utf-8").decode("utf-8"))
                mxfile=ET.Element("mxfile",{"host":"app.diagrams.net"})
                diagram=ET.SubElement(mxfile,"diagram",{"name":self.name,"id":uuid.uuid4().hex[:12]})
                diagram.text=enc
                return ET.tostring(mxfile,encoding="utf-8",xml_declaration=True).decode("utf-8")

        g=G()
        ledgers=sorted(set(df_bu["Ledger"])|set(df_co["Ledger"])) or [""]
        ledger_ids={}; y=Y_LEDGER
        for L in ledgers:
            ledger_ids[L]=g.add_node(L if L else "(No Ledger)",X_LEDGER,y,W,H,S_NODE_LEDGER); y+=STEP

        df_le=(pd.concat([df_bu[["Ledger","LegalEntityName"]],df_co[["Ledger","LegalEntityName"]]],ignore_index=True)
               .drop_duplicates().sort_values(["Ledger","LegalEntityName"]))
        le_ids={}; last=None; y_le=Y_LE
        for _,r in df_le.iterrows():
            L,E=r["Ledger"],r["LegalEntityName"] or "(Unnamed LE)"
            if L!=last: y_le=Y_LE; last=L
            nid=g.add_node(E,X_LE,y_le,W,H,S_NODE_LE); le_ids[(L,E)]=nid
            g.add_edge(ledger_ids.get(L,ledger_ids.get("",next(iter(ledger_ids.values())))),nid,S_LEDGER_LE)
            y_le+=STEP

        for (L,E),grp in df_bu.groupby(["Ledger","LegalEntityName"],dropna=False):
            if (L,E or "(Unnamed LE)") not in le_ids: continue
            y_bu=Y_BU
            for bu in sorted({b for b in grp["BusinessUnitName"].map(_norm) if b}):
                bn=g.add_node(bu,X_BU,y_bu,W,H,S_NODE_BU); g.add_edge(le_ids[(L,E or "(Unnamed LE)")],bn,S_LE_BU); y_bu+=STEP

        for (L,E),grp in df_co.groupby(["Ledger","LegalEntityName"],dropna=False):
            if (L,E or "(Unnamed LE)") not in le_ids: continue
            y_co=Y_CO
            for co in sorted({c for c in grp["CostOrganization"].map(_norm) if c}):
                cn=g.add_node(co,X_CO,y_co,W,H,S_NODE_CO); g.add_edge(le_ids[(L,E or "(Unnamed LE)")],cn,S_LE_CO); y_co+=STEP

        return g.xml()

    xml = drawio_from_frames(df_bu, df_co)
    st.download_button("⬇️ Download Diagram (.drawio)",data=xml.encode("utf-8"),
        file_name="EnterpriseStructure.drawio",mime="application/xml")
    b64 = base64.b64encode(xml.encode("utf-8")).decode("utf-8")
    st.markdown(f"[🔗 Open in draw.io](https://app.diagrams.net/#R{_deflate_base64(xml)})")
