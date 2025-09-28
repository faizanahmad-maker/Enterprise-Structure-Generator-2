import io, zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Enterprise Structure Generator", page_icon="📊", layout="wide")
st.title("Enterprise Structure Generator — Excel + draw.io (with Inventory Orgs)")

st.markdown("""
Upload up to **9 Oracle export ZIPs** (any order):
- `Manage General Ledger` → **GL_PRIMARY_LEDGER.csv**
- `Manage Legal Entities` → **XLE_ENTITY_PROFILE.csv**
- `Assign Legal Entities` → **ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv**
- *(optional backup)* Journal config detail → **ORA_GL_JOURNAL_CONFIG_DETAIL.csv**
- `Manage Business Units` → **FUN_BUSINESS_UNIT.csv**
- `Manage Cost Organizations` → **CST_COST_ORGANIZATION.csv**
- `Manage Cost Organization Relationships` → **CST_COST_ORG_BOOK.csv**
- `Manage Inventory Organizations` → **INV_ORGANIZATION_PARAMETER.csv**
- `Cost Org ↔ Inventory Org relationships` → **ORA_CST_COST_ORG_INV.csv**
""")

uploads = st.file_uploader("Drop your ZIPs here", type="zip", accept_multiple_files=True)

# ---------- helpers ----------
def read_csv_from_zip(zf, name):
    if name not in zf.namelist():
        return None
    with zf.open(name) as fh:
        return pd.read_csv(fh, dtype=str)

def pick_col(df, candidates):
    """Return the first matching column from `candidates` (exact > case-insensitive > substring)."""
    cols = list(df.columns)
    for c in candidates:
        if c in cols:
            return c
    lower_map = {c.lower(): c for c in cols}
    for c in candidates:
        if c.lower() in lower_map:
            return lower_map[c.lower()]
    for c in candidates:
        for existing in cols:
            if c.lower() in existing.lower():
                return existing
    return None

if not uploads:
    st.info("Upload your ZIPs to generate the Excel & diagram.")
else:
    # ------------ Collectors ------------
    ledger_names = set()
    legal_entity_names = set()
    ledger_to_idents = {}            # ledger -> {LE identifier}
    ident_to_le_name = {}            # LE identifier -> LE name
    bu_rows = []                     # BU rows (for Tab 1 only)

    # Cost Orgs (MASTER)
    costorg_rows = []                # [{Name, LegalEntityIdentifier, JoinKey}]
    costorg_name_to_joinkeys = {}    # Name -> {JoinKey}
    # Cost Books: JoinKey -> {CostBookCode}
    books_by_joinkey = {}

    # Inventory Orgs (MASTER)
    invorg_rows = []                 # [{Code, Name, LEIdent, BUName, PCBU, Mfg}]
    # IO↔CostOrg relationships: InvOrgCode -> CostOrgJoinKey
    invorg_rel = {}

    # ------------ Scan uploads ------------
    for up in uploads:
        try:
            z = zipfile.ZipFile(up)
        except Exception as e:
            st.error(f"Could not open `{up.name}` as a ZIP: {e}")
            continue

        # Ledgers
        df = read_csv_from_zip(z, "GL_PRIMARY_LEDGER.csv")
        if df is not None:
            col = pick_col(df, ["ORA_GL_PRIMARY_LEDGER_CONFIG.Name"])
            if col:
                ledger_names |= set(df[col].dropna().map(str).str.strip())
            else:
                st.warning("`GL_PRIMARY_LEDGER.csv` missing `ORA_GL_PRIMARY_LEDGER_CONFIG.Name`.")

        # Legal Entities
        df = read_csv_from_zip(z, "XLE_ENTITY_PROFILE.csv")
        if df is not None:
            name_col  = pick_col(df, ["Name"])
            ident_col = pick_col(df, ["LegalEntityIdentifier"])
            if name_col and ident_col:
                for _, r in df[[name_col, ident_col]].dropna(how="all").iterrows():
                    le_name = str(r[name_col]).strip()
                    le_ident = str(r[ident_col]).strip()
                    if le_name:
                        legal_entity_names.add(le_name)
                    if le_ident and le_name:
                        ident_to_le_name[le_ident] = le_name
            else:
                st.warning(f"`XLE_ENTITY_PROFILE.csv` missing needed columns. Found: {list(df.columns)}")

        # Ledger ↔ LE identifier
        df = read_csv_from_zip(z, "ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv")
        if df is not None:
            led_col   = pick_col(df, ["GL_LEDGER.Name"])
            ident_col = pick_col(df, ["LegalEntityIdentifier"])
            if led_col and ident_col:
                for _, r in df[[led_col, ident_col]].dropna(how="all").iterrows():
                    led = str(r[led_col]).strip()
                    ident = str(r[ident_col]).strip()
                    if led and ident:
                        ledger_to_idents.setdefault(led, set()).add(ident)
            else:
                st.warning(f"`ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv` missing needed columns. Found: {list(df.columns)}")

        # Backup map for identifier -> LE name
        df = read_csv_from_zip(z, "ORA_GL_JOURNAL_CONFIG_DETAIL.csv")
        if df is not None:
            ident_col = pick_col(df, ["LegalEntityIdentifier"])
            obj_col   = pick_col(df, ["ObjectName"])
            if ident_col and obj_col:
                for _, r in df[[ident_col, obj_col]].dropna(how="all").iterrows():
                    ident = str(r[ident_col]).strip()
                    obj   = str(r[obj_col]).strip()
                    if ident and obj and ident not in ident_to_le_name:
                        ident_to_le_name[ident] = obj

        # Business Units (for Tab 1)
        df = read_csv_from_zip(z, "FUN_BUSINESS_UNIT.csv")
        if df is not None:
            bu_col  = pick_col(df, ["Name"])
            led_col = pick_col(df, ["PrimaryLedgerName"])
            le_col  = pick_col(df, ["LegalEntityName"])
            if bu_col and led_col and le_col:
                for _, r in df[[bu_col, led_col, le_col]].dropna(how="all").iterrows():
                    bu_rows.append({
                        "Name": str(r[bu_col]).strip(),
                        "PrimaryLedgerName": str(r[led_col]).strip(),
                        "LegalEntityName": str(r[le_col]).strip()
                    })
            else:
                st.warning(f"`FUN_BUSINESS_UNIT.csv` missing needed columns. Found: {list(df.columns)}")

        # Cost Orgs (MASTER)
        df = read_csv_from_zip(z, "CST_COST_ORGANIZATION.csv")
        if df is not None:
            name_col   = pick_col(df, ["Name"])
            ident_col  = pick_col(df, ["LegalEntityIdentifier"])
            join_col   = pick_col(df, ["OrgInformation2"])  # join to BOOKS + IO relationships
            if name_col and ident_col and join_col:
                for _, r in df[[name_col, ident_col, join_col]].dropna(how="all").iterrows():
                    name  = str(r[name_col]).strip()
                    ident = str(r[ident_col]).strip()
                    joink = str(r[join_col]).strip()
                    costorg_rows.append({"Name": name, "LegalEntityIdentifier": ident, "JoinKey": joink})
                    if name and joink:
                        costorg_name_to_joinkeys.setdefault(name, set()).add(joink)
            else:
                st.warning(f"`CST_COST_ORGANIZATION.csv` missing needed columns (Name, LegalEntityIdentifier, OrgInformation2). Found: {list(df.columns)}")

        # Cost Books — JoinKey(CostOrgCode) -> {CostBookCode}
        df = read_csv_from_zip(z, "CST_COST_ORG_BOOK.csv")
        if df is not None:
            key_col  = pick_col(df, ["ORA_CST_ACCT_COST_ORG.CostOrgCode", "CostOrgCode"])
            book_col = pick_col(df, ["CostBookCode"])
            if key_col and book_col:
                for _, r in df[[key_col, book_col]].dropna(how="all").iterrows():
                    joink = str(r[key_col]).strip()
                    book  = str(r[book_col]).strip()
                    if joink and book:
                        books_by_joinkey.setdefault(joink, set()).add(book)
            else:
                st.warning(f"`CST_COST_ORG_BOOK.csv` missing needed columns (CostOrgCode, CostBookCode). Found: {list(df.columns)}")

        # Inventory Orgs (MASTER)
        df = read_csv_from_zip(z, "INV_ORGANIZATION_PARAMETER.csv")
        if df is not None:
            code_col  = pick_col(df, ["OrganizationCode"])
            name_col  = pick_col(df, ["Name", "OrganizationName"])
            le_col    = pick_col(df, ["LegalEntityIdentifier", "LEIdentifier"])
            bu_col    = pick_col(df, ["BusinessUnitName"])
            pcbu_col  = pick_col(df, ["ProfitCenterBuName"])
            mfg_col   = pick_col(df, ["MfgPlantFlag"])
            if code_col and name_col:
                for _, r in df.dropna(how="all").iterrows():
                    invorg_rows.append({
                        "Code": str(r.get(code_col, "")).strip(),
                        "Name": str(r.get(name_col, "")).strip(),
                        "LEIdent": str(r.get(le_col, "")).strip(),
                        "BUName": str(r.get(bu_col, "")).strip(),
                        "PCBU": str(r.get(pcbu_col, "")).strip(),
                        "Mfg": "Yes" if str(r.get(mfg_col, "")).strip().upper() == "Y" else ""
                    })
            else:
                st.warning(f"`INV_ORGANIZATION_PARAMETER.csv` missing needed columns. Found: {list(df.columns)}")

        # Cost Org ↔ Inventory Org relationships
        df = read_csv_from_zip(z, "ORA_CST_COST_ORG_INV.csv")
        if df is not None:
            inv_col  = pick_col(df, ["OrganizationCode", "InventoryOrganizationCode"])
            co_col   = pick_col(df, ["ORA_CST_ACCT_COST_ORG.CostOrgCode", "CostOrgCode"])
            if inv_col and co_col:
                for _, r in df[[inv_col, co_col]].dropna(how="all").iterrows():
                    inv_code, co_code = str(r[inv_col]).strip(), str(r[co_col]).strip()
                    if inv_code and co_code:
                        invorg_rel[inv_code] = co_code
            else:
                st.warning(f"`ORA_CST_COST_ORG_INV.csv` missing needed columns (OrganizationCode, CostOrgCode). Found: {list(df.columns)}")

    # ------------ Derived maps ------------
    ident_to_ledgers = {}
    for led, idents in ledger_to_idents.items():
        for ident in idents:
            ident_to_ledgers.setdefault(ident, set()).add(led)

    ledger_to_le_names = {}
    for led, idents in ledger_to_idents.items():
        for ident in idents:
            le_name = ident_to_le_name.get(ident, "").strip()
            if le_name:
                ledger_to_le_names.setdefault(led, set()).add(le_name)

    known_pairs = set()
    for led, idents in ledger_to_idents.items():
        for ident in idents:
            le_name = ident_to_le_name.get(ident, "").strip()
            if le_name:
                known_pairs.add((led, le_name))

    # ===================================================
    # Tab 1: Ledger – Legal Entity – Business Unit (unchanged semantics)
    # ===================================================
    rows1, seen_triples, seen_ledgers_with_bu = [], set(), set()

    # Emit BU-driven rows (no heuristics; you asked for strictness)
    for r in bu_rows:
        bu  = r["Name"]
        led = r["PrimaryLedgerName"]
        le  = r["LegalEntityName"]
        rows1.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": bu})
        seen_triples.add((led, le, bu))
        if led:
            seen_ledgers_with_bu.add(led)

    # Add ledger–LE pairs from mapping that have no BU
    seen_pairs = {(a, b) for (a, b, _) in seen_triples}
    for led, le in sorted(known_pairs):
        if (led, le) not in seen_pairs:
            rows1.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": ""})

    # Orphan ledgers (exist, but no mapping & no BU)
    mapped_ledgers = set(ledger_to_le_names.keys())
    for led in sorted(ledger_names - mapped_ledgers - seen_ledgers_with_bu):
        rows1.append({"Ledger Name": led, "Legal Entity": "", "Business Unit": ""})

    df1 = pd.DataFrame(rows1).drop_duplicates().reset_index(drop=True)
    df1["__LedgerEmpty"] = (df1["Ledger Name"] == "").astype(int)
    df1 = (
        df1.sort_values(
            ["__LedgerEmpty", "Ledger Name", "Legal Entity", "Business Unit"],
            ascending=[True, True, True, True]
        )
        .drop(columns="__LedgerEmpty")
        .reset_index(drop=True)
    )
    df1.insert(0, "Assignment", range(1, len(df1) + 1))

    # ===================================================
    # Tab 2: Ledger – LE – Cost Org – Cost Book – Inventory Org – ProfitCenter BU – Management BU – Mfg Plant
    #  - Emits all IOs (even if unassigned)
    #  - Sorts so "hanging" (missing Ledger or Cost Org) drop to bottom
    # ===================================================
    rows2 = []

    # Helper: Cost Org name from JoinKey
    co_name_by_joinkey = {r["JoinKey"]: r["Name"] for r in costorg_rows if r.get("JoinKey")}

    for inv in invorg_rows:
        code = inv.get("Code", "")
        name = inv.get("Name", "")
        le_ident = inv.get("LEIdent", "")
        le_name  = ident_to_le_name.get(le_ident, "") if le_ident else ""
        leds     = ident_to_ledgers.get(le_ident, set()) if le_ident else set()

        co_key  = invorg_rel.get(code, "")
        co_name = co_name_by_joinkey.get(co_key, "") if co_key else ""
        books   = "; ".join(sorted(books_by_joinkey.get(co_key, []))) if co_key else ""

        # Emit rows; if multiple ledgers for the LE, fan out
        if leds:
            for led in sorted(leds):
                rows2.append({
                    "Ledger Name": led,
                    "Legal Entity": le_name,
                    "Cost Organization": co_name,
                    "Cost Book": books,
                    "Inventory Org": name,
                    "Profit Center BU": inv.get("PCBU", ""),
                    "Management BU": inv.get("BUName", ""),
                    "Manufacturing Plant": inv.get("Mfg", "")
                })
        else:
            # Hanging IO (no ledger alignment)
            rows2.append({
                "Ledger Name": "",
                "Legal Entity": le_name,
                "Cost Organization": co_name,
                "Cost Book": books,
                "Inventory Org": name,
                "Profit Center BU": inv.get("PCBU", ""),
                "Management BU": inv.get("BUName", ""),
                "Manufacturing Plant": inv.get("Mfg", "")
            })

    df2 = pd.DataFrame(rows2).drop_duplicates().reset_index(drop=True)

    # ---------- Hanging-first sort controls (push unassigned to bottom) ----------
    # Ledger empty → bottom; Cost Org empty → bottom; then normal alpha
    if not df2.empty:
        df2["__LedgerEmpty"] = (df2["Ledger Name"].fillna("") == "").astype(int)
        df2["__COEmpty"]     = (df2["Cost Organization"].fillna("") == "").astype(int)
        df2["__HasIO"]       = 1  # reserved for future pair-padding; currently always 1
        df2 = (
            df2.sort_values(
                ["__LedgerEmpty", "Ledger Name", "Legal Entity", "__COEmpty", "Cost Organization", "Inventory Org"],
                ascending=[True, True, True, True, True, True]
            )
            .drop(columns=["__LedgerEmpty", "__COEmpty", "__HasIO"])
            .reset_index(drop=True)
        )

    df2.insert(0, "Assignment", range(1, len(df2) + 1))

    # ------------ Excel Output ------------
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        df1.to_excel(writer, index=False, sheet_name="Ledger_LE_BU_Assignments")
        df2.to_excel(writer, index=False, sheet_name="Ledger_LE_CostOrg_IOs")

    st.success(f"Built {len(df1)} BU rows and {len(df2)} Inventory Org rows (hanging handled).")
    st.dataframe(df1.head(25), use_container_width=True, height=280)
    st.dataframe(df2.head(25), use_container_width=True, height=320)

    st.download_button(
        "⬇️ Download Excel (EnterpriseStructure.xlsx)",
        data=excel_buf.getvalue(),
        file_name="EnterpriseStructure.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Diagram section intentionally deferred until Excel is signed off.
# ===================== DRAW.IO DIAGRAM BLOCK (with Inventory Orgs) =====================
if (
    "df1" in locals() and isinstance(df1, pd.DataFrame) and not df1.empty and
    "df2" in locals() and isinstance(df2, pd.DataFrame)
):
    import xml.etree.ElementTree as ET
    import zlib, base64, uuid

    def _make_drawio_xml(df_bu: pd.DataFrame, df_io: pd.DataFrame) -> str:
        # --- layout & spacing ---
        W, H       = 180, 48
        X_STEP     = 230
        PAD_GROUP  = 60
        LEFT_PAD   = 260
        RIGHT_PAD  = 200

        Y_LEDGER   = 150
        Y_LE       = 310
        Y_BU       = 470
        Y_CO       = 630
        Y_CB       = 790
        Y_IO       = 950  # Inventory Orgs sit *below* Cost Books

        # --- styles ---
        S_LEDGER = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE6E6;strokeColor=#C86868;fontSize=12;"
        S_LE     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE2C2;strokeColor=#A66000;fontSize=12;"
        S_BU     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFF1B3;strokeColor=#B38F00;fontSize=12;"
        S_CO     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#E2F7E2;strokeColor=#3D8B3D;fontSize=12;"
        S_CB     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#7FBF7F;strokeColor=#2F7D2F;fontSize=12;"
        # Inventory Orgs
        S_IO      = "rounded=1;whiteSpace=wrap;html=1;fillColor=#D6EFFF;strokeColor=#2F71A8;fontSize=12;"
        S_IO_PLANT = "rounded=1;whiteSpace=wrap;html=1;fillColor=#D6EFFF;strokeColor=#1F4D7A;strokeWidth=2;fontSize=12;"

        S_EDGE   = ("endArrow=block;rounded=1;edgeStyle=orthogonalEdgeStyle;orthogonal=1;"
                    "jettySize=auto;strokeColor=#666666;exitX=0.5;exitY=0;entryX=0.5;entryY=1;")

        # --- normalize inputs ---
        df_io = df_io.fillna("")
        df_bu = df_bu.fillna("")

        ledgers_all = sorted([x for x in set(df_bu["Ledger Name"]) | set(df_io["Ledger Name"]) if x])

        # maps
        le_map, bu_map, co_map, cb_map, io_map = {}, {}, {}, {}, {}

        for _, r in pd.concat([df_bu[["Ledger Name","Legal Entity"]],
                               df_io[["Ledger Name","Legal Entity"]]]).drop_duplicates().iterrows():
            if r["Ledger Name"] and r["Legal Entity"]:
                le_map.setdefault(r["Ledger Name"], set()).add(r["Legal Entity"])

        for _, r in df_bu.iterrows():
            if r["Ledger Name"] and r["Legal Entity"] and r["Business Unit"]:
                bu_map.setdefault((r["Ledger Name"], r["Legal Entity"]), set()).add(r["Business Unit"])

        for _, r in df_io.iterrows():
            if r["Ledger Name"] and r["Legal Entity"] and r["Cost Organization"]:
                co_map.setdefault((r["Ledger Name"], r["Legal Entity"]), set()).add(r["Cost Organization"])

        for _, r in df_io.iterrows():
            if r["Ledger Name"] and r["Legal Entity"] and r["Cost Organization"] and r["Cost Book"]:
                cb_map.setdefault((r["Ledger Name"], r["Legal Entity"], r["Cost Organization"]), set()).add(r["Cost Book"])

        for _, r in df_io.iterrows():
            if r["Ledger Name"] and r["Legal Entity"] and r["Cost Organization"] and r["Cost Book"] and r["Inventory Org"]:
                io_map.setdefault((r["Ledger Name"], r["Legal Entity"], r["Cost Organization"], r["Cost Book"]), []).append({
                    "Name": r["Inventory Org"],
                    "Mfg": str(r.get("Manufacturing Plant", "")).strip()
                })

        # --- XML skeleton ---
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

        def add_vertex(label, style, x, y):
            vid = uuid.uuid4().hex[:8]
            c = ET.SubElement(root, "mxCell", attrib={
                "id": vid, "value": label, "style": style, "vertex": "1", "parent": "1"})
            ET.SubElement(c, "mxGeometry", attrib={
                "x": str(int(x)), "y": str(int(y)), "width": str(W), "height": str(H), "as": "geometry"})
            return vid

        def add_edge(src, tgt, dashed=False):
            eid = uuid.uuid4().hex[:8]
            style = S_EDGE + ("dashed=1;" if dashed else "")
            c = ET.SubElement(root, "mxCell", attrib={
                "id": eid, "value": "", "style": style, "edge": "1", "parent": "1",
                "source": src, "target": tgt})
            ET.SubElement(c, "mxGeometry", attrib={"relative": "1", "as": "geometry"})

        # --- build nodes ---
        id_map = {}
        next_x = LEFT_PAD

        for L in ledgers_all:
            id_map[("L", L)] = add_vertex(L, S_LEDGER, next_x, Y_LEDGER)
            les = sorted(le_map.get(L, []))
            for le in les:
                id_map[("E", L, le)] = add_vertex(le, S_LE, next_x, Y_LE)
                add_edge(id_map[("E", L, le)], id_map[("L", L)])
                # BUs
                for b in sorted(bu_map.get((L, le), [])):
                    id_map[("B", L, le, b)] = add_vertex(b, S_BU, next_x, Y_BU)
                    add_edge(id_map[("B", L, le, b)], id_map[("E", L, le)])
                # Cost Orgs
                for c in sorted(co_map.get((L, le), [])):
                    id_map[("C", L, le, c)] = add_vertex(c, S_CO, next_x, Y_CO)
                    add_edge(id_map[("C", L, le, c)], id_map[("E", L, le)])
                    # Cost Books
                    for cb in sorted(cb_map.get((L, le, c), [])):
                        id_map[("CB", L, le, c, cb)] = add_vertex(cb, S_CB, next_x, Y_CB)
                        add_edge(id_map[("CB", L, le, c, cb)], id_map[("C", L, le, c)])
                        # Inventory Orgs under CB
                        for io in io_map.get((L, le, c, cb), []):
                            io_label = f"🏭 {io['Name']}" if io["Mfg"].lower() == "yes" else io["Name"]
                            style = S_IO_PLANT if io["Mfg"].lower() == "yes" else S_IO
                            id_map[("IO", L, le, c, cb, io["Name"])] = add_vertex(io_label, style, next_x, Y_IO)
                            add_edge(id_map[("IO", L, le, c, cb, io["Name"])], id_map[("C", L, le, c)])  # link to Cost Org
            next_x += X_STEP + PAD_GROUP

        return ET.tostring(mxfile, encoding="utf-8", method="xml").decode("utf-8")

    def _drawio_url_from_xml(xml: str) -> str:
        raw = zlib.compress(xml.encode("utf-8"), level=9)[2:-4]
        b64 = base64.b64encode(raw).decode("ascii")
        return f"https://app.diagrams.net/?title=EnterpriseStructure.drawio#R{b64}"

    _xml = _make_drawio_xml(df1, df2)

    st.download_button(
        "⬇️ Download diagram (.drawio)",
        data=_xml.encode("utf-8"),
        file_name="EnterpriseStructure.drawio",
        mime="application/xml",
        use_container_width=True
    )
    st.markdown(f"[🔗 Open in draw.io (preview)]({_drawio_url_from_xml(_xml)})")
