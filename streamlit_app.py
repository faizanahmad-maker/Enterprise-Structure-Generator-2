import io, zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Enterprise Structure Generator 2", page_icon="🧭", layout="wide")
st.title("Enterprise Structure Generator 2 — Core + Cost Org lane")

st.markdown("""
Upload up to **6 Oracle export ZIPs** (any order):
- Core (same as v1): `GL_PRIMARY_LEDGER.csv`, `XLE_ENTITY_PROFILE.csv`,
  `ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv`, `ORA_GL_JOURNAL_CONFIG_DETAIL.csv`, `FUN_BUSINESS_UNIT.csv`
- Costing (for Sheet 2): `CST_COST_ORGANIZATION.csv` (Name, LegalEntityIdentifier, OrgInformation2),
  **either** `ORA_CST_ACCT_COST_ORG.csv` (CostOrgCode→Ledger) **or** `CST_COST_ORG_BOOK.csv` (book→ledger)
""")

uploads = st.file_uploader("Drop your ZIPs here", type="zip", accept_multiple_files=True)

def read_csv_from_zip(zf, name):
    if name not in zf.namelist():
        return None
    with zf.open(name) as fh:
        return pd.read_csv(fh, dtype=str)

if not uploads:
    st.info("Upload your ZIPs to generate outputs.")
else:
    # collectors (unchanged core)
    ledger_names = set()            # GL_PRIMARY_LEDGER.csv :: ORA_GL_PRIMARY_LEDGER_CONFIG.Name
    legal_entity_names = set()      # XLE_ENTITY_PROFILE.csv :: Name
    ledger_to_idents = {}           # ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv :: GL_LEDGER.Name -> {LegalEntityIdentifier}
    ident_to_le_name = {}           # ORA_GL_JOURNAL_CONFIG_DETAIL.csv     :: LegalEntityIdentifier -> ObjectName
    bu_rows = []                    # FUN_BUSINESS_UNIT.csv :: Name, PrimaryLedgerName, LegalEntityName

    # >>> NEW: Cost Org collectors
    costorg_rows = []   # from CST_COST_ORGANIZATION.csv (Name, LegalEntityIdentifier, OrgInformation2)
    cstorgcode_to_ledger = set()  # tuples (CostOrgCode, LedgerName)
    costbook_lookup = []          # optional book-derived mapping (CostOrgName or Code -> LedgerName)

    # scan all uploaded zips
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
            else:
                st.warning(f"`GL_PRIMARY_LEDGER.csv` missing `{col}`. Found: {list(df.columns)}")

        # Legal Entities
        df = read_csv_from_zip(z, "XLE_ENTITY_PROFILE.csv")
        if df is not None:
            col = "Name"
            if col in df.columns:
                legal_entity_names |= set(df[col].dropna().map(str).str.strip())
            else:
                st.warning(f"`XLE_ENTITY_PROFILE.csv` missing `Name`. Found: {list(df.columns)}")

        # Ledger ↔ LE identifier
        df = read_csv_from_zip(z, "ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv")
        if df is not None:
            need = ["GL_LEDGER.Name", "LegalEntityIdentifier"]
            missing = [c for c in need if c not in df.columns]
            if missing:
                st.warning(f"`ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv` missing {missing}. Found: {list(df.columns)}")
            else:
                for _, r in df[need].dropna().iterrows():
                    led = str(r["GL_LEDGER.Name"]).strip()
                    ident = str(r["LegalEntityIdentifier"]).strip()
                    if led and ident:
                        ledger_to_idents.setdefault(led, set()).add(ident)

        # Identifier ↔ LE name
        df = read_csv_from_zip(z, "ORA_GL_JOURNAL_CONFIG_DETAIL.csv")
        if df is not None:
            need = ["LegalEntityIdentifier", "ObjectName"]
            missing = [c for c in need if c not in df.columns]
            if missing:
                st.warning(f"`ORA_GL_JOURNAL_CONFIG_DETAIL.csv` missing {missing}. Found: {list(df.columns)}")
            else:
                for _, r in df[need].dropna().iterrows():
                    ident = str(r["LegalEntityIdentifier"]).strip()
                    obj = str(r["ObjectName"]).strip()
                    if ident:
                        ident_to_le_name[ident] = obj

        # Business Units
        df = read_csv_from_zip(z, "FUN_BUSINESS_UNIT.csv")
        if df is not None:
            need = ["Name", "PrimaryLedgerName", "LegalEntityName"]
            missing = [c for c in need if c not in df.columns]
            if missing:
                st.warning(f"`FUN_BUSINESS_UNIT.csv` missing {missing}. Found: {list(df.columns)}")
            else:
                for c in need:
                    df[c] = df[c].astype(str).map(lambda x: x.strip() if x else "")
                bu_rows += df[need].to_dict(orient="records")

        # >>> NEW: Cost Orgs master (CST_COST_ORGANIZATION.csv)
        df = read_csv_from_zip(z, "CST_COST_ORGANIZATION.csv")
        if df is not None:
            need_any = ["Name", "LegalEntityIdentifier", "OrgInformation2"]
            miss = [c for c in need_any if c not in df.columns]
            if miss:
                st.warning(f"`CST_COST_ORGANIZATION.csv` missing {miss}. Found: {list(df.columns)}")
            else:
                tmp = df[need_any].copy()
                for c in need_any:
                    tmp[c] = tmp[c].astype(str).str.strip()
                tmp.rename(columns={
                    "Name":"CostOrgName",
                    "LegalEntityIdentifier":"LE_Ident",
                    "OrgInformation2":"CostOrgCode"
                }, inplace=True)
                costorg_rows += tmp.to_dict(orient="records")

        # >>> NEW: Cost Org → Ledger via accounting relationship (preferred)
        df = read_csv_from_zip(z, "ORA_CST_ACCT_COST_ORG.csv")
        if df is not None:
            # accept common column variants
            code_col = "CostOrgCode" if "CostOrgCode" in df.columns else None
            if not code_col and "CST_COST_ORGANIZATION.Code" in df.columns:
                code_col = "CST_COST_ORGANIZATION.Code"
            led_col = None
            for c in ["LedgerName", "GL_LedgerName", "PrimaryLedgerName"]:
                if c in df.columns: led_col = c; break
            if code_col and led_col:
                for _, r in df[[code_col, led_col]].dropna().iterrows():
                    ccode = str(r[code_col]).strip()
                    lnam  = str(r[led_col]).strip()
                    if ccode and lnam:
                        cstorgcode_to_ledger.add((ccode, lnam))
            else:
                st.warning("`ORA_CST_ACCT_COST_ORG.csv` present but missing CostOrgCode or Ledger column.")

        # >>> NEW: Fallback via Cost Book (CST_COST_ORG_BOOK.csv)
        df = read_csv_from_zip(z, "CST_COST_ORG_BOOK.csv")
        if df is not None:
            # try to harvest cost org identity + ledger
            code_col = None
            for c in ["CostOrgCode","CST_COST_ORGANIZATION.Code","Cost Organization Code"]:
                if c in df.columns: code_col = c; break
            name_col = "Name" if "Name" in df.columns else None
            led_col  = None
            for c in ["LedgerName","GL_LedgerName","PrimaryLedgerName"]:
                if c in df.columns: led_col = c; break
            if led_col and (code_col or name_col):
                for _, r in df.iterrows():
                    lnam = str(r[led_col]).strip() if pd.notna(r.get(led_col, "")) else ""
                    ccode = str(r[code_col]).strip() if code_col and pd.notna(r.get(code_col, "")) else ""
                    cname = str(r[name_col]).strip() if name_col and pd.notna(r.get(name_col, "")) else ""
                    if lnam and (ccode or cname):
                        costbook_lookup.append({"CostOrgCode": ccode, "CostOrgName": cname, "LedgerName": lnam})

    # ---------------- build mappings (core stays same) ----------------
    ledger_to_le_names = {}
    for led, idents in ledger_to_idents.items():
        for ident in idents:
            le_name = ident_to_le_name.get(ident, "").strip()
            if le_name:
                ledger_to_le_names.setdefault(led, set()).add(le_name)

    le_to_ledgers = {}
    for led, le_set in ledger_to_le_names.items():
        for le in le_set:
            le_to_ledgers.setdefault(le, set()).add(led)

    # Build final rows: Ledger Name | Legal Entity | Business Unit (v1 behavior)
    rows = []
    seen_triples = set()
    seen_ledgers_with_bu = set()
    seen_les_with_bu = set()

    # 1) BU-driven rows with smart back-fill
    for r in bu_rows:
        bu = r["Name"]
        led = r["PrimaryLedgerName"] if r["PrimaryLedgerName"] in ledger_names else ""
        le  = r["LegalEntityName"]  if r["LegalEntityName"]  in legal_entity_names else ""

        # back-fill from unique relationships
        if not led and le and le in le_to_ledgers and len(le_to_ledgers[le]) == 1:
            led = next(iter(le_to_ledgers[le]))
        if not le and led and led in ledger_to_le_names and len(ledger_to_le_names[led]) == 1:
            le = next(iter(ledger_to_le_names[led]))

        rows.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": bu})
        seen_triples.add((led, le, bu))
        if led: seen_ledgers_with_bu.add(led)
        if le:  seen_les_with_bu.add(le)

    # 2) Ledger–LE pairs with no BU
    seen_pairs = {(a, b) for (a, b, _) in seen_triples}
    for led, le_set in ledger_to_le_names.items():
        if not le_set:
            if led not in seen_ledgers_with_bu:
                rows.append({"Ledger Name": led, "Legal Entity": "", "Business Unit": ""})
            continue
        for le in le_set:
            if (led, le) not in seen_pairs:
                rows.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": ""})

    # 3) Orphan ledgers
    for led in sorted(ledger_names - set(ledger_to_le_names.keys()) - seen_ledgers_with_bu):
        rows.append({"Ledger Name": led, "Legal Entity": "", "Business Unit": ""})

    # 4) Orphan LEs
    for le in sorted(legal_entity_names - seen_les_with_bu):
        if le in le_to_ledgers and len(le_to_ledgers[le]) == 1:
            led = next(iter(le_to_ledgers[le]))
        else:
            led = ""
        rows.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": ""})

    df = pd.DataFrame(rows).drop_duplicates().reset_index(drop=True)
    df["__LedgerEmpty"] = (df["Ledger Name"] == "").astype(int)
    df = df.sort_values(["__LedgerEmpty","Ledger Name","Legal Entity","Business Unit"],
                        ascending=[True, True, True, True]).drop(columns="__LedgerEmpty").reset_index(drop=True)
    df.insert(0, "Assignment", range(1, len(df)+1))

    st.success(f"Sheet 1 (Core): Built {len(df)} rows.")
    st.dataframe(df, use_container_width=True, height=400)

    # ===================== SHEET 2: Ledger - Legal Entity - Cost Org (row-per-cost-org) =====================
    # >>> NEW: build cost org master with names and mapped LE names
    cost_master = pd.DataFrame(costorg_rows) if costorg_rows else pd.DataFrame(columns=["CostOrgName","LE_Ident","CostOrgCode"])
    for c in ["CostOrgName","LE_Ident","CostOrgCode"]:
        if c not in cost_master.columns: cost_master[c] = ""
        cost_master[c] = cost_master[c].fillna("").map(str).str.strip()

    # map LE_Ident -> LE_Name
    cost_master["Legal Entity"] = cost_master["LE_Ident"].map(lambda x: ident_to_le_name.get(x, "").strip() if x else "")

    # build CostOrgCode -> LedgerName from preferred file + fallback from books
    code_to_ledger = {}
    for ccode, lnam in cstorgcode_to_ledger:
        code_to_ledger.setdefault(ccode, set()).add(lnam)
    for row in costbook_lookup:
        ccode = row.get("CostOrgCode","").strip()
        lnam  = row.get("LedgerName","").strip()
        if ccode and lnam:
            code_to_ledger.setdefault(ccode, set()).add(lnam)
        # if only name present and no code, we cannot safely map—skip for Sheet 2 (no guessing)

    # explode to rows: one per CostOrg x Ledger (orphan if no ledger)
    out_rows = []
    if not cost_master.empty:
        for _, r in cost_master.iterrows():
            cname = r["CostOrgName"].strip()
            ccode = r["CostOrgCode"].strip()
            le    = r["Legal Entity"].strip()
            ledgers = sorted(code_to_ledger.get(ccode, []))
            if ledgers:
                for L in ledgers:
                    # only assign BU if we can prove it—per requirement, keep blanks rather than guess
                    out_rows.append({
                        "Ledger Name": L,
                        "Legal Entity": le if le in legal_entity_names else le,  # keep whatever LE name we resolved (may be blank)
                        "Business Unit": "",   # not required for this sheet
                        "Cost Organization": cname if cname else ccode
                    })
            else:
                # orphan row (no ledger mapping)
                out_rows.append({
                    "Ledger Name": "",
                    "Legal Entity": le if le in legal_entity_names else le,
                    "Business Unit": "",
                    "Cost Organization": cname if cname else ccode
                })
    df_costorg = pd.DataFrame(out_rows if out_rows else [], columns=["Ledger Name","Legal Entity","Business Unit","Cost Organization"]).drop_duplicates().reset_index(drop=True)

    st.success(f"Sheet 2 (Cost Orgs): Built {len(df_costorg)} rows (row-per-cost-org; orphans included).")
    st.dataframe(df_costorg, use_container_width=True, height=320)

    # -------------------- Excel download (two sheets) --------------------
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Core_Ledger_LE_BU")
        df_costorg.to_excel(writer, index=False, sheet_name="Ledger_LE_CostOrg")
    st.download_button(
        "⬇️ Download Excel (EnterpriseStructure_v2.xlsx)",
        data=excel_buf.getvalue(),
        file_name="EnterpriseStructure_v2.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ===================== DRAW.IO DIAGRAM BLOCK (updated with Cost Org lane) =====================
    # ======= DRAW.IO DIAGRAM (bus on left, cost orgs on right) =======
    if "df" in locals() and isinstance(df, pd.DataFrame) and not df.empty:
        import xml.etree.ElementTree as ET
        import zlib, base64, uuid

        def _make_drawio_xml(df_core: pd.DataFrame, df_costorg: pd.DataFrame) -> str:
            # --- layout & spacing ---
            LEFT_PAD   = 260               # leave room for legend
            W, H       = 180, 48
            X_STEP     = 230
            PAD_GROUP  = 60
            RIGHT_PAD  = 160

            # vertical positions
            Y_LEDGER   = 170
            Y_LE       = 330
            Y_BU       = 490
            Y_COSTORG  = 560             # >>> NEW: Cost Orgs sit one row lower than BU
            BUS_Y      = 250              # LE↔Ledger bus

            # styles
            S_LEDGER = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE6E6;strokeColor=#C86868;fontSize=12;"
            S_LE     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE2C2;strokeColor=#A66000;fontSize=12;"
            S_BU     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFF1B3;strokeColor=#B38F00;fontSize=12;"
            S_COST   = "rounded=1;whiteSpace=wrap;html=1;fillColor=#DDEBFF;strokeColor=#3B82F6;fontSize=12;"  # >>> NEW blue

            S_EDGE_OTHER  = (
                "endArrow=block;rounded=1;"
                "edgeStyle=orthogonalEdgeStyle;orthogonal=1;jettySize=auto;"
                "strokeColor=#666666;"
                "exitX=0.5;exitY=0;entryX=0.5;entryY=1;"
            )
            S_EDGE_LEDGER = (
                "endArrow=block;rounded=1;"
                "edgeStyle=orthogonalEdgeStyle;orthogonal=1;jettySize=auto;"
                "strokeColor=#444444;"
                "exitX=0.5;exitY=0;entryX=0.5;entryY=1;"
            )
            S_EDGE_COST = (
                "endArrow=block;rounded=1;"
                "edgeStyle=orthogonalEdgeStyle;orthogonal=1;jettySize=auto;"
                "strokeColor=#3B82F6;"   # match cost blue
                "exitX=0.5;exitY=0;entryX=0.5;entryY=1;"
            )

            # --- normalize inputs ---
            df_core = df_core[["Ledger Name","Legal Entity","Business Unit"]].copy()
            for c in df_core.columns:
                df_core[c] = df_core[c].fillna("").map(str).str.strip()

            df_cost = df_costorg[["Ledger Name","Legal Entity","Cost Organization"]].copy() if not df_costorg.empty else pd.DataFrame(columns=["Ledger Name","Legal Entity","Cost Organization"])
            for c in df_cost.columns:
                df_cost[c] = df_cost[c].fillna("").map(str).str.strip()

            # ledgers and LEs used in diagram
            ledgers = sorted([x for x in df_core["Ledger Name"].unique() if x])
            # map ledger->LEs from core
            led_to_les = {}
            for _, r in df_core.iterrows():
                L,E = r["Ledger Name"], r["Legal Entity"]
                if L and E:
                    led_to_les.setdefault(L, set()).add(E)
            # add LEs that appear only in cost sheet (to draw their LE nodes too)
            for _, r in df_cost.iterrows():
                L,E = r["Ledger Name"], r["Legal Entity"]
                if L and E:
                    led_to_les.setdefault(L, set()).add(E)

            # BU placement (left lane)
            le_to_bus = {}
            for _, r in df_core.iterrows():
                L,E,B = r["Ledger Name"], r["Legal Entity"], r["Business Unit"]
                if L and E and B:
                    le_to_bus.setdefault((L,E), set()).add(B)

            # Cost Org placement (right lane) — allow orphans (no ledger or no LE)
            le_to_cost = {}
            orphan_cost = []  # (CostOrgName) with no L or no E
            for _, r in df_cost.iterrows():
                L,E,C = r["Ledger Name"], r["Legal Entity"], r["Cost Organization"]
                if C and L and E:
                    le_to_cost.setdefault((L,E), set()).add(C)
                elif C:
                    orphan_cost.append(C)

            # ---- compute X positions: left lane for BU, right lane for Cost Orgs
            next_x_left = LEFT_PAD
            next_x_right = LEFT_PAD

            led_x, le_x = {}, {}
            bu_x = {}
            cost_x = {}

            # First pass: allocate BU lane left, Cost lane right (after a separator)
            SEP = 180  # center gap between lanes

            for L in ledgers:
                les = sorted(list(led_to_les.get(L, set())))
                # BU lane (left)
                lane_left_xs = []
                for E in les:
                    buses = sorted(list(le_to_bus.get((L,E), set())))
                    # if no BU, reserve a spot for LE’s left lane center
                    if buses:
                        for b in buses:
                            if b not in bu_x:
                                bu_x[b] = next_x_left
                                next_x_left += X_STEP
                        lane_left_xs += [bu_x[b] for b in buses]
                    else:
                        # reserve a left-lane slot for symmetry
                        lane_left_xs.append(next_x_left)
                        next_x_left += X_STEP

                # Cost lane (right)
                # push right lane to start after max(left lane) + SEP for this ledger cluster
                start_right = max(next_x_left + SEP, LEFT_PAD + SEP)
                if next_x_right < start_right:
                    next_x_right = start_right

                lane_right_xs = []
                for E in les:
                    costs = sorted(list(le_to_cost.get((L,E), set())))
                    if costs:
                        for c in costs:
                            if c not in cost_x:
                                cost_x[c] = next_x_right
                                next_x_right += X_STEP
                        lane_right_xs += [cost_x[c] for c in costs]
                    else:
                        # reserve a right-lane slot for symmetry
                        lane_right_xs.append(next_x_right)
                        next_x_right += X_STEP

                # center LE above the midpoint between left & right lane spans
                all_xs = (lane_left_xs or []) + (lane_right_xs or [])
                if all_xs:
                    le_center = int(sum(all_xs) / len(all_xs))
                else:
                    # if no children at all, anchor LE on left lane progression
                    le_center = next_x_left
                    next_x_left += X_STEP
                le_x[(L,E)] = le_center

                # center ledger over its LEs
                # (we'll average LE centers for this ledger)
                xs_this_ledger = [le_x[(L, e)] for e in les] if les else [le_center]
                led_x[L] = int(sum(xs_this_ledger) / len(xs_this_ledger))

                # gap after each ledger cluster
                next_x_left += PAD_GROUP
                next_x_right += PAD_GROUP

            # allocate orphans (Cost Orgs with no L/LE) on the far right
            next_x_right += RIGHT_PAD
            for c in orphan_cost:
                if c not in cost_x:
                    cost_x[c] = next_x_right
                    next_x_right += X_STEP

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

            def add_vertex(label, style, x, y, w=W, h=H):
                vid = uuid.uuid4().hex[:8]
                c = ET.SubElement(root, "mxCell", attrib={"id": vid, "value": label, "style": style, "vertex": "1", "parent": "1"})
                ET.SubElement(c, "mxGeometry", attrib={"x": str(int(x)), "y": str(int(y)), "width": str(w), "height": str(h), "as": "geometry"})
                return vid

            def add_edge(src, tgt, style=S_EDGE_OTHER, points=None):
                eid = uuid.uuid4().hex[:8]
                c = ET.SubElement(root, "mxCell", attrib={
                    "id": eid, "value": "", "style": style, "edge": "1", "parent": "1",
                    "source": src, "target": tgt
                })
                g = ET.SubElement(c, "mxGeometry", attrib={"relative": "1", "as": "geometry"})
                if points:
                    arr = ET.SubElement(g, "Array", attrib={"as": "points"})
                    for px,py in points:
                        ET.SubElement(arr, "mxPoint", attrib={"x": str(int(px)), "y": str(int(py))})

            def add_bus_edge(src_id, src_center_x, tgt_id, tgt_center_x):
                add_edge(src_id, tgt_id, style=S_EDGE_LEDGER, points=[(src_center_x, BUS_Y), (tgt_center_x, BUS_Y)])

            # vertices
            id_map = {}
            # Ledgers
            for L in ledgers:
                id_map[("L", L)] = add_vertex(L, S_LEDGER, led_x[L], Y_LEDGER)

            # LEs
            for L, les in led_to_les.items():
                for E in sorted(list(les)):
                    id_map[("E", L, E)] = add_vertex(E, S_LE, le_x[(L,E)], Y_LE)

            # BUs (left lane)
            for (L,E), buses in le_to_bus.items():
                for b in sorted(list(buses)):
                    id_map[("B", b)] = add_vertex(b, S_BU, bu_x[b], Y_BU)

            # Cost Orgs (right lane, blue)
            for (L,E), costs in le_to_cost.items():
                for c in sorted(list(costs)):
                    id_map[("C", c)] = add_vertex(c, S_COST, cost_x[c], Y_COSTORG)

            # orphan Cost Orgs (no L/LE)
            for c in orphan_cost:
                id_map[("C", c)] = add_vertex(c, S_COST, cost_x[c], Y_COSTORG)

            # edges BU → LE (left lane)
            drawn = set()
            for (L,E), buses in le_to_bus.items():
                for b in sorted(list(buses)):
                    if (("B", b) in id_map) and (("E", L, E) in id_map):
                        k = ("B2E", b, L, E)
                        if k not in drawn:
                            add_edge(id_map[("B", b)], id_map[("E", L, E)], style=S_EDGE_OTHER)
                            drawn.add(k)

            # edges Cost Org → LE (right lane, blue)
            for (L,E), costs in le_to_cost.items():
                for c in sorted(list(costs)):
                    if (("C", c) in id_map) and (("E", L, E) in id_map):
                        k = ("C2E", c, L, E)
                        if k not in drawn:
                            add_edge(id_map[("C", c)], id_map[("E", L, E)], style=S_EDGE_COST)
                            drawn.add(k)

            # edges LE → Ledger via forced bus waypoints
            for L, les in led_to_les.items():
                for E in les:
                    if (("E", L, E) in id_map) and (("L", L) in id_map):
                        k = ("E2L", L, E)
                        if k not in drawn:
                            src_x_center = le_x[(L, E)] + W/2
                            tgt_x_center = led_x[L] + W/2
                            add_bus_edge(id_map[("E", L, E)], src_x_center, id_map[("L", L)], tgt_x_center)
                            drawn.add(k)

            # legend
            def add_legend(x=20, y=20):
                panel_w, panel_h = 230, 150
                panel = add_vertex("", "rounded=1;fillColor=#FFFFFF;strokeColor=#CBD5E1;", x, y, panel_w, panel_h)
                def swatch(lbl, color, gy):
                    box = add_vertex("", f"rounded=1;fillColor={color};strokeColor=#666666;", x+12, y+gy, 18, 12)
                    txt = add_vertex(lbl, "text;align=left;verticalAlign=middle;fontSize=12;", x+36, y+gy-4, 160, 20)
                swatch("Ledger", "#FFE6E6", 36)
                swatch("Legal Entity", "#FFE2C2", 62)
                swatch("Business Unit (left lane)", "#FFF1B3", 88)
                swatch("Cost Organization (right lane)", "#DDEBFF", 114)

            add_legend()
            return ET.tostring(mxfile, encoding="utf-8", method="xml").decode("utf-8")

        def _drawio_url_from_xml(xml: str) -> str:
            raw = zlib.compress(xml.encode("utf-8"), level=9)[2:-4]
            b64 = base64.b64encode(raw).decode("ascii")
            return f"https://app.diagrams.net/?title=EnterpriseStructure.drawio#R{b64}"

        _xml = _make_drawio_xml(df, df_costorg)

        st.download_button(
            "⬇️ Download diagram (.drawio)",
            data=_xml.encode("utf-8"),
            file_name="EnterpriseStructure.drawio",
            mime="application/xml",
            use_container_width=True
        )
        st.markdown(f"[🔗 Open in draw.io (preview)]({_drawio_url_from_xml(_xml)})")
        st.caption("Left lane = BU, Right lane = Cost Org (blue). LE bridges both; Ledger sits above via the bus. File → Save to persist.")

