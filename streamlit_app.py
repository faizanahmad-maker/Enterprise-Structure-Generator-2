import io, zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Enterprise Structure Generator", page_icon="📊", layout="wide")
st.title("Enterprise Structure Generator — Excel + draw.io (BUs, Cost Orgs, Cost Books)")

st.markdown("""
Upload up to **6 Oracle export ZIPs** (any order):
- `Manage General Ledger` (Ledgers) → **GL_PRIMARY_LEDGER.csv**
- `Manage Legal Entities` (Legal Entities) → **XLE_ENTITY_PROFILE.csv**
- `Assign Legal Entities` (Ledger↔LE mapping) → **ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv**
- *(optional backup)* Journal config detail → **ORA_GL_JOURNAL_CONFIG_DETAIL.csv**
- `Manage Business Units` (Business Units) → **FUN_BUSINESS_UNIT.csv**
- `Manage Cost Organizations` (Cost Orgs) → **CST_COST_ORGANIZATION.csv**
- `Manage Cost Organization Relationships` (Cost Books) → **CST_COST_ORG_BOOK.csv**
""")

uploads = st.file_uploader("Drop your ZIPs here", type="zip", accept_multiple_files=True)

def read_csv_from_zip(zf, name):
    if name not in zf.namelist():
        return None
    with zf.open(name) as fh:
        return pd.read_csv(fh, dtype=str)

def pick_col(df, candidates):
    """
    Return the first matching column from `candidates` (list of exact names or substrings).
    - Exact match wins over substring.
    - Case-sensitive exact; then case-insensitive exact; then substring (case-insensitive).
    """
    cols = list(df.columns)
    # exact case-sensitive
    for c in candidates:
        if c in cols:
            return c
    # exact case-insensitive
    lower_map = {c.lower(): c for c in cols}
    for c in candidates:
        if c.lower() in lower_map:
            return lower_map[c.lower()]
    # substring (case-insensitive)
    for c in candidates:
        for existing in cols:
            if c.lower() in existing.lower():
                return existing
    return None

if not uploads:
    st.info("Upload your ZIPs to generate the Excel & diagram.")
else:
    # ------------ Collectors ------------
    ledger_names = set()                 # GL_PRIMARY_LEDGER.csv :: ORA_GL_PRIMARY_LEDGER_CONFIG.Name
    legal_entity_names = set()           # XLE_ENTITY_PROFILE.csv :: Name
    ledger_to_idents = {}                # ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv :: GL_LEDGER.Name -> {LegalEntityIdentifier}
    ident_to_le_name = {}                # XLE_ENTITY_PROFILE / ORA_GL_JOURNAL_CONFIG_DETAIL
    bu_rows = []                         # FUN_BUSINESS_UNIT.csv :: Name, PrimaryLedgerName, LegalEntityName

    # Cost Org master (with code + name + LE ident)
    # We'll keep both name and code so we can bind books via code reliably
    costorg_rows = []                    # {Name, LegalEntityIdentifier, CostOrgCode?}
    costorg_name_to_code = {}            # Name -> code (best effort)
    costorg_code_to_name = {}            # code -> Name (authoritative if present)

    # Cost Books (from relationships ZIP)
    books_by_costorg_code = {}           # CostOrgCode -> set([CostBookName,...])

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
                st.warning(f"`GL_PRIMARY_LEDGER.csv` missing `ORA_GL_PRIMARY_LEDGER_CONFIG.Name`. Found: {list(df.columns)}")

        # Legal Entities (primary ident -> name)
        df = read_csv_from_zip(z, "XLE_ENTITY_PROFILE.csv")
        if df is not None:
            name_col = pick_col(df, ["Name"])
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
                st.warning(f"`XLE_ENTITY_PROFILE.csv` missing needed cols. Found: {list(df.columns)}")

        # Ledger ↔ LE identifier
        df = read_csv_from_zip(z, "ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv")
        if df is not None:
            led_col = pick_col(df, ["GL_LEDGER.Name"])
            ident_col = pick_col(df, ["LegalEntityIdentifier"])
            if led_col and ident_col:
                for _, r in df[[led_col, ident_col]].dropna(how="all").iterrows():
                    led = str(r[led_col]).strip()
                    ident = str(r[ident_col]).strip()
                    if led and ident:
                        ledger_to_idents.setdefault(led, set()).add(ident)
            else:
                st.warning(f"`ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv` missing needed cols. Found: {list(df.columns)}")

        # Identifier ↔ LE name (backup)
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

        # Business Units
        df = read_csv_from_zip(z, "FUN_BUSINESS_UNIT.csv")
        if df is not None:
            bu_col   = pick_col(df, ["Name"])
            led_col  = pick_col(df, ["PrimaryLedgerName"])
            le_col   = pick_col(df, ["LegalEntityName"])
            if bu_col and led_col and le_col:
                for _, r in df[[bu_col, led_col, le_col]].dropna(how="all").iterrows():
                    bu_rows.append({
                        "Name": str(r[bu_col]).strip(),
                        "PrimaryLedgerName": str(r[led_col]).strip(),
                        "LegalEntityName": str(r[le_col]).strip()
                    })
            else:
                st.warning(f"`FUN_BUSINESS_UNIT.csv` missing needed cols. Found: {list(df.columns)}")

        # Cost Orgs (master list with code)
        df = read_csv_from_zip(z, "CST_COST_ORGANIZATION.csv")
        if df is not None:
            name_col  = pick_col(df, ["Name"])
            ident_col = pick_col(df, ["LegalEntityIdentifier"])
            code_col  = pick_col(df, ["CostOrgCode", "ORA_CST_ACCT_COST_ORG.CostOrgCode", "Cost Org Code"])
            need = [name_col, ident_col]
            if all(need):
                cols = [name_col, ident_col] + ([code_col] if code_col else [])
                for _, r in df[cols].dropna(how="all").iterrows():
                    name  = str(r[name_col]).strip() if name_col else ""
                    ident = str(r[ident_col]).strip() if ident_col else ""
                    code  = str(r[code_col]).strip() if code_col and pd.notna(r[code_col]) else ""
                    costorg_rows.append({"Name": name, "LegalEntityIdentifier": ident, "CostOrgCode": code})
                    if name and code:
                        costorg_name_to_code.setdefault(name, set()).add(code)
                    if code and name:
                        costorg_code_to_name[code] = name
            else:
                st.warning(f"`CST_COST_ORGANIZATION.csv` missing needed cols. Found: {list(df.columns)}")

        # Cost Books (relationships)
        df = read_csv_from_zip(z, "CST_COST_ORG_BOOK.csv")
        if df is not None:
            code_col = pick_col(df, ["ORA_CST_ACCT_COST_ORG.CostOrgCode", "CostOrgCode"])
            name_col = pick_col(df, ["Name", "CostBookName"])
            if code_col and name_col:
                for _, r in df[[code_col, name_col]].dropna(how="all").iterrows():
                    co_code = str(r[code_col]).strip()
                    book_nm = str(r[name_col]).strip()
                    if co_code and book_nm:
                        books_by_costorg_code.setdefault(co_code, set()).add(book_nm)
            else:
                st.warning(f"`CST_COST_ORG_BOOK.csv` missing needed cols. Found: {list(df.columns)}")

    # ------------ Build maps (identifier-first) ------------
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

    known_pairs = set()  # (ledger, LE name) from mapping table
    for led, idents in ledger_to_idents.items():
        for ident in idents:
            le_name = ident_to_le_name.get(ident, "").strip()
            if le_name:
                known_pairs.add((led, le_name))

    # Legacy name-based map (for cautious back-fill in Tab 1 only)
    le_to_ledgers_namekey = {}
    for led, le_set in ledger_to_le_names.items():
        for le in le_set:
            le_to_ledgers_namekey.setdefault(le, set()).add(led)

    # ===================================================
    # Tab 1: Ledger – Legal Entity – Business Unit
    # ===================================================
    rows1 = []
    seen_triples = set()
    seen_ledgers_with_bu = set()

    # 1) BU-driven rows with smart back-fill (using name only if unique)
    for r in bu_rows:
        bu  = r["Name"]
        led = r["PrimaryLedgerName"] if r["PrimaryLedgerName"] in ledger_names else ""
        le  = r["LegalEntityName"]  if r["LegalEntityName"]  in legal_entity_names else ""

        if not led and le and le in le_to_ledgers_namekey and len(le_to_ledgers_namekey[le]) == 1:
            led = next(iter(le_to_ledgers_namekey[le]))
        if not le and led and led in ledger_to_le_names and len(ledger_to_le_names[led]) == 1:
            le = next(iter(ledger_to_le_names[led]))

        rows1.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": bu})
        seen_triples.add((led, le, bu))
        if led:
            seen_ledgers_with_bu.add(led)

    # 2) Ledger–LE pairs with no BU (from mapping)
    seen_pairs = {(a, b) for (a, b, _) in seen_triples}
    for led, le in sorted(known_pairs):
        if (led, le) not in seen_pairs:
            rows1.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": ""})

    # 3) Orphan ledgers in master list (no mapping & no BUs)
    mapped_ledgers = set(ledger_to_le_names.keys())
    for led in sorted(ledger_names - mapped_ledgers - seen_ledgers_with_bu):
        rows1.append({"Ledger Name": led, "Legal Entity": "", "Business Unit": ""})

    # 4) TRUE UNASSIGNED LEs (present in XLE profile, but nowhere else)
    le_names_in_pairs = {le for (_, le) in known_pairs}
    le_names_in_bu    = {r["LegalEntityName"] for r in bu_rows if r.get("LegalEntityName")}
    for le in sorted(legal_entity_names - le_names_in_pairs - le_names_in_bu):
        rows1.append({"Ledger Name": "", "Legal Entity": le, "Business Unit": ""})

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
    # Tab 2: Ledger – Legal Entity – Cost Organization – Cost Book
    # ===================================================
    rows2 = []
    seen_pairs2 = set()   # (ledger, LE) we've already emitted via CO rows

    # Build quick lookup: cost org name -> book list via code(s)
    def books_for_costorg_name(co_name: str):
        co_name = (co_name or "").strip()
        if not co_name:
            return []
        codes = sorted(costorg_name_to_code.get(co_name, []))
        acc = set()
        for c in codes:
            acc |= set(books_by_costorg_code.get(c, []))
        # Fallback: if no code mapping, try reverse (rare)
        if not acc:
            # if a single code maps back to this name
            for c, nm in costorg_code_to_name.items():
                if nm == co_name and c in books_by_costorg_code:
                    acc |= set(books_by_costorg_code[c])
        return sorted(acc)

    # 0) Base pairs coming from Tab 1 (so both tabs align)
    base_pairs = {
        (r["Ledger Name"], r["Legal Entity"])
        for _, r in df1.iterrows()
        if str(r["Ledger Name"]).strip() and str(r["Legal Entity"]).strip()
    }

    # 1) From cost orgs → ident → ledgers (one row per ledger if identifier maps to multiple ledgers)
    for r in costorg_rows:
        co = r.get("Name", "").strip()
        ident = r.get("LegalEntityIdentifier", "").strip()
        le = ident_to_le_name.get(ident, "").strip()
        leds = ident_to_ledgers.get(ident, set())
        cb_names = "; ".join(books_for_costorg_name(co)) if co else ""

        if leds:
            for led in sorted(leds):
                rows2.append({
                    "Ledger Name": led, "Legal Entity": le,
                    "Cost Organization": co, "Cost Book": cb_names
                })
                seen_pairs2.add((led, le))
        else:
            # no mapping to any ledger found; emit with blank ledger so user sees the orphan
            rows2.append({
                "Ledger Name": "", "Legal Entity": le,
                "Cost Organization": co, "Cost Book": cb_names
            })

    # 2) Ensure *all* Ledger–LE pairs from Tab 1 appear here (even if no COs / no mapping)
    for led, le in sorted(base_pairs):
        if (led, le) not in seen_pairs2:
            rows2.append({
                "Ledger Name": led, "Legal Entity": le,
                "Cost Organization": "", "Cost Book": ""
            })
            seen_pairs2.add((led, le))

    # 3) Also add mapping-known pairs that didn't appear yet (paranoia/safety)
    for led, le in sorted(known_pairs):
        if (led, le) not in seen_pairs2:
            rows2.append({
                "Ledger Name": led, "Legal Entity": le,
                "Cost Organization": "", "Cost Book": ""
            })
            seen_pairs2.add((led, le))

    # 4) Completely orphan ledgers (exist in masters, but no CO rows and no base pair)
    seen_ledgers_any_co = {row["Ledger Name"] for row in rows2 if row["Ledger Name"]}
    for led in sorted(ledger_names - seen_ledgers_any_co):
        rows2.append({
            "Ledger Name": led, "Legal Entity": "",
            "Cost Organization": "", "Cost Book": ""
        })

    df2 = pd.DataFrame(rows2).drop_duplicates().reset_index(drop=True)
    df2["__LedgerEmpty"] = (df2["Ledger Name"] == "").astype(int)
    df2 = (
        df2.sort_values(
            ["__LedgerEmpty", "Ledger Name", "Legal Entity", "Cost Organization", "Cost Book"],
            ascending=[True, True, True, True, True]
        )
        .drop(columns="__LedgerEmpty")
        .reset_index(drop=True)
    )
    df2.insert(0, "Assignment", range(1, len(df2) + 1))

    # ------------ Excel Output ------------
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        df1.to_excel(writer, index=False, sheet_name="Ledger_LE_BU_Assignments")
        df2.to_excel(writer, index=False, sheet_name="Ledger_LE_CostOrg_Books")

    st.success(f"Built {len(df1)} BU rows and {len(df2)} Cost Org/Book rows.")
    st.dataframe(df1.head(25), use_container_width=True, height=280)
    st.dataframe(df2.head(25), use_container_width=True, height=280)

    st.download_button(
        "⬇️ Download Excel (EnterpriseStructure.xlsx)",
        data=excel_buf.getvalue(),
        file_name="EnterpriseStructure.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ===================== DRAW.IO DIAGRAM BLOCK (with parking lots + Cost Books) =====================
    if (
        "df1" in locals() and isinstance(df1, pd.DataFrame) and not df1.empty and
        "df2" in locals() and isinstance(df2, pd.DataFrame)
    ):
        import xml.etree.ElementTree as ET
        import zlib, base64, uuid

        def _make_drawio_xml(df_bu: pd.DataFrame, df_co_books: pd.DataFrame) -> str:
            # --- layout & spacing ---
            W, H       = 180, 48
            X_STEP     = 230
            PAD_GROUP  = 60
            LEFT_PAD   = 260
            RIGHT_PAD  = 200

            Y_LEDGER   = 150
            Y_LE       = 310
            Y_BU       = 470
            Y_CO       = 630   # cost orgs row
            Y_CB       = 790   # cost book row (new)

            # --- styles ---
            S_LEDGER = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE6E6;strokeColor=#C86868;fontSize=12;"
            S_LE     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE2C2;strokeColor=#A66000;fontSize=12;"
            S_BU     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFF1B3;strokeColor=#B38F00;fontSize=12;"
            S_CO     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#E2F7E2;strokeColor=#3D8B3D;fontSize=12;"
            S_CB     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#CFF5CF;strokeColor=#2F7D2F;fontSize=12;"  # deeper green
            S_EDGE   = ("endArrow=block;rounded=1;edgeStyle=orthogonalEdgeStyle;orthogonal=1;"
                        "jettySize=auto;strokeColor=#666666;exitX=0.5;exitY=0;entryX=0.5;entryY=1;")
            S_HDR    = "text;align=left;verticalAlign=middle;fontSize=13;fontStyle=1;"

            # --- normalize input ---
            df_bu = df_bu[["Ledger Name", "Legal Entity", "Business Unit"]].copy()
            df_co_books = df_co_books[["Ledger Name", "Legal Entity", "Cost Organization", "Cost Book"]].copy()
            for df in (df_bu, df_co_books):
                for c in df.columns:
                    df[c] = df[c].fillna("").map(str).str.strip()

            ledgers_all = sorted([x for x in set(df_bu["Ledger Name"]) | set(df_co_books["Ledger Name"]) if x])

            # maps
            le_map = {}
            for _, r in pd.concat([df_bu[["Ledger Name","Legal Entity"]],
                                   df_co_books[["Ledger Name","Legal Entity"]]]).drop_duplicates().iterrows():
                if r["Ledger Name"] and r["Legal Entity"]:
                    le_map.setdefault(r["Ledger Name"], set()).add(r["Legal Entity"])

            bu_map = {}
            for _, r in df_bu.iterrows():
                if r["Ledger Name"] and r["Legal Entity"] and r["Business Unit"]:
                    bu_map.setdefault((r["Ledger Name"], r["Legal Entity"]), set()).add(r["Business Unit"])

            co_map = {}
            for _, r in df_co_books.iterrows():
                if r["Ledger Name"] and r["Legal Entity"] and r["Cost Organization"]:
                    co_map.setdefault((r["Ledger Name"], r["Legal Entity"]), set()).add(r["Cost Organization"])

            # cost books map (under a specific Cost Org label)
            cb_map = {}
            for _, r in df_co_books.iterrows():
                L = r["Ledger Name"]; E = r["Legal Entity"]; C = r["Cost Organization"]; B = r["Cost Book"]
                if L and E and C and B:
                    # split if multiple books joined by "; "
                    for book in [b.strip() for b in B.split(";") if b.strip()]:
                        cb_map.setdefault((L, E, C), set()).add(book)

            # -------- parking-lot sets --------
            orphan_ledgers = sorted([L for L in ledgers_all if not le_map.get(L)])
            unassigned_les = sorted(
                set(df_bu.loc[(df_bu["Ledger Name"] == "") & (df_bu["Legal Entity"] != ""), "Legal Entity"].unique())
                | set(df_co_books.loc[(df_co_books["Ledger Name"] == "") & (df_co_books["Legal Entity"] != ""), "Legal Entity"].unique())
            )
            all_bus = set(df_bu.loc[df_bu["Business Unit"] != "", "Business Unit"].unique())
            assigned_bus = set(
                df_bu.loc[
                    (df_bu["Ledger Name"] != "") & (df_bu["Legal Entity"] != "") & (df_bu["Business Unit"] != ""),
                    "Business Unit"
                ].unique()
            )
            unassigned_bus = sorted(all_bus - assigned_bus)

            # --- x-coordinates ---
            next_x = LEFT_PAD
            led_x, le_x, bu_x, co_x, cb_x = {}, {}, {}, {}, {}

            for L in ledgers_all:
                if L in orphan_ledgers:
                    continue  # show in parking lot, not grid
                les = sorted(le_map.get(L, []))
                for le in les:
                    buses = sorted(bu_map.get((L, le), []))
                    cos   = sorted(co_map.get((L, le), []))
                    # pre-allocate children X for BUs and COs
                    all_children = (buses + cos) if (buses or cos) else [le]

                    for child in all_children:
                        if (child in buses) and (child not in bu_x):
                            bu_x[child] = next_x
                            next_x += X_STEP
                        elif (child in cos) and (child not in co_x):
                            co_x[child] = next_x
                            next_x += X_STEP

                    if buses or cos:
                        xs = []
                        if buses: xs += [bu_x[b] for b in buses]
                        if cos:   xs += [co_x[c] for c in cos]
                        le_x[(L, le)] = int(sum(xs)/len(xs))
                    else:
                        le_x[(L, le)] = next_x
                        next_x += X_STEP

                    # cost books sit under each cost org; allocate beneath its X, expanding if multiple
                    for c in cos:
                        books = sorted(cb_map.get((L, le, c), []))
                        if books:
                            # center books around the cost org X if multiple
                            base_x = co_x[c]
                            start_x = base_x - (len(books)-1) * (X_STEP//2)
                            for i, bk in enumerate(books):
                                cb_x[(L, le, c, bk)] = start_x + i*(X_STEP)
                        # else: no CBs; nothing to place

                if les:
                    xs = [le_x[(L, le)] for le in les]
                    led_x[L] = int(sum(xs)/len(xs))
                next_x += PAD_GROUP

            # allocate parking lots to the right
            next_x += RIGHT_PAD
            # Unassigned LEs
            for e in unassigned_les:
                le_x[("UNASSIGNED", e)] = next_x
                next_x += X_STEP
            next_x += PAD_GROUP
            # Unassigned BUs
            for b in unassigned_bus:
                if b not in bu_x:
                    bu_x[b] = next_x
                    next_x += X_STEP
            # Orphan Ledgers (place last)
            next_x += PAD_GROUP
            for L in orphan_ledgers:
                led_x[("ORPHAN", L)] = next_x
                next_x += X_STEP

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

            def add_edge(src, tgt):
                eid = uuid.uuid4().hex[:8]
                c = ET.SubElement(root, "mxCell", attrib={
                    "id": eid, "value": "", "style": S_EDGE, "edge": "1", "parent": "1",
                    "source": src, "target": tgt})
                ET.SubElement(c, "mxGeometry", attrib={"relative": "1", "as": "geometry"})

            def add_text(text, x, y):
                tid = uuid.uuid4().hex[:8]
                t = ET.SubElement(root, "mxCell", attrib={
                    "id": tid, "value": text, "style": S_HDR, "vertex": "1", "parent": "1"})
                ET.SubElement(t, "mxGeometry", attrib={
                    "x": str(int(x)), "y": str(int(y)), "width": "260", "height": "22", "as": "geometry"})
                return tid

            # --- vertices ---
            id_map = {}
            # main grid
            for L in ledgers_all:
                if L in orphan_ledgers:
                    continue
                id_map[("L", L)] = add_vertex(L, S_LEDGER, led_x[L], Y_LEDGER)
                for le in sorted(le_map.get(L, [])):
                    id_map[("E", L, le)] = add_vertex(le, S_LE, le_x[(L, le)], Y_LE)
                    for b in sorted(bu_map.get((L, le), [])):
                        id_map[("B", L, le, b)] = add_vertex(b, S_BU, bu_x[b], Y_BU)
                    for c in sorted(co_map.get((L, le), [])):
                        id_map[("C", L, le, c)] = add_vertex(c, S_CO, co_x[c], Y_CO)
                        # books for this cost org
                        for bk in sorted(cb_map.get((L, le, c), [])):
                            id_map[("CB", L, le, c, bk)] = add_vertex(bk, S_CB, cb_x.get((L, le, c, bk), co_x[c]), Y_CB)

            # parking lot headers + vertices
            if unassigned_les:
                add_text("Unassigned LEs", le_x[("UNASSIGNED", unassigned_les[0])] - 40, Y_LE - 40)
                for e in unassigned_les:
                    id_map[("E_UN", e)] = add_vertex(e, S_LE, le_x[("UNASSIGNED", e)], Y_LE)

            if unassigned_bus:
                add_text("Unassigned BUs", bu_x[unassigned_bus[0]] - 40, Y_BU - 40)
                for b in unassigned_bus:
                    id_map[("B_UN", b)] = add_vertex(b, S_BU, bu_x[b], Y_BU)

            if orphan_ledgers:
                # place header roughly above their area
                any_x = led_x[("ORPHAN", orphan_ledgers[0])]
                add_text("Orphan Ledgers", any_x - 60, Y_LEDGER - 40)
                for L in orphan_ledgers:
                    id_map[("L_ORPHAN", L)] = add_vertex(L, S_LEDGER, led_x[("ORPHAN", L)], Y_LEDGER)

            # --- edges (assigned only, child→parent) ---
            for L in ledgers_all:
                if L in orphan_ledgers:
                    continue
                for le in sorted(le_map.get(L, [])):
                    if ("E", L, le) in id_map and ("L", L) in id_map:
                        add_edge(id_map[("E", L, le)], id_map[("L", L)])
                    for b in sorted(bu_map.get((L, le), [])):
                        if ("B", L, le, b) in id_map:
                            add_edge(id_map[("B", L, le, b)], id_map[("E", L, le)])
                    for c in sorted(co_map.get((L, le), [])):
                        if ("C", L, le, c) in id_map:
                            add_edge(id_map[("C", L, le, c)], id_map[("E", L, le)])
                            # books up to CO
                            for bk in sorted(cb_map.get((L, le, c), [])):
                                if ("CB", L, le, c, bk) in id_map:
                                    add_edge(id_map[("CB", L, le, c, bk)], id_map[("C", L, le, c)])

            # Legend
            def add_legend(x=20, y=20):
                def swatch(lbl, color, offset):
                    box = ET.SubElement(root, "mxCell", attrib={
                        "id": uuid.uuid4().hex[:8], "value": "",
                        "style": f"rounded=1;fillColor={color};strokeColor=#666666;",
                        "vertex": "1", "parent": "1"})
                    ET.SubElement(box, "mxGeometry", attrib={
                        "x": str(x+12), "y": str(y+offset), "width": "18", "height": "12", "as": "geometry"})
                    txt = ET.SubElement(root, "mxCell", attrib={
                        "id": uuid.uuid4().hex[:8], "value": lbl,
                        "style": "text;align=left;verticalAlign=middle;fontSize=12;",
                        "vertex": "1", "parent": "1"})
                    ET.SubElement(txt, "mxGeometry", attrib={
                        "x": str(x+36), "y": str(y+offset-4), "width": "170", "height": "20", "as": "geometry"})
                swatch("Ledger", "#FFE6E6", 36)
                swatch("Legal Entity", "#FFE2C2", 62)
                swatch("Business Unit", "#FFF1B3", 88)
                swatch("Cost Org", "#E2F7E2", 114)
                swatch("Cost Book", "#CFF5CF", 140)  # NEW
            add_legend()

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
