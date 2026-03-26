"""
RFP Supplier Agent — Streamlit app
Deploy on share.streamlit.io
"""

import json, time, io, re
import streamlit as st
import anthropic
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="RFP Supplier Agent", page_icon="🔍", layout="wide")
st.markdown("<style>.block-container{padding-top:2rem;}</style>", unsafe_allow_html=True)

st.title("🔍 RFP Supplier Agent")
st.caption("Two-stage search: finds relevant supplier companies, then locates the right commercial contact at each one.")
st.divider()

with st.sidebar:
    st.header("Settings")
    api_key = st.text_input("Anthropic API Key", type="password", placeholder="sk-ant-api03-...")
    st.caption("Stored in session only — never saved to disk.")
    supplier_count_override = st.number_input(
        "Suppliers per RFP (override)", min_value=0, max_value=50, value=0,
        help="Set to 0 to use value from the Excel file. Start with 3-5 to test quality first."
    )
    st.divider()
    st.markdown("**How it works:**")
    st.markdown("**Stage 1** — Finds N relevant companies via web search")
    st.markdown("**Stage 2** — Finds the right commercial contact at each company")
    st.markdown("_(Sales Manager, KAM, BD Manager, Tender Manager...)_")
    st.divider()
    st.markdown("Tip: Start with 3 suppliers to verify quality before running 20.")


def normalize_header(text):
    s = str(text or "").strip().lower()
    s = re.sub(r"[^a-z0-9]+", "_", s)
    return s.strip("_")


def parse_excel(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))

    if "RFP_Input" not in wb.sheetnames:
        st.error('Sheet "RFP_Input" not found.')
        return [], []

    ws = wb["RFP_Input"]
    rows_raw = list(ws.iter_rows(values_only=True))

    header_row_idx = -1
    for i, row in enumerate(rows_raw):
        if normalize_header(row[0]) == "rfp_id":
            header_row_idx = i
            break

    if header_row_idx < 0:
        st.error('Header "RFP ID" not found in RFP_Input sheet.')
        return [], []

    headers = [normalize_header(c) for c in rows_raw[header_row_idx]]
    idx = {h: i for i, h in enumerate(headers)}

    def get(row, *keys):
        for k in keys:
            if k in idx:
                v = str(row[idx[k]] or "").strip()
                if v:
                    return v
        return ""

    rfp_rows = []
    for row in rows_raw[header_row_idx + 2:]:
        if not row or row[0] is None:
            continue
        rid = str(row[idx.get("rfp_id", 0)] or "").strip()
        if not rid or re.match(r"^(unique|e\.g\.|links|\s*$)", rid, re.I):
            continue
        sc = get(row, "supplier_count")
        try:
            supplier_count = int(float(sc)) if sc else 20
        except ValueError:
            supplier_count = 20

        rfp_rows.append({
            "rfp_id":         rid,
            "title":          get(row, "rfp_title", "title"),
            "description":    get(row, "description"),
            "country":        get(row, "country"),
            "region":         get(row, "region_city", "region", "city"),
            "phase":          get(row, "rfp_phase", "phase"),
            "keywords":       get(row, "keywords_tags", "keywords", "tags"),
            "supplier_count": supplier_count,
            "notes":          get(row, "notes"),
            "output_sheet":   get(row, "output_sheet") or ("Results_" + rid),
        })

    spec_rows = []
    if "SPEC_Items" in wb.sheetnames:
        ws2 = wb["SPEC_Items"]
        rows2 = list(ws2.iter_rows(values_only=True))
        h2_idx = -1
        for i, row in enumerate(rows2):
            if normalize_header(row[0]) == "rfp_id":
                h2_idx = i
                break
        if h2_idx >= 0:
            heads2 = [normalize_header(c) for c in rows2[h2_idx]]
            i2 = {h: i for i, h in enumerate(heads2)}
            for row in rows2[h2_idx + 2:]:
                if not row or row[0] is None:
                    continue
                sid = str(row[i2.get("rfp_id", 0)] or "").strip()
                if not sid or re.match(r"links to|rfp_id", sid, re.I):
                    continue
                spec_rows.append({
                    "rfp_id":      sid,
                    "item_name":   str(row[i2.get("item_name", 1)] or "").strip(),
                    "description": str(row[i2.get("item_description", 3)] or "").strip(),
                    "quantity":    str(row[i2.get("quantity", 4)] or "").strip(),
                    "unit":        str(row[i2.get("unit_of_measure", 5)] or "").strip(),
                })

    return rfp_rows, spec_rows


def call_claude(client, prompt, max_tokens, single_object):
    messages = [{"role": "user", "content": prompt}]
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=max_tokens,
        tools=[{"type": "web_search_20250305", "name": "web_search"}],
        messages=messages,
    )
    full_text = "".join(b.text for b in response.content if hasattr(b, "text"))

    if not full_text.strip():
        tool_results = [
            {"type": "tool_result", "tool_use_id": b.id, "content": "Search completed."}
            for b in response.content if b.type == "tool_use"
        ]
        if tool_results:
            follow_up = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=max_tokens,
                tools=[{"type": "web_search_20250305", "name": "web_search"}],
                messages=messages + [
                    {"role": "assistant", "content": response.content},
                    {"role": "user", "content": tool_results},
                ],
            )
            full_text = "".join(b.text for b in follow_up.content if hasattr(b, "text"))

    if not full_text.strip():
        raise ValueError("Agent returned empty response.")

    if single_object:
        m = re.search(r"\{[\s\S]*\}", full_text)
        if not m:
            raise ValueError(f"No JSON object found. Preview: {full_text[:200]}")
        return json.loads(m.group())
    else:
        m = re.search(r"\[[\s\S]*\]", full_text)
        if not m:
            raise ValueError(f"No JSON array found. Preview: {full_text[:200]}")
        result = json.loads(m.group())
        if not isinstance(result, list):
            raise ValueError("Response is not a list.")
        return result


def find_companies(client, rfp, specs, count):
    has_spec = len(specs) > 0
    spec_text = ""
    if has_spec:
        lines = []
        for i, s in enumerate(specs, 1):
            line = f"  {i}. {s['item_name']}"
            if s["quantity"]:
                line += f" | Qty: {s['quantity']} {s['unit']}"
            lines.append(line)
        spec_text = "\n\nSPECIFICATION:\n" + "\n".join(lines)

    prompt = f"""You are a procurement specialist. Find {count} real companies that can supply what is described.

TENDER:
- Title: {rfp['title']}
- Country: {rfp['country']}{(', ' + rfp['region']) if rfp['region'] else ''}
- Keywords: {rfp['keywords']}
- Description: {rfp['description'][:800]}{spec_text}
{'NOTE: No specification provided - use title, description, and keywords.' if not has_spec else ''}

Return ONLY a JSON array with exactly {count} objects. No explanation, no markdown.

[{{"company_name":"...","website":"https://...","country":"...","city":"...","company_description":"2-3 sentences","why_relevant":"1-2 sentences","relevance_score":8}}]"""

    return call_claude(client, prompt, 5000, False)


def find_contact(client, company, rfp):
    prompt = f"""Find a commercial contact person at this company.

COMPANY: {company['company_name']}
WEBSITE: {company.get('website', 'unknown')}
COUNTRY: {company.get('country', rfp['country'])}
CONTEXT: Inviting to a tender for: {rfp['title']} ({rfp['keywords']})

Look for: Sales Manager, Key Account Manager, Business Development Manager, Sales Director, Commercial Director, Tender Manager, Bid Manager.
NOT: secretary, HR, support, accountant.
Contact priority: 1) direct mobile  2) personal email (NOT info@)  3) LinkedIn  4) general contacts

Return ONLY a JSON object:
{{"contact_found":true,"contact_person":"First Last","job_title":"Sales Manager","phone":"","email":"","linkedin":"","general_email":"","general_phone":"","contact_note":"where found"}}"""

    return call_claude(client, prompt, 2000, True)


THIN = Side(style="thin", color="B8CCE4")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MID_BLUE = "2E5FA3"
DARK_BLUE = "1F3864"
LIGHT_BLUE = "D6E4F0"
WHITE = "FFFFFF"
GREY_BG = "F2F2F2"


def build_output_excel(original_bytes, rfp_rows, results):
    wb_orig = openpyxl.load_workbook(io.BytesIO(original_bytes))
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    for sheet_name in ["RFP_Input", "SPEC_Items"]:
        if sheet_name in wb_orig.sheetnames:
            ws_src = wb_orig[sheet_name]
            ws_dst = wb_out.create_sheet(sheet_name)
            ws_dst.sheet_properties.tabColor = "2E5FA3"
            for row in ws_src.iter_rows():
                for cell in row:
                    ws_dst.cell(row=cell.row, column=cell.column, value=cell.value)
            for col in ws_src.column_dimensions:
                ws_dst.column_dimensions[col].width = ws_src.column_dimensions[col].width

    HEADERS = [
        "#", "Company Name", "Website", "Country", "City",
        "Contact Person", "Job Title",
        "Direct Phone", "Direct Email", "LinkedIn",
        "General Email", "General Phone",
        "Company Description", "Why Relevant", "Relevance Score (1-10)",
        "Persona Match?", "Contact Notes",
    ]
    COL_WIDTHS = [4, 26, 26, 13, 15, 22, 24, 18, 28, 28, 24, 18, 36, 36, 12, 18, 30]

    for rfp in rfp_rows:
        suppliers = results.get(rfp["rfp_id"], [])
        sheet_name = rfp["output_sheet"][:31]
        ws = wb_out.create_sheet(sheet_name)
        ws.sheet_properties.tabColor = "27AE60"
        ws.sheet_view.showGridLines = False

        ws.merge_cells(f"A1:{get_column_letter(len(HEADERS))}1")
        banner = ws["A1"]
        banner.value = (
            f"SUPPLIER LIST FOR SSM  |  ID {rfp['rfp_id']}: {rfp['title']}"
            f"  |  {rfp['country']}"
            + (f"  |  Phase: {rfp['phase']}" if rfp["phase"] else "")
        )
        banner.font = Font(name="Arial", bold=True, color=WHITE, size=12)
        banner.fill = PatternFill("solid", start_color=DARK_BLUE)
        banner.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 28

        ws.merge_cells(f"A2:{get_column_letter(len(HEADERS))}2")
        ctx = ws["A2"]
        with_contact = sum(1 for s in suppliers if s.get("contact_found"))
        ctx.value = (
            f"Keywords: {rfp['keywords']}  |  "
            f"Companies found: {len(suppliers)}  |  "
            f"With persona contact: {with_contact}"
        )
        ctx.font = Font(name="Arial", italic=True, color="1F3864", size=9)
        ctx.fill = PatternFill("solid", start_color=LIGHT_BLUE)
        ctx.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[2].height = 20

        for ci, h in enumerate(HEADERS, 1):
            c = ws.cell(row=3, column=ci, value=h)
            c.font = Font(name="Arial", bold=True, color=WHITE, size=10)
            c.fill = PatternFill("solid", start_color=MID_BLUE)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = BORDER
            ws.column_dimensions[get_column_letter(ci)].width = COL_WIDTHS[ci - 1]
        ws.row_dimensions[3].height = 30

        if not suppliers:
            ws.cell(row=4, column=1).value = "No suppliers found — check the log for errors."
        else:
            for ri, s in enumerate(suppliers):
                r = 4 + ri
                bg = WHITE if ri % 2 == 0 else GREY_BG
                row_data = [
                    ri + 1,
                    s.get("company_name", ""),
                    s.get("website", ""),
                    s.get("country", ""),
                    s.get("city", ""),
                    s.get("contact_person", ""),
                    s.get("job_title", ""),
                    s.get("phone", ""),
                    s.get("email", ""),
                    s.get("linkedin", ""),
                    s.get("general_email", ""),
                    s.get("general_phone", ""),
                    s.get("company_description", ""),
                    s.get("why_relevant", ""),
                    s.get("relevance_score", ""),
                    "Yes" if s.get("contact_found") else "No (general contacts)",
                    s.get("contact_note", ""),
                ]
                for ci, val in enumerate(row_data, 1):
                    c = ws.cell(row=r, column=ci, value=val)
                    c.font = Font(name="Arial", size=10)
                    c.fill = PatternFill("solid", start_color=bg)
                    c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                    c.border = BORDER
                ws.row_dimensions[r].height = 40

        ws.freeze_panes = "B4"

    buf = io.BytesIO()
    wb_out.save(buf)
    return buf.getvalue()


# ── Main UI ──
st.subheader("Upload RFP File")
uploaded = st.file_uploader(
    "Drop your RFP_Supplier_Template.xlsx here",
    type=["xlsx", "xls"],
    label_visibility="collapsed",
)

rfp_rows, spec_rows, original_bytes = [], [], None

if uploaded:
    original_bytes = uploaded.read()
    rfp_rows, spec_rows = parse_excel(original_bytes)

    if rfp_rows:
        st.success(f"Found **{len(rfp_rows)} RFP(s)** · {len(spec_rows)} specification lines")
        import pandas as pd
        preview_data = []
        for r in rfp_rows:
            preview_data.append({
                "RFP ID":    r["rfp_id"],
                "Phase":     r["phase"] or "—",
                "Title":     r["title"][:60] + ("..." if len(r["title"]) > 60 else ""),
                "Country":   r["country"],
                "Keywords":  r["keywords"][:50] + ("..." if len(r["keywords"]) > 50 else ""),
                "Suppliers": r["supplier_count"],
            })
        st.dataframe(pd.DataFrame(preview_data), use_container_width=True, hide_index=True)

st.divider()

can_run = bool(rfp_rows) and bool(api_key) and api_key.startswith("sk-")

if not api_key:
    st.info("Enter your Anthropic API key in the sidebar to get started.")
elif not api_key.startswith("sk-"):
    st.error("API key must start with sk-ant-...")
elif not rfp_rows:
    st.info("Upload an RFP Excel file above to continue.")

if st.button("Find Suppliers", type="primary", disabled=not can_run):
    client = anthropic.Anthropic(api_key=api_key)
    results = {}
    overall_bar = st.progress(0, text="Starting...")
    log_area = st.empty()
    log_lines = []

    def add_log(msg):
        log_lines.append(msg)
        log_area.markdown(
            "<div style='background:#0d1117;border:1px solid #30363d;border-radius:6px;"
            "padding:12px 16px;font-family:monospace;font-size:12px;max-height:320px;"
            "overflow-y:auto;line-height:1.8;color:#8b949e'>"
            + "<br>".join(log_lines[-40:])
            + "</div>",
            unsafe_allow_html=True,
        )

    total = len(rfp_rows)
    for i, rfp in enumerate(rfp_rows):
        count = supplier_count_override or rfp["supplier_count"] or 20
        specs = [s for s in spec_rows if s["rfp_id"] == rfp["rfp_id"]]
        overall_bar.progress(i / total, text=f"Processing {rfp['rfp_id']} ({i+1}/{total})")

        add_log(f"[{i+1}/{total}] {rfp['rfp_id']} -- {rfp['title'] or '(no title)'}")
        add_log(f"  Keywords: {rfp['keywords'] or '(empty - check Excel file)'}")
        if specs:
            add_log(f"  Specification: {len(specs)} line items")
        else:
            add_log(f"  No specification -- using title & description")

        add_log(f"  [Stage 1] Searching {count} companies...")
        companies = []
        try:
            companies = find_companies(client, rfp, specs, count)
            add_log(f"  [Stage 1] Done: {len(companies)} companies found")
        except Exception as e:
            add_log(f"  [Stage 1] ERROR: {e}")
            results[rfp["rfp_id"]] = []
            if i < total - 1:
                time.sleep(5)
            continue

        add_log(f"  [Stage 2] Finding contacts for {len(companies)} companies...")
        enriched = []
        for j, co in enumerate(companies):
            try:
                contact = find_contact(client, co, rfp)
                enriched.append({**co, **contact})
                if contact.get("contact_found"):
                    add_log(f"    OK {co['company_name']}: {contact.get('contact_person')} ({contact.get('job_title')})")
                else:
                    add_log(f"    ~ {co['company_name']}: no persona -- general contacts saved")
            except Exception as e:
                add_log(f"    ERROR {co['company_name']}: {e}")
                enriched.append({
                    **co,
                    "contact_found": False, "contact_person": "", "job_title": "",
                    "email": "", "phone": "", "linkedin": "",
                    "general_email": "", "general_phone": "", "contact_note": str(e),
                })
            if j < len(companies) - 1:
                time.sleep(3)

        results[rfp["rfp_id"]] = enriched
        with_contact = sum(1 for s in enriched if s.get("contact_found"))
        add_log(f"  [Stage 2] Done: persona contacts {with_contact}/{len(enriched)}")

        if i < total - 1:
            add_log("  Pausing 5s before next RFP...")
            time.sleep(5)

    overall_bar.progress(1.0, text="Done!")
    add_log("=" * 50)
    add_log("All RFPs processed.")

    st.divider()
    st.subheader("Results")
    cols = st.columns(min(len(rfp_rows), 4))
    for i, rfp in enumerate(rfp_rows):
        sup = results.get(rfp["rfp_id"], [])
        wc = sum(1 for s in sup if s.get("contact_found"))
        with cols[i % len(cols)]:
            st.metric(
                label=rfp["rfp_id"],
                value=f"{len(sup)} companies",
                delta=f"{wc} with persona contact",
            )

    st.divider()
    with st.spinner("Building Excel file..."):
        out_bytes = build_output_excel(original_bytes, rfp_rows, results)

    from datetime import date
    st.download_button(
        label="Download Excel for SSM",
        data=out_bytes,
        file_name=f"RFP_Suppliers_{date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
