import streamlit as st
import pandas as pd
import openpyxl
import requests
import json
import io
from datetime import datetime

st.set_page_config(page_title="BOQ Auto Pricer", page_icon="📊", layout="wide")

st.markdown("""
<style>
.main-header { background: #1e3a5f; color: white; padding: 1.2rem 2rem; border-radius: 12px; margin-bottom: 1.5rem; }
.main-header h1 { font-size: 24px; margin: 0; }
.main-header p  { font-size: 13px; opacity: 0.75; margin: 4px 0 0; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
  <h1>📊 BOQ Auto Pricer</h1>
  <p>ارفع أي ملف BOQ Excel — سعّره بالذكاء الاصطناعي — حمّل النتيجة</p>
</div>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ الإعدادات")

    api_key = st.text_input(
        "🔑 Gemini API Key (مجاني)",
        type="password",
        help="احصل عليه من aistudio.google.com/app/apikey"
    )
    st.caption("[احصل على مفتاح مجاني ←](https://aistudio.google.com/app/apikey)")
    st.divider()

    region = st.selectbox("📍 المنطقة", [
        "Madinah, Saudi Arabia", "Riyadh, Saudi Arabia", "Jeddah, Saudi Arabia",
        "Cairo, Egypt", "Alexandria, Egypt",
        "Dubai, UAE", "Abu Dhabi, UAE", "Doha, Qatar", "Kuwait City, Kuwait",
    ])
    currency = st.selectbox("💰 العملة", ["SAR", "EGP", "AED", "USD"])
    proj_type = st.selectbox("🏗 نوع المشروع", [
        "residential", "commercial", "mixed use", "industrial", "hospitality"
    ])
    quality = st.selectbox("⭐ مستوى الجودة", ["standard", "high-end", "luxury"], index=1)
    st.divider()
    st.caption("v3.0 — BOQ Auto Pricer")

# ── Smart extractor ───────────────────────────────────────────────
def extract_items(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    all_items = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if not ws.max_row or ws.max_row < 3:
            continue

        # Auto-detect header row and columns
        col_item = col_desc = col_unit = col_qty = -1
        header_row = -1

        for r in range(1, min(25, ws.max_row + 1)):
            for c in range(1, min(ws.max_column + 1, 20)):
                v = ws.cell(r, c).value
                if not v:
                    continue
                v_lo = str(v).lower().strip()
                if any(k in v_lo for k in ['item no', 'item no.', 'بند', 'رقم البند', 'no.']):
                    col_item = c; header_row = r
                if any(k in v_lo for k in ['description', 'desc', 'وصف', 'particular', 'work item', 'particulars']):
                    col_desc = c; header_row = r
                if v_lo in ['unit', 'وحدة', 'units', 'unit of']:
                    col_unit = c
                if any(k in v_lo for k in ['qty', 'quantity', 'quant', 'كمية', 'quantities']):
                    col_qty = c
            if header_row > 0 and col_desc > 0 and col_unit > 0 and col_qty > 0:
                break

        # Fallback to default columns
        if col_item < 0: col_item = 1
        if col_desc < 0: col_desc = 2
        if col_unit < 0: col_unit = 3
        if col_qty  < 0: col_qty  = 4
        start = header_row + 1 if header_row > 0 else 3

        # Read all rows
        rows = []
        for r in range(start, ws.max_row + 1):
            rows.append({
                'item': ws.cell(r, col_item).value,
                'desc': ws.cell(r, col_desc).value,
                'unit': ws.cell(r, col_unit).value,
                'qty':  ws.cell(r, col_qty).value,
            })

        # Extract items — handle multi-row pattern
        for i, row in enumerate(rows):
            if not row['unit'] or not row['qty']:
                continue
            try:
                qty_f = float(row['qty'])
                if qty_f <= 0:
                    continue
            except:
                continue

            desc    = str(row['desc'] or '').strip()
            item_no = str(row['item'] or '').strip()

            # Look backwards up to 5 rows for desc/item_no
            if not desc or not item_no:
                for j in range(i - 1, max(-1, i - 5), -1):
                    if not desc and rows[j]['desc']:
                        desc = str(rows[j]['desc']).strip()
                    if not item_no and rows[j]['item']:
                        item_no = str(rows[j]['item']).strip()
                    if desc and item_no:
                        break

            # Skip header/total/summary rows
            lo = (item_no + ' ' + desc).lower()
            skip = ['div.', 'total', 'summary', 'carried', 'subtotal',
                    'item no', 'description', 'unit price', 'remarks', 'collection']
            if any(k in lo for k in skip):
                continue

            all_items.append({
                'sheet':       sheet_name,
                'item_no':     item_no,
                'description': desc[:150],
                'unit':        str(row['unit']).strip(),
                'qty':         qty_f,
                'price':       0.0,
                'total':       0.0,
            })

    return all_items

# ── Gemini pricing ────────────────────────────────────────────────
def price_with_gemini(items, api_key, region, currency, proj_type, quality):
    # Try multiple models
    models = [
        "gemini-2.0-flash-lite",
        "gemini-1.5-flash",
        "gemini-1.5-flash-latest",
        "gemini-pro",
    ]

    BATCH = 20
    progress = st.progress(0)
    status   = st.empty()
    priced   = 0

    for start in range(0, len(items), BATCH):
        batch = items[start: start + BATCH]
        pct   = int(start / len(items) * 95)
        progress.progress(pct)
        status.info(f"⏳ جاري تسعير البنود {start+1}–{min(start+BATCH, len(items))} من {len(items)}...")

        numbered = "\n".join(
            f"{start+j}. {it['description']} (unit: {it['unit']}, qty: {it['qty']})"
            for j, it in enumerate(batch)
        )

        prompt = f"""You are a senior quantity surveyor expert in {region}.
Estimate realistic {quality} market unit prices in {currency} for a {proj_type} construction project in {region} (year 2025/2026).
Include labor, materials, equipment, and contractor overhead in all unit rates.

Items (index. description):
{numbered}

Return ONLY a valid JSON object mapping each index to its unit price as a number.
Example: {{"0": 185, "1": 95}}
No markdown, no explanation — just raw JSON."""

        success = False
        for model in models:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
            try:
                resp = requests.post(url, json={
                    "contents": [{"parts": [{"text": prompt}]}],
                    "generationConfig": {"temperature": 0.1, "maxOutputTokens": 2048}
                }, timeout=30)

                if resp.status_code == 404:
                    continue  # try next model

                resp.raise_for_status()
                text = resp.json()['candidates'][0]['content']['parts'][0]['text']
                text = text.replace('```json', '').replace('```', '').strip()
                s = text.find('{')
                e = text.rfind('}')
                if s >= 0 and e >= 0:
                    result = json.loads(text[s:e+1])
                    for k, v in result.items():
                        idx = int(k)
                        if 0 <= idx < len(items) and float(v) > 0:
                            items[idx]['price'] = float(v)
                            items[idx]['total'] = float(v) * items[idx]['qty']
                            priced += 1
                    success = True
                    break  # model worked, stop trying others

            except Exception as ex:
                if "404" not in str(ex):
                    st.warning(f"⚠ خطأ: {ex}")
                continue

        if not success:
            st.warning(f"⚠ فشلت دفعة البنود {start+1}–{min(start+BATCH, len(items))}")

    progress.progress(100)
    status.success(f"✅ تم تسعير {priced} بند من أصل {len(items)}")
    return items

# ── Export Excel ──────────────────────────────────────────────────
def to_excel(items, currency):
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BOQ Priced"

    headers = ['رقم البند', 'الشيت', 'الوصف', 'الوحدة', 'الكمية',
               f'سعر الوحدة ({currency})', f'الإجمالي ({currency})']

    hdr_font = Font(name='Arial', bold=True, color='FFFFFF', size=11)
    hdr_fill = PatternFill('solid', fgColor='1E3A5F')
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    ai_fill  = PatternFill('solid', fgColor='EFF6FF')
    num_font = Font(name='Arial', size=10, color='1E3A5F', bold=True)
    nrm_font = Font(name='Arial', size=10)

    for i, it in enumerate(items, 2):
        ws.cell(i, 1, it['item_no']).font      = nrm_font
        ws.cell(i, 2, it['sheet']).font        = nrm_font
        ws.cell(i, 3, it['description']).font  = nrm_font
        ws.cell(i, 4, it['unit']).font         = nrm_font
        ws.cell(i, 5, it['qty']).font          = num_font
        pc = ws.cell(i, 6, round(it['price'], 2))
        pc.font = num_font; pc.fill = ai_fill; pc.number_format = '#,##0.00'
        tc = ws.cell(i, 7, f'=F{i}*E{i}')
        tc.font = num_font; tc.fill = ai_fill; tc.number_format = '#,##0.00'

    last = len(items) + 2
    ws.cell(last, 6, 'الإجمالي الكلي').font = Font(name='Arial', bold=True, size=11)
    gc = ws.cell(last, 7, f'=SUM(G2:G{last-1})')
    gc.font = Font(name='Arial', bold=True, size=12, color='1E3A5F')
    gc.number_format = '#,##0.00'
    gc.fill = PatternFill('solid', fgColor='DBEAFE')

    for col, w in zip('ABCDEFG', [12, 18, 52, 10, 10, 20, 20]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Main UI ───────────────────────────────────────────────────────
uploaded = st.file_uploader("📂 ارفع ملف BOQ", type=['xlsx', 'xls'], label_visibility='collapsed')

if uploaded:
    file_bytes = uploaded.read()
    with st.spinner("جاري قراءة الملف..."):
        items = extract_items(file_bytes)

    if not items:
        st.error("⚠ لم يتم اكتشاف أي بنود — تأكد إن الملف يحتوي على أعمدة الوصف والوحدة والكمية")
    else:
        st.success(f"✅ تم اكتشاف **{len(items)} بند** من {len(set(i['sheet'] for i in items))} شيت")

        df_preview = pd.DataFrame(items)[['item_no', 'sheet', 'description', 'unit', 'qty']]
        df_preview.columns = ['رقم البند', 'الشيت', 'الوصف', 'الوحدة', 'الكمية']
        st.dataframe(df_preview, use_container_width=True, height=280)

        st.divider()

        if not api_key:
            st.warning("⚠ أدخل Gemini API Key في الشريط الجانبي أولاً")
        else:
            if st.button("✨ ابدأ التسعير التلقائي", use_container_width=True):
                items = price_with_gemini(items, api_key, region, currency, proj_type, quality)
                st.session_state['priced_items'] = items
                st.session_state['currency']     = currency

if 'priced_items' in st.session_state:
    items    = st.session_state['priced_items']
    currency = st.session_state['currency']

    st.divider()
    st.subheader("📋 النتائج")

    total_priced = sum(1 for it in items if it['price'] > 0)
    grand_total  = sum(it['price'] * it['qty'] for it in items)

    c1, c2, c3 = st.columns(3)
    c1.metric("إجمالي البنود",   len(items))
    c2.metric("بنود مسعّرة",     f"{total_priced} / {len(items)}")
    c3.metric(f"الإجمالي ({currency})", f"{grand_total:,.0f}")

    df = pd.DataFrame(items)[['item_no', 'description', 'unit', 'qty', 'price', 'total']]
    df.columns = ['رقم البند', 'الوصف', 'الوحدة', 'الكمية', 'سعر الوحدة', 'الإجمالي']

    st.data_editor(
        df,
        use_container_width=True,
        height=400,
        column_config={
            'سعر الوحدة': st.column_config.NumberColumn(format=f"%.2f {currency}", min_value=0),
            'الإجمالي':   st.column_config.NumberColumn(format=f"%.2f {currency}", disabled=True),
        },
        disabled=['رقم البند', 'الوصف', 'الوحدة', 'الكمية', 'الإجمالي']
    )

    excel_buf = to_excel(items, currency)
    st.download_button(
        label="📥 تحميل Excel",
        data=excel_buf,
        file_name=f"BOQ_priced_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
