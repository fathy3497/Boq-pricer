import streamlit as st
import pandas as pd
import openpyxl
import requests
import json
import io
from datetime import datetime


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="BOQ Auto Pricer",
    page_icon="📊",
    layout="wide"
)


# ============================================================
# STYLING
# ============================================================

st.markdown("""
<style>

[data-testid="stAppViewContainer"] {
    background: #f0f4f8;
}

.main-header {
    background: #1e3a5f;
    color: white;
    padding: 1.2rem 2rem;
    border-radius: 12px;
    margin-bottom: 1.5rem;
}

.main-header h1 {
    font-size: 24px;
    margin: 0;
}

.main-header p {
    font-size: 13px;
    opacity: 0.75;
    margin: 4px 0 0;
}

.metric-card {
    background: white;
    border: 1px solid #e5e7eb;
    border-radius: 10px;
    padding: 1rem;
    text-align: center;
}

.metric-card .val {
    font-size: 22px;
    font-weight: 700;
    color: #1e3a5f;
}

.metric-card .lbl {
    font-size: 12px;
    color: #9ca3af;
    margin-top: 4px;
}

.stButton > button {
    background: #1e3a5f !important;
    color: white !important;
    border-radius: 8px !important;
    border: none !important;
    font-weight: 600 !important;
}

.stButton > button:hover {
    background: #16304f !important;
}

</style>
""", unsafe_allow_html=True)


# ============================================================
# HEADER
# ============================================================

st.markdown("""
<div class="main-header">
    <h1>📊 BOQ Auto Pricer</h1>
    <p>
        ارفع أي ملف BOQ Excel — حلّل البنود بالذكاء الاصطناعي —
        احصل على أسعار تقديرية — حمّل النتيجة
    </p>
</div>
""", unsafe_allow_html=True)


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header("⚙️ الإعدادات")

    api_key = st.text_input(
        "🔑 Gemini API Key",
        type="password",
        help="احصل على المفتاح من Google AI Studio"
    )

    st.caption(
        "احصل على المفتاح من: "
        "[Google AI Studio](https://aistudio.google.com/app/apikey)"
    )

    st.divider()

    region = st.selectbox(
        "📍 المنطقة",
        [
            "Madinah, Saudi Arabia",
            "Riyadh, Saudi Arabia",
            "Jeddah, Saudi Arabia",
            "Dammam, Saudi Arabia",
            "Makkah, Saudi Arabia",
            "Cairo, Egypt",
            "Alexandria, Egypt",
            "Dubai, UAE",
            "Abu Dhabi, UAE",
            "Doha, Qatar",
            "Kuwait City, Kuwait"
        ]
    )

    currency = st.selectbox(
        "💰 العملة",
        ["SAR", "EGP", "AED", "USD"]
    )

    proj_type = st.selectbox(
        "🏗 نوع المشروع",
        [
            "residential",
            "commercial",
            "mixed use",
            "industrial",
            "hospitality",
            "infrastructure"
        ]
    )

    quality = st.selectbox(
        "⭐ مستوى الجودة",
        [
            "standard",
            "high-end",
            "luxury"
        ],
        index=1
    )

    st.divider()

    st.caption("v2.1 — AI-Powered BOQ Pricer")


# ============================================================
# BOQ EXTRACTION
# ============================================================

def extract_items(file_bytes):

    wb = openpyxl.load_workbook(
        io.BytesIO(file_bytes),
        data_only=True
    )

    all_items = []

    for sheet_name in wb.sheetnames:

        ws = wb[sheet_name]

        if ws.max_row is None or ws.max_row < 3:
            continue

        # ----------------------------------------------------
        # Find Header
        # ----------------------------------------------------

        col_item = -1
        col_desc = -1
        col_unit = -1
        col_qty = -1
        header_row = -1

        for r in range(
            1,
            min(25, ws.max_row + 1)
        ):

            row_found = False

            for c in range(
                1,
                min(ws.max_column + 1, 20)
            ):

                value = ws.cell(r, c).value

                if value is None:
                    continue

                value_lower = str(value).lower().strip()

                # Item number
                if any(
                    keyword in value_lower
                    for keyword in [
                        "item no",
                        "item no.",
                        "item number",
                        "بند",
                        "رقم البند"
                    ]
                ):
                    col_item = c
                    header_row = r
                    row_found = True

                # Description
                if any(
                    keyword in value_lower
                    for keyword in [
                        "description",
                        "desc",
                        "وصف",
                        "particular",
                        "work item",
                        "scope"
                    ]
                ):
                    col_desc = c
                    header_row = r
                    row_found = True

                # Unit
                if value_lower in [
                    "unit",
                    "وحدة",
                    "units"
                ]:
                    col_unit = c
                    header_row = r
                    row_found = True

                # Quantity
                if any(
                    keyword in value_lower
                    for keyword in [
                        "qty",
                        "quantity",
                        "quant",
                        "كمية",
                        "quantities"
                    ]
                ):
                    col_qty = c
                    header_row = r
                    row_found = True

            if (
                row_found
                and col_desc > 0
                and col_unit > 0
                and col_qty > 0
            ):
                break

        # ----------------------------------------------------
        # Fallback
        # ----------------------------------------------------

        if col_item < 0:
            col_item = 1

        if col_desc < 0:
            col_desc = 2

        if col_unit < 0:
            col_unit = 3

        if col_qty < 0:
            col_qty = 4

        start_row = (
            header_row + 1
            if header_row > 0
            else 3
        )

        # ----------------------------------------------------
        # Read Rows
        # ----------------------------------------------------

        rows = []

        for r in range(
            start_row,
            ws.max_row + 1
        ):

            rows.append({

                "r": r,

                "item": ws.cell(
                    r,
                    col_item
                ).value,

                "desc": ws.cell(
                    r,
                    col_desc
                ).value,

                "unit": ws.cell(
                    r,
                    col_unit
                ).value,

                "qty": ws.cell(
                    r,
                    col_qty
                ).value

            })

        # ----------------------------------------------------
        # Process Items
        # ----------------------------------------------------

        for i, row in enumerate(rows):

            if not row["unit"] or not row["qty"]:
                continue

            # Quantity must be numeric
            try:

                qty_f = float(
                    str(row["qty"]).replace(",", "")
                )

            except Exception:
                continue

            if qty_f <= 0:
                continue

            desc = str(
                row["desc"] or ""
            ).strip()

            item_no = str(
                row["item"] or ""
            ).strip()

            # ------------------------------------------------
            # Look backwards for missing values
            # ------------------------------------------------

            if not desc or not item_no:

                for j in range(
                    i - 1,
                    max(-1, i - 5),
                    -1
                ):

                    if (
                        not desc
                        and rows[j]["desc"]
                    ):
                        desc = str(
                            rows[j]["desc"]
                        ).strip()

                    if (
                        not item_no
                        and rows[j]["item"]
                    ):
                        item_no = str(
                            rows[j]["item"]
                        ).strip()

                    if desc and item_no:
                        break

            # ------------------------------------------------
            # Skip non-item rows
            # ------------------------------------------------

            lo = (
                item_no + " " + desc
            ).lower()

            skip_keywords = [
                "div.",
                "total",
                "summary",
                "carried",
                "subtotal",
                "item no",
                "description",
                "unit price",
                "remarks",
                "grand total",
                "المجموع",
                "الإجمالي"
            ]

            if any(
                keyword in lo
                for keyword in skip_keywords
            ):
                continue

            # ------------------------------------------------
            # Add Item
            # ------------------------------------------------

            all_items.append({

                "sheet": sheet_name,

                "item_no": item_no,

                "description": desc[:300],

                "unit": str(
                    row["unit"]
                ).strip(),

                "qty": qty_f,

                "price": 0.0,

                "total": 0.0

            })

    return all_items


# ============================================================
# GEMINI PRICING
# ============================================================

def price_with_gemini(
    items,
    api_key,
    region,
    currency,
    proj_type,
    quality
):

    BATCH_SIZE = 20

    # --------------------------------------------------------
    # Correct Gemini API URL
    # --------------------------------------------------------

    url = (
        "https://generativelanguage.googleapis.com/"
        "v1beta/models/gemini-1.5-flash:generateContent"
        f"?key={api_key}"
    )

    progress = st.progress(0)
    status = st.empty()

    priced = 0
    total_items = len(items)

    # --------------------------------------------------------
    # Process batches
    # --------------------------------------------------------

    for batch_start in range(
        0,
        total_items,
        BATCH_SIZE
    ):

        batch = items[
            batch_start:
            batch_start + BATCH_SIZE
        ]

        batch_end = min(
            batch_start + BATCH_SIZE,
            total_items
        )

        progress_pct = int(
            batch_start / total_items * 100
        )

        progress.progress(progress_pct)

        status.info(
            f"⏳ جاري تسعير البنود "
            f"{batch_start + 1}–{batch_end} "
            f"من {total_items}..."
        )

        # ----------------------------------------------------
        # Local batch numbering
        # ----------------------------------------------------

        numbered_items = []

        for local_index, item in enumerate(batch):

            numbered_items.append(
                f"{local_index}: "
                f"{item['description']} "
                f"(unit: {item['unit']}, "
                f"qty: {item['qty']})"
            )

        numbered = "\n".join(
            numbered_items
        )

        # ----------------------------------------------------
        # Prompt
        # ----------------------------------------------------

        prompt = f"""
You are a senior Quantity Surveyor and Construction Cost Estimator
working in {region}.

Estimate realistic current market unit rates for the following BOQ items.

Project location:
{region}

Project type:
{proj_type}

Quality level:
{quality}

Currency:
{currency}

Pricing year:
2025/2026

IMPORTANT:

1. Return a realistic contractor unit rate.
2. Include materials.
3. Include labor.
4. Include equipment where applicable.
5. Include normal contractor overhead.
6. Do NOT include VAT unless it is explicitly part of the unit rate.
7. Use the correct unit of measurement.
8. Consider typical Saudi/Gulf construction market conditions.
9. Do not return total cost.
10. Return UNIT PRICE only.
11. If an item is unclear, estimate the most reasonable market rate.
12. Prices must be numeric.

BOQ ITEMS:

{numbered}

Return ONLY valid JSON.

The JSON keys MUST be the local item indexes:
0, 1, 2, 3...

Example:

{{
  "0": 185.00,
  "1": 95.50,
  "2": 1250.00
}}

No markdown.
No explanation.
No comments.
Only JSON.
"""

        # ----------------------------------------------------
        # API Request
        # ----------------------------------------------------

        try:

            response = requests.post(

                url,

                json={
                    "contents": [
                        {
                            "parts": [
                                {
                                    "text": prompt
                                }
                            ]
                        }
                    ],

                    "generationConfig": {

                        "temperature": 0.1,

                        "maxOutputTokens": 2048

                    }

                },

                timeout=60

            )

            # ------------------------------------------------
            # HTTP error
            # ------------------------------------------------

            response.raise_for_status()

            data = response.json()

            # ------------------------------------------------
            # Extract Gemini text
            # ------------------------------------------------

            candidates = data.get(
                "candidates",
                []
            )

            if not candidates:
                raise Exception(
                    "Gemini لم يرجع أي نتيجة"
                )

            content = candidates[0].get(
                "content",
                {}
            )

            parts = content.get(
                "parts",
                []
            )

            if not parts:
                raise Exception(
                    "Gemini لم يرجع نص"
                )

            text = parts[0].get(
                "text",
                ""
            ).strip()

            # ------------------------------------------------
            # Clean JSON
            # ------------------------------------------------

            text = text.replace(
                "```json",
                ""
            )

            text = text.replace(
                "```",
                ""
            )

            text = text.strip()

            # ------------------------------------------------
            # Extract JSON object
            # ------------------------------------------------

            json_start = text.find("{")
            json_end = text.rfind("}")

            if (
                json_start < 0
                or json_end < 0
            ):
                raise Exception(
                    f"لم يتم العثور على JSON في رد Gemini: {text[:300]}"
                )

            json_text = text[
                json_start:
                json_end + 1
            ]

            result = json.loads(
                json_text
            )

            # ------------------------------------------------
            # Apply prices to correct batch
            # ------------------------------------------------

            for key, value in result.items():

                try:

                    local_index = int(key)

                    price = float(value)

                except Exception:

                    continue

                if (
                    local_index < 0
                    or local_index >= len(batch)
                ):
                    continue

                if price <= 0:
                    continue

                global_index = (
                    batch_start
                    + local_index
                )

                items[
                    global_index
                ]["price"] = price

                items[
                    global_index
                ]["total"] = (
                    price
                    * items[
                        global_index
                    ]["qty"]
                )

                priced += 1

        except requests.exceptions.HTTPError as ex:

            try:
                error_data = response.json()

                error_message = error_data.get(
                    "error",
                    {}
                ).get(
                    "message",
                    str(ex)
                )

            except Exception:

                error_message = str(ex)

            st.error(
                f"❌ Gemini API Error: {error_message}"
            )

        except requests.exceptions.Timeout:

            st.warning(
                "⚠ انتهت مهلة الاتصال بـ Gemini "
                "في هذه الدفعة."
            )

        except Exception as ex:

            st.warning(
                f"⚠ خطأ في دفعة "
                f"{batch_start + 1}–{batch_end}: "
                f"{ex}"
            )

    # --------------------------------------------------------
    # Finish
    # --------------------------------------------------------

    progress.progress(100)

    status.success(
        f"✅ تم تسعير {priced} بند "
        f"من أصل {total_items}"
    )

    return items


# ============================================================
# EXCEL EXPORT
# ============================================================

def to_excel(
    items,
    currency
):

    wb = openpyxl.Workbook()

    ws = wb.active

    ws.title = "BOQ Priced"

    # --------------------------------------------------------
    # Headers
    # --------------------------------------------------------

    headers = [
        "رقم البند",
        "الشيت",
        "الوصف",
        "الوحدة",
        "الكمية",
        f"سعر الوحدة ({currency})",
        f"الإجمالي ({currency})"
    ]

    bold_white = openpyxl.styles.Font(
        name="Arial",
        bold=True,
        color="FFFFFF",
        size=11
    )

    fill_header = openpyxl.styles.PatternFill(
        "solid",
        fgColor="1E3A5F"
    )

    center_alignment = openpyxl.styles.Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )

    for col_index, header in enumerate(
        headers,
        1
    ):

        cell = ws.cell(
            1,
            col_index,
            header
        )

        cell.font = bold_white

        cell.fill = fill_header

        cell.alignment = center_alignment

    # --------------------------------------------------------
    # Styles
    # --------------------------------------------------------

    fill_ai = openpyxl.styles.PatternFill(
        "solid",
        fgColor="EFF6FF"
    )

    font_normal = openpyxl.styles.Font(
        name="Arial",
        size=10
    )

    font_number = openpyxl.styles.Font(
        name="Arial",
        size=10,
        color="1E3A5F",
        bold=True
    )

    # --------------------------------------------------------
    # Rows
    # --------------------------------------------------------

    for row_index, item in enumerate(
        items,
        2
    ):

        ws.cell(
            row_index,
            1,
            item["item_no"]
        ).font = font_normal

        ws.cell(
            row_index,
            2,
            item["sheet"]
        ).font = font_normal

        ws.cell(
            row_index,
            3,
            item["description"]
        ).font = font_normal

        ws.cell(
            row_index,
            4,
            item["unit"]
        ).font = font_normal

        ws.cell(
            row_index,
            5,
            item["qty"]
        ).font = font_number

        price_cell = ws.cell(
            row_index,
            6,
            round(
                item["price"],
                2
            )
        )

        price_cell.font = font_number

        price_cell.fill = fill_ai

        price_cell.number_format = "#,##0.00"

        total_cell = ws.cell(
            row_index,
            7,
            f"=F{row_index}*E{row_index}"
        )

        total_cell.font = font_number

        total_cell.fill = fill_ai

        total_cell.number_format = "#,##0.00"

    # --------------------------------------------------------
    # Grand Total
    # --------------------------------------------------------

    last_data_row = len(items) + 1

    total_row = last_data_row + 1

    ws.cell(
        total_row,
        6,
        "الإجمالي الكلي"
    ).font = openpyxl.styles.Font(
        name="Arial",
        bold=True,
        size=11
    )

    grand_total = ws.cell(
        total_row,
        7,
        f"=SUM(G2:G{last_data_row})"
    )

    grand_total.font = openpyxl.styles.Font(
        name="Arial",
        bold=True,
        size=12,
        color="1E3A5F"
    )

    grand_total.number_format = "#,##0.00"

    grand_total.fill = openpyxl.styles.PatternFill(
        "solid",
        fgColor="DBEAFE"
    )

    # --------------------------------------------------------
    # Column Widths
    # --------------------------------------------------------

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 60
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 22
    ws.column_dimensions["G"].width = 22

    # --------------------------------------------------------
    # Freeze Header
    # --------------------------------------------------------

    ws.freeze_panes = "A2"

    # --------------------------------------------------------
    # Auto Filter
    # --------------------------------------------------------

    ws.auto_filter.ref = (
        f"A1:G{last_data_row}"
    )

    # --------------------------------------------------------
    # Save
    # --------------------------------------------------------

    buffer = io.BytesIO()

    wb.save(buffer)

    buffer.seek(0)

    return buffer


# ============================================================
# MAIN UI
# ============================================================

uploaded = st.file_uploader(
    "📂 ارفع ملف BOQ",
    type=["xlsx"],
    label_visibility="collapsed"
)


# ============================================================
# FILE PROCESSING
# ============================================================

if uploaded:

    file_bytes = uploaded.read()

    with st.spinner(
        "جاري قراءة وتحليل ملف BOQ..."
    ):

        try:

            items = extract_items(
                file_bytes
            )

        except Exception as ex:

            items = []

            st.error(
                f"❌ حدث خطأ أثناء قراءة الملف: {ex}"
            )

    # --------------------------------------------------------
    # No Items
    # --------------------------------------------------------

    if not items:

        st.error(
            "⚠ لم يتم اكتشاف أي بنود. "
            "تأكد أن الملف يحتوي على أعمدة "
            "Description / Unit / Quantity."
        )

    else:

        # ----------------------------------------------------
        # Success
        # ----------------------------------------------------

        st.success(
            f"✅ تم اكتشاف **{len(items)} بند** "
            f"من **{len(set(i['sheet'] for i in items))} شيت**"
        )

        # ----------------------------------------------------
        # Preview
        # ----------------------------------------------------

        df_preview = pd.DataFrame(
            items
        )[
            [
                "item_no",
                "sheet",
                "description",
                "unit",
                "qty"
            ]
        ].copy()

        df_preview.columns = [
            "رقم البند",
            "الشيت",
            "الوصف",
            "الوحدة",
            "الكمية"
        ]

        st.dataframe(
            df_preview,
            use_container_width=True,
            height=280
        )

        st.divider()

        # ----------------------------------------------------
        # API Key Check
        # ----------------------------------------------------

        if not api_key:

            st.warning(
                "⚠ أدخل Gemini API Key "
                "في الشريط الجانبي أولاً."
            )

        elif not api_key.startswith(
            "AIza"
        ):

            st.warning(
                "⚠ مفتاح Gemini غير صحيح. "
                "المفتاح يجب أن يبدأ بـ AIza."
            )

        else:

            # ------------------------------------------------
            # Pricing Button
            # ------------------------------------------------

            if st.button(
                "✨ ابدأ التسعير التلقائي",
                use_container_width=True
            ):

                # Clear previous result
                st.session_state.pop(
                    "priced_items",
                    None
                )

                # Price
                priced_items = price_with_gemini(

                    items,

                    api_key,

                    region,

                    currency,

                    proj_type,

                    quality

                )

                # Save
                st.session_state[
                    "priced_items"
                ] = priced_items

                st.session_state[
                    "currency"
                ] = currency

                # Rerun
                st.rerun()


# ============================================================
# PRICED RESULTS
# ============================================================

if "priced_items" in st.session_state:

    items = st.session_state[
        "priced_items"
    ]

    currency = st.session_state[
        "currency"
    ]

    st.divider()

    st.subheader(
        "📋 النتائج"
    )

    # --------------------------------------------------------
    # KPIs
    # --------------------------------------------------------

    total_items = len(items)

    total_priced = sum(
        1
        for item in items
        if item["price"] > 0
    )

    grand_total = sum(
        item["price"] * item["qty"]
        for item in items
    )

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "إجمالي البنود",
        total_items
    )

    c2.metric(
        "بنود مسعّرة",
        f"{total_priced} / {total_items}"
    )

    c3.metric(
        f"الإجمالي ({currency})",
        f"{grand_total:,.0f}"
    )

    st.divider()

    # --------------------------------------------------------
    # Editable DataFrame
    # --------------------------------------------------------

    df = pd.DataFrame(
        items
    )[
        [
            "item_no",
            "description",
            "unit",
            "qty",
            "price"
        ]
    ].copy()

    df["total"] = (
        df["qty"]
        * df["price"]
    )

    df.columns = [
        "رقم البند",
        "الوصف",
        "الوحدة",
        "الكمية",
        "سعر الوحدة",
        "الإجمالي"
    ]

    # --------------------------------------------------------
    # Data Editor
    # --------------------------------------------------------

    edited = st.data_editor(

        df,

        use_container_width=True,

        height=450,

        num_rows="fixed",

        column_config={

            "سعر الوحدة":
                st.column_config.NumberColumn(
                    format=f"%.2f {currency}",
                    min_value=0
                ),

            "الإجمالي":
                st.column_config.NumberColumn(
                    format=f"%.2f {currency}"
                )

        },

        disabled=[
            "رقم البند",
            "الوصف",
            "الوحدة",
            "الكمية",
            "الإجمالي"
        ],

        key="priced_editor"

    )

    # --------------------------------------------------------
    # Save Edited Prices
    # --------------------------------------------------------

    edited_prices = edited[
        "سعر الوحدة"
    ].tolist()

    for index, price in enumerate(
        edited_prices
    ):

        try:

            price = float(
                price or 0
            )

        except Exception:

            price = 0

        items[index][
            "price"
        ] = price

        items[index][
            "total"
        ] = (
            price
            * items[index]["qty"]
        )

    # Update session state
    st.session_state[
        "priced_items"
    ] = items

    # --------------------------------------------------------
    # Recalculate Total
    # --------------------------------------------------------

    grand_total = sum(
        item["price"]
        * item["qty"]
        for item in items
    )

    st.info(
        f"💰 **الإجمالي الحالي: "
        f"{grand_total:,.2f} {currency}**"
    )

    # --------------------------------------------------------
    # Export
    # --------------------------------------------------------

    excel_buffer = to_excel(
        items,
        currency
    )

    st.download_button(

        label="📥 تحميل Excel",

        data=excel_buffer,

        file_name=(
            f"BOQ_priced_"
            f"{datetime.now().strftime('%Y%m%d_%H%M')}"
            f".xlsx"
        ),

        mime=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),

        use_container_width=True

                  )
