import streamlit as st
import pandas as pd
import io
import os
import zipfile
from datetime import date
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Image as RLImage
from reportlab.lib.utils import ImageReader

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Accurex Distributor Report Generator",
    page_icon="📋",
    layout="wide"
)

DEFAULT_LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "accurex_logo.png")


# ── Font registration ─────────────────────────────────────────────────────────
@st.cache_resource
def register_fonts():
    import os
    font_paths = [
        '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
        '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',
    ]
    # Try system fonts first, then bundled
    for path in font_paths:
        if not os.path.exists(path):
            st.error(f"Font not found: {path}. Please ensure DejaVu fonts are installed.")
            return False
    pdfmetrics.registerFont(TTFont('DVSans', font_paths[0]))
    pdfmetrics.registerFont(TTFont('DVSans-Bold', font_paths[1]))
    pdfmetrics.registerFontFamily('DVSans', normal='DVSans', bold='DVSans-Bold')
    return True

register_fonts()

# ── Colours ───────────────────────────────────────────────────────────────────
BLUE   = colors.HexColor('#1A56A0')
RED    = colors.HexColor('#A32D2D')
GREEN  = colors.HexColor('#3B6D11')
AMBER  = colors.HexColor('#854F0B')
DGRAY  = colors.HexColor('#5F5E5A')
LGRAY  = colors.HexColor('#F5F5F3')
MGRAY  = colors.HexColor('#E0DED8')
WHITE  = colors.white
RED_BG = colors.HexColor('#FCEBEB')
GRN_BG = colors.HexColor('#EAF3DE')
AMB_BG = colors.HexColor('#FAEEDA')
BLU_BG = colors.HexColor('#E6F1FB')

def PS(name, **kw):
    base = dict(fontName='DVSans', fontSize=8.5, leading=12, textColor=colors.HexColor('#2C2C2A'))
    base.update(kw)
    return ParagraphStyle(name, **base)

def inr(v):
    return f'\u20b9{int(round(v)):,}'

def ph(txt, align=TA_LEFT):
    return Paragraph(txt, PS(f'th_{txt[:6]}', fontName='DVSans-Bold', fontSize=7.5,
                              textColor=DGRAY, alignment=align, leading=10))

def pc(txt, color=None, align=TA_LEFT):
    c = color or '#2C2C2A'
    return Paragraph(f'<font color="{c}"><font name="DVSans">{txt}</font></font>',
                     PS(f'tc_{txt[:6]}', fontSize=7.5, leading=10, alignment=align))

# ── Parse Excel sheet ─────────────────────────────────────────────────────────
def parse_sheet(df):
    fy26, fy27 = {}, {}
    for i in range(2, len(df)):
        row = df.iloc[i]
        try:
            ly_p = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
            ly_q = float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0
            ly_a = float(str(row.iloc[2]).replace(',','').replace('₹','').replace(' ','')) if pd.notna(row.iloc[2]) and str(row.iloc[2]).strip() not in ('','nan') else 0
            cy_p = str(row.iloc[4]).strip() if len(row) > 4 and pd.notna(row.iloc[4]) else ''
            cy_q = float(row.iloc[5]) if len(row) > 5 and pd.notna(row.iloc[5]) else 0
            cy_a = float(str(row.iloc[6]).replace(',','').replace('₹','').replace(' ','')) if len(row) > 6 and pd.notna(row.iloc[6]) and str(row.iloc[6]).strip() not in ('','nan') else 0
            if ly_p and ly_a > 0 and 'total' not in ly_p.lower():
                fy26[ly_p] = (int(ly_q), ly_a)
            if cy_p and cy_a > 0 and 'total' not in cy_p.lower():
                fy27[cy_p] = (int(cy_q), cy_a)
        except:
            continue
    return fy26, fy27

# ── Generate PDF (returns bytes) ──────────────────────────────────────────────
def generate_pdf(dist_name, fy26, fy27, today_str, month_str, logo_bytes=None):
    buf = io.BytesIO()

    total_ly   = sum(a for _, a in fy26.values())
    total_cy   = sum(a for _, a in fy27.values())
    rev_drop   = round((total_cy - total_ly) / total_ly * 100, 1) if total_ly else 0
    order_kpi  = (total_ly - total_cy) * 1.10
    miss_count = sum(1 for p in fy26 if p not in fy27)
    new_count  = sum(1 for p in fy27 if p not in fy26)

    # Top 5 value degrowth
    degrowth_prods = []
    for p, (q6, a6) in fy26.items():
        if p in fy27:
            q7, a7 = fy27[p]
            if a7 < a6:
                degrowth_prods.append((p, q6, a6, q7, a7, a7 - a6, (a7 - a6) / a6 * 100))
    degrowth_prods.sort(key=lambda x: x[5])
    top5 = degrowth_prods[:5]

    # Orders table — simple YTD comparison (no monthly averaging)
    orders_table = []
    for p, (q6, a6) in fy26.items():
        cy_q, cy_a = fy27[p] if p in fy27 else (0, 0)
        order_q = (q6 - cy_q) * 1.10
        order_a = (a6 - cy_a) * 1.10
        orders_table.append((p, q6, a6, cy_q, cy_a, order_q, order_a))
    # Value-wise: highest order-to-be-placed first
    orders_table.sort(key=lambda x: -x[6])

    # Management summary top degrowth string
    if top5:
        top_str = ', '.join([f'{p} ({inr(vd)} / {pct:.1f}%)' for p, _, _, _, _, vd, pct in top5[:3]])
    else:
        top_str = 'No de-growth products found'

    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=14*mm, rightMargin=14*mm,
                            topMargin=13*mm, bottomMargin=13*mm)
    story = []

    # ── Header ──
    header_content = [
        Paragraph(f'<b>{dist_name}</b>',
                  PS('hn', fontName='DVSans-Bold', fontSize=17, textColor=BLUE, leading=21)),
    ]
    if logo_bytes:
        logo_buf = io.BytesIO(logo_bytes)
        max_w, max_h = 38*mm, 14.5*mm
        img_reader = ImageReader(logo_buf)
        img_w, img_h = img_reader.getSize()
        scale = min(max_w / img_w, max_h / img_h)
        logo = RLImage(logo_buf, width=img_w*scale, height=img_h*scale)
        hdr = Table([[header_content[0], logo]], colWidths=[142*mm, 38*mm])
    else:
        hdr = Table([[header_content[0], Paragraph('Accurex Biomedical',
                      PS('hb', fontName='DVSans-Bold', fontSize=9, textColor=DGRAY, alignment=TA_RIGHT))]],
                    colWidths=[142*mm, 38*mm])

    hdr.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), BLU_BG),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (1,0), (1,0), 'RIGHT'),
        ('TOPPADDING', (0,0), (-1,-1), 8), ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (0,0), 10), ('RIGHTPADDING', (1,0), (1,0), 8),
    ]))
    story.append(hdr)

    sub = Table([[Paragraph(f'Distributor Order Report Card for {today_str}',
                  PS('sub1', fontSize=9, textColor=DGRAY))]], colWidths=[180*mm])
    sub.setStyle(TableStyle([
        ('TOPPADDING', (0,0), (-1,-1), 5), ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 2),
    ]))
    story.append(sub)
    story.append(HRFlowable(width='100%', thickness=0.5, color=MGRAY, spaceAfter=8))

    # ── KPI cards ──
    CW = 36*mm
    kpi_bgs   = [RED_BG, AMB_BG, GRN_BG, RED_BG, GRN_BG]
    kpi_clrs  = [RED, AMBER, GREEN, RED, GREEN]
    kpi_bords = [colors.HexColor(c) for c in ['#F7C1C1','#FAC775','#C0DD97','#F7C1C1','#C0DD97']]
    kpi_labels = [f'Sales Revenue YTD till {month_str}', f'Last Year Revenue (FY 2025\u201326)',
                  'Orders to be Placed (+10% Growth)', 'Missing Products', 'New Products Added']
    kpi_vals  = [inr(total_cy), inr(total_ly), inr(order_kpi),
                 f'{miss_count} products', f'{new_count} products']
    kpi_subs  = [f'As of {today_str}', f'\u2193 {abs(rev_drop)}% de-growth vs LY',
                 '(LY \u2212 CY) \u00d7 1.10', 'Zero orders in FY 2026\u201327', 'Freshly onboarded in FY27']

    row0 = [Paragraph(kpi_labels[i], PS(f'kl{i}', fontName='DVSans-Bold', fontSize=7, textColor=DGRAY, leading=9)) for i in range(5)]
    row1 = [Paragraph(f'<b><font name="DVSans">{kpi_vals[i]}</font></b>',
                      PS(f'kv{i}', fontName='DVSans-Bold', fontSize=13, textColor=kpi_clrs[i], leading=18)) for i in range(5)]
    row2 = [Paragraph(kpi_subs[i], PS(f'ks{i}', fontSize=6.5, textColor=DGRAY, leading=9)) for i in range(5)]

    kpi_row = Table([row0, row1, row2], colWidths=[CW]*5)
    kpi_row.setStyle(TableStyle([
        *[('BACKGROUND', (i,0), (i,2), kpi_bgs[i]) for i in range(5)],
        *[('BOX', (i,0), (i,2), 0.5, kpi_bords[i]) for i in range(5)],
        ('TOPPADDING', (0,0), (-1,0), 7), ('BOTTOMPADDING', (0,0), (-1,0), 3),
        ('TOPPADDING', (0,1), (-1,1), 2), ('BOTTOMPADDING', (0,1), (-1,1), 2),
        ('TOPPADDING', (0,2), (-1,2), 2), ('BOTTOMPADDING', (0,2), (-1,2), 7),
        ('LEFTPADDING', (0,0), (-1,-1), 7), ('RIGHTPADDING', (0,0), (-1,-1), 7),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LINEAFTER', (0,0), (3,2), 0.5, WHITE),
    ]))
    story.append(kpi_row)
    story.append(Spacer(1, 10))

    # ── Table: Order to be Placed ──
    story.append(Paragraph(
        'Order to be Placed',
        PS('sh2', fontName='DVSans-Bold', fontSize=8.5, textColor=DGRAY, spaceAfter=5)))

    t2_hdr = [
        ph('Product'),
        ph('Last Year\nQty', TA_RIGHT), ph('Last Year\nAmount', TA_RIGHT),
        ph('This Year\nQty', TA_RIGHT), ph('This Year\nAmount', TA_RIGHT),
        ph('Qty to be\nOrdered', TA_RIGHT),
        ph('Amount to be\nOrdered (+10%)', TA_RIGHT),
        ph('Action', TA_CENTER),
    ]
    t2_data = [t2_hdr]
    for p, q6, a6, cy_q, cy_a, order_q, order_a in orders_table:
        no_action    = order_a <= 0
        order_color  = '#3B6D11' if no_action else '#A32D2D'
        action       = 'No Action' if no_action else 'Place Order'
        t2_data.append([
            Paragraph(p, PS(f'p2_{p[:6]}', fontSize=7, leading=9)),
            pc(str(q6), align=TA_RIGHT),
            pc(inr(a6), align=TA_RIGHT),
            pc(str(cy_q) if cy_q else '\u2014', align=TA_RIGHT),
            pc(inr(cy_a) if cy_a else '\u2014', align=TA_RIGHT),
            pc(str(int(round(order_q))), order_color, TA_RIGHT),
            pc(inr(order_a), order_color, TA_RIGHT),
            pc(action, order_color, TA_CENTER),
        ])

    t2 = Table(t2_data,
               colWidths=[50*mm, 14*mm, 20*mm, 14*mm, 20*mm, 18*mm, 24*mm, 20*mm],
               repeatRows=1)
    t2.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), LGRAY), ('LINEBELOW', (0,0), (-1,0), 0.5, MGRAY),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [WHITE, LGRAY]),
        ('LINEBELOW', (0,1), (-1,-2), 0.3, MGRAY), ('BOX', (0,0), (-1,-1), 0.5, MGRAY),
        ('TOPPADDING', (0,0), (-1,-1), 3), ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 4), ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(t2)
    story.append(Spacer(1, 10))

    # ── Management summary ──
    story.append(Paragraph('Management Summary',
                            PS('msh', fontName='DVSans-Bold', fontSize=8.5, textColor=DGRAY, spaceAfter=5)))
    points = [
        ('#A32D2D', f'Sales revenue YTD till {month_str}: {inr(total_cy)} against last year {inr(total_ly)} \u2014 de-growth of {inr(total_cy - total_ly)} ({rev_drop}%).'),
        ('#A32D2D', f'Top value de-growth: {top_str}.'),
        ('#854F0B', f'{miss_count} products with last year movement have zero orders in FY 2026\u201327 \u2014 urgent collection required.'),
        ('#3B6D11', f'{new_count} new products onboarded in FY 2026\u201327 \u2014 potential for basket expansion.'),
        ('#A32D2D', 'Immediate field visit recommended to recover pending orders and arrest the value decline.'),
    ]
    for dot, text in points:
        row = Table([[
            Table([[Paragraph('\u25cf', PS('dot', fontName='DVSans-Bold', fontSize=10,
                    textColor=colors.HexColor(dot), leading=13))]],
                  colWidths=[5*mm]),
            Paragraph(text, PS('mt', fontSize=8.5, leading=13)),
        ]], colWidths=[7*mm, 169*mm])
        row.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('TOPPADDING', (0,0), (-1,-1), 2), ('BOTTOMPADDING', (0,0), (-1,-1), 2),
            ('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ]))
        story.append(row)

    story.append(Spacer(1, 8))
    story.append(HRFlowable(width='100%', thickness=0.5, color=MGRAY))
    story.append(Paragraph('Accurex Biomedical  |  Confidential  |  For internal use only',
                            PS('footer', fontSize=7.5, textColor=DGRAY, alignment=TA_CENTER, spaceBefore=5)))

    doc.build(story)
    return buf.getvalue()


# ── Streamlit UI ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
.block-container{padding-top:2rem}
.stButton>button{background:#1A56A0;color:white;border:none;border-radius:8px;padding:0.5rem 1.5rem;font-weight:500}
.stButton>button:hover{background:#154a8e}
.stDownloadButton>button{background:#1A56A0;color:white;border:none;border-radius:8px}
</style>
""", unsafe_allow_html=True)

col1, col2 = st.columns([3,1])
with col1:
    st.title("Distributor Order Report Card Generator")
    st.caption("Accurex Biomedical — generates PDFs identical to the finalized report format")
with col2:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/8/8e/Accurex_logo.png/200px-Accurex_logo.png",
             width=120, use_container_width=False)

st.divider()

# Sidebar settings
with st.sidebar:
    st.header("Settings")
    today = date.today()
    today_str = today.strftime("%d %B %Y")
    month_str = today.strftime("%B %Y")
    st.info(f"Report date: **{today_str}**")
    st.divider()
    st.caption("**Column structure expected:**\n- Col A: LY Product\n- Col B: LY Qty\n- Col C: LY Amount\n- Col E: CY Product\n- Col F: CY Qty\n- Col G: CY Amount\n- Row 1: Headers\n- Row 2: Totals (skipped)\n- Row 3+: Products")
    logo_file = st.file_uploader("Override logo for this run (optional)", type=['jpg','jpeg','png'],
                                  help="Leave empty to use the default Accurex logo bundled with the app")
    if logo_file:
        logo_bytes = logo_file.read()
    elif os.path.exists(DEFAULT_LOGO_PATH):
        with open(DEFAULT_LOGO_PATH, 'rb') as f:
            logo_bytes = f.read()
    else:
        logo_bytes = None
        st.warning("No default logo found in repo (accurex_logo.png) — reports will show text fallback until you add one or upload here.")

# Main upload
uploaded = st.file_uploader("Upload your sales Excel file", type=['xlsx','xls'],
                              help="Each sheet should be one distributor")

if uploaded:
    xl = pd.ExcelFile(uploaded)
    sheets_data = {}
    with st.spinner("Parsing sheets..."):
        for sheet in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=sheet, header=None)
            fy26, fy27 = parse_sheet(df)
            if fy26 or fy27:
                sheets_data[sheet] = (fy26, fy27)

    st.success(f"Found **{len(sheets_data)}** distributor sheets")

    # Distributor selection
    st.subheader("Select distributors")
    cols = st.columns(4)
    selected = {}
    for i, (name, (fy26, fy27)) in enumerate(sheets_data.items()):
        with cols[i % 4]:
            checked = st.checkbox(f"**{name}**\n{len(fy26)} LY · {len(fy27)} CY", value=True, key=name)
            selected[name] = checked

    sel_names = [n for n, v in selected.items() if v]
    st.caption(f"{len(sel_names)} of {len(sheets_data)} distributors selected")

    st.divider()

    if st.button(f"Generate {len(sel_names)} PDF report(s)", disabled=len(sel_names)==0):
        generated = {}
        progress = st.progress(0)
        status = st.empty()

        for i, name in enumerate(sel_names):
            status.text(f"Generating: {name}...")
            fy26, fy27 = sheets_data[name]
            try:
                pdf_bytes = generate_pdf(name, fy26, fy27, today_str, month_str, logo_bytes)
                generated[name] = pdf_bytes
            except Exception as e:
                st.error(f"Error generating {name}: {e}")
            progress.progress((i+1)/len(sel_names))

        status.text("All done!")
        st.success(f"Generated {len(generated)} reports successfully!")

        # Individual downloads
        st.subheader("Download reports")
        dl_cols = st.columns(3)
        for i, (name, pdf_bytes) in enumerate(generated.items()):
            with dl_cols[i % 3]:
                fname = f"{name.replace(' ','_')}_Order_Report_{month_str.replace(' ','_')}.pdf"
                st.download_button(
                    label=f"⬇ {name}",
                    data=pdf_bytes,
                    file_name=fname,
                    mime='application/pdf',
                    key=f"dl_{name}"
                )

        # Download all as ZIP
        if len(generated) > 1:
            st.divider()
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                for name, pdf_bytes in generated.items():
                    fname = f"{name.replace(' ','_')}_Order_Report_{month_str.replace(' ','_')}.pdf"
                    zf.writestr(fname, pdf_bytes)
            st.download_button(
                label=f"⬇ Download all {len(generated)} PDFs as ZIP",
                data=zip_buf.getvalue(),
                file_name=f"Accurex_Order_Reports_{month_str.replace(' ','_')}.zip",
                mime='application/zip'
            )
