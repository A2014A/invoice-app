from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_RIGHT, TA_CENTER, TA_LEFT
import io, os, arabic_reshaper
from bidi.algorithm import get_display

FONT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'static', 'fonts')

def setup_fonts():
    try:
        pdfmetrics.registerFont(TTFont('Alef', os.path.join(FONT_DIR, 'Alef-Regular.ttf')))
        pdfmetrics.registerFont(TTFont('Alef-Bold', os.path.join(FONT_DIR, 'Alef-Bold.ttf')))
        return 'Alef'
    except:
        return 'Helvetica'

def heb(text):
    """Convert Hebrew text for proper RTL display in PDF"""
    try:
        reshaped = arabic_reshaper.reshape(str(text))
        return get_display(reshaped)
    except:
        return str(text)

def build_pdf(data):
    setup_fonts()
    font = 'Alef'
    font_bold = 'Alef-Bold'

    doc_type  = data.get('doc_type', '')
    doc_num   = data.get('doc_num', '')
    date_str  = data.get('date', '')
    client    = data.get('client', '')
    client_id = data.get('client_id', '')
    address   = data.get('address', '')
    desc      = data.get('desc', '')

    def sf(v):
        try: return float(v) if v else 0.0
        except: return 0.0

    inv_amt  = sf(data.get('inv_amt'))
    bank_amt = sf(data.get('bank_amt'))
    nik_amt  = sf(data.get('nik_amt'))
    pct_nik  = sf(data.get('pct_nik')) or 0.2
    rec_total = bank_amt + nik_amt

    # צבעים לפי סוג
    if doc_type == 'חשבונית עסקה':
        title_color = colors.HexColor('#2563EB')
        header_color = colors.HexColor('#2563EB')
    elif doc_type == 'קבלה':
        title_color = colors.HexColor('#16A34A')
        header_color = colors.HexColor('#16A34A')
    else:
        title_color = colors.HexColor('#7C3AED')
        header_color = colors.HexColor('#7C3AED')

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
        rightMargin=2*cm, leftMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm)

    story = []

    # סגנונות
    def style(size=11, bold=False, color=colors.black, align=TA_RIGHT):
        return ParagraphStyle('s',
            fontName=font_bold if bold else font,
            fontSize=size, textColor=color,
            alignment=align, leading=size*1.4,
            wordWrap='RTL')

    # כותרת
    story.append(Paragraph(heb(doc_type), style(24, True, title_color, TA_CENTER)))
    story.append(Spacer(1, 6))
    story.append(Paragraph(heb('יהודה קורץ'), style(18, True, colors.HexColor('#1E3A5F'), TA_CENTER)))
    story.append(Spacer(1, 4))
    story.append(Paragraph(heb('עוסק פטור מס׳ 027394865'), style(12, False, colors.HexColor('#374151'), TA_CENTER)))
    story.append(Spacer(1, 16))

    # מספר ותאריך
    num_date = Table([
        [Paragraph(heb(f'תאריך: {date_str}'), style(11, False, colors.HexColor('#374151'))),
         Paragraph(heb(f'מספר מסמך:  {doc_num}'), style(13, True, colors.HexColor('#2563EB')))]
    ], colWidths=[8*cm, 8*cm])
    num_date.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (0,0), colors.HexColor('#F8FAFC')),
        ('BACKGROUND', (1,0), (1,0), colors.HexColor('#EFF6FF')),
        ('BOX', (0,0), (0,0), 1, colors.HexColor('#CCCCCC')),
        ('BOX', (1,0), (1,0), 1, colors.HexColor('#2563EB')),
        ('ALIGN', (0,0), (-1,-1), 'RIGHT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
    ]))
    story.append(num_date)
    story.append(Spacer(1, 12))

    # פרטי לקוח
    client_data = [
        [Paragraph(heb('פרטי הלקוח'), style(12, True, colors.HexColor('#16A34A')))],
        [Table([
            [Paragraph(heb(f'מס׳ עוסק: {client_id}'), style(10)),
             Paragraph(heb(f'שם לקוח: {client}'), style(10))],
            [Paragraph(heb(f'תיאור: {desc}'), style(10)),
             Paragraph(heb(f'כתובת: {address}'), style(10))],
        ], colWidths=[8*cm, 8*cm])],
    ]
    client_table = Table(client_data, colWidths=[16*cm])
    client_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F0FDF4')),
        ('BOX', (0,0), (-1,-1), 1.5, colors.HexColor('#16A34A')),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
    ]))
    story.append(client_table)
    story.append(Spacer(1, 12))

    # חשבונית
    if doc_type in ('חשבונית עסקה', 'חשבונית עסקה + קבלה'):
        inv_table = Table([
            [Paragraph(heb('סכום ₪'), style(11, True, colors.white, TA_CENTER)),
             Paragraph(heb('כמות'), style(11, True, colors.white, TA_CENTER)),
             Paragraph(heb('תיאור השירות / הפריט'), style(11, True, colors.white, TA_CENTER))],
            [Paragraph(f'{inv_amt:,.2f}', style(11, False, colors.HexColor('#1E3A5F'), TA_CENTER)),
             Paragraph('1', style(11, False, colors.HexColor('#1E3A5F'), TA_CENTER)),
             Paragraph(heb(desc or ''), style(11, False, colors.HexColor('#1E3A5F'), TA_CENTER))],
        ], colWidths=[4*cm, 3*cm, 9*cm])
        inv_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), header_color),
            ('BACKGROUND', (0,1), (-1,1), colors.HexColor('#EFF6FF')),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#2563EB')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#DBEAFE')),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ]))
        story.append(inv_table)
        story.append(Spacer(1, 8))

        total_table = Table([
            [Paragraph(f'₪ {inv_amt:,.2f}', style(14, True, colors.HexColor('#1E3A5F'), TA_CENTER)),
             Paragraph(heb('סה"כ לתשלום:'), style(13, True, colors.white, TA_CENTER))]
        ], colWidths=[8*cm, 8*cm])
        total_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (0,0), colors.HexColor('#DBEAFE')),
            ('BACKGROUND', (1,0), (1,0), colors.HexColor('#2563EB')),
            ('BOX', (0,0), (-1,-1), 1.5, colors.HexColor('#2563EB')),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 10),
            ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ]))
        story.append(total_table)
        story.append(Spacer(1, 12))

    # קבלה
    if doc_type in ('קבלה', 'חשבונית עסקה + קבלה'):
        rec_rows = [
            [Paragraph(f'₪ {rec_total:,.2f}', style(12, True, colors.HexColor('#16A34A'), TA_CENTER)),
             Paragraph(heb('סה"כ שולם:'), style(11, True, colors.HexColor('#16A34A')))],
            [Paragraph(f'₪ {bank_amt:,.2f}', style(11, True, colors.HexColor('#15803D'), TA_CENTER)),
             Paragraph(heb('התקבל בבנק:'), style(11, True, colors.HexColor('#15803D')))],
            [Paragraph(f'₪ {nik_amt:,.2f}', style(11, True, colors.HexColor('#7C3AED'), TA_CENTER)),
             Paragraph(heb(f'ניכוי במקור ({int(pct_nik*100)}%):'), style(11, True, colors.HexColor('#7C3AED')))],
            [Paragraph('____________', style(11)),
             Paragraph(heb('אמצעי תשלום:'), style(11, True, colors.HexColor('#374151')))],
        ]
        rec_table = Table(rec_rows, colWidths=[8*cm, 8*cm])
        rec_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F0FDF4')),
            ('BACKGROUND', (0,1), (-1,1), colors.HexColor('#F0FDF4')),
            ('BACKGROUND', (0,2), (-1,2), colors.HexColor('#F5F3FF')),
            ('BACKGROUND', (0,3), (-1,3), colors.HexColor('#F8FAFC')),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#16A34A')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E5E7EB')),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ('RIGHTPADDING', (0,0), (-1,-1), 10),
        ]))
        story.append(rec_table)
        story.append(Spacer(1, 12))

    # הערה + חתימות
    story.append(Paragraph(heb('* עוסק פטור — אינו מחייב במע"מ'), style(9, False, colors.HexColor('#6B7280'), TA_CENTER)))
    story.append(Spacer(1, 16))

    sig_table = Table([
        [Paragraph(heb('חתימת הלקוח: ____________________'), style(11)),
         Paragraph(heb('חתימת המוכר: ____________________'), style(11))]
    ], colWidths=[8*cm, 8*cm])
    sig_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#E5E7EB')),
        ('TOPPADDING', (0,0), (-1,-1), 10),
        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
    ]))
    story.append(sig_table)

    doc.build(story)
    buf.seek(0)
    return buf
