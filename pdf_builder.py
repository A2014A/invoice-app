from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
import io, os
 
FONT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'static', 'fonts')
 
def setup_fonts():
    try:
        pdfmetrics.registerFont(TTFont('Alef', os.path.join(FONT_DIR, 'Alef-Regular.ttf')))
        pdfmetrics.registerFont(TTFont('Alef-Bold', os.path.join(FONT_DIR, 'Alef-Bold.ttf')))
        return True
    except:
        return False
 
def s(size=11, bold=False, color=colors.black, align=TA_RIGHT):
    has_font = setup_fonts()
    fn = 'Alef-Bold' if bold and has_font else ('Alef' if has_font else ('Helvetica-Bold' if bold else 'Helvetica'))
    return ParagraphStyle('x', fontName=fn, fontSize=size, textColor=color,
        alignment=align, leading=size*1.5, wordWrap='CJK')
 
def build_pdf(data):
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
 
    inv_amt   = sf(data.get('inv_amt'))
    bank_amt  = sf(data.get('bank_amt'))
    nik_amt   = sf(data.get('nik_amt'))
    pct_nik   = sf(data.get('pct_nik')) or 0.2
    rec_total = bank_amt + nik_amt
 
    if doc_type == 'חשבונית עסקה':
        title_color = colors.HexColor('#2563EB')
    elif doc_type == 'קבלה':
        title_color = colors.HexColor('#16A34A')
    else:
        title_color = colors.HexColor('#7C3AED')
 
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
        rightMargin=2*cm, leftMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm)
 
    story = []
 
    # כותרת
    story.append(Paragraph(doc_type, s(24, True, title_color, TA_CENTER)))
    story.append(Spacer(1, 6))
    story.append(Paragraph('יהודה קורץ', s(18, True, colors.HexColor('#1E3A5F'), TA_CENTER)))
    story.append(Spacer(1, 4))
    story.append(Paragraph('עוסק פטור מס׳ 027394865', s(12, False, colors.HexColor('#374151'), TA_CENTER)))
    story.append(Spacer(1, 16))
 
    # מספר ותאריך
    t1 = Table([
        [Paragraph(f'תאריך: {date_str}', s(11)),
         Paragraph(f'מספר מסמך: {doc_num}', s(13, True, colors.HexColor('#2563EB')))]
    ], colWidths=[8*cm, 8*cm])
    t1.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (0,0), colors.HexColor('#F8FAFC')),
        ('BACKGROUND', (1,0), (1,0), colors.HexColor('#EFF6FF')),
        ('BOX', (0,0), (0,0), 0.5, colors.HexColor('#CCCCCC')),
        ('BOX', (1,0), (1,0), 1.5, colors.HexColor('#2563EB')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
    ]))
    story.append(t1)
    story.append(Spacer(1, 12))
 
    # פרטי לקוח
    t2 = Table([
        [Paragraph('פרטי הלקוח', s(12, True, colors.HexColor('#16A34A')))],
        [Table([
            [Paragraph(f'מס׳ עוסק: {client_id}', s(10)),
             Paragraph(f'שם לקוח: {client}', s(10))],
            [Paragraph(f'תיאור: {desc}', s(10)),
             Paragraph(f'כתובת: {address}', s(10))],
        ], colWidths=[8*cm, 8*cm])],
    ], colWidths=[16*cm])
    t2.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F0FDF4')),
        ('BOX', (0,0), (-1,-1), 1.5, colors.HexColor('#16A34A')),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
    ]))
    story.append(t2)
    story.append(Spacer(1, 12))
 
    # חשבונית
    if doc_type in ('חשבונית עסקה', 'חשבונית עסקה + קבלה'):
        t3 = Table([
            [Paragraph('סכום ₪', s(11, True, colors.white, TA_CENTER)),
             Paragraph('כמות', s(11, True, colors.white, TA_CENTER)),
             Paragraph('תיאור השירות / הפריט', s(11, True, colors.white, TA_CENTER))],
            [Paragraph(f'{inv_amt:,.2f}', s(11, False, colors.HexColor('#1E3A5F'), TA_CENTER)),
             Paragraph('1', s(11, False, colors.HexColor('#1E3A5F'), TA_CENTER)),
             Paragraph(desc or '', s(11, False, colors.HexColor('#1E3A5F'), TA_CENTER))],
        ], colWidths=[4*cm, 3*cm, 9*cm])
        t3.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2563EB')),
            ('BACKGROUND', (0,1), (-1,1), colors.HexColor('#EFF6FF')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#DBEAFE')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ]))
        story.append(t3)
        story.append(Spacer(1, 8))
 
        t4 = Table([
            [Paragraph(f'₪ {inv_amt:,.2f}', s(14, True, colors.HexColor('#1E3A5F'), TA_CENTER)),
             Paragraph('סה"כ לתשלום:', s(13, True, colors.white, TA_CENTER))]
        ], colWidths=[8*cm, 8*cm])
        t4.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (0,0), colors.HexColor('#DBEAFE')),
            ('BACKGROUND', (1,0), (1,0), colors.HexColor('#2563EB')),
            ('BOX', (0,0), (-1,-1), 1.5, colors.HexColor('#2563EB')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 10),
            ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ]))
        story.append(t4)
        story.append(Spacer(1, 12))
 
    # קבלה
    if doc_type in ('קבלה', 'חשבונית עסקה + קבלה'):
        t5 = Table([
            [Paragraph(f'₪ {rec_total:,.2f}', s(12, True, colors.HexColor('#16A34A'), TA_CENTER)),
             Paragraph('סה"כ שולם:', s(11, True, colors.HexColor('#16A34A')))],
            [Paragraph(f'₪ {bank_amt:,.2f}', s(11, True, colors.HexColor('#15803D'), TA_CENTER)),
             Paragraph('התקבל בבנק:', s(11, True, colors.HexColor('#15803D')))],
            [Paragraph(f'₪ {nik_amt:,.2f}', s(11, True, colors.HexColor('#7C3AED'), TA_CENTER)),
             Paragraph(f'ניכוי במקור ({int(pct_nik*100)}%):', s(11, True, colors.HexColor('#7C3AED')))],
            [Paragraph('____________', s(11)),
             Paragraph('אמצעי תשלום:', s(11, True, colors.HexColor('#374151')))],
        ], colWidths=[8*cm, 8*cm])
        t5.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,1), colors.HexColor('#F0FDF4')),
            ('BACKGROUND', (0,2), (-1,2), colors.HexColor('#F5F3FF')),
            ('BACKGROUND', (0,3), (-1,3), colors.HexColor('#F8FAFC')),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#16A34A')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E5E7EB')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ('RIGHTPADDING', (0,0), (-1,-1), 10),
        ]))
        story.append(t5)
        story.append(Spacer(1, 12))
 
    # הערה
    story.append(Paragraph('* עוסק פטור — אינו מחייב במע"מ', s(9, False, colors.HexColor('#6B7280'), TA_CENTER)))
    story.append(Spacer(1, 16))
 
    # חתימות
    t6 = Table([
        [Paragraph('חתימת הלקוח: ____________________', s(11)),
         Paragraph('חתימת המוכר: ____________________', s(11))]
    ], colWidths=[8*cm, 8*cm])
    t6.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#E5E7EB')),
        ('TOPPADDING', (0,0), (-1,-1), 10),
        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
    ]))
    story.append(t6)
 
    doc.build(story)
    buf.seek(0)
    return buf
