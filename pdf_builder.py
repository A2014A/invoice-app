from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
import io, os, urllib.request

FONT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'static', 'fonts')
FONTS_READY = False

def ensure_fonts():
    global FONTS_READY
    if FONTS_READY:
        return True
    os.makedirs(FONT_DIR, exist_ok=True)
    reg = os.path.join(FONT_DIR, 'Alef-Regular.ttf')
    bold = os.path.join(FONT_DIR, 'Alef-Bold.ttf')
    try:
        if not os.path.exists(reg):
            urllib.request.urlretrieve(
                'https://github.com/google/fonts/raw/main/ofl/alef/Alef-Regular.ttf', reg)
        if not os.path.exists(bold):
            urllib.request.urlretrieve(
                'https://github.com/google/fonts/raw/main/ofl/alef/Alef-Bold.ttf', bold)
        pdfmetrics.registerFont(TTFont('Alef', reg))
        pdfmetrics.registerFont(TTFont('Alef-Bold', bold))
        FONTS_READY = True
        return True
    except Exception as e:
        print(f"Font error: {e}")
        return False

import arabic_reshaper
from bidi.algorithm import get_display

def h(text):
    """עברית תקינה ב-PDF"""
    try:
        reshaped = arabic_reshaper.reshape(str(text))
        return get_display(reshaped)
    except:
        return str(text)

def st(size=11, bold=False, color=colors.black, align=TA_RIGHT):
    has = ensure_fonts()
    fn = ('Alef-Bold' if bold else 'Alef') if has else ('Helvetica-Bold' if bold else 'Helvetica')
    return ParagraphStyle('x', fontName=fn, fontSize=size, textColor=color,
        alignment=align, leading=size*1.5)

def build_pdf(data):
    ensure_fonts()
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
        tc = colors.HexColor('#2563EB')
    elif doc_type == 'קבלה':
        tc = colors.HexColor('#16A34A')
    else:
        tc = colors.HexColor('#7C3AED')

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
        rightMargin=2*cm, leftMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm)
    story = []

    story.append(Paragraph(h(doc_type), st(24, True, tc, TA_CENTER)))
    story.append(Spacer(1, 6))
    story.append(Paragraph(h('יהודה קורץ'), st(18, True, colors.HexColor('#1E3A5F'), TA_CENTER)))
    story.append(Spacer(1, 4))
    story.append(Paragraph(h('עוסק פטור מס׳ 027394865'), st(12, False, colors.HexColor('#374151'), TA_CENTER)))
    story.append(Spacer(1, 16))

    t1 = Table([[
        Paragraph(h(f'תאריך: {date_str}'), st(11)),
        Paragraph(h(f'מספר מסמך: {doc_num}'), st(13, True, colors.HexColor('#2563EB')))
    ]], colWidths=[8*cm, 8*cm])
    t1.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(0,0),colors.HexColor('#F8FAFC')),
        ('BACKGROUND',(1,0),(1,0),colors.HexColor('#EFF6FF')),
        ('BOX',(0,0),(0,0),0.5,colors.HexColor('#CCCCCC')),
        ('BOX',(1,0),(1,0),1.5,colors.HexColor('#2563EB')),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('TOPPADDING',(0,0),(-1,-1),8),('BOTTOMPADDING',(0,0),(-1,-1),8),
        ('RIGHTPADDING',(0,0),(-1,-1),10),
    ]))
    story.append(t1)
    story.append(Spacer(1,12))

    t2 = Table([
        [Paragraph(h('פרטי הלקוח'), st(12, True, colors.HexColor('#16A34A')))],
        [Table([[
            Paragraph(h(f'מס׳ עוסק: {client_id}'), st(10)),
            Paragraph(h(f'שם לקוח: {client}'), st(10))],[
            Paragraph(h(f'תיאור: {desc}'), st(10)),
            Paragraph(h(f'כתובת: {address}'), st(10)),
        ]], colWidths=[8*cm,8*cm])],
    ], colWidths=[16*cm])
    t2.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1),colors.HexColor('#F0FDF4')),
        ('BOX',(0,0),(-1,-1),1.5,colors.HexColor('#16A34A')),
        ('TOPPADDING',(0,0),(-1,-1),8),('BOTTOMPADDING',(0,0),(-1,-1),8),
        ('RIGHTPADDING',(0,0),(-1,-1),10),
    ]))
    story.append(t2)
    story.append(Spacer(1,12))

    if doc_type in ('חשבונית עסקה','חשבונית עסקה + קבלה'):
        t3 = Table([
            [Paragraph(h('סכום ₪'),st(11,True,colors.white,TA_CENTER)),
             Paragraph(h('כמות'),st(11,True,colors.white,TA_CENTER)),
             Paragraph(h('תיאור השירות / הפריט'),st(11,True,colors.white,TA_CENTER))],
            [Paragraph(f'{inv_amt:,.2f}',st(11,False,colors.HexColor('#1E3A5F'),TA_CENTER)),
             Paragraph('1',st(11,False,colors.HexColor('#1E3A5F'),TA_CENTER)),
             Paragraph(h(desc or ''),st(11,False,colors.HexColor('#1E3A5F'),TA_CENTER))],
        ], colWidths=[4*cm,3*cm,9*cm])
        t3.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#2563EB')),
            ('BACKGROUND',(0,1),(-1,1),colors.HexColor('#EFF6FF')),
            ('GRID',(0,0),(-1,-1),0.5,colors.HexColor('#DBEAFE')),
            ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
            ('TOPPADDING',(0,0),(-1,-1),8),('BOTTOMPADDING',(0,0),(-1,-1),8),
        ]))
        story.append(t3)
        story.append(Spacer(1,8))

        t4 = Table([[
            Paragraph(f'₪ {inv_amt:,.2f}',st(14,True,colors.HexColor('#1E3A5F'),TA_CENTER)),
            Paragraph(h('סה"כ לתשלום:'),st(13,True,colors.white,TA_CENTER))
        ]], colWidths=[8*cm,8*cm])
        t4.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(0,0),colors.HexColor('#DBEAFE')),
            ('BACKGROUND',(1,0),(1,0),colors.HexColor('#2563EB')),
            ('BOX',(0,0),(-1,-1),1.5,colors.HexColor('#2563EB')),
            ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
            ('TOPPADDING',(0,0),(-1,-1),10),('BOTTOMPADDING',(0,0),(-1,-1),10),
        ]))
        story.append(t4)
        story.append(Spacer(1,12))

    if doc_type in ('קבלה','חשבונית עסקה + קבלה'):
        t5 = Table([
            [Paragraph(f'₪ {rec_total:,.2f}',st(12,True,colors.HexColor('#16A34A'),TA_CENTER)),
             Paragraph(h('סה"כ שולם:'),st(11,True,colors.HexColor('#16A34A')))],
            [Paragraph(f'₪ {bank_amt:,.2f}',st(11,True,colors.HexColor('#15803D'),TA_CENTER)),
             Paragraph(h('התקבל בבנק:'),st(11,True,colors.HexColor('#15803D')))],
            [Paragraph(f'₪ {nik_amt:,.2f}',st(11,True,colors.HexColor('#7C3AED'),TA_CENTER)),
             Paragraph(h(f'ניכוי במקור ({int(pct_nik*100)}%):'),st(11,True,colors.HexColor('#7C3AED')))],
            [Paragraph('____________',st(11)),
             Paragraph(h('אמצעי תשלום:'),st(11,True,colors.HexColor('#374151')))],
        ], colWidths=[8*cm,8*cm])
        t5.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#F0FDF4')),
            ('BACKGROUND',(0,2),(-1,2),colors.HexColor('#F5F3FF')),
            ('BACKGROUND',(0,3),(-1,3),colors.HexColor('#F8FAFC')),
            ('BOX',(0,0),(-1,-1),1,colors.HexColor('#16A34A')),
            ('GRID',(0,0),(-1,-1),0.5,colors.HexColor('#E5E7EB')),
            ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
            ('TOPPADDING',(0,0),(-1,-1),8),('BOTTOMPADDING',(0,0),(-1,-1),8),
            ('RIGHTPADDING',(0,0),(-1,-1),10),
        ]))
        story.append(t5)
        story.append(Spacer(1,12))

    story.append(Paragraph(h('* עוסק פטור — אינו מחייב במע"מ'),
        st(9,False,colors.HexColor('#6B7280'),TA_CENTER)))
    story.append(Spacer(1,16))

    t6 = Table([[
        Paragraph(h('חתימת הלקוח: ____________________'),st(11)),
        Paragraph(h('חתימת המוכר: ____________________'),st(11))
    ]], colWidths=[8*cm,8*cm])
    t6.setStyle(TableStyle([
        ('BOX',(0,0),(-1,-1),0.5,colors.HexColor('#E5E7EB')),
        ('TOPPADDING',(0,0),(-1,-1),10),('BOTTOMPADDING',(0,0),(-1,-1),10),
    ]))
    story.append(t6)

    doc.build(story)
    buf.seek(0)
    return buf
