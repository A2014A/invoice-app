from flask import Flask, render_template, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io, os, datetime, json

app = Flask(__name__)

OWNER_NAME = "יהודה קורץ"
OWNER_TAX  = "027394865"

# סדרות מספור — נשמרות בזיכרון (אפשר להחליף ב-DB בעתיד)
counters = {
    "חשבונית עסקה": 2600166,
    "קבלה": 2601166,
    "חשבונית עסקה + קבלה": 26002166,
}

# ===== פונקציות בניית Word =====

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_rtl(para):
    pPr = para._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    pPr.append(bidi)

def make_border(cell, color="CCCCCC"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ['top','left','bottom','right']:
        border = OxmlElement(f'w:{side}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:color'), color)
        tcBorders.append(border)
    tcPr.append(tcBorders)

def add_para(cell, text, bold=False, size=11, color=None, align=WD_ALIGN_PARAGRAPH.RIGHT):
    para = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    para.alignment = align
    set_rtl(para)
    run = para.add_run(text)
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    return para

def build_doc(data):
    doc_type   = data.get('doc_type', '')
    doc_num    = data.get('doc_num', '')
    date_str   = data.get('date', datetime.date.today().strftime('%d/%m/%Y'))
    client     = data.get('client', '')
    client_id  = data.get('client_id', '')
    address    = data.get('address', '')
    desc       = data.get('desc', '')
    inv_amt    = data.get('inv_amt')
    rec_total  = data.get('rec_total')
    bank_amt   = data.get('bank_amt')
    pct_nik    = data.get('pct_nik', 0.2)
    nik_amt    = data.get('nik_amt')

    def sf(v):
        try: return float(v) if v else None
        except: return None

    inv_amt   = sf(inv_amt)
    rec_total = sf(rec_total)
    bank_amt  = sf(bank_amt)
    pct_nik   = sf(pct_nik) or 0.2
    nik_amt   = sf(nik_amt)

    doc = Document()
    section = doc.sections[0]
    section.page_width  = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = section.right_margin = Cm(2)
    section.top_margin  = section.bottom_margin = Cm(2)

    # RTL
    doc_settings = doc.settings.element
    doc_settings.append(OxmlElement('w:bidi'))

    # צבע כותרת לפי סוג
    if doc_type == 'חשבונית עסקה':
        title_color = '2563EB'
    elif doc_type == 'קבלה':
        title_color = '16A34A'
    else:
        title_color = '7C3AED'

    # כותרת
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_rtl(p)
    r = p.add_run(doc_type)
    r.font.name = 'Arial'; r.font.size = Pt(26); r.font.bold = True
    r.font.color.rgb = RGBColor.from_string(title_color)

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_rtl(p2)
    r2 = p2.add_run(OWNER_NAME)
    r2.font.name = 'Arial'; r2.font.size = Pt(18); r2.font.bold = True
    r2.font.color.rgb = RGBColor.from_string('1E3A5F')

    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_rtl(p3)
    r3 = p3.add_run(f"עוסק פטור מס׳ {OWNER_TAX}")
    r3.font.name = 'Arial'; r3.font.size = Pt(12)
    r3.font.color.rgb = RGBColor.from_string('374151')

    doc.add_paragraph()

    # מספר + תאריך
    t1 = doc.add_table(rows=1, cols=2)
    t1.style = 'Table Grid'
    c_num = t1.rows[0].cells[0]
    set_cell_bg(c_num, "EFF6FF")
    make_border(c_num, "2563EB")
    p_num = c_num.paragraphs[0]
    p_num.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_rtl(p_num)
    r_l = p_num.add_run("מספר מסמך:  ")
    r_l.font.name = 'Arial'; r_l.font.size = Pt(12); r_l.font.bold = True
    r_l.font.color.rgb = RGBColor.from_string('2563EB')
    r_n = p_num.add_run(str(doc_num))
    r_n.font.name = 'Arial'; r_n.font.size = Pt(16); r_n.font.bold = True
    r_n.font.color.rgb = RGBColor.from_string('1E3A5F')

    c_date = t1.rows[0].cells[1]
    set_cell_bg(c_date, "F8FAFC")
    make_border(c_date, "CCCCCC")
    add_para(c_date, f"תאריך: {date_str}", size=12)

    doc.add_paragraph()

    # פרטי לקוח
    t2 = doc.add_table(rows=1, cols=1)
    t2.style = 'Table Grid'
    c_cl = t2.rows[0].cells[0]
    set_cell_bg(c_cl, "F0FDF4")
    make_border(c_cl, "16A34A")
    add_para(c_cl, "פרטי הלקוח", bold=True, size=12, color='16A34A')
    c_cl.add_paragraph()
    t_in = c_cl.add_table(rows=2, cols=2)
    cells = [(0,0,f"שם לקוח: {client}"), (0,1,f"מס׳ עוסק: {client_id}"),
             (1,0,f"כתובת: {address}"), (1,1,f"תיאור: {desc}")]
    for ri, ci, txt in cells:
        add_para(t_in.rows[ri].cells[ci], txt, size=11)

    doc.add_paragraph()

    # חשבונית
    if doc_type in ('חשבונית עסקה', 'חשבונית עסקה + קבלה'):
        t3 = doc.add_table(rows=2, cols=3)
        t3.style = 'Table Grid'
        for i, (h, bg) in enumerate(zip(
            ["תיאור השירות / הפריט", "כמות", "סכום ₪"],
            ["2563EB", "2563EB", "2563EB"])):
            c = t3.rows[0].cells[i]
            set_cell_bg(c, bg)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_rtl(p)
            r = p.add_run(h)
            r.font.name = 'Arial'; r.font.size = Pt(11); r.font.bold = True
            r.font.color.rgb = RGBColor.from_string('FFFFFF')
        for i, val in enumerate([desc or "", "1", f"{inv_amt:,.2f}" if inv_amt else ""]):
            c = t3.rows[1].cells[i]
            set_cell_bg(c, "EFF6FF")
            add_para(c, val, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)

        doc.add_paragraph()

        t4 = doc.add_table(rows=1, cols=2)
        t4.style = 'Table Grid'
        cl = t4.rows[0].cells[0]
        cr = t4.rows[0].cells[1]
        set_cell_bg(cl, "2563EB"); make_border(cl, "2563EB")
        p_l = cl.paragraphs[0]; p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER; set_rtl(p_l)
        r_l = p_l.add_run('סה"כ לתשלום:')
        r_l.font.name = 'Arial'; r_l.font.size = Pt(13); r_l.font.bold = True
        r_l.font.color.rgb = RGBColor.from_string('FFFFFF')
        set_cell_bg(cr, "DBEAFE"); make_border(cr, "2563EB")
        p_r = cr.paragraphs[0]; p_r.alignment = WD_ALIGN_PARAGRAPH.CENTER; set_rtl(p_r)
        r_r = p_r.add_run(f"₪ {inv_amt:,.2f}" if inv_amt else "₪ 0.00")
        r_r.font.name = 'Arial'; r_r.font.size = Pt(14); r_r.font.bold = True
        r_r.font.color.rgb = RGBColor.from_string('1E3A5F')

        doc.add_paragraph()

    # קבלה
    if doc_type in ('קבלה', 'חשבונית עסקה + קבלה'):
        t5 = doc.add_table(rows=4, cols=2)
        t5.style = 'Table Grid'
        rows_data = [
            ("סה\"כ שולם:", f"₪ {rec_total:,.2f}" if rec_total else "₪ 0.00", "F0FDF4", "16A34A"),
            ("התקבל בבנק:", f"₪ {bank_amt:,.2f}" if bank_amt else "₪ 0.00", "F0FDF4", "15803D"),
            (f"ניכוי במקור ({int(pct_nik*100)}%):", f"₪ {nik_amt:,.2f}" if nik_amt else "₪ 0.00", "F5F3FF", "7C3AED"),
            ("אמצעי תשלום:", "____________", "F8FAFC", "374151"),
        ]
        for i, (label, val, bg, col) in enumerate(rows_data):
            cl = t5.rows[i].cells[0]; cr = t5.rows[i].cells[1]
            set_cell_bg(cl, bg); set_cell_bg(cr, bg)
            add_para(cl, label, bold=True, size=11, color=col)
            add_para(cr, val, bold=True, size=12, color=col, align=WD_ALIGN_PARAGRAPH.CENTER)

        doc.add_paragraph()

    # הערת עוסק פטור
    p_note = doc.add_paragraph()
    p_note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_rtl(p_note)
    r_note = p_note.add_run("* עוסק פטור — אינו מחייב במע\"מ")
    r_note.font.name = 'Arial'; r_note.font.size = Pt(9)
    r_note.font.color.rgb = RGBColor.from_string('6B7280')
    r_note.font.italic = True

    doc.add_paragraph()

    # חתימות
    t6 = doc.add_table(rows=1, cols=2)
    t6.style = 'Table Grid'
    add_para(t6.rows[0].cells[0], "חתימת המוכר: ____________________", size=11)
    add_para(t6.rows[0].cells[1], "חתימת הלקוח: ____________________", size=11)

    return doc


# ===== נתיבי Flask =====

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.json

        doc_type = data.get('doc_type', '')
        if doc_type not in counters:
            return jsonify({'error': 'סוג מסמך לא חוקי'}), 400

        # מספור אוטומטי
        doc_num = counters[doc_type]
        counters[doc_type] += 1
        data['doc_num'] = doc_num

        # בנה מסמך
        doc = build_doc(data)

        # שמור לזיכרון
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)

        safe_client = (data.get('client','') or '').replace(' ','_')
        safe_type   = doc_type.replace(' ','_').replace('+','-')
        filename    = f"{safe_type}_{doc_num}_{safe_client}.docx"

        return send_file(
            buf,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/clients', methods=['GET', 'POST'])
def clients():
    clients_file = 'clients.json'
    if request.method == 'GET':
        if os.path.exists(clients_file):
            with open(clients_file, 'r', encoding='utf-8') as f:
                return jsonify(json.load(f))
        return jsonify([])
    else:
        data = request.json
        existing = []
        if os.path.exists(clients_file):
            with open(clients_file, 'r', encoding='utf-8') as f:
                existing = json.load(f)
        existing.append(data)
        with open(clients_file, 'w', encoding='utf-8') as f:
            json.dump(existing, f, ensure_ascii=False, indent=2)
        return jsonify({'ok': True})

if __name__ == '__main__':
    app.run(debug=True)
