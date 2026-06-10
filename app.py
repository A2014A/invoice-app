from flask import Flask, render_template, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io, os, datetime, sqlite3, smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from pdf_builder import build_pdf
 
app = Flask(__name__)
 
OWNER_NAME = "יהודה קורץ"
OWNER_TAX  = "027394865"
DB_PATH    = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.db")
 
# ===== מסד נתונים =====
 
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn
 
def init_db():
    conn = get_db()
    conn.execute('''CREATE TABLE IF NOT EXISTS clients (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        tax_id TEXT,
        address TEXT,
        phone TEXT
    )''')
    conn.execute('''CREATE TABLE IF NOT EXISTS documents (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        doc_type TEXT NOT NULL,
        doc_num INTEGER NOT NULL,
        date TEXT NOT NULL,
        client_name TEXT NOT NULL,
        client_id TEXT,
        address TEXT,
        description TEXT,
        inv_amt REAL DEFAULT 0,
        bank_amt REAL DEFAULT 0,
        nik_amt REAL DEFAULT 0,
        created_at TEXT DEFAULT (datetime('now','localtime'))
    )''')
    conn.execute('''CREATE TABLE IF NOT EXISTS counters (
        doc_type TEXT PRIMARY KEY,
        next_num INTEGER NOT NULL
    )''')
    for t, n in [("חשבונית עסקה", 2600166), ("קבלה", 2601166), ("חשבונית עסקה + קבלה", 26002166)]:
        conn.execute("INSERT OR IGNORE INTO counters VALUES (?, ?)", (t, n))
    conn.commit()
    conn.close()
 
init_db()
 
def next_doc_num(conn, doc_type):
    row = conn.execute("SELECT next_num FROM counters WHERE doc_type=?", (doc_type,)).fetchone()
    num = row["next_num"]
    conn.execute("UPDATE counters SET next_num=? WHERE doc_type=?", (num+1, doc_type))
    return num
 
# ===== בניית Word =====
 
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
    doc_type  = data.get('doc_type', '')
    doc_num   = data.get('doc_num', '')
    date_str  = data.get('date', datetime.date.today().strftime('%d/%m/%Y'))
    client    = data.get('client', '')
    client_id = data.get('client_id', '')
    address   = data.get('address', '')
    desc      = data.get('desc', '')
    inv_amt   = data.get('inv_amt')
    bank_amt  = data.get('bank_amt')
    nik_amt   = data.get('nik_amt')
    pct_nik   = data.get('pct_nik', 0.2)
 
    def sf(v):
        try: return float(v) if v else 0.0
        except: return 0.0
 
    inv_amt = sf(inv_amt)
    bank_amt = sf(bank_amt)
    nik_amt  = sf(nik_amt)
    pct_nik  = sf(pct_nik) or 0.2
    rec_total = bank_amt + nik_amt
 
    doc = Document()
    section = doc.sections[0]
    section.page_width  = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = section.right_margin = Cm(2)
    section.top_margin  = section.bottom_margin = Cm(2)
    doc.settings.element.append(OxmlElement('w:bidi'))
 
    # הגדרת RTL לכל המסמך
    for style_name in ['Normal', 'Table Grid']:
        try:
            s = doc.styles[style_name]
            pPr = s.element.get_or_add_pPr()
            bidi = OxmlElement('w:bidi')
            pPr.append(bidi)
        except: pass
 
    title_color = {'חשבונית עסקה':'2563EB','קבלה':'16A34A'}.get(doc_type,'7C3AED')
 
    # כותרת
    for txt, sz, bold, col in [
        (doc_type, 26, True, title_color),
        (OWNER_NAME, 18, True, '1E3A5F'),
        (f"עוסק פטור מס׳ {OWNER_TAX}", 12, False, '374151'),
    ]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_rtl(p)
        r = p.add_run(txt)
        r.font.name='Arial'; r.font.size=Pt(sz); r.font.bold=bold
        r.font.color.rgb = RGBColor.from_string(col)
 
    doc.add_paragraph()
 
    # מספר + תאריך — מספר מימין, תאריך משמאל
    t1 = doc.add_table(rows=1, cols=2)
    t1.style = 'Table Grid'
    t1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
 
    # תאריך — תא שמאלי
    c_date = t1.rows[0].cells[0]
    set_cell_bg(c_date, "F8FAFC")
    p_date = c_date.paragraphs[0]
    p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_rtl(p_date)
    rd = p_date.add_run(f"תאריך: {date_str}")
    rd.font.name='Arial'; rd.font.size=Pt(12)
 
    # מספר מסמך — תא ימני
    c_num = t1.rows[0].cells[1]
    set_cell_bg(c_num, "EFF6FF")
    make_border(c_num, "2563EB")
    p_num = c_num.paragraphs[0]
    p_num.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_rtl(p_num)
    rl = p_num.add_run("מספר מסמך:  ")
    rl.font.name='Arial'; rl.font.size=Pt(12); rl.font.bold=True
    rl.font.color.rgb = RGBColor.from_string('2563EB')
    rn = p_num.add_run(str(doc_num))
    rn.font.name='Arial'; rn.font.size=Pt(16); rn.font.bold=True
    rn.font.color.rgb = RGBColor.from_string('1E3A5F')
 
    doc.add_paragraph()
 
    # פרטי לקוח — טבלה פשוטה ללא קווים כפולים
    t2 = doc.add_table(rows=3, cols=2)
    t2.style = 'Table Grid'
 
    # כותרת
    c_title = t2.rows[0].cells[0]
    c_title.merge(t2.rows[0].cells[1])
    set_cell_bg(c_title, "E8F5E9")
    p_title = c_title.paragraphs[0]
    p_title.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_rtl(p_title)
    rt = p_title.add_run("פרטי הלקוח")
    rt.font.name='Arial'; rt.font.size=Pt(12); rt.font.bold=True
    rt.font.color.rgb = RGBColor.from_string('16A34A')
 
    # שורה 1: שם לקוח | מס׳ עוסק
    for ci, (label, val) in enumerate([(f"מס׳ עוסק: {client_id}", f"שם לקוח: {client}")]):
        pass
    # שם לקוח מימין
    c_name = t2.rows[1].cells[1]
    set_cell_bg(c_name, "F8FAFC")
    p_name = c_name.paragraphs[0]; p_name.alignment=WD_ALIGN_PARAGRAPH.RIGHT; set_rtl(p_name)
    r_name = p_name.add_run(f"שם לקוח: {client}")
    r_name.font.name='Arial'; r_name.font.size=Pt(11)
 
    # מס׳ עוסק משמאל
    c_id = t2.rows[1].cells[0]
    set_cell_bg(c_id, "F8FAFC")
    p_id = c_id.paragraphs[0]; p_id.alignment=WD_ALIGN_PARAGRAPH.RIGHT; set_rtl(p_id)
    r_id = p_id.add_run(f"מס׳ עוסק: {client_id}")
    r_id.font.name='Arial'; r_id.font.size=Pt(11)
 
    # שורה 2: כתובת | תיאור
    c_addr = t2.rows[2].cells[1]
    set_cell_bg(c_addr, "F8FAFC")
    p_addr = c_addr.paragraphs[0]; p_addr.alignment=WD_ALIGN_PARAGRAPH.RIGHT; set_rtl(p_addr)
    r_addr = p_addr.add_run(f"כתובת: {address}")
    r_addr.font.name='Arial'; r_addr.font.size=Pt(11)
 
    c_desc = t2.rows[2].cells[0]
    set_cell_bg(c_desc, "F8FAFC")
    p_desc = c_desc.paragraphs[0]; p_desc.alignment=WD_ALIGN_PARAGRAPH.RIGHT; set_rtl(p_desc)
    r_desc = p_desc.add_run(f"תיאור: {desc}")
    r_desc.font.name='Arial'; r_desc.font.size=Pt(11)
 
    doc.add_paragraph()
 
    # חשבונית
    if doc_type in ('חשבונית עסקה','חשבונית עסקה + קבלה'):
        # כותרת עמודות: סכום | כמות | תיאור (מימין לשמאל)
        t3 = doc.add_table(rows=2, cols=3)
        t3.style = 'Table Grid'
        headers = ["סכום ₪", "כמות", "תיאור השירות / הפריט"]
        values  = [f"{inv_amt:,.2f}", "1", desc or ""]
        for i, h in enumerate(headers):
            c = t3.rows[0].cells[i]
            set_cell_bg(c, "2563EB")
            p = c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; set_rtl(p)
            r = p.add_run(h)
            r.font.name='Arial'; r.font.size=Pt(11); r.font.bold=True
            r.font.color.rgb = RGBColor.from_string('FFFFFF')
        for i, val in enumerate(values):
            c = t3.rows[1].cells[i]
            set_cell_bg(c, "EFF6FF")
            p = c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.RIGHT; set_rtl(p)
            r = p.add_run(val)
            r.font.name='Arial'; r.font.size=Pt(11)
 
        doc.add_paragraph()
 
        # סה"כ — ימין=סכום, שמאל=תווית
        t4 = doc.add_table(rows=1, cols=2)
        t4.style = 'Table Grid'
        cr = t4.rows[0].cells[0]  # שמאל — תווית
        cl = t4.rows[0].cells[1]  # ימין — סכום
        set_cell_bg(cr, "2563EB")
        pr = cr.paragraphs[0]; pr.alignment=WD_ALIGN_PARAGRAPH.CENTER; set_rtl(pr)
        rr = pr.add_run('סה"כ לתשלום:')
        rr.font.name='Arial'; rr.font.size=Pt(13); rr.font.bold=True
        rr.font.color.rgb = RGBColor.from_string('FFFFFF')
        set_cell_bg(cl, "DBEAFE")
        pl = cl.paragraphs[0]; pl.alignment=WD_ALIGN_PARAGRAPH.CENTER; set_rtl(pl)
        rl2 = pl.add_run(f"₪ {inv_amt:,.2f}")
        rl2.font.name='Arial'; rl2.font.size=Pt(14); rl2.font.bold=True
        rl2.font.color.rgb = RGBColor.from_string('1E3A5F')
        doc.add_paragraph()
 
    # קבלה
    if doc_type in ('קבלה','חשבונית עסקה + קבלה'):
        rows_data = [
            ('סה"כ שולם:', f"₪ {rec_total:,.2f}", "F0FDF4","16A34A"),
            ("התקבל בבנק:", f"₪ {bank_amt:,.2f}", "F0FDF4","15803D"),
            (f"ניכוי במקור ({int(pct_nik*100)}%):", f"₪ {nik_amt:,.2f}", "F5F3FF","7C3AED"),
            ("אמצעי תשלום:", "____________", "F8FAFC","374151"),
        ]
        t5 = doc.add_table(rows=4, cols=2)
        t5.style = 'Table Grid'
        for i,(label,val,bg,col) in enumerate(rows_data):
            # ימין — תווית, שמאל — סכום
            c_label = t5.rows[i].cells[1]
            c_val   = t5.rows[i].cells[0]
            set_cell_bg(c_label, bg); set_cell_bg(c_val, bg)
            p_l = c_label.paragraphs[0]; p_l.alignment=WD_ALIGN_PARAGRAPH.RIGHT; set_rtl(p_l)
            rl = p_l.add_run(label)
            rl.font.name='Arial'; rl.font.size=Pt(11); rl.font.bold=True
            rl.font.color.rgb = RGBColor.from_string(col)
            p_v = c_val.paragraphs[0]; p_v.alignment=WD_ALIGN_PARAGRAPH.CENTER; set_rtl(p_v)
            rv = p_v.add_run(val)
            rv.font.name='Arial'; rv.font.size=Pt(12); rv.font.bold=True
            rv.font.color.rgb = RGBColor.from_string(col)
        doc.add_paragraph()
