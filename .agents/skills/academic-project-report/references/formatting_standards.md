# Academic Formatting Standards & Cross-Platform Automation Reference

คู่มืออ้างอิงเชิงลึกและแม่แบบโค้ดมาตรฐานสำหรับการจัดทำรูปเล่มรายงานวิชาการ โครงงานวิจัย และปริญญานิพนธ์ (Academic Project Report & Thesis) ตามมาตรฐานบัณฑิตวิทยาลัยและระเบียบงานสารบรรณ มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ (มจพ.) และสถาบันอุดมศึกษาไทยอย่างสมบูรณ์แบบ 100%

---

## 1. การตั้งค่าระยะขอบกระดาษและเซกชัน (Page Setup in twips)

มาตรฐานขนาดกระดาษคือ A4 (ความกว้าง 8.27 นิ้ว / 11906 twips, ความสูง 11.69 นิ้ว / 16838 twips)
*สูตรแปลงหน่วย: 1 นิ้ว = 1440 twips = 72 pt = 2.54 ซม.*

```python
from docx.shared import Inches

# 1.1 ส่วนหน้าปกนอก (Cover Page Section - Zero Margin Bias)
# ปรับขอบซ้าย-ขวาเท่ากัน 1.0 นิ้ว เพื่อให้เนื้อหาและรายชื่อสมาชิกอยู่กึ่งกลางหน้ากระดาษอย่างแท้จริง
cover_sec = doc.sections[0]
cover_sec.top_margin = Inches(1.5)       # 2160 twips
cover_sec.bottom_margin = Inches(1.0)    # 1440 twips
cover_sec.left_margin = Inches(1.0)      # 1440 twips
cover_sec.right_margin = Inches(1.0)     # 1440 twips
cover_sec.different_first_page_header_footer = True

# 1.2 ส่วนเนื้อหาหลักและส่วนนำ (Body & Front Matter Sections)
# ขอบซ้าย 1.5 นิ้ว เพื่อเผื่อระยะเข้าเล่ม/เย็บเล่มสันปก
body_sec = doc.add_section()
body_sec.top_margin = Inches(1.5)       # 2160 twips
body_sec.bottom_margin = Inches(1.0)    # 1440 twips
body_sec.left_margin = Inches(1.5)      # 2160 twips (เผื่อเย็บเล่ม)
body_sec.right_margin = Inches(1.0)     # 1440 twips
body_sec.header_distance = Inches(0.75) # 1080 twips
body_sec.footer_distance = Inches(0.5)  # 720 twips
body_sec.header.is_linked_to_previous = False
```

---

## 2. การตัดคำภาษาไทยและกฎเหล็กห้ามเครื่องหมายวรรคตอนห้อยท้าย (Zero Dangling Punctuation)

การใช้ Zero-Width Space (`\u200b`) ช่วยให้ Microsoft Word และ LibreOffice ตัดบรรทัดภาษาไทยได้อย่างเป็นธรรมชาติและแม่นยำ แต่ **ต้องกำจัด ZWSP หน้าเครื่องหมายปิดและหลังเครื่องหมายเปิดทั้งหมด** เพื่อป้องกันไม่ให้เครื่องหมายวรรคตอน วงเล็บปิด หรือเครื่องหมายเปอร์เซ็นต์ หลุดไปขึ้นต้นบรรทัดใหม่อย่างโดดเดี่ยว

```python
import re
from pythainlp.tokenize import word_tokenize

NO_BREAK_BEFORE = r'[,\.\:\;\!\?\)\]\}\%\"\'\sฯๆ/”’]'
NO_BREAK_AFTER = r'[\(\[\{\s\"\'/“‘]'

def thai_zwsp(text):
    """
    Inserts zero-width spaces (\u200b) between Thai syllables using PyThaiNLP
    while strictly removing ZWSP adjacent to punctuation, quotes, brackets, and percent signs
    to enforce the Zero Dangling Punctuation Rule.
    """
    if not text:
        return ""
    lines = str(text).split('\n')
    processed = []
    for line in lines:
        if not line:
            processed.append("")
            continue
        tokens = word_tokenize(line, engine='newmm')
        joined = '\u200b'.join(tokens)
        # ลบ ZWSP หน้าเครื่องหมายปิด วงเล็บปิด % และเครื่องหมายวรรคตอน
        joined = re.sub(r'\u200b+(' + NO_BREAK_BEFORE + ')', r'\1', joined)
        # ลบ ZWSP หลังเครื่องหมายเปิด วงเล็บเปิด
        joined = re.sub('(' + NO_BREAK_AFTER + r')\u200b+', r'\1', joined)
        # ลบ ZWSP ที่ซ้ำซ้อน
        joined = re.sub(r'\u200b+', '\u200b', joined)
        processed.append(joined)
    return '\n'.join(processed)
```

---

## 3. การกำหนดฟอนต์ TH SarabunPSK ทุกมิติใน OpenXML

เพื่อป้องกันไม่ให้ฟอนต์สลับกลับเป็น Times New Roman, Arial หรือ Calibri เมื่อเปิดบนเครื่องอื่นหรือระบบปฏิบัติการอื่น ต้องกำหนดทั้ง 4 มิติ (`ascii`, `hAnsi`, `cs`, `eastAsia`) ผ่าน OpenXML:

```python
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt, RGBColor

FONT_NAME = 'TH SarabunPSK'
CLR_BLACK = RGBColor(0, 0, 0)

def set_run_font(run, size_pt=16, bold=False, italic=False, underline=False, color_rgb=CLR_BLACK):
    run.font.name = FONT_NAME
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    run.underline = underline
    run.font.color.rgb = color_rgb
    
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:ascii'), FONT_NAME)
    rFonts.set(qn('w:hAnsi'), FONT_NAME)
    rFonts.set(qn('w:cs'), FONT_NAME)
    rFonts.set(qn('w:eastAsia'), FONT_NAME)
    
    sz_val = int(round(size_pt * 2))
    for sz in rPr.findall(qn('w:sz')):
        rPr.remove(sz)
    for szCs in rPr.findall(qn('w:szCs')):
        rPr.remove(szCs)
    rPr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="{sz_val}"/>'))
    rPr.append(parse_xml(f'<w:szCs {nsdecls("w")} w:val="{sz_val}"/>'))

    for b in rPr.findall(qn('w:b')):
        rPr.remove(b)
    for bCs in rPr.findall(qn('w:bCs')):
        rPr.remove(bCs)
    if bold:
        rPr.append(parse_xml(f'<w:b {nsdecls("w")}/>'))
        rPr.append(parse_xml(f'<w:bCs {nsdecls("w")}/>'))
    else:
        rPr.append(parse_xml(f'<w:b {nsdecls("w")} w:val="0"/>'))
        rPr.append(parse_xml(f'<w:bCs {nsdecls("w")} w:val="0"/>'))

    for i_elem in rPr.findall(qn('w:i')):
        rPr.remove(i_elem)
    for iCs in rPr.findall(qn('w:iCs')):
        rPr.remove(iCs)
    if italic:
        rPr.append(parse_xml(f'<w:i {nsdecls("w")}/>'))
        rPr.append(parse_xml(f'<w:iCs {nsdecls("w")}/>'))
```

---

## 4. การจัดย่อหน้าและรายการข้อย่อยตามระเบียบงานสารบรรณไทย

### 4.1 ย่อหน้าเนื้อหาทั่วไป (Body Paragraphs)
* **First-line Indent:** **0.5 นิ้ว** (1.27 ซม. / 720 twips)
* **Alignment:** `WD_ALIGN_PARAGRAPH.THAI_JUSTIFY` เพื่อให้ขอบซ้ายและขวาเรียบเสมอกัน
* **Line Spacing:** 1.15 เท่า, Space Before 0pt, Space After 2–3pt

### 4.2 การจัดแท็บรายการข้อย่อย (Numbered & Bullet Lists)
ตามมาตรฐานงานสารบรรณไทยและคู่มือ มจพ.:
* **บรรทัดที่ 1 (ตัวเลข/สัญลักษณ์ข้อ):** ย่อหน้าลึกกว่าปกติที่ระยะ 0.8 นิ้ว (`left_indent = Inches(0.5)`, `first_line_indent = Inches(0.3)`)
* **บรรทัดที่ 2 เป็นต้นไป:** ชิดแนวระยะ 1 Tab หรือ 0.5 นิ้ว (`left_indent = Inches(0.5)`)
* **การจัดข้อความ:** ใช้ `WD_ALIGN_PARAGRAPH.THAI_JUSTIFY`

```python
def add_numbered_item(doc, num_label, text, bold_prefix=None, space_after=2, keep_with_next=False):
    p = doc.add_paragraph()
    format_paragraph(p, space_before=0, space_after=space_after, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.THAI_JUSTIFY, keep_with_next=keep_with_next)
    p.paragraph_format.left_indent = Inches(0.5)
    p.paragraph_format.first_line_indent = Inches(0.3)
    
    r_num = p.add_run(f"{num_label}  ")
    set_run_font(r_num, size_pt=16, bold=False)
    if bold_prefix:
        r_bold = p.add_run(thai_zwsp(bold_prefix))
        set_run_font(r_bold, size_pt=16, bold=True)
    r_text = p.add_run(thai_zwsp(text))
    set_run_font(r_text, size_pt=16, bold=False)
    return p
```

---

## 5. ตารางวิชาการมาตรฐาน APA (APA 3-Line Table)

กฎเกณฑ์เชิงวิชาการ:
1. มีเฉพาะเส้นแนวนอน 3 เส้น: เส้นบนสุด (1.5pt / sz=12), เส้นใต้หัวตาราง (1.0pt / sz=8), และเส้นล่างสุด (1.5pt / sz=12)
2. **ห้ามมีเส้นขอบแนวตั้งเด็ดขาด (No Vertical Borders)**
3. ป้องกันการตัดแถวแยกข้ามหน้าด้วย `<w:cantSplit/>` บนทุกแถว
4. ตั้งค่าหัวตารางซ้ำอัตโนมัติด้วย `<w:tblHeader/>` บนแถวแรก
5. ควบคุมระยะห่างภายในเซลล์ด้วย `<w:tcMar>` (dxa/twips) เพื่อความกะทัดรัดและประณีต

```python
from docx.enum.table import WD_TABLE_ALIGNMENT

def add_styled_table(doc, table_num, caption, headers, data, col_widths, font_size_pt=12.0, padding_twips=25, page_break_before=False):
    p_cap = doc.add_paragraph()
    format_paragraph(p_cap, space_before=10, space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT, keep_with_next=True)
    p_cap.paragraph_format.first_line_indent = Inches(0)
    if page_break_before:
        p_cap.paragraph_format.page_break_before = True
        
    r_lbl = p_cap.add_run(thai_zwsp(f"ตารางที่ {table_num}  "))
    set_run_font(r_lbl, size_pt=14, bold=True)
    r_cap = p_cap.add_run(thai_zwsp(caption))
    set_run_font(r_cap, size_pt=14, bold=False)

    tbl = doc.add_table(rows=len(data) + 1, cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False

    # เส้นขอบแบบ APA 3-Line Table
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
        f'  <w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
        f'  <w:insideH w:val="none"/><w:insideV w:val="none"/>'
        f'  <w:left w:val="none"/><w:right w:val="none"/>'
        f'</w:tblBorders>'
    )
    tbl._tbl.tblPr.append(tblBorders)

    for i, row in enumerate(tbl.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        if i == 0:
            trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

    # แถวหัวตาราง
    for c_idx, h_text in enumerate(headers):
        cell = tbl.cell(0, c_idx)
        cell.width = col_widths[c_idx]
        p = cell.paragraphs[0]
        format_paragraph(p, space_before=3, space_after=3, line_spacing=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(thai_zwsp(h_text))
        set_run_font(r, size_pt=font_size_pt, bold=True)
        tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/></w:tcBorders>')
        cell._tc.get_or_add_tcPr().append(tcBorders)

    # แถวข้อมูล
    for r_idx, row_values in enumerate(data):
        for c_idx, val in enumerate(row_values):
            cell = tbl.cell(r_idx + 1, c_idx)
            cell.width = col_widths[c_idx]
            p = cell.paragraphs[0]
            align = WD_ALIGN_PARAGRAPH.CENTER if str(val).isdigit() or str(val).endswith('%') else WD_ALIGN_PARAGRAPH.LEFT
            format_paragraph(p, space_before=2, space_after=2, line_spacing=1.0, align=align)
            p.paragraph_format.first_line_indent = Inches(0)
            r = p.add_run(thai_zwsp(str(val)))
            set_run_font(r, size_pt=font_size_pt, bold=False)

    # ตั้งค่า Cell Padding
    for row in tbl.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcMar = parse_xml(
                f'<w:tcMar {nsdecls("w")}>'
                f'  <w:top w:w="{padding_twips}" w:type="dxa"/>'
                f'  <w:bottom w:w="{padding_twips}" w:type="dxa"/>'
                f'  <w:left w:w="40" w:type="dxa"/>'
                f'  <w:right w:w="40" w:type="dxa"/>'
                f'</w:tcMar>'
            )
            tcPr.append(tcMar)
    return tbl
```

---

## 6. ระบบสารบัญซิงค์พิกัดหน้าจริงแบบ 2 รอบ (Two-Pass Dynamic TOC Engine)

เพื่อแก้ไขปัญหาคลาสสิกของเล่มรายงานที่ **"เลขหน้าในสารบัญไม่ตรงกับหน้าจริงในเล่ม"**:
1. **Pass 1:** บิลด์เล่มฉบับร่าง (.docx) โดยใช้เลขหน้าสมมุติ แล้วแปลงเป็น PDF
2. **Scan:** ใช้ `PyMuPDF (fitz)` สแกนอ่าน PDF เพื่อตรวจจับหัวข้อ, `ตารางที่ X.X`, และ `รูปที่ X.X` พร้อมบันทึกพิกัดหน้าจริง
3. **Pass 2:** นำพิกัดหน้าจริงมาเขียนทับหน้าสารบัญใน .docx ด้วย Right Tab Stop ที่ระยะ **8300 dxa (5.76 นิ้ว)** พร้อมจุดไข่ปลา แล้วแปลงเป็น Final PDF

```python
def add_toc_line(doc, title_text, page_num_str, indent=0.0, is_bold=False):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Inches(indent)
    p.paragraph_format.first_line_indent = Inches(0)
    pPr = p._p.get_or_add_pPr()
    tab_xml = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="right" w:leader="dot" w:pos="8300"/></w:tabs>')
    pPr.append(tab_xml)
    r1 = p.add_run(thai_zwsp(title_text))
    set_run_font(r1, size_pt=14, bold=is_bold)
    r2 = p.add_run(f"\t{page_num_str}")
    set_run_font(r2, size_pt=14, bold=is_bold)
    return p

def scan_pdf_toc_pages(pdf_path, search_patterns, body_start_page=7):
    import fitz
    doc = fitz.open(pdf_path)
    results = {}
    for p_idx, page in enumerate(doc):
        text = page.get_text()
        for key, pattern in search_patterns.items():
            if key not in results and re.search(pattern, text):
                actual_page = (p_idx + 1) - body_start_page + 1
                results[key] = str(max(1, actual_page))
    doc.close()
    return results
```

---

## 7. เลขหน้าอัตโนมัติในส่วนหัวกระดาษ (Native Word XML PAGE Field)

การแทรก Field โค้ด `PAGE` เพื่อให้ Microsoft Word คำนวณเลขหน้าอัตโนมัติที่มุมขวาบนของทุกหน้า:

```python
def add_header_page_field(header_obj):
    h_p = header_obj.paragraphs[0]
    format_paragraph(h_p, space_before=0, space_after=0, align=WD_ALIGN_PARAGRAPH.RIGHT)
    h_p.paragraph_format.first_line_indent = Inches(0)
    h_p.paragraph_format.left_indent = Inches(0)
    h_run = h_p.add_run()
    set_run_font(h_run, size_pt=14)
    h_run._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w')))
    h_run._r.append(parse_xml(r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls('w')))
    h_run._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w')))
    h_run._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w')))
```

---

## 8. การจัดตารางรายชื่อนักศึกษาหน้าปกนอก (Invisible 3-Column Table)

การใช้ตารางล่องหน 3 คอลัมน์ ไร้เส้นขอบ เพื่อการันตีว่ารายชื่อและรหัสนักศึกษาจะเรียงตรงกัน 100%:
- คอลัมน์ 1 (กว้าง 1.5 นิ้ว): คำนำหน้า + ชื่อจริง (ชิดซ้าย)
- คอลัมน์ 2 (กว้าง 1.7 นิ้ว): นามสกุล (ชิดซ้าย)
- คอลัมน์ 3 (กว้าง 1.5 นิ้ว): รหัสนักศึกษา 13 หลัก (ชิดขวา)

```python
def build_cover_page(doc, logo_path, title_en, title_th, submission_lines, group_title, students):
    sec = doc.sections[0]
    sec.top_margin = Inches(1.5)
    sec.bottom_margin = Inches(1.0)
    sec.left_margin = Inches(1.0)
    sec.right_margin = Inches(1.0)
    sec.different_first_page_header_footer = True

    # ตราสัญลักษณ์และชื่อเรื่อง...
    # ตารางรายชื่อนักศึกษา 3 คอลัมน์
    if students:
        tbl = doc.add_table(rows=len(students), cols=3)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        tbl_borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="none"/><w:bottom w:val="none"/><w:left w:val="none"/>'
            f'<w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/>'
            f'</w:tblBorders>'
        )
        tbl._tbl.tblPr.append(tbl_borders)
        col_widths = [Inches(1.5), Inches(1.7), Inches(1.5)]
        for r_idx, (pfx_first, last, sid) in enumerate(students):
            row = tbl.rows[r_idx]
            for c_idx, text in enumerate([pfx_first, last, sid]):
                cell = row.cells[c_idx]
                cell.width = col_widths[c_idx]
                p = cell.paragraphs[0]
                align = WD_ALIGN_PARAGRAPH.LEFT if c_idx < 2 else WD_ALIGN_PARAGRAPH.RIGHT
                format_paragraph(p, space_before=0, space_after=1, align=align)
                p.paragraph_format.first_line_indent = Inches(0)
                r = p.add_run(thai_zwsp(text))
                set_run_font(r, size_pt=14, bold=False)
```

---

## 9. การแปลงไฟล์เอกสารข้ามระบบปฏิบัติการ (Cross-Platform Headless Engine)

รองรับทั้ง Linux และ Windows อย่างราบรื่น:

```python
import subprocess
import sys

def convert_to_pdf(doc_path, pdf_path):
    doc_path = os.path.abspath(doc_path)
    pdf_path = os.path.abspath(pdf_path)
    out_dir = os.path.dirname(pdf_path)

    if sys.platform == 'win32':
        import win32com.client
        if doc_path.lower().endswith('.docx'):
            word = win32com.client.Dispatch("Word.Application")
            doc = word.Documents.Open(doc_path)
            doc.SaveAs(pdf_path, FileFormat=17) # 17 = wdExportFormatPDF
            doc.Close(False)
            word.Quit()
        elif doc_path.lower().endswith('.pptx'):
            ppt = win32com.client.Dispatch("PowerPoint.Application")
            prs = ppt.Presentations.Open(doc_path, WithWindow=False)
            prs.SaveAs(pdf_path, 32) # 32 = ppSaveAsPDF
            prs.Close()
            ppt.Quit()
    else:
        cmd = ['soffice', '--headless', '--convert-to', 'pdf', doc_path, '--outdir', out_dir]
        subprocess.run(cmd, check=True)
```
