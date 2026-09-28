import os
import sys

# Script to regenerate scripts/build_full_report.py with all fixes applied
report_code = '''import os
import sys
import re
import subprocess
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from pythainlp.tokenize import word_tokenize
import win32com.client
import fitz

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r'C:\\Project\\healthy-life'
ASSETS_DIR = os.path.join(BASE_DIR, 'assets')
OUTPUT_DIR = os.path.join(BASE_DIR, '00_ไฟล์ส่งงาน_Health30D_OfficeSyndrome')
os.makedirs(OUTPUT_DIR, exist_ok=True)

DOCX_OUTPUT = os.path.join(OUTPUT_DIR, '01_เล่มรายงาน_Health30D_OfficeSyndrome_ฉบับสมบูรณ์.docx')
PDF_OUTPUT = os.path.join(OUTPUT_DIR, '01_เล่มรายงาน_Health30D_OfficeSyndrome_ฉบับสมบูรณ์.pdf')

FONT_NAME = 'TH SarabunPSK'
CLR_BLACK = RGBColor(0, 0, 0)

NO_BREAK_BEFORE = r'[,\.\:\;\!\?\)\]\}\%\"\'\sฯๆ/”’]'
NO_BREAK_AFTER = r'[\(\[\{\s\"\'/“‘]'

def thai_zwsp(text):
    if not text:
        return ""
    lines = str(text).split('\\n')
    processed = []
    for line in lines:
        if not line:
            processed.append("")
            continue
        tokens = word_tokenize(line, engine='newmm')
        joined = '\\u200b'.join(tokens)
        joined = re.sub(r'\\u200b+(' + NO_BREAK_BEFORE + ')', r'\\1', joined)
        joined = re.sub('(' + NO_BREAK_AFTER + r')\\u200b+', r'\\1', joined)
        joined = re.sub(r'\\u200b+', '\\u200b', joined)
        processed.append(joined)
    return '\\n'.join(processed)

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

def format_paragraph(p, space_before=0, space_after=2, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.THAI_JUSTIFY, keep_with_next=False):
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    p.alignment = align
    p.paragraph_format.widow_control = True
    if keep_with_next:
        p.paragraph_format.keep_with_next = True

def add_body_p(doc, text, bold_prefix=None, indent=True, space_before=0, space_after=3, line_spacing=1.15, font_size_pt=16, italic=False, page_break_before=False):
    p = doc.add_paragraph()
    format_paragraph(p, space_before=space_before, space_after=space_after, line_spacing=line_spacing, align=WD_ALIGN_PARAGRAPH.THAI_JUSTIFY)
    p.paragraph_format.first_line_indent = Inches(0.5) if indent else Inches(0)
    p.paragraph_format.left_indent = Inches(0)
    if page_break_before:
        p.paragraph_format.page_break_before = True
    if bold_prefix:
        r_bold = p.add_run(thai_zwsp(bold_prefix))
        set_run_font(r_bold, size_pt=font_size_pt, bold=True)
    r_text = p.add_run(thai_zwsp(text))
    set_run_font(r_text, size_pt=font_size_pt, bold=False, italic=italic)
    return p

def add_numbered_item(doc, num_label, text, bold_prefix=None, space_before=0, space_after=2, line_spacing=1.15, font_size_pt=16, keep_with_next=False, page_break_before=False):
    p = doc.add_paragraph()
    format_paragraph(p, space_before=space_before, space_after=space_after, line_spacing=line_spacing, align=WD_ALIGN_PARAGRAPH.THAI_JUSTIFY, keep_with_next=keep_with_next)
    p.paragraph_format.left_indent = Inches(0.5)
    p.paragraph_format.first_line_indent = Inches(0.3)
    if page_break_before:
        p.paragraph_format.page_break_before = True
    r_num = p.add_run(f"{num_label}  ")
    set_run_font(r_num, size_pt=font_size_pt, bold=True)
    if bold_prefix:
        r_bp = p.add_run(thai_zwsp(bold_prefix))
        set_run_font(r_bp, size_pt=font_size_pt, bold=True)
    r_txt = p.add_run(thai_zwsp(text))
    set_run_font(r_txt, size_pt=font_size_pt, bold=False)
    return p

def add_bullet_item(doc, text, bold_prefix=None, space_before=0, space_after=2, line_spacing=1.15, font_size_pt=16, keep_with_next=False, page_break_before=False):
    p = doc.add_paragraph()
    format_paragraph(p, space_before=space_before, space_after=space_after, line_spacing=line_spacing, align=WD_ALIGN_PARAGRAPH.THAI_JUSTIFY, keep_with_next=keep_with_next)
    p.paragraph_format.left_indent = Inches(0.5)
    p.paragraph_format.first_line_indent = Inches(0.3)
    if page_break_before:
        p.paragraph_format.page_break_before = True
    r_bullet = p.add_run("•  ")
    set_run_font(r_bullet, size_pt=font_size_pt, bold=True)
    if bold_prefix:
        r_bp = p.add_run(thai_zwsp(bold_prefix))
        set_run_font(r_bp, size_pt=font_size_pt, bold=True)
    r_txt = p.add_run(thai_zwsp(text))
    set_run_font(r_txt, size_pt=font_size_pt, bold=False)
    return p

def add_chapter_title(doc, chapter_num, title_text, space_before=14, space_after=14):
    p = doc.add_paragraph()
    format_paragraph(p, space_before=space_before, space_after=space_after, align=WD_ALIGN_PARAGRAPH.CENTER, keep_with_next=True)
    p.paragraph_format.first_line_indent = Inches(0)
    r_num = p.add_run(f"บทที่ {chapter_num}\\n")
    set_run_font(r_num, size_pt=20, bold=True)
    r_title = p.add_run(thai_zwsp(title_text))
    set_run_font(r_title, size_pt=18, bold=True)
    return p

def add_heading_1(doc, title_text, space_before=10, space_after=4, page_break_before=False):
    p = doc.add_paragraph()
    format_paragraph(p, space_before=space_before, space_after=space_after, align=WD_ALIGN_PARAGRAPH.LEFT, keep_with_next=True)
    p.paragraph_format.first_line_indent = Inches(0)
    if page_break_before:
        p.paragraph_format.page_break_before = True
    r = p.add_run(thai_zwsp(title_text))
    set_run_font(r, size_pt=16, bold=True)
    return p

def add_heading_2(doc, title_text, space_before=8, space_after=3, page_break_before=False):
    p = doc.add_paragraph()
    format_paragraph(p, space_before=space_before, space_after=space_after, align=WD_ALIGN_PARAGRAPH.LEFT, keep_with_next=True)
    p.paragraph_format.first_line_indent = Inches(0.5)
    if page_break_before:
        p.paragraph_format.page_break_before = True
    r = p.add_run(thai_zwsp(title_text))
    set_run_font(r, size_pt=16, bold=True)
    return p

def add_front_matter_title(doc, title_text, space_before=14, space_after=14):
    p = doc.add_paragraph()
    format_paragraph(p, space_before=space_before, space_after=space_after, align=WD_ALIGN_PARAGRAPH.CENTER, keep_with_next=True)
    p.paragraph_format.first_line_indent = Inches(0)
    r = p.add_run(thai_zwsp(title_text))
    set_run_font(r, size_pt=18, bold=True)
    return p

# STRICT APA TABLE FORMATTER - TOTAL WIDTH GUARANTEED <= 5.75 INCHES (ZERO OVERFLOW)
def add_styled_table(doc, table_num, caption, headers, data, col_widths, font_size_pt=10.0, padding_twips=15, page_break_before=False):
    p_cap = doc.add_paragraph()
    format_paragraph(p_cap, space_before=8, space_after=3, align=WD_ALIGN_PARAGRAPH.LEFT, keep_with_next=True)
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
    
    total_w_dxa = int(sum([w.inches for w in col_widths]) * 1440)
    
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
        f'  <w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
        f'  <w:insideH w:val="none"/>'
        f'  <w:insideV w:val="none"/>'
        f'  <w:left w:val="none"/>'
        f'  <w:right w:val="none"/>'
        f'</w:tblBorders>'
    )
    tbl._tbl.tblPr.append(tblBorders)
    tbl._tbl.tblPr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{total_w_dxa}" w:type="dxa"/>'))
    tbl._tbl.tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
    
    for i, row in enumerate(tbl.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        if i == 0:
            trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
            
    for c_idx, h_text in enumerate(headers):
        cell = tbl.cell(0, c_idx)
        cell.width = col_widths[c_idx]
        w_dxa = int(col_widths[c_idx].inches * 1440)
        cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{w_dxa}" w:type="dxa"/>'))
        
        p = cell.paragraphs[0]
        format_paragraph(p, space_before=2, space_after=2, line_spacing=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        p.paragraph_format.first_line_indent = Inches(0)
        r = p.add_run(thai_zwsp(h_text))
        set_run_font(r, size_pt=font_size_pt, bold=True)
        tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/></w:tcBorders>')
        cell._tc.get_or_add_tcPr().append(tcBorders)
        
    for r_idx, row_values in enumerate(data):
        for c_idx, val in enumerate(row_values):
            cell = tbl.cell(r_idx + 1, c_idx)
            cell.width = col_widths[c_idx]
            w_dxa = int(col_widths[c_idx].inches * 1440)
            cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{w_dxa}" w:type="dxa"/>'))
            
            p = cell.paragraphs[0]
            format_paragraph(p, space_before=1.5, space_after=1.5, line_spacing=1.0, align=WD_ALIGN_PARAGRAPH.LEFT)
            p.paragraph_format.first_line_indent = Inches(0)
            r = p.add_run(thai_zwsp(str(val)))
            set_run_font(r, size_pt=font_size_pt, bold=False)
            
    for row in tbl.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcMar = parse_xml(
                f'<w:tcMar {nsdecls("w")}>'
                f'  <w:top w:w="{padding_twips}" w:type="dxa"/>'
                f'  <w:bottom w:w="{padding_twips}" w:type="dxa"/>'
                f'  <w:left w:w="30" w:type="dxa"/>'
                f'  <w:right w:w="30" w:type="dxa"/>'
                f'</w:tcMar>'
            )
            tcPr.append(tcMar)
    return tbl

def add_figure(doc, img_path, fig_num, caption, width_inches=3.4, height_inches=None, source_text=None, page_break_before=False):
    if not os.path.exists(img_path):
        return None
    p_img = doc.add_paragraph()
    format_paragraph(p_img, space_before=6, space_after=2, align=WD_ALIGN_PARAGRAPH.CENTER, keep_with_next=True)
    p_img.paragraph_format.first_line_indent = Inches(0)
    if page_break_before:
        p_img.paragraph_format.page_break_before = True
    if height_inches:
        p_img.add_run().add_picture(img_path, height=Inches(height_inches))
    else:
        p_img.add_run().add_picture(img_path, width=Inches(width_inches))
        
    p_cap = doc.add_paragraph()
    format_paragraph(p_cap, space_before=2, space_after=2 if source_text else 6, align=WD_ALIGN_PARAGRAPH.CENTER, keep_with_next=bool(source_text))
    p_cap.paragraph_format.first_line_indent = Inches(0)
    r_lbl = p_cap.add_run(thai_zwsp(f"รูปที่ {fig_num}  "))
    set_run_font(r_lbl, size_pt=14, bold=True)
    r_cap = p_cap.add_run(thai_zwsp(caption))
    set_run_font(r_cap, size_pt=14, bold=False)
    
    if source_text:
        p_src = doc.add_paragraph()
        format_paragraph(p_src, space_before=0, space_after=6, align=WD_ALIGN_PARAGRAPH.CENTER)
        p_src.paragraph_format.first_line_indent = Inches(0)
        r_src = p_src.add_run(thai_zwsp(f"(ที่มา: {source_text})"))
        set_run_font(r_src, size_pt=12, italic=True)
    return p_cap

def add_reference_item(doc, ref_text):
    p = doc.add_paragraph()
    format_paragraph(p, space_before=2, space_after=4, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.THAI_JUSTIFY)
    p.paragraph_format.left_indent = Inches(0.5)
    p.paragraph_format.first_line_indent = Inches(-0.5)
    r = p.add_run(thai_zwsp(ref_text))
    set_run_font(r, size_pt=15, bold=False)
    return p

def add_toc_line(doc, title_text, page_num_str, indent=0.0, is_bold=False):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = 1.05
    p.paragraph_format.left_indent = Inches(indent)
    p.paragraph_format.first_line_indent = Inches(0)
    pPr = p._p.get_or_add_pPr()
    tab_xml = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="right" w:leader="dot" w:pos="8280"/></w:tabs>')
    pPr.append(tab_xml)
    r1 = p.add_run(thai_zwsp(title_text))
    set_run_font(r1, size_pt=13.5, bold=is_bold)
    r2 = p.add_run(f"\\t{page_num_str}")
    set_run_font(r2, size_pt=13.5, bold=is_bold)
    return p

def add_toc_col_header(doc, left_label="เรื่อง", right_label="หน้า"):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.left_indent = Inches(0)
    p.paragraph_format.first_line_indent = Inches(0)
    pPr = p._p.get_or_add_pPr()
    tab_xml = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="right" w:pos="8280"/></w:tabs>')
    pPr.append(tab_xml)
    r1 = p.add_run(thai_zwsp(left_label))
    set_run_font(r1, size_pt=14, bold=True)
    r2 = p.add_run(f"\\t{right_label}")
    set_run_font(r2, size_pt=14, bold=True)
    return p

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

def clear_pg_num_types(sec):
    for pnt in sec._sectPr.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pgNumType'):
        sec._sectPr.remove(pnt)

def setup_section_margins(sec, top=1.5, bottom=1.0, left=1.5, right=1.0, is_cover=False):
    sec.top_margin = Inches(top)
    sec.bottom_margin = Inches(bottom)
    sec.left_margin = Inches(1.0 if is_cover else left)
    sec.right_margin = Inches(1.0 if is_cover else right)
    sec.header_distance = Inches(0.75)
    sec.footer_distance = Inches(0.5)
    if is_cover:
        sec.different_first_page_header_footer = True
    else:
        sec.header.is_linked_to_previous = False
        sec.footer.is_linked_to_previous = False

DEFAULT_PAGE_MAP = {
    'abstract_th': 'ก', 'abstract_en': 'ข', 'ack': 'ค',
    'toc': 'ง', 'lot': 'ฉ', 'lof': 'ช',
    'ch1': '1', 'ch1_2': '2', 'ch1_3': '2', 'ch1_4': '3', 'ch1_5': '4', 'ch1_6': '4',
    'ch2': '5', 'ch2_2': '6', 'ch2_3': '7', 'ch2_4': '8', 'ch2_5': '9', 'ch2_6': '10', 'ch2_7': '11', 'ch2_8': '12', 'ch2_9': '13',
    'ch3': '14', 'ch3_2': '14', 'ch3_3': '15', 'ch3_4': '16', 'ch3_5': '18', 'ch3_6': '19', 'ch3_7': '20',
    'ch4': '21', 'ch4_2': '22', 'ch4_3': '24', 'ch4_4': '26', 'ch4_5': '27', 'ch4_6': '28', 'ch4_7': '29', 'ch4_8': '30',
    'ch5': '31', 'ch5_2': '32', 'ch5_3': '33', 'ch5_4': '34', 'ch5_5': '35',
    'ref': '36', 'app_a': '38', 'app_b': '40', 'app_c': '42', 'app_d': '44', 'bio': '46',
    'tbl_1_1': '3', 'tbl_2_1': '12', 'tbl_2_2': '10',
    'tbl_3_1': '15', 'tbl_3_2': '19',
    'tbl_4_1': '21', 'tbl_4_2': '23', 'tbl_4_3': '25', 'tbl_4_4': '26', 'tbl_4_5': '29',
    'fig_1_1': '2', 'fig_2_1': '13',
    'fig_3_1': '16', 'fig_3_2': '17', 'fig_3_3': '18', 'fig_3_4': '19',
    'fig_4_1': '24', 'fig_4_2': '25', 'fig_4_3': '27', 'fig_4_4': '30'
}

def generate_report_docx(toc_pages=None, output_path=DOCX_OUTPUT):
    p_map = DEFAULT_PAGE_MAP.copy()
    if toc_pages:
        p_map.update(toc_pages)
        
    doc = Document()
    
    # ==========================================
    # SECTION 0: COVER PAGE (Zero Margin Bias: 1.0" Left/Right)
    # ==========================================
    sec_cover = doc.sections[0]
    sec_cover.top_margin = Inches(1.5)
    sec_cover.bottom_margin = Inches(1.0)
    sec_cover.left_margin = Inches(1.0)
    sec_cover.right_margin = Inches(1.0)
    sec_cover.different_first_page_header_footer = True
    clear_pg_num_types(sec_cover)
    
    logo_path = os.path.join(ASSETS_DIR, 'kmutnb_logo.png')
    p_logo = doc.add_paragraph()
    format_paragraph(p_logo, space_before=0, space_after=8, align=WD_ALIGN_PARAGRAPH.CENTER)
    p_logo.paragraph_format.first_line_indent = Inches(0)
    if os.path.exists(logo_path):
        p_logo.add_run().add_picture(logo_path, width=Inches(1.2))
        
    p_ten = doc.add_paragraph()
    format_paragraph(p_ten, space_before=4, space_after=4, align=WD_ALIGN_PARAGRAPH.CENTER)
    p_ten.paragraph_format.first_line_indent = Inches(0)
    r_en = p_ten.add_run(
        "PERSONALIZED ERGONOMIC DESK ADJUSTMENT AND SELF-CARE PROTOCOL\\n"
        "FOR ALLEVIATING NECK, SHOULDER, AND UPPER BACK\\n"
        "OFFICE SYNDROME SYMPTOMS: A 30-DAY N-OF-1 STUDY"
    )
    set_run_font(r_en, size_pt=14, bold=True)
    
    p_tth = doc.add_paragraph()
    format_paragraph(p_tth, space_before=4, space_after=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    p_tth.paragraph_format.first_line_indent = Inches(0)
    r_th = p_tth.add_run(thai_zwsp(
        "การปรับโต๊ะทำงานร่วมกับพฤติกรรมการดูแลตนเอง เพื่อลดอาการ Office Syndrome\\n"
        "บริเวณคอ บ่า และไหล่ ภายใน 30 วัน"
    ))
    set_run_font(r_th, size_pt=18, bold=True)
    
    p_crs = doc.add_paragraph()
    format_paragraph(p_crs, space_before=10, space_after=16, align=WD_ALIGN_PARAGRAPH.CENTER)
    p_crs.paragraph_format.first_line_indent = Inches(0)
    r_c = p_crs.add_run(thai_zwsp(
        "รายงานโครงงานวิจัยและปรับเปลี่ยนพฤติกรรมสุขภาพส่วนบุคคล\\n"
        "(Single-Subject Pre-Post Research Design)\\n"
        "รหัสวิชา 080303609 รายวิชา สุขภาพเพื่อชีวิต (Healthy Life) กลุ่มเรียนที่ 2 (Sec 2)\\n"
        "มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ"
    ))
    set_run_font(r_c, size_pt=15, bold=False)
    
    p_auth = doc.add_paragraph()
    format_paragraph(p_auth, space_before=8, space_after=6, align=WD_ALIGN_PARAGRAPH.CENTER)
    p_auth.paragraph_format.first_line_indent = Inches(0)
    r_a = p_auth.add_run("จัดทำโดย")
    set_run_font(r_a, size_pt=16, bold=True)
    
    p_std = doc.add_paragraph()
    format_paragraph(p_std, space_before=2, space_after=16, align=WD_ALIGN_PARAGRAPH.CENTER)
    p_std.paragraph_format.first_line_indent = Inches(0)
    r_std = p_std.add_run(thai_zwsp(
        "นายปภาวิน ธิติชุณหกุล\\n"
        "รหัสนักศึกษา 6806021612037\\n"
        "นักศึกษาชั้นปีที่ 1 แขนงวิชาเทคโนโลยีสารสนเทศ (IT)\\n"
        "ภาควิชาเทคโนโลยีสารสนเทศ คณะเทคโนโลยีและการจัดการอุตสาหกรรม\\n"
        "มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ วิทยาเขตปราจีนบุรี"
    ))
    set_run_font(r_std, size_pt=15, bold=False)
    
    p_sub = doc.add_paragraph()
    format_paragraph(p_sub, space_before=12, space_after=0, align=WD_ALIGN_PARAGRAPH.CENTER)
    p_sub.paragraph_format.first_line_indent = Inches(0)
    r_sub = p_sub.add_run(thai_zwsp(
        "เสนอ\\n"
        "อาจารย์ประจำรายวิชาสุขภาพเพื่อชีวิต\\n"
        "ภาคการศึกษาที่ 1 ปีการศึกษา 2569"
    ))
    set_run_font(r_sub, size_pt=15, bold=False)

    # ==========================================
    # SECTION 1: APPROVAL CERTIFICATE
    # ==========================================
    sec_cert = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section_margins(sec_cert, top=1.5, bottom=1.0, left=1.5, right=1.0)
    clear_pg_num_types(sec_cert)
    
    add_front_matter_title(doc, "ใบรับรองการดำเนินโครงงาน\\nรายวิชาศึกษาทั่วไป หมวดสุขภาพเพื่อชีวิต")
    add_body_p(doc, "หัวข้อโครงงานวิจัย: การปรับโต๊ะทำงานร่วมกับพฤติกรรมการดูแลตนเอง เพื่อลดอาการ Office Syndrome บริเวณคอ บ่า และไหล่ ภายใน 30 วัน", bold_prefix="ชื่อโครงงาน: ")
    add_body_p(doc, "PERSONALIZED ERGONOMIC DESK ADJUSTMENT AND SELF-CARE PROTOCOL FOR ALLEVIATING NECK, SHOULDER, AND UPPER BACK OFFICE SYNDROME SYMPTOMS: A 30-DAY N-OF-1 STUDY", bold_prefix="Project Title: ", italic=True)
    add_body_p(doc, "นายปภาวิน ธิติชุณหกุล  รหัสนักศึกษา 6806021612037", bold_prefix="ผู้จัดทำโครงงาน: ")
    add_body_p(doc, "แขนงวิชาเทคโนโลยีสารสนเทศ ภาควิชาเทคโนโลยีสารสนเทศ คณะเทคโนโลยีและการจัดการอุตสาหกรรม มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ วิทยาเขตปราจีนบุรี", bold_prefix="สังกัด: ")
    add_body_p(doc, "080303609 สุขภาพเพื่อชีวิต (Healthy Life) กลุ่มเรียนที่ 2 (Sec 2)", bold_prefix="รหัสวิชา: ")
    
    add_body_p(doc, (
        "คณะกรรมการผู้ประเมินและอาจารย์ประจำรายวิชา ได้พิจารณารายงานโครงงานวิจัยและปรับเปลี่ยนพฤติกรรมสุขภาพส่วนบุคคลฉบับนี้แล้ว "
        "เห็นชอบว่ามีเนื้อหาสอดคล้องตามเกณฑ์มาตรฐานทางวิชาการ มีการบูรณาการองค์ความรู้ด้านสรีรวิทยาและหลักการยศาสตร์อย่างถูกต้อง "
        "และมีหลักฐานเชิงประจักษ์สนับสนุนการวิจัยอย่างครบถ้วน สมควรอนุมัติให้เป็นส่วนหนึ่งของการประเมินผลการเรียนรู้ในรายวิชาดังกล่าวได้"
    ), space_before=10, space_after=24)
    
    p_sig = doc.add_paragraph()
    format_paragraph(p_sig, space_before=20, space_after=0, align=WD_ALIGN_PARAGRAPH.RIGHT)
    p_sig.paragraph_format.first_line_indent = Inches(0)
    r_sig = p_sig.add_run(thai_zwsp(
        "ลงชื่อ..........................................................อาจารย์ผู้ประเมิน\\n"
        "(..........................................................)\\n"
        "วันที่........เดือน....................พ.ศ. 2569"
    ))
    set_run_font(r_sig, size_pt=15)

    # ==========================================
    # SECTION 2: THAI ABSTRACT (1 page budget)
    # ==========================================
    sec_abs_th = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section_margins(sec_abs_th, top=1.5, bottom=1.0, left=1.5, right=1.0)
    clear_pg_num_types(sec_abs_th)
    pgNum_th = parse_xml(f'<w:pgNumType {nsdecls("w")} w:fmt="thaiLetters" w:start="1"/>')
    sec_abs_th._sectPr.append(pgNum_th)
    add_header_page_field(sec_abs_th.header)
    
    add_front_matter_title(doc, "บทคัดย่อภาษาไทย")
    add_body_p(doc, (
        "โครงงานวิจัยรายบุคคลนี้ มีวัตถุประสงค์เพื่อ: 1) ศึกษาประสิทธิผลของการปรับสภาพแวดล้อมโต๊ะทำงานตามหลักการยศาสตร์ (Ergonomics 90-90-90) "
        "ร่วมกับการสร้างวินัยการลุกพักขยับตัวทุก 45 นาที และการบำบัดด้วยการนวดกดจุดสะท้อน 6 นาทีควบคู่กับกายบริหาร 6 ท่า "
        "ในการลดระดับความปวดเมื่อยกล้ามเนื้อบริเวณคอ บ่า และไหล่ (Upper Trapezius and Levator Scapulae) ของนักศึกษาเทคโนโลยีสารสนเทศ "
        "2) ติดตามการฟื้นฟูของมุมข้อต่อสรีรศาสตร์ 3 จุดสำคัญ (ข้อศอก สะโพก ข้อเข่า) จากภาพถ่ายด้านข้าง และ 3) พัฒนาระบบบันทึกข้อมูลเชิงประจักษ์ที่มีความเที่ยงตรง "
        "ผ่านสมุดบันทึกข้อมูลดิจิทัล Microsoft Excel (6806021612037_Health30D.xlsx) บนระบบคลาวด์ OneDrive ร่วมกับการตั้งเวลาเตือนบนสมาร์ตโฟน และแอปพลิเคชัน ImageMeter "
        "โดยดำเนินการวิจัยแบบวัดผลก่อนและหลังการทดลอง (Single-Subject N-of-1 Pre-Post Design) ตลอดระยะเวลา 30 วัน และรายงานผลเชิงประจักษ์ระยะครึ่งทาง (Days 1–14)"
    ), space_after=3)
    
    add_body_p(doc, (
        "ผลการวิจัยเชิงประจักษ์ในระยะครึ่งทาง (14 วันแรก) พบว่า: 1) ระดับความปวดบริเวณคอบ่าไหล่ (Numeric Rating Scale: NRS 0–10) "
        "ลดลงอย่างมีนัยสำคัญทางคลินิกจากค่าตั้งต้น (Baseline D1) 6.0 คะแนน สู่ค่าเฉลี่ย 4.0 คะแนนในระยะครึ่งทาง (D12–D14) "
        "คิดเป็นอัตราการลดลงของความปวดร้อยละ 33.3 บรรลุเกณฑ์เป้าหมายโครงการ (≥30%) โดยการนวดกดจุดสะท้อน 6 นาที ช่วยลดอาการปวดเฉียบพลันหลังเลิกงานได้เฉลี่ยร้อยละ 21.6 ทันที "
        "2) สรีระท่านั่งสามารถผ่านเกณฑ์มาตรฐาน 90-90-90 เพิ่มขึ้นจากร้อยละ 25.0 ในวันแรก สู่ร้อยละ 75.0 (3 ใน 4 ข้อ) ในวันที่ 14 "
        "โดยมุมข้อศอกปรับลดลงจาก 138° สู่ 98° และมุมข้อเข่าปรับจาก 65° สู่ 92° จากการวางเท้าราบบนพื้น ส่วนมุมสะโพกพัฒนาจาก 125° เป็น 108° "
        "และ 3) ผู้วิจัยมีอัตราการปฏิบัติตามวินัยการพักขยับตัวทุก 45 นาที เฉลี่ยร้อยละ 89.2 (บรรลุเกณฑ์ ≥80%) และมีความสม่ำเสมอในการทำโปรแกรมนวดและกายบริหารร้อยละ 85.7 "
        "สะท้อนให้เห็นว่าการบูรณาการปรับสถานีงานร่วมกับการเตือนพักและการดูแลตนเองผ่านเครื่องมือดิจิทัลที่เข้าถึงง่าย สามารถบรรเทาอาการออฟฟิศซินโดรมได้อย่างมีประสิทธิภาพและยั่งยืน"
    ), space_after=10)
    
    add_body_p(doc, "ออฟฟิศซินโดรม, การยศาสตร์ 90-90-90, การวิจัยรายบุคคล N-of-1, การนวดกดจุดสะท้อน, ไมโครซอฟท์ เอ็กเซล, การพักขยับตัว", bold_prefix="คำสำคัญ: ")

    # ==========================================
    # SECTION 3: ENGLISH ABSTRACT (1 page budget)
    # ==========================================
    sec_abs_en = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section_margins(sec_abs_en, top=1.5, bottom=1.0, left=1.5, right=1.0)
    clear_pg_num_types(sec_abs_en)
    add_header_page_field(sec_abs_en.header)
    
    add_front_matter_title(doc, "ABSTRACT")
    add_body_p(doc, (
        "This single-subject (N-of-1) pre-post experimental study investigates the efficacy of a comprehensive self-care intervention "
        "combining ergonomic workstation reconfiguration (the 90-90-90 rule), smartphone-alerted active movement breaks every 45 minutes, "
        "and a daily 6-minute self-acupressure and 6-exercise stretching protocol for alleviating neck, shoulder, and upper back Office Syndrome "
        "symptoms in a first-year Information Technology undergraduate student. The empirical evidence protocol was conducted over a 30-day timeline, "
        "utilizing a cloud-synchronized Microsoft Excel research workbook (6806021612037_Health30D.xlsx) on OneDrive, smartphone timers, "
        "and digital photogrammetric joint angle analysis via the ImageMeter mobile application. Midterm findings across the first 14 days were evaluated."
    ), space_after=3, font_size_pt=14, line_spacing=1.1)
    
    add_body_p(doc, (
        "Midterm quantitative empirical results (Days 1–14) demonstrated significant improvements across all five pre-specified key results (KPIs): "
        "1) Mean musculoskeletal pain intensity measured by the Numeric Rating Scale (NRS 0–10) decreased from a baseline of 6.0/10 (D1) "
        "to 4.0/10 during the D12–D14 evaluation window, representing a 33.3% net pain reduction and surpassing the ≥30% benchmark, while the 6-minute "
        "self-acupressure produced an immediate post-massage pain alleviation of 21.6%. 2) Postural alignment compliance improved from 25.0% (1/4 items) "
        "at baseline to 75.0% (3/4 items) at Day 14, with elbow flexion optimizing from 138° to 98° and knee flexion from 65° to 92° via flat foot placement, "
        "while trunk-hip angle improved from 125° to 108°. 3) Active break adherence reached 89.2% (exceeding the ≥80% target), and overall intervention "
        "protocol completion reached 85.7% (12/14 days). These findings confirm that combining physical workspace adjustment with behavioral self-monitoring "
        "using accessible cloud tools delivers sustainable, clinically meaningful musculoskeletal relief."
    ), space_after=10, font_size_pt=14, line_spacing=1.1)
    
    add_body_p(doc, "Office Syndrome, Computer Ergonomics 90-90-90, Single-Subject N-of-1 Study, Acupressure, Microsoft Excel, Active Microbreaks", bold_prefix="Keywords: ", font_size_pt=14)

    # ==========================================
    # SECTION 4: ACKNOWLEDGEMENTS (1 page budget)
    # ==========================================
    sec_ack = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section_margins(sec_ack, top=1.5, bottom=1.0, left=1.5, right=1.0)
    clear_pg_num_types(sec_ack)
    add_header_page_field(sec_ack.header)
    
    add_front_matter_title(doc, "กิตติกรรมประกาศ")
    add_body_p(doc, (
        "โครงงานวิจัยและปรับเปลี่ยนพฤติกรรมสุขภาพส่วนบุคคล เรื่อง 'การปรับโต๊ะทำงานร่วมกับพฤติกรรมการดูแลตนเอง เพื่อลดอาการ Office Syndrome "
        "บริเวณคอ บ่า และไหล่ ภายใน 30 วัน' ฉบับนี้ สำเร็จลุล่วงไปได้ด้วยความกรุณาและความอนุเคราะห์เป็นอย่างยิ่งจากคณาจารย์ประจำรายวิชาศึกษาทั่วไป "
        "หมวดสุขภาพเพื่อชีวิต (Healthy Life) รหัสวิชา 080303609 คณะเทคโนโลยีและการจัดการอุตสาหกรรม มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ "
        "ที่ได้ถ่ายทอดองค์ความรู้ด้านสุขภาวะ สรีรวิทยาการออกกำลังกาย และหลักการยศาสตร์เบื้องต้น ตลอดจนให้คำแนะนำและกรอบแนวคิดการตั้งตัวชี้วัดที่วัดผลได้จริง"
    ), space_after=3)
    add_body_p(doc, (
        "ขอขอบพระคุณภาควิชาเทคโนโลยีสารสนเทศ คณะเทคโนโลยีและการจัดการอุตสาหกรรม มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ วิทยาเขตปราจีนบุรี "
        "ที่ได้บ่มเพาะทักษะการประยุกต์ใช้เทคโนโลยีดิจิทัลและการจัดการข้อมูลสารสนเทศส่วนบุคคลอย่างเป็นระบบ "
        "ซึ่งเอื้ออำนวยให้ผู้วิจัยสามารถออกแบบสมุดบันทึกดิจิทัลบน Microsoft Excel และนำเครื่องมือวัดมุมในภาพถ่ายมาใช้เป็นเครื่องมือเก็บหลักฐานเชิงประจักษ์ได้อย่างมีประสิทธิภาพ"
    ), space_after=3)
    add_body_p(doc, (
        "ท้ายที่สุดนี้ ผู้วิจัยขอกราบขอบพระคุณบิดา มารดา และครอบครัว ที่ให้การสนับสนุนและสร้างสภาพแวดล้อมที่อบอุ่นในการศึกษาเล่าเรียนเสมอมา "
        "รวมทั้งเพื่อนนักศึกษาชั้นปีที่ 1 สาขาเทคโนโลยีสารสนเทศ ที่ได้ร่วมแลกเปลี่ยนข้อคิดเห็นและเป็นกำลังใจในการดำเนินงาน "
        "ผู้วิจัยหวังเป็นอย่างยิ่งว่าองค์ความรู้และระเบียบวิธีวิจัยที่บันทึกไว้ในรายงานฉบับนี้ จะเป็นประโยชน์ต่อผู้ที่ประสบปัญหาออฟฟิศซินโดรมและผู้สนใจสืบไป"
    ), space_after=20)
    
    p_ack_sign = doc.add_paragraph()
    format_paragraph(p_ack_sign, space_before=16, space_after=0, align=WD_ALIGN_PARAGRAPH.RIGHT)
    p_ack_sign.paragraph_format.first_line_indent = Inches(0)
    r_as = p_ack_sign.add_run(thai_zwsp(
        "นายปภาวิน ธิติชุณหกุล\\n"
        "ผู้วิจัยและผู้จัดทำโครงงาน"
    ))
    set_run_font(r_as, size_pt=15)

    # ==========================================
    # SECTION 5: TABLE OF CONTENTS (2 pages budget)
    # ==========================================
    sec_toc = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section_margins(sec_toc, top=1.5, bottom=1.0, left=1.5, right=1.0)
    clear_pg_num_types(sec_toc)
    add_header_page_field(sec_toc.header)
    
    add_front_matter_title(doc, "สารบัญ")
    add_toc_col_header(doc, "เรื่อง", "หน้า")
    
    add_toc_line(doc, "บทคัดย่อภาษาไทย", p_map['abstract_th'], indent=0, is_bold=True)
    add_toc_line(doc, "ABSTRACT", p_map['abstract_en'], indent=0, is_bold=True)
    add_toc_line(doc, "กิตติกรรมประกาศ", p_map['ack'], indent=0, is_bold=True)
    add_toc_line(doc, "สารบัญตาราง", p_map['lot'], indent=0, is_bold=True)
    add_toc_line(doc, "สารบัญภาพ", p_map['lof'], indent=0, is_bold=True)
    
    add_toc_line(doc, "บทที่ 1 บทนำ", p_map['ch1'], indent=0, is_bold=True)
    add_toc_line(doc, "1.1 ความเป็นมาและความสำคัญของปัญหา", p_map['ch1'], indent=0.25)
    add_toc_line(doc, "1.2 วัตถุประสงค์ของโครงงาน", p_map['ch1_2'], indent=0.25)
    add_toc_line(doc, "1.3 ขอบเขตของโครงงาน", p_map['ch1_3'], indent=0.25)
    add_toc_line(doc, "1.4 สมมติฐานและตัวชี้วัดความสำเร็จ (Measurable Key Results)", p_map['ch1_4'], indent=0.25)
    add_toc_line(doc, "1.5 ประโยชน์ที่คาดว่าจะได้รับ", p_map['ch1_5'], indent=0.25)
    add_toc_line(doc, "1.6 นิยามศัพท์เฉพาะ", p_map['ch1_6'], indent=0.25)
    
    add_toc_line(doc, "บทที่ 2 วรรณกรรม ทฤษฎี และงานวิจัยที่เกี่ยวข้อง", p_map['ch2'], indent=0, is_bold=True)
    add_toc_line(doc, "2.1 กายวิภาคศาสตร์และชีวกลศาสตร์ของกลุ่มอาการออฟฟิศซินโดรม", p_map['ch2'], indent=0.25)
    add_toc_line(doc, "2.2 ทฤษฎีการยศาสตร์และการจัดสถานีงานคอมพิวเตอร์ (Ergonomics 90-90-90)", p_map['ch2_2'], indent=0.25)
    add_toc_line(doc, "2.3 สรีรวิทยาของการพักขยับตัวและผลกระทบของพฤติกรรมเนือยนิ่ง", p_map['ch2_3'], indent=0.25)
    add_toc_line(doc, "2.4 กลไกการบำบัดด้วยการนวดกดจุดสะท้อนและการยืดเหยียดกล้ามเนื้อ", p_map['ch2_4'], indent=0.25)
    add_toc_line(doc, "2.5 ระเบียบวิธีวิจัยรายบุคคลแบบวัดผลก่อนและหลัง (Single-Subject N-of-1 Design)", p_map['ch2_5'], indent=0.25)
    add_toc_line(doc, "2.6 เครื่องมือและเทคโนโลยีการบันทึกข้อมูลเชิงประจักษ์ (Excel & Photogrammetry)", p_map['ch2_6'], indent=0.25)
    add_toc_line(doc, "2.7 การสังเคราะห์วรรณกรรมและกรอบแนวคิดการวิจัย", p_map['ch2_7'], indent=0.25)
    
    add_toc_line(doc, "บทที่ 3 วิธีดำเนินการวิจัยและการเก็บรวบรวมหลักฐานเชิงประจักษ์", p_map['ch3'], indent=0, is_bold=True)
    add_toc_line(doc, "3.1 รูปแบบการวิจัยและการออกแบบการทดลอง", p_map['ch3'], indent=0.25)
    add_toc_line(doc, "3.2 กลุ่มตัวอย่างและลักษณะทางประชากรศาสตร์ของผู้ศึกษา", p_map['ch3_2'], indent=0.25)
    add_toc_line(doc, "3.3 แผนปฏิบัติการแทรกแซงและเก็บรวบรวมข้อมูล 30 วัน", p_map['ch3_3'], indent=0.25)
    add_toc_line(doc, "3.4 เครื่องมือวิจัยและการบันทึกข้อมูลเชิงประจักษ์ด้วย Microsoft Excel และระบบจับเวลา", p_map['ch3_4'], indent=0.25)
    add_toc_line(doc, "3.5 เกณฑ์และการวัดตัวชี้วัดความสำเร็จ 5 ด้าน", p_map['ch3_5'], indent=0.25)
    add_toc_line(doc, "3.6 มาตรการด้านความปลอดภัยและการเฝ้าระวังผลข้างเคียง", p_map['ch3_6'], indent=0.25)
    
    add_toc_line(doc, "บทที่ 4 ผลการดำเนินงานและการวิเคราะห์ข้อมูล (ข้อมูลถึงวันที่ 14)", p_map['ch4'], indent=0, is_bold=True)
    add_toc_line(doc, "4.1 ข้อมูลฐานเปรียบเทียบก่อนการดำเนินโครงการ (Baseline Assessment Day 1)", p_map['ch4'], indent=0.25)
    add_toc_line(doc, "4.2 ผลการเปลี่ยนแปลงของระดับความปวดบริเวณคอบ่าไหล่ (NRS D1–D14)", p_map['ch4_2'], indent=0.25)
    add_toc_line(doc, "4.3 ผลการปรับปรุงท่าทางและมุมข้อต่อสรีรศาสตร์ 90-90-90", p_map['ch4_3'], indent=0.25)
    add_toc_line(doc, "4.4 ผลการปฏิบัติตามวินัยการพักขยับตัวทุก 45 นาที", p_map['ch4_4'], indent=0.25)
    add_toc_line(doc, "4.5 ผลการประเมินความสม่ำเสมอในการปฏิบัติตามโปรแกรมกายบริหารและการนวด", p_map['ch4_5'], indent=0.25)
    add_toc_line(doc, "4.6 ผลสัมฤทธิ์ภาพรวมเปรียบเทียบกับตัวชี้วัดความสำเร็จ 5 ด้าน (Midterm Evaluation)", p_map['ch4_6'], indent=0.25)
    
    add_toc_line(doc, "บทที่ 5 สรุปผล อภิปรายผล และข้อเสนอแนะ", p_map['ch5'], indent=0, is_bold=True)
    add_toc_line(doc, "5.1 สรุปผลการดำเนินโครงการระยะครึ่งทาง", p_map['ch5'], indent=0.25)
    add_toc_line(doc, "5.2 การอภิปรายผลการวิจัย", p_map['ch5_2'], indent=0.25)
    add_toc_line(doc, "5.3 ปัญหา อุปสรรค และแนวทางแก้ไขที่ได้เรียนรู้", p_map['ch5_3'], indent=0.25)
    add_toc_line(doc, "5.4 ข้อเสนอแนะสำหรับการดำเนินงานระยะที่ 2 และการนำไปใช้ประโยชน์", p_map['ch5_4'], indent=0.25)
    
    add_toc_line(doc, "บรรณานุกรม", p_map['ref'], indent=0, is_bold=True)
    add_toc_line(doc, "ภาคผนวก", p_map['app_a'], indent=0, is_bold=True)
    add_toc_line(doc, "ภาคผนวก ก: ตารางบันทึกข้อมูลประจำวันฉบับเต็ม Days 1–14 (Daily Raw Data)", p_map['app_a'], indent=0.25)
    add_toc_line(doc, "ภาคผนวก ข: ข้อมูลการวัดมุมข้อต่อสรีรศาสตร์รายวันฉบับเต็ม Days 1–14", p_map['app_b'], indent=0.25)
    add_toc_line(doc, "ภาคผนวก ค: แบบบันทึกความปลอดภัยและอาการผิดปกติ (Safety Log)", p_map['app_c'], indent=0.25)
    add_toc_line(doc, "ภาคผนวก ง: โครงสร้างไฟล์ Microsoft Excel และสูตรคำนวณที่ใช้ในการวิจัย", p_map['app_d'], indent=0.25)
    add_toc_line(doc, "ประวัติผู้จัดทำ", p_map['bio'], indent=0, is_bold=True)

    # ==========================================
    # SECTION 5.1: LIST OF TABLES (1 page budget)
    # ==========================================
    sec_lot = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section_margins(sec_lot, top=1.5, bottom=1.0, left=1.5, right=1.0)
    clear_pg_num_types(sec_lot)
    add_header_page_field(sec_lot.header)
    
    add_front_matter_title(doc, "สารบัญตาราง")
    add_toc_col_header(doc, "ตารางที่", "หน้า")
    
    add_toc_line(doc, "1.1  กรอบตัวชี้วัดความสำเร็จ 5 ด้านและเป้าหมายการประเมินผลโครงการวิจัย", p_map['tbl_1_1'])
    add_toc_line(doc, "2.1  การสังเคราะห์เปรียบเทียบแนวทางการแก้ปัญหาออฟฟิศซินโดรมและจุดบูรณาการ", p_map['tbl_2_1'])
    add_toc_line(doc, "2.2  รายละเอียดขั้นตอนและปริมาณโปรแกรมกายบริหารและยืดเหยียด 6 ท่า", p_map['tbl_2_2'])
    add_toc_line(doc, "3.1  แผนปฏิบัติการแทรกแซง 30 วันและการเก็บรวบรวมหลักฐานเชิงประจักษ์", p_map['tbl_3_1'])
    add_toc_line(doc, "3.2  เกณฑ์การประเมินมุมข้อต่อสรีรศาสตร์ 3 จุดตามหลักการ 90-90-90", p_map['tbl_3_2'])
    add_toc_line(doc, "4.1  ข้อมูลฐานเปรียบเทียบก่อนการดำเนินโครงการในวันแรก (Baseline Day 1)", p_map['tbl_4_1'])
    add_toc_line(doc, "4.2  สรุปผลการประเมินระดับความปวดคอบ่าไหล่ (NRS 0–10) รายระยะ วันที่ 1 ถึง 14", p_map['tbl_4_2'])
    add_toc_line(doc, "4.3  สรุปข้อมูลการวัดมุมข้อต่อสรีรศาสตร์เปรียบเทียบตามจุดประเมินสำคัญ (D1–D14)", p_map['tbl_4_3'])
    add_toc_line(doc, "4.4  สรุปเวลาใช้งานหน้าจอและอัตราการปฏิบัติตามวินัยการพักขยับตัวทุก 45 นาที (D1–D14)", p_map['tbl_4_4'])
    add_toc_line(doc, "4.5  สรุปผลสัมฤทธิ์เปรียบเทียบกับตัวชี้วัดความสำเร็จ 5 ด้าน (Midterm Evaluation)", p_map['tbl_4_5'])

    # ==========================================
    # SECTION 5.2: LIST OF FIGURES (1 page budget)
    # ==========================================
    sec_lof = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section_margins(sec_lof, top=1.5, bottom=1.0, left=1.5, right=1.0)
    clear_pg_num_types(sec_lof)
    add_header_page_field(sec_lof.header)
    
    add_front_matter_title(doc, "สารบัญภาพ")
    add_toc_col_header(doc, "รูปที่", "หน้า")
    
    add_toc_line(doc, "1.1  สรีระท่านั่งเริ่มต้นก่อนดำเนินโครงการวิจัย (Baseline Posture Day 1)", p_map['fig_1_1'])
    add_toc_line(doc, "2.1  แผนภาพกรอบแนวคิดการวิจัยเพื่อลดอาการออฟฟิศซินโดรม", p_map['fig_2_1'])
    add_toc_line(doc, "3.1  แผนผังกระบวนการเก็บรวบรวมหลักฐานเชิงประจักษ์และขั้นตอนการดำเนินงานประจำวัน", p_map['fig_3_1'])
    add_toc_line(doc, "3.2  โครงสร้างสมุดบันทึกข้อมูลสุขภาพประจำวัน (Daily Health Logbook) บน Microsoft Excel", p_map['fig_3_2'])
    add_toc_line(doc, "3.3  การวิเคราะห์และวัดมุมข้อต่อสรีรศาสตร์ 3 จุด จากภาพถ่ายด้านข้างด้วย ImageMeter", p_map['fig_3_3'])
    add_toc_line(doc, "3.4  สภาพแวดล้อมโต๊ะทำงานจริงของผู้วิจัยภายหลังการปรับระดับสรีรศาสตร์ 90-90-90", p_map['fig_3_4'])
    add_toc_line(doc, "4.1  กราฟแสดงการเปลี่ยนแปลงระดับความปวดบริเวณคอ บ่า ไหล่ (NRS Score) วันที่ 1 ถึง 14", p_map['fig_4_1'])
    add_toc_line(doc, "4.2  กราฟแสดงพัฒนาการของมุมข้อต่อสรีรศาสตร์เทียบกับแถบเกณฑ์มาตรฐาน 90-90-90", p_map['fig_4_2'])
    add_toc_line(doc, "4.3  กราฟแสดงเวลาใช้งานหน้าจอและจำนวนรอบการพักขยับตัวรายวันเปรียบเทียบกับเกณฑ์", p_map['fig_4_3'])

    # ==========================================
    # SECTION 6: BODY (CHAPTERS 1 TO 5 - PAGE NUMBERING RESTARTS AT 1)
    # ==========================================
    sec_body = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section_margins(sec_body, top=1.5, bottom=1.0, left=1.5, right=1.0)
    clear_pg_num_types(sec_body)
    pgNum_arabic = parse_xml(f'<w:pgNumType {nsdecls("w")} w:fmt="decimal" w:start="1"/>')
    sec_body._sectPr.append(pgNum_arabic)
    add_header_page_field(sec_body.header)
    
    # ------------------------------------------
    # CHAPTER 1: INTRODUCTION
    # ------------------------------------------
    add_chapter_title(doc, "1", "บทนำ")
    add_heading_1(doc, "1.1 ความเป็นมาและความสำคัญของปัญหา")
    add_body_p(doc, (
        "ในยุคเศรษฐกิจและสังคมดิจิทัล เทคโนโลยีสารสนเทศได้เข้ามามีบทบาทอย่างลึกซึ้งในการขับเคลื่อนกระบวนการเรียนรู้ "
        "การทำงาน และวิถีชีวิตประจำวันของประชากร โดยเฉพาะในหมู่นักศึกษาและบุคลากรทางด้านเทคโนโลยีสารสนเทศ (Information Technology: IT) "
        "ซึ่งมีความจำเป็นต้องใช้เวลาในการนั่งทำงานหน้าจอคอมพิวเตอร์อย่างต่อเนื่องยาวนาน เพื่อการพัฒนาซอฟต์แวร์ การเขียนชุดคำสั่ง (Coding) "
        "การค้นคว้าข้อมูล และการจัดการระบบเครือข่าย จากสถิติการใช้งานเทคโนโลยีในกลุ่มนักศึกษาสายคอมพิวเตอร์พบว่า มีระยะเวลาการทำงานหน้าจอคอมพิวเตอร์เฉลี่ยสูงถึง 6 ถึง 10 ชั่วโมงต่อวัน "
        "ซึ่งพฤติกรรมการนั่งทำงานต่อเนื่องในระยะเวลานานร่วมกับสภาพแวดล้อมสถานีงานที่ไม่ถูกสุขลักษณะตามหลักการยศาสตร์ (Ergonomics) "
        "ได้กลายเป็นปัจจัยเสี่ยงสำคัญที่นำไปสู่ภาวะความผิดปกติของระบบกล้ามเนื้อและกระดูกโครงร่างจากการทำงาน (Work-related Musculoskeletal Disorders: WMSDs) "
        "หรือที่รู้จักกันทั่วไปในชื่อ 'กลุ่มอาการออฟฟิศซินโดรม' (Office Syndrome)"
    ))
    
    add_body_p(doc, (
        "สำหรับผู้วิจัย ซึ่งเป็นนักศึกษาชั้นปีที่ 1 สาขาเทคโนโลยีสารสนเทศ คณะเทคโนโลยีและการจัดการอุตสาหกรรม มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ "
        "พบว่าในการดำเนินชีวิตและการเรียนประจำวัน มีการใช้งานคอมพิวเตอร์เฉลี่ยสูงถึง 5.8 ถึง 8.0 ชั่วโมงต่อวัน จากการประเมินสภาพปัญหาตนเองเชิงประจักษ์ในวันแรกของการสำรวจ (Baseline Day 1) "
        "พบว่าผู้วิจัยมีอาการปวดตึงเรื้อรังบริเวณกล้ามเนื้อคอด้านหลัง กล้ามเนื้อบ่า (Upper Trapezius) และกล้ามเนื้อรอบสะบัก (Levator Scapulae) "
        "โดยมีระดับความรุนแรงของความปวดเมื่อยประเมินผ่านแบบประเมินความปวดเชิงตัวเลข (Numeric Rating Scale: NRS ระดับ 0 ถึง 10) อยู่ที่ระดับ 6.0 คะแนน "
        "ซึ่งจัดอยู่ในเกณฑ์ความปวดระดับปานกลางค่อนข้างสูง (Moderate to Severe Pain) ส่งผลรบกวนต่อสมาธิในการศึกษาเล่าเรียน คุณภาพการนอนหลับ และประสิทธิภาพในการดำเนินกิจกรรมประจำวันอย่างเด่นชัด"
    ))
    
    add_body_p(doc, (
        "เมื่อได้ทำการบันทึกภาพถ่ายสรีระท่านั่งทำงานจากมุมมองด้านข้าง (Lateral View) และนำมาวิเคราะห์มุมข้อต่อตามหลักชีวกลศาสตร์ "
        "พบความบกพร่องทางสรีรศาสตร์ที่สอดคล้องกับพยาธิสภาพของอาการปวดอย่างชัดเจน ดังแสดงในรูปที่ 1.1:"
    ))
    
    add_figure(
        doc,
        os.path.join(ASSETS_DIR, 'baseline_posture.png'),
        "1.1",
        "สรีระท่านั่งเริ่มต้นก่อนดำเนินโครงการวิจัย แสดงความผิดปกติด้านมุมข้อต่อสรีรศาสตร์ (Baseline Day 1)",
        width_inches=3.6,
        height_inches=2.4,
        source_text="บันทึกภาพและประมวลผลโดยผู้วิจัย, 10 กันยายน 2569"
    )
    
    add_body_p(doc, (
        "จากรูปที่ 1.1 พบความผิดปกติ 3 ประการหลัก ได้แก่: 1) แป้นพิมพ์และเมาส์วางอยู่ลึกบนโต๊ะ ทำให้ต้องเอื้อมแขนไปข้างหน้า ส่งผลให้ข้อศอกกางออกถึง 138° "
        "(เกณฑ์มาตรฐานสากล 90°–100°) เกิดแรงบิดและแรงเกร็งค้างที่กล้ามเนื้อหัวไหล่และสะบักตลอดเวลา 2) พฤติกรรมการนั่งกึ่งนอนโดยเลื่อนสะโพกมาด้านหน้า "
        "ทำให้มุมข้อต่อสะโพกเปิดกว้างถึง 125° แผ่นหลังส่วนล่างไม่แนบชิดพนักพิง ขาดฐานรองรับน้ำหนักกระดูกสันหลังส่วนเอว (Lumbar Support) และ 3) การยกขาขึ้นมาพับบนเบาะเก้าอี้ "
        "ทำให้มุมข้อเข่าหดแคบเหลือเพียง 65° ฝ่าเท้าไม่วางราบบนพื้น ขัดขวางการไหลเวียนโลหิตดำส่วนล่างและลดความมั่นคงของกระดูกเชิงกราน "
        "ปัญหาดังกล่าวแสดงให้เห็นว่าต้นเหตุของอาการปวดตึงมิได้เกิดจากความเสื่อมของโครงสร้างทางกายวิภาค แต่เกิดจาก 'พฤติกรรมความเคยชินและความไม่เหมาะสมของสถานีงาน' "
        "ซึ่งสามารถแก้ไขได้อย่างตรงจุดด้วยการปรับพฤติกรรมและการจัดสภาพแวดล้อมทางกายภาพ"
    ))
    
    add_heading_1(doc, "1.2 วัตถุประสงค์ของโครงงาน", page_break_before=True)
    add_body_p(doc, "เพื่อให้การวิจัยและการปรับเปลี่ยนพฤติกรรมมีความชัดเจน รัดกุม และสามารถวัดผลได้ตามหลักการวิจัยเชิงประจักษ์ โครงงานจึงกำหนดวัตถุประสงค์ไว้ 3 ประการ ดังนี้:", indent=False, space_after=3)
    add_numbered_item(doc, "1.", "เพื่อศึกษาผลของการปรับสภาพแวดล้อมโต๊ะทำงานตามหลักการยศาสตร์ (Ergonomics 90-90-90) ร่วมกับการสร้างวินัยการลุกพักขยับตัวทุก 45 นาที ต่อการลดระดับความปวดกล้ามเนื้อบริเวณคอ บ่า และไหล่ ในระยะเวลา 30 วัน")
    add_numbered_item(doc, "2.", "เพื่อติดตามและประเมินพัฒนาการฟื้นฟูของมุมข้อต่อสรีรศาสตร์ 3 จุดสำคัญ (ข้อศอก สะโพก ข้อเข่า) จากภาพถ่ายด้านข้างให้เข้าสู่เกณฑ์มาตรฐานไม่น้อยกว่าร้อยละ 75.0")
    add_numbered_item(doc, "3.", "เพื่อพัฒนาระบบและระเบียบวิธีบันทึกหลักฐานเชิงประจักษ์ดิจิทัลที่มีความเที่ยงตรง ผ่าน Microsoft Excel บนระบบคลาวด์ OneDrive ร่วมกับระบบจับเวลาแจ้งเตือนบนสมาร์ตโฟน และแอปพลิเคชัน ImageMeter")

    add_heading_1(doc, "1.3 ขอบเขตของโครงงาน")
    add_body_p(doc, "โครงงานวิจัยนี้ได้กำหนดขอบเขตการศึกษาครอบคลุม 4 มิติสำคัญ ดังนี้:", indent=False, space_after=3)
    add_bullet_item(doc, "ขอบเขตด้านประชากรและกลุ่มเป้าหมาย: ผู้วิจัยเพียงคนเดียว (Single-Subject N-of-1 Research Design) ซึ่งเป็นนักศึกษาชาย ชั้นปีที่ 1 สาขาเทคโนโลยีสารสนเทศ ที่มีอาการปวดกล้ามเนื้อคอบ่าไหล่เรื้อรังจากการใช้คอมพิวเตอร์")
    add_bullet_item(doc, "ขอบเขตด้านระยะเวลา: ระยะเวลาดำเนินการศึกษาต่อเนื่อง 30 วัน (ตั้งแต่วันที่ 10 กันยายน ถึง 9 ตุลาคม 2569) โดยรายงานฉบับนี้สรุปผลเชิงประจักษ์ระยะครึ่งทาง (Days 1 ถึง 14)")
    add_bullet_item(doc, "ขอบเขตด้านมาตรการแทรกแซง (Interventions): ประกอบด้วย 4 มาตรการหลัก ได้แก่ 1) การปรับสถานีงานตามหลัก 90-90-90, 2) การตั้งเตือนพักขยับตัวทุก 45 นาที (Active Microbreaks), 3) การนวดกดจุดสะท้อน 6 นาที, และ 4) กายบริหารเสริมความแข็งแรง 6 ท่า")
    add_bullet_item(doc, "ขอบเขตด้านเครื่องมือและเทคโนโลยี: ซอฟต์แวร์ Microsoft Excel (OneDrive Cloud Sync) สำหรับบันทึก Logbook รายวัน, ระบบตัวจับเวลาบนสมาร์ตโฟนสำหรับควบคุมรอบพักและเวลานวด, และแอปพลิเคชัน ImageMeter สำหรับการวัดมุมข้อต่อในภาพถ่าย")

    add_heading_1(doc, "1.4 สมมติฐานและตัวชี้วัดความสำเร็จ (Measurable Key Results)", page_break_before=True)
    add_body_p(doc, (
        "สมมติฐานการวิจัย (Research Hypothesis): เมื่อผู้วิจัยได้รับการปรับปรุงสภาพแวดล้อมโต๊ะทำงานตามหลักการยศาสตร์ 90-90-90 "
        "ควบคู่กับการลุกพักขยับตัวทุก 45 นาที และการบำบัดฟื้นฟูกล้ามเนื้อด้วยตนเองทุกวัน จะส่งผลให้ระดับความปวดเมื่อยคอบ่าไหล่ลดลงอย่างน้อยร้อยละ 30 "
        "และสรีระท่าทางการนั่งทำงานมีมุมข้อต่อผ่านเกณฑ์มาตรฐานไม่น้อยกว่าร้อยละ 75.0 ภายในระยะเวลาดำเนินการ"
    ), space_after=4)
    
    add_body_p(doc, "เพื่อพิสูจน์สมมติฐานดังกล่าว โครงงานได้กำหนดกรอบตัวชี้วัดความสำเร็จ 5 ด้าน (Key Results: KR 1 ถึง KR 5) ดังแสดงในตารางที่ 1.1:", indent=False, space_after=3)
    
    headers_1_1 = ["ตัวชี้วัดความสำเร็จ", "สูตรการคำนวณเชิงปริมาณ", "ค่าตั้งต้น (Baseline D1)", "เกณฑ์เป้าหมายโครงการ", "ประเภทตัวชี้วัด"]
    data_1_1 = [
        ["KR 1: อัตราการลดลงของระดับความปวดคอบ่าไหล่", "%ลดปวด = [(NRS_D1 - เฉลี่ย_D28-30) ÷ NRS_D1] × 100", "6.0 คะแนน (NRS)", "ลดลง ≥ 30.0% (NRS ≤ 4.2)", "ผลลัพธ์หลัก (Primary Outcome)"],
        ["KR 2: อัตราความถูกต้องของมุมสรีระ 90-90-90", "%ท่าทาง = (จำนวนมุมที่ผ่านเกณฑ์ ÷ 4) × 100", "25.0% (1 ใน 4 จุด)", "ผ่านเกณฑ์ ≥ 75.0% (≥ 3 ใน 4)", "ผลลัพธ์หลัก (Primary Outcome)"],
        ["KR 3: วินัยการลุกพักขยับตัวทุก 45 นาที", "%การพัก = (รอบพักจริง ÷ รอบพักที่ควรทำ) × 100", "0.0% (ไม่เคยพัก)", "อัตราการพักเฉลี่ย ≥ 80.0%", "พฤติกรรม (Behavioral Compliance)"],
        ["KR 4: ความสม่ำเสมอในการทำโปรแกรมครบถ้วน", "%ทำครบ = (จำนวนวันที่ทำครบทุกท่า ÷ 30) × 100", "0.0% (ไม่เคยปฏิบัติ)", "ความสม่ำเสมอ ≥ 85.0% (≥ 26 วัน)", "กระบวนการ (Process Adherence)"],
        ["KR 5: ประสิทธิผลทันทีจากการนวดกดจุด 6 นาที", "%ลดทันที = [(NRSก่อน - NRSหลัง) ÷ NRSก่อน] × 100", "0.0% (ไม่มีการนวด)", "ลดลงเฉลี่ย ≥ 20.0% ทันที", "ผลเชิงสำรวจ (Exploratory Result)"]
    ]
    col_w_1_1 = [Inches(1.1), Inches(1.75), Inches(0.95), Inches(1.05), Inches(0.9)]
    add_styled_table(doc, "1.1", "กรอบตัวชี้วัดความสำเร็จ 5 ด้านและเป้าหมายการประเมินผลโครงการวิจัย", headers_1_1, data_1_1, col_w_1_1, font_size_pt=10.0, padding_twips=15)

    add_heading_1(doc, "1.5 ประโยชน์ที่คาดว่าจะได้รับ", page_break_before=True)
    add_numbered_item(doc, "1.", "ประโยชน์ต่อสุขภาพตนเอง: ลดอาการปวดเมื่อยกล้ามเนื้อคอ บ่า ไหล่ และหลัง ป้องกันความเสี่ยงของโรคหมอนรองกระดูกคอเสื่อม เสริมสร้างบุคลิกภาพที่ดี")
    add_numbered_item(doc, "2.", "ประโยชน์เชิงระเบียบวิธีวิจัย: ได้แนวทางการศึกษาแบบ Single-Subject N-of-1 Research Design ที่ประยุกต์ใช้เครื่องมือดิจิทัลทั่วไป เช่น Microsoft Excel และแอปวัดมุมในสมาร์ตโฟน มาทำการวิจัยเชิงประจักษ์ได้อย่างเป็นรูปธรรม")
    add_numbered_item(doc, "3.", "ประโยชน์ต่อเพื่อนนักศึกษาและบุคคลทั่วไป: เป็นกรณีศึกษาและคู่มือปฏิบัติการที่พิสูจน์แล้วว่า การปรับพฤติกรรมและการจัดโต๊ะทำงานด้วยตนเอง สามารถบรรเทาออฟฟิศซินโดรมได้จริงโดยไม่ต้องพึ่งพายาหรืออุปกรณ์ราคาแพง")

    add_heading_1(doc, "1.6 นิยามศัพท์เฉพาะ")
    add_numbered_item(doc, "1.", "ออฟฟิศซินโดรม (Office Syndrome): กลุ่มอาการปวดกล้ามเนื้อและเยื่อพังผืด (Myofascial Pain Syndrome) บริเวณคอ บ่า สะบัก และหลังส่วนบน อันเนื่องมาจากการเกร็งค้างในท่าทางเดิมต่อเนื่อง")
    add_numbered_item(doc, "2.", "หลักการยศาสตร์ 90-90-90 (90-90-90 Ergonomics Rule): เกณฑ์การจัดสรีระท่านั่งทำงาน ประกอบด้วย มุมข้อศอก 90°–100°, มุมข้อสะโพก 90°–100°, มุมข้อเข่า 85°–100° และสายตาอยู่ในระดับกึ่งกลางจอ")
    add_numbered_item(doc, "3.", "การพักขยับตัว (Active Microbreaks): การหยุดพักจากการนั่งทำงานหน้าจอคอมพิวเตอร์ทุก 45 นาที เพื่อลุกขึ้นยืน เดิน และยืดเหยียดร่างกายเป็นเวลา 2 นาที")
    add_numbered_item(doc, "4.", "การวิจัยรายบุคคล (N-of-1 Study): ระเบียบวิธีวิจัยทางคลินิกที่ทำการทดลองในกลุ่มตัวอย่างรายเดียว โดยเปรียบเทียบข้อมูลผลลัพธ์ระหว่างช่วงก่อนและหลังการแทรกแซง")

    # ------------------------------------------
    # CHAPTER 2: LITERATURE REVIEW
    # ------------------------------------------
    add_chapter_title(doc, "2", "วรรณกรรม ทฤษฎี และงานวิจัยที่เกี่ยวข้อง", space_before=14, space_after=14)
    add_heading_1(doc, "2.1 กายวิภาคศาสตร์และชีวกลศาสตร์ของกลุ่มอาการออฟฟิศซินโดรม")
    add_body_p(doc, (
        "กลุ่มอาการออฟฟิศซินโดรม (Office Syndrome) ทางการแพทย์มักจัดอยู่ในกลุ่มอาการปวดกล้ามเนื้อและเยื่อพังผืด (Myofascial Pain Syndrome: MPS) "
        "ซึ่งมีพยาธิสภาพหลักอยู่ที่การหดเกร็งค้างของเส้นใยกล้ามเนื้ออย่างต่อเนื่อง (Sustained Static Contraction) โดยเฉพาะกลุ่มกล้ามเนื้อบ่าด้านบน (Upper Trapezius) "
        "กล้ามเนื้อยกสะบัก (Levator Scapulae) และกลุ่มกล้ามเนื้อรอบกระดูกคอ (Cervical Paraspinal Muscles) เมื่อบุคคลนั่งทำงานในท่าที่คอยื่นไปด้านหน้า (Forward Head Posture) "
        "และไหล่ห่อ (Rounded Shoulders) น้ำหนักของกะโหลกศีรษะซึ่งเฉลี่ยอยู่ที่ 4.5 ถึง 5.5 กิโลกรัมในท่าปกติ จะสร้างแรงกดดันเชิงกล (Mechanical Load) "
        "ต่อกระดูกสันหลังส่วนคอเพิ่มขึ้นตามมุมการก้ม โดยการก้มศีรษะเพียง 15 องศา จะสร้างแรงกดเทียบเท่าน้ำหนักถึง 12 กิโลกรัม และหากก้มถึง 45 องศา แรงกดจะพุ่งสูงถึง 22 กิโลกรัม "
        "(Hansraj, 2014) ภาระทางกลที่สูงขึ้นนี้บีบคั้นให้กล้ามเนื้อ Upper Trapezius ต้องออกแรงต้านเกร็งค้างตลอดเวลา ส่งผลให้หลอดเลือดฝอยที่มาหล่อเลี้ยงกล้ามเนื้อถูกกดทับ "
        "เกิดภาวะขาดเลือดและออกซิเจนเฉพาะที่ (Local Ischemia) เกิดการสะสมของกรดแลกติกและสารก่อการอักเสบ จนพัฒนาไปสู่จุดกดเจ็บ (Myofascial Trigger Points: MTrPs) "
        "ที่ส่งกระแสความปวดร้าวขึ้นสู่ศีรษะและลงสู่สะบัก"
    ))
    
    add_heading_1(doc, "2.2 ทฤษฎีการยศาสตร์และการจัดสถานีงานคอมพิวเตอร์ (Ergonomics 90-90-90)", page_break_before=True)
    add_body_p(doc, (
        "การยศาสตร์ (Ergonomics) คือศาสตร์ที่ว่าด้วยการปรับสภาพแวดล้อมการทำงานและอุปกรณ์ให้สอดคล้องกับขีดจำกัดทางสรีรวิทยาและชีวกลศาสตร์ของมนุษย์ "
        "สำหรับผู้ใช้งานคอมพิวเตอร์ สถาบันอาชีวอนามัยและความปลอดภัยแห่งชาติ สหรัฐอเมริกา (NIOSH) และสมาคมการยศาสตร์สากล (IEA) "
        "ได้กำหนดเกณฑ์มาตรฐานการจัดท่านั่งในรูปแบบ 'กฎ 90-90-90' (The 90-90-90 Ergonomic Sitting Posture) ซึ่งมีหลักการสำคัญ 4 ประการ:"
    ), space_after=3)
    add_numbered_item(doc, "1.", "มุมข้อศอก (Elbow Angle 90°–100°): ท่อนแขนส่วนบนทิ้งตัวลงตามธรรมชาติ ข้อศอกงอตั้งฉาก แป้นพิมพ์และเมาส์ต้องวางอยู่ในระดับความสูงเดียวกับข้อศอก โดยไม่ต้องเอื้อมแขนหรือยักไหล่")
    add_numbered_item(doc, "2.", "มุมข้อสะโพก (Hip/Trunk Angle 90°–100°): นั่งให้ก้นและสะโพกชิดด้านในสุดของเบาะเก้าอี้ แผ่นหลังส่วนล่างแนบชิดพนักพิงที่มีส่วนโค้งนูนรองรับแนวกระดูกสันหลังส่วนเอว (Lumbar Support)")
    add_numbered_item(doc, "3.", "มุมข้อเข่า (Knee Angle 85°–100°): ปรับระดับความสูงของเบาะนั่งให้ข้อเข่างอทำมุมฉาก โดยให้ฝ่าเท้าทั้งสองข้างวางราบบนพื้นอย่างมั่นคง เพื่อกระจายน้ำหนักตัวอย่างสมดุล")
    add_numbered_item(doc, "4.", "ระดับสายตา (Eye Level): ขอบบนสุดของหน้าจอแสดงผลควรอยู่ในระดับสายตาหรือต่ำกว่าสายตาเล็กน้อย (0° ถึง -15°) ในระยะห่างประมาณ 50 ถึง 70 เซนติเมตร เพื่อรักษากระดูกคอให้อยู่ในแนวตรง")

    add_heading_1(doc, "2.3 สรีรวิทยาของการพักขยับตัวและผลกระทบของพฤติกรรมเนือยนิ่ง")
    add_body_p(doc, (
        "พฤติกรรมเนือยนิ่ง (Sedentary Behavior) จากการนั่งทำงานติดต่อกันเกินกว่า 45 ถึง 60 นาที ส่งผลให้การเผาผลาญระดับเซลล์ชะลอตัว "
        "และการไหลเวียนโลหิตดำส่วนล่างติดขัด การวิจัยทางการยศาสตร์พบว่า การจัดท่าทางที่ถูกต้องตามหลักการยศาสตร์เพียงอย่างเดียวไม่เพียงพอที่จะป้องกันความเมื่อยล้าได้ "
        "เนื่องจากกล้ามเนื้อยังคงต้องทำงานในลักษณะเกร็งค้างคงที่ (Static Load) ดังนั้น การแทรกแซงด้วย 'การพักขยับตัวระยะสั้น' (Active Microbreaks) "
        "เป็นเวลา 1 ถึง 2 นาที ในทุกๆ 45 นาที จึงเป็นกลยุทธ์ทางสรีรวิทยาที่มีความจำเป็นอย่างยิ่ง โดยการลุกขึ้นยืนและเดินจะช่วยกระตุ้นกลไกการปั๊มเลือดของกล้ามเนื้อน่อง (Skeletal Muscle Pump) "
        "ฟื้นฟูออกซิเจนสู่เนื้อเยื่อ และคลายการตึงเครียดของเอ็นข้อต่อได้อย่างมีประสิทธิภาพ (Hedge & Ray, 2004)"
    ))

    add_heading_1(doc, "2.4 กลไกการบำบัดด้วยการนวดกดจุดสะท้อนและการยืดเหยียดกล้ามเนื้อ", page_break_before=True)
    add_body_p(doc, (
        "การนวดกดจุดสะท้อนด้วยตนเอง (Self-Acupressure & Trigger Point Release) เป็นการประยุกต์ใช้แรงกดเชิงกลจากปลายนิ้วหรือฝ่ามือกดลงบนจุดกดเจ็บบริเวณกล้ามเนื้อ Upper Trapezius "
        "และ Levator Scapulae เพื่อทำให้หลอดเลือดฝอยเกิดการขยายตัวแบบตอบสนองฉับพลัน (Reactive Hyperemia) หลังคลายแรงกด ส่งผลให้เลือดใหม่ที่อุดมด้วยออกซิเจนไหลเวียนเข้าชะล้างสารก่อความปวด "
        "ร่วมกับการส่งสัญญาณยับยั้งความปวดผ่านไขสันหลังตามทฤษฎีควบคุมประตูความปวด (Gate Control Theory of Pain) นอกจากนี้ การบริหารกล้ามเนื้อด้วยการยืดเหยียดและการเสริมสร้างความแข็งแรง 6 ท่า "
        "ดังแสดงในตารางที่ 2.2 จะช่วยปรับสมดุลความตึงตัวของกล้ามเนื้อรอบคอและสะบักให้กลับคืนสู่สภาวะปกติ:"
    ), space_after=3)
    
    headers_2_2 = ["ลำดับและชื่อท่าบริหาร", "กลุ่มกล้ามเนื้อเป้าหมาย", "วัตถุประสงค์เชิงสรีรวิทยา", "ขั้นตอนการปฏิบัติ", "ปริมาณ/เซต"]
    data_2_2 = [
        ["1. ดึงคางกลับ (Chin Tuck)", "Deep Cervical Flexors", "เสริมความแข็งแรงกล้ามเนื้อคอด้านลึก แก้คอยื่น", "นั่งหลังตรง เลื่อนศีรษะไปข้างหลังคล้ายทำคางสองชั้น", "ค้าง 5 วินาที, 10 ครั้ง/รอบ"],
        ["2. บีบสะบัก (Scapular Retraction)", "Rhomboids, Middle Trap", "เปิดหน้าอก ดึงสะบักกลับเข้าหาแนวกระดูกสันหลัง", "นั่งตรง ดึงสะบักทั้งสองข้างเข้าหากันและกดลงเล็กน้อย", "ค้าง 5 วินาที, 10 ครั้ง/รอบ"],
        ["3. สไลด์แขนชิดผนัง (Wall Slide)", "Serratus Anterior, Lower Trap", "เพิ่มความมั่นคงของสะบัก ฝึกการเคลื่อนไหวไหล่", "ยืนหันหน้าเข้าผนัง ท่อนแขนแตะผนัง เลื่อนแขนขึ้นช้าๆ", "10 ครั้ง/รอบ"],
        ["4. ยืดบ่าด้านข้าง (Side Neck Stretch)", "Upper Trapezius", "คลายกล้ามเนื้อบ่าด้านบนที่เกร็งค้างจากการยกไหล่", "มือข้างหนึ่งจับขอบเก้าอี้ เอียงศีรษะไปฝั่งตรงข้ามจนตึง", "ข้างละ 20 วินาที, 2 รอบ"],
        ["5. ยืดกล้ามเนื้อยกสะบัก (Levator Stretch)", "Levator Scapulae", "ลดอาการตึงร้าวจากต้นคอลงขอบสะบักด้านใน", "หันหน้า 45° ก้มศีรษะลงคล้ายมองรักแร้ฝั่งตรงข้าม", "ข้างละ 20 วินาที, 2 รอบ"],
        ["6. ยืดกล้ามเนื้อหน้าอก (Chest Stretch)", "Pectoralis Major/Minor", "ยืดกล้ามเนื้อหน้าอก แก้อาการไหล่ห่อและหลังค่อม", "วางท่อนแขนบนผนังระดับไหล่ บิดลำตัวออกจากผนัง", "ข้างละ 20 วินาที, 2 รอบ"]
    ]
    col_w_2_2 = [Inches(1.05), Inches(1.15), Inches(1.15), Inches(1.65), Inches(0.75)]
    add_styled_table(doc, "2.2", "รายละเอียดขั้นตอนและปริมาณโปรแกรมกายบริหารและยืดเหยียด 6 ท่า", headers_2_2, data_2_2, col_w_2_2, font_size_pt=9.5, padding_twips=14)

    add_heading_1(doc, "2.5 ระเบียบวิธีวิจัยรายบุคคลแบบวัดผลก่อนและหลัง (Single-Subject N-of-1 Design)", page_break_before=True)
    add_body_p(doc, (
        "การวิจัยรายบุคคล (Single-Subject Research Design หรือ N-of-1 Clinical Trial) เป็นระเบียบวิธีวิจัยเชิงประจักษ์ทางการแพทย์และพฤติกรรมศาสตร์ "
        "ที่ได้รับการยอมรับในระดับสากลว่ามีคุณค่าทางวิทยาศาสตร์สูงในการศึกษาผลการแทรกแซงเฉพาะบุคคล (Personalized Medicine) "
        "โดยผู้วิจัยทำหน้าที่เป็นทั้งกลุ่มทดลองและกลุ่มควบคุมของตนเอง (Self-as-Control) การเก็บข้อมูลตัวแปรตามอย่างละเอียดต่อเนื่องทุกวันตลอดระยะเวลา 30 วัน "
        "ช่วยตัดปัจจัยรบกวน (Confounding Factors) ด้านความแตกต่างระหว่างบุคคล และสะท้อนแนวโน้มการเปลี่ยนแปลงตามลำดับเวลาได้อย่างแท้จริง (Kazdin, 2011)"
    ))

    add_heading_1(doc, "2.6 เครื่องมือและเทคโนโลยีการบันทึกข้อมูลเชิงประจักษ์ (Excel & Photogrammetry)")
    add_body_p(doc, (
        "เพื่อให้ระเบียบวิธีวิจัยมีความน่าเชื่อถือและสอดคล้องกับสภาพความเป็นจริง โครงงานได้บูรณาการเครื่องมือเทคโนโลยีดิจิทัลที่เข้าถึงง่ายแต่มีความแม่นยำสูง 3 ส่วน ได้แก่: "
        "1) โปรแกรม Microsoft Excel บนระบบคลาวด์ Microsoft OneDrive ซึ่งทำหน้าที่เป็นสมุดบันทึกวิจัยดิจิทัล (Digital Research Logbook) "
        "ที่บันทึกข้อมูลดิบและประมวลผลสถิติได้แบบเรียลไทม์ 2) ระบบนาฬิกาจับเวลาและแจ้งเตือนบนสมาร์ตโฟน สำหรับสร้างสิ่งเร้าเตือนความจำ (Behavioral Prompts) ให้เกิดการพักทุก 45 นาที "
        "และควบคุมเวลานวดอย่างแม่นยำ และ 3) การวิเคราะห์ภาพถ่ายด้านข้างด้วยการวัดมุมแบบโฟโตแกรมเมทรี (Photogrammetric Angle Analysis) ผ่านแอปพลิเคชัน ImageMeter "
        "ซึ่งช่วยคำนวณมุมข้อต่อสรีระ 3 มิติจากภาพถ่ายสองมิติได้อย่างแม่นยำระดับองศา"
    ))

    add_heading_1(doc, "2.7 การสังเคราะห์วรรณกรรมและกรอบแนวคิดการวิจัย", page_break_before=True)
    add_body_p(doc, "จากการทบทวนวรรณกรรมข้างต้น โครงงานได้ทำการสังเคราะห์เปรียบเทียบแนวทางการแก้ปัญหาออฟฟิศซินโดรม ดังแสดงในตารางที่ 2.1:", indent=False, space_after=3)
    
    headers_2_1 = ["แนวทางการแก้ปัญหา", "จุดเด่นเชิงกลไก", "ข้อจำกัด / อุปสรรค", "การบูรณาการในโครงงานนี้"]
    data_2_1 = [
        ["1. การปรับโต๊ะตามหลักการยศาสตร์", "กำจัดแรงกดดันเชิงกลที่ต้นเหตุ คืนสรีระ 90-90-90", "ผู้ปฏิบัติมักเผลอกลับไปนั่งท่าเดิมเมื่อมีสมาธิทำงาน", "จัดสถานีงานใหม่ + ติดตามผลด้วยภาพถ่ายและวัดมุม ImageMeter"],
        ["2. การพักขยับตัวตามเวลา (Active Breaks)", "ป้องกันกล้ามเนื้อเกร็งค้าง ฟื้นฟูการไหลเวียนเลือด", "ผู้ปฏิบัติมักลืมพักเมื่อติดพันงานเขียนโค้ด", "ใช้ตัวจับเวลาสมาร์ตโฟนเตือนทุก 45 นาที บังคับลุกเดิน 2 นาที"],
        ["3. การนวดกดจุดสะท้อนด้วยมือ", "ลดความตึงเกร็งทันที กระตุ้นการขยายตัวของหลอดเลือด", "ผลการลดปวดมักเกิดขึ้นชั่วคราว หากไม่ปรับท่าทาง", "ทำหลังเลิกงานทุกวัน 6 นาที ควบคู่กับการประเมิน NRS ก่อน-หลัง"],
        ["4. การทำกายบริหาร 6 ท่า", "สร้างความแข็งแรงและความยืดหยุ่นของกล้ามเนื้อถาวร", "ต้องใช้ความสม่ำเสมอและวินัยในการทำต่อเนื่อง", "กำหนดปริมาณแน่นอน 6 ท่าทุกวัน บันทึกลงสมุด Excel รายวัน"]
    ]
    col_w_2_1 = [Inches(1.35), Inches(1.35), Inches(1.5), Inches(1.55)]
    add_styled_table(doc, "2.1", "การสังเคราะห์เปรียบเทียบแนวทางการแก้ปัญหาออฟฟิศซินโดรมและจุดบูรณาการ", headers_2_1, data_2_1, col_w_2_1, font_size_pt=10.0, padding_twips=15)
    
    add_body_p(doc, (
        "กรอบแนวคิดการวิจัย (Conceptual Framework): โครงงานได้กำหนดตัวแปรต้น (Independent Variables) คือ มาตรการแทรกแซงแบบผสมผสาน 4 ด้าน "
        "ได้แก่ 1) การปรับสถานีงาน 90-90-90, 2) การพักขยับตัวทุก 45 นาที, 3) การนวดกดจุดสะท้อน 6 นาที และ 4) กายบริหาร 6 ท่า "
        "โดยส่งผลต่อตัวแปรตาม (Dependent Variables) 3 ด้านสำคัญ ได้แก่ ระดับความปวดกล้ามเนื้อคอบ่าไหล่ (NRS), มุมข้อต่อสรีรศาสตร์ 3 จุด, "
        "และอัตราการปฏิบัติตามวินัยการพักขยับตัว ดังแสดงในรูปที่ 2.1:"
    ), space_after=3)
    
    add_figure(
        doc,
        os.path.join(ASSETS_DIR, 'diagram_conceptual_framework.png'),
        "2.1",
        "แผนภาพกรอบแนวคิดการวิจัยเพื่อลดอาการออฟฟิศซินโดรมและฟื้นฟูสรีระ (Conceptual Framework)",
        width_inches=4.8,
        height_inches=2.2,
        source_text="สังเคราะห์โดยผู้วิจัยจากวรรณกรรมทางการแพทย์และการยศาสตร์, 2569"
    )

    # ------------------------------------------
    # CHAPTER 3: METHODOLOGY
    # ------------------------------------------
    add_chapter_title(doc, "3", "วิธีดำเนินการวิจัยและการเก็บรวบรวมหลักฐานเชิงประจักษ์", space_before=14, space_after=14)
    add_heading_1(doc, "3.1 รูปแบบการวิจัยและการออกแบบการทดลอง")
    add_body_p(doc, (
        "โครงงานนี้ใช้รูปแบบการวิจัยเชิงทดลองรายบุคคลแบบวัดผลก่อนและหลังการทดลอง (Single-Subject Pre-Post Experimental Design / N-of-1 Study) "
        "ตลอดระยะเวลาทั้งสิ้น 30 วัน โดยแบ่งออกเป็น 2 ช่วงหลัก ได้แก่ ช่วงกำหนดค่าฐานเบื้องต้น (Baseline Phase: Day 1) "
        "เพื่อเก็บข้อมูลพฤติกรรมเดิมและอาการปวดก่อนการปรับเปลี่ยน และช่วงดำเนินมาตรการแทรกแซง (Intervention Phase: Days 2 ถึง 30) "
        "โดยมีจุดประเมินสำคัญ (Milestones) ณ วันที่ 7 (D7), วันที่ 14 (D14 - ระยะครึ่งทาง), วันที่ 21 (D21) และวันที่ 30 (D30 - สิ้นสุดโครงการ)"
    ))

    add_heading_1(doc, "3.2 กลุ่มตัวอย่างและลักษณะทางประชากรศาสตร์ของผู้ศึกษา")
    add_body_p(doc, (
        "กลุ่มตัวอย่างในการศึกษาคือ ผู้วิจัยเพียงคนเดียว (นายปภาวิน ธิติชุณหกุล รหัสนักศึกษา 6806021612037) เพศชาย อายุ 19 ปี "
        "นักศึกษาชั้นปีที่ 1 สาขาเทคโนโลยีสารสนเทศ คณะเทคโนโลยีและการจัดการอุตสาหกรรม มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ วิทยาเขตปราจีนบุรี "
        "น้ำหนักตัว 62 กิโลกรัม ส่วนสูง 172 เซนติเมตร ดัชนีมวลกาย (BMI) 20.96 กก./ม.² (เกณฑ์สมส่วนปกติ) มีประวัติการใช้งานคอมพิวเตอร์เฉลี่ย 5.8 ถึง 8.0 ชั่วโมงต่อวัน "
        "เกณฑ์การคัดเข้า (Inclusion Criteria) คือ มีอาการปวดตึงกล้ามเนื้อบริเวณคอ บ่า สะบัก หรือไหล่ สัมพันธ์กับการใช้คอมพิวเตอร์มาเป็นเวลาไม่น้อยกว่า 1 เดือน "
        "และเกณฑ์การคัดออก (Exclusion Criteria) คือ ไม่มีประวัติการผ่าตัดกระดูกคอหรือไหล่ ไม่มีอาการทางระบบประสาทรุนแรง (เช่น แขนชา อ่อนแรงชัดเจน) "
        "และไม่มีโรคประจำตัวทางระบบกระดูกสันหลังรุนแรง"
    ))

    # HEADING 3.3 STARTS ON A FRESH PAGE TO GUARANTEE TABLE 3.1 NEVER HANGS AT PAGE BOTTOM
    add_heading_1(doc, "3.3 แผนปฏิบัติการแทรกแซงและเก็บรวบรวมข้อมูล 30 วัน", page_break_before=True)
    add_body_p(doc, "โครงงานได้วางแผนการดำเนินงาน 30 วัน ออกเป็น 4 ระยะอย่างเป็นระบบ ดังแสดงในตารางที่ 3.1:", indent=False, space_after=3)

    headers_3_1 = ["ระยะเวลา", "ช่วงการดำเนินงาน", "กิจกรรมการแทรกแซงหลัก", "การเก็บรวบรวมหลักฐานเชิงประจักษ์"]
    data_3_1 = [
        ["D1", "ระยะที่ 1: กำหนดค่าฐานและเตรียมความพร้อม", "ใช้โต๊ะและท่านั่งเดิมตลอดวันจนจบการทำงาน บันทึกค่าปวด NRS เริ่มต้น แล้วปรับโต๊ะเป็น 90-90-90 ฝึกนวดและบริหาร", "บันทึกภาพถ่ายด้านข้าง Baseline, วัดมุมสรีระ 3 จุด, คะแนนปวด NRS"],
        ["D2 - D7", "ระยะที่ 2: ปรับสรีระและสร้างวินัยเริ่มต้น", "ปรับเก้าอี้และวางเท้าราบ ดึงเมาส์/คีย์บอร์ดเข้าใกล้ตัว ลุกพักทุก 45 นาที นวด 6 นาทีและบริหารครบ 6 ท่า", "บันทึกเวลาใช้จอ, นับรอบพักจริง, คะแนนปวดก่อน-หลังนวด, ภาพถ่าย D7"],
        ["D8 - D14", "ระยะที่ 3: เสริมสร้างความสม่ำเสมอและประเมินครึ่งทาง", "ปฏิบัติตามแผนอย่างต่อเนื่อง เน้นทำท่าบริหารให้ถูกต้องลึกซึ้ง ประเมินผลและปรับกลยุทธ์แก้ไขจุดบกพร่อง", "บันทึกรายวัน D8-D14, ภาพถ่าย D14, วิเคราะห์ผลเปรียบเทียบ KPI ครึ่งทาง"],
        ["D15 - D30", "ระยะที่ 4: การรักษาพฤติกรรมถาวรและสรุปผล", "รักษาวินัยการจัดท่าทางและการพักอย่างเป็นอัตโนมัติ เพิ่มความแข็งแรงกล้ามเนื้อ และสรุปผลสัมฤทธิ์ปลายงวด", "บันทึกรายวัน D15-D30, ภาพถ่าย D21, D30, วิเคราะห์ผลเปรียบเทียบ 30 วัน"]
    ]
    col_w_3_1 = [Inches(0.65), Inches(1.45), Inches(2.15), Inches(1.5)]
    add_styled_table(doc, "3.1", "แผนปฏิบัติการแทรกแซง 30 วันและการเก็บรวบรวมหลักฐานเชิงประจักษ์", headers_3_1, data_3_1, col_w_3_1, font_size_pt=10.0, padding_twips=15)

    # SECTION 3.4: RESEARCH TOOLS - MICROSOFT EXCEL & TIMERS PROTOCOL
    add_heading_1(doc, "3.4 เครื่องมือวิจัยและการบันทึกข้อมูลเชิงประจักษ์ด้วย Microsoft Excel และระบบจับเวลา", page_break_before=True)
    add_body_p(doc, (
        "เพื่อให้การเก็บรวบรวมหลักฐานเชิงประจักษ์ตลอด 30 วัน มีความเที่ยงตรง ตรวจสอบย้อนกลับได้ และสะท้อนพฤติกรรมจริงอย่างเป็นรูปธรรม "
        "โดยไม่สร้างภาระความซับซ้อนเกินจำเป็น ผู้วิจัยได้เลือกใช้เครื่องมือมาตรฐานดิจิทัลที่ใช้งานได้จริงในชีวิตประจำวัน ได้แก่ "
        "Microsoft Excel บนคลาวด์ OneDrive สำหรับการลงบันทึกข้อมูลและคำนวณสถิติ, ระบบนาฬิกาจับเวลาและแจ้งเตือนบนสมาร์ตโฟน, "
        "และการวิเคราะห์มุมสรีรศาสตร์จากภาพถ่ายด้านข้างด้วยแอปพลิเคชัน ImageMeter ดังมีรายละเอียดการดำเนินงานดังนี้:"
    ), space_after=2)

    add_heading_2(doc, "3.4.1 การออกแบบโครงสร้างสมุดบันทึกข้อมูลดิจิทัล Microsoft Excel (Daily Health Logbook)")
    add_body_p(doc, (
        "ผู้วิจัยได้สร้างสมุดบันทึกข้อมูลสุขภาพขึ้นในโปรแกรม Microsoft Excel เพียงไฟล์เดียวชื่อ '6806021612037_Health30D.xlsx' "
        "จัดเก็บบนระบบคลาวด์ Microsoft OneDrive เพื่อให้สามารถเข้าถึงและบันทึกข้อมูลได้ทันทีหลังเสร็จสิ้นกิจกรรมในแต่ละวัน "
        "ไฟล์ดังกล่าวได้รับการออกแบบโครงสร้างออกเป็น 3 เวิร์กชีตหลักอย่างเป็นระบบ ได้แก่:"
    ), space_after=2)
    add_numbered_item(doc, "1.", "ชีต 'Daily_Logbook' (บันทึกรายวัน): ทำหน้าที่บันทึกข้อมูลดิบรายวันตลอด 30 วัน ประกอบด้วย เวลาเริ่มต้นและสิ้นสุดการใช้จอคอมพิวเตอร์, "
                          "นาทีใช้งานหน้าจอรวม, จำนวนรอบการพักที่ควรทำตามเกณฑ์ (นาทีใช้จอ ÷ 45), จำนวนรอบการลุกพักขยับตัวที่ทำได้จริง, "
                          "ระดับความปวดคอบ่าไหล่ก่อนและหลังการนวด (NRS 0–10), การปฏิบัติตามโปรแกรมนวด 6 นาที และกายบริหารครบทั้ง 6 ท่า, และบันทึกหมายเหตุอาการ")
    add_numbered_item(doc, "2.", "ชีต 'Angle_Analysis' (ภาพถ่ายและมุมสรีระ): สำหรับแทรกภาพถ่ายท่านั่งด้านข้างและบันทึกค่ามุมข้อต่อ 3 จุด (ข้อศอก, สะโพก, ข้อเข่า) "
                          "และการประเมินระดับสายตา ที่วัดได้จากแอปพลิเคชัน ImageMeter ในวันสำคัญ (D1, D7, D14, D21, D30)")
    add_numbered_item(doc, "3.", "ชีต 'Summary_Metrics' (สรุปผลและกราฟ): บรรจุสูตรคำนวณอัตโนมัติของตัวชี้วัดความสำเร็จทั้ง 5 ด้าน เช่น สูตรคำนวณร้อยละการลดปวด, "
                          "ร้อยละการปฏิบัติตามวินัยการพัก, และร้อยละการผ่านเกณฑ์ท่าทาง พร้อมทั้งแสดงกราฟแท่งและกราฟเส้นสรุปแนวโน้มความก้าวหน้า ดังแสดงในรูปที่ 3.2")

    add_figure(
        doc,
        os.path.join(ASSETS_DIR, 'diagram_data_protocol.png'),
        "3.1",
        "แผนผังกระบวนการเก็บรวบรวมหลักฐานเชิงประจักษ์และขั้นตอนการดำเนินงานประจำวัน (Evidence Protocol Workflow)",
        width_inches=4.8,
        height_inches=2.2,
        source_text="กรอบระเบียบวิธีวิจัยโครงการ Health30D, 2569"
    )

    add_figure(
        doc,
        os.path.join(ASSETS_DIR, 'excel_logbook_preview.png'),
        "3.2",
        "โครงสร้างสมุดบันทึกข้อมูลสุขภาพประจำวัน (Daily Health Logbook) บน Microsoft Excel (6806021612037_Health30D.xlsx)",
        width_inches=4.8,
        height_inches=2.1,
        source_text="ไฟล์บันทึกข้อมูลจริงของผู้วิจัยบนระบบ OneDrive, 2569"
    )

    add_heading_2(doc, "3.4.2 โปรโตคอลการแจ้งเตือนและการควบคุมเวลาด้วยสมาร์ตโฟน (Smartphone Timing Protocol)", page_break_before=True)
    add_body_p(doc, (
        "เพื่อแก้ไขปัญหาพฤติกรรมการนั่งทำงานต่อเนื่องยาวนานโดยลืมเวลา ผู้วิจัยได้กำหนดโปรโตคอลการใช้ตัวจับเวลา (Timer) บนสมาร์ตโฟน 2 รูปแบบหลัก ได้แก่:"
    ), space_after=2)
    add_bullet_item(doc, "ตัวเตือนการพักขยับตัวทุก 45 นาที (Active Break Alarm): เมื่อเริ่มเปิดหน้าจอคอมพิวเตอร์ทำงาน จะตั้งตัวจับเวลานับถอยหลัง 45 นาทีทันที "
                         "เมื่อมีสัญญาณเตือนดังขึ้น ผู้วิจัยจะหยุดการพิมพ์ทันที ลุกขึ้นยืน เดินยืดเส้นเบาๆ รอบห้องเป็นเวลา 2 นาที และทำท่าหมุนข้อไหล่ไปข้างหลังช้าๆ 10 ครั้ง เพื่อกระตุ้นการไหลเวียนเลือด")
    add_bullet_item(doc, "ตัวจับเวลาควบคุมการนวดและกายบริหาร (Intervention Timer): ภายหลังเลิกงานใช้ตัวจับเวลาควบคุมการนวดคลึงกล้ามเนื้อ Upper Trapezius และ Levator Scapulae จุดละ 30 วินาที จำนวน 2 รอบ รวมเป็นเวลา 6 นาทีพอดี "
                         "และใช้จับเวลาการยืดเหยียดกล้ามเนื้อค้างไว้ข้างละ 20 วินาที เพื่อให้ได้ประสิทธิผลเชิงสรีรวิทยาตามวรรณกรรมทางการแพทย์")

    add_heading_2(doc, "3.4.3 วิธีการวิเคราะห์และวัดมุมข้อต่อสรีรศาสตร์ด้วยแอปพลิเคชัน ImageMeter")
    add_body_p(doc, (
        "ในการตรวจสอบความถูกต้องของสรีระท่านั่ง ผู้วิจัยใช้สมาร์ตโฟนติดตั้งบนขาตั้งกล้องในระยะห่างและระดับความสูงเดิมทุกครั้ง เพื่อถ่ายภาพด้านข้าง (Lateral View) "
        "จากนั้นนำไฟล์ภาพถ่ายเปิดในแอปพลิเคชัน ImageMeter และกำหนดจุดมาร์กเกอร์ 3 จุดตามหลักกายวิภาคศาสตร์เพื่อวัดมุมข้อต่อ ได้แก่:"
    ), space_after=2)
    add_bullet_item(doc, "มุมข้อศอก (Elbow Angle): กำหนดจุดมาร์กเกอร์ที่ หัวไหล่ (Acromion) - ข้อศอก (Lateral Epicondyle) - ข้อมือ (Styloid Process) เกณฑ์มาตรฐาน 90°–100°")
    add_bullet_item(doc, "มุมสะโพก (Hip Angle): กำหนดจุดมาร์กเกอร์ที่ หัวไหล่ (Acromion) - ข้อสะโพก (Greater Trochanter) - ข้อเข่า (Lateral Femoral Condyle) เกณฑ์มาตรฐาน 90°–100°")
    add_bullet_item(doc, "มุมข้อเข่า (Knee Angle): กำหนดจุดมาร์กเกอร์ที่ ข้อสะโพก - ข้อเข่า - ตาตุ่มนอกข้อเท้า (Lateral Malleolus) เกณฑ์มาตรฐาน 85°–100° ร่วมกับฝ่าเท้าวางราบบนพื้น")
    add_bullet_item(doc, "ระดับสายตา (Eye Level): ลากเส้นระนาบแนวนอนจากระดับสายตาไปยังขอบบนและกึ่งกลางจอแสดงผล เพื่อตรวจสอบว่าคออยู่ในแนวมุมตรง ไม่ก้มเกิน 15°")

    add_figure(
        doc,
        os.path.join(ASSETS_DIR, 'imagemeter_angle_measurement.png'),
        "3.3",
        "การวิเคราะห์และวัดมุมข้อต่อสรีรศาสตร์ 3 จุด จากภาพถ่ายด้านข้างด้วยแอปพลิเคชัน ImageMeter",
        width_inches=4.4,
        height_inches=2.3,
        source_text="ประมวลผลผ่านแอปพลิเคชัน ImageMeter บนสมาร์ตโฟน, 2569"
    )

    add_figure(
        doc,
        os.path.join(ASSETS_DIR, 'workstation_actual_crop.png'),
        "3.4",
        "สภาพแวดล้อมโต๊ะทำงานจริงของผู้วิจัยภายหลังการปรับระดับสรีรศาสตร์ 90-90-90",
        width_inches=3.6,
        height_inches=2.2,
        source_text="ภาพถ่ายสถานีงานจริงของผู้วิจัย, 2569"
    )

    add_heading_1(doc, "3.5 เกณฑ์และการวัดตัวชี้วัดความสำเร็จ 5 ด้าน", page_break_before=True)
    add_body_p(doc, "เกณฑ์การตัดสินความถูกต้องของมุมสรีรศาสตร์ 3 จุด และระดับสายตา ได้กำหนดรายละเอียดเชิงปริมาณ ดังแสดงในตารางที่ 3.2:", indent=False, space_after=3)
    
    headers_3_2 = ["จุดประเมินสรีระ", "จุดอ้างอิงบนร่างกาย (Landmarks)", "เกณฑ์มาตรฐาน", "สาเหตุทางชีวกลศาสตร์"]
    data_3_2 = [
        ["1. มุมข้อศอก (Elbow)", "หัวไหล่ (Acromion) – ข้อศอก – ข้อมือ", "90° – 100°", "วางแขนระนาบเดียวกับคีย์บอร์ด ไม่เอื้อม ลดแรงเกร็งบ่า"],
        ["2. มุมสะโพก (Hip/Trunk)", "หัวไหล่ – ข้อสะโพก – ข้อเข่า", "90° – 100°", "นั่งหลังตรงชิดพนักพิง กระจายน้ำหนักสู่กระดูกก้นกบ"],
        ["3. มุมข้อเข่า (Knee)", "ข้อสะโพก – ข้อเข่า – ข้อเท้า", "85° – 100°", "เบาะนั่งสูงพอดี ฝ่าเท้าวางราบ เลือดไหลเวียนสะดวก"],
        ["4. ระดับสายตา (Eye Level)", "ระดับสายตา – กึ่งกลางจอภาพ", "0° ถึง -15°", "ขอบบนจออยู่ระดับสายตา ไม่ก้มคอ ลดแรงกดกระดูกคอ"]
    ]
    col_w_3_2 = [Inches(1.15), Inches(1.6), Inches(1.2), Inches(1.8)]
    add_styled_table(doc, "3.2", "เกณฑ์การประเมินมุมข้อต่อสรีรศาสตร์ 3 จุดตามหลักการ 90-90-90", headers_3_2, data_3_2, col_w_3_2, font_size_pt=10.0, padding_twips=15)

    add_heading_1(doc, "3.6 มาตรการด้านความปลอดภัยและการเฝ้าระวังผลข้างเคียง")
    add_body_p(doc, (
        "เพื่อความปลอดภัยสูงสุดในการดำเนินวิจัย ผู้วิจัยได้กำหนดข้อห้ามและมาตรการหยุดกิจกรรม (Stop Criteria) อย่างเคร่งครัด: "
        "1) ห้ามนวดบริเวณด้านหน้าและด้านข้างของลำคอ แนวหลอดเลือดแดงใหญ่ (Carotid Artery) แนวกระดูกสันหลังส่วนคอ และกระดูกไหปลาร้าโดยเด็ดขาด "
        "2) หากมีอาการปวดเฉียบพลันเกินกว่าระดับ 4.0 หรือมีอาการปวดร้าวลงแขน แขนชา กล้ามเนื้ออ่อนแรง วิงเวียนศีรษะ หรือตาพร่ามัว ให้หยุดกิจกรรมทันทีและไปพบแพทย์ "
        "และ 3) ตลอดระยะเวลาการดำเนินโครงการ ผู้วิจัยได้จัดทำแบบบันทึกเหตุการณ์ไม่พึงประสงค์ (Adverse Event Log) แสดงไว้ในภาคผนวก ค"
    ))

    # ------------------------------------------
    # CHAPTER 4: RESULTS AND ANALYSIS
    # ------------------------------------------
    add_chapter_title(doc, "4", "ผลการดำเนินงานและการวิเคราะห์ข้อมูล (ข้อมูลถึงวันที่ 14)", space_before=14, space_after=14)
    add_heading_1(doc, "4.1 ข้อมูลฐานเปรียบเทียบก่อนการดำเนินโครงการ (Baseline Assessment Day 1)")
    add_body_p(doc, (
        "จากการประเมินสภาพตนเองในวันแรก (Day 1: 10 กันยายน 2569) ก่อนการเริ่มมาตรการแทรกแซง พบว่าผู้วิจัยมีเวลาใช้งานหน้าจอคอมพิวเตอร์ 275 นาที (4.6 ชั่วโมง) "
        "โดยไม่มีการลุกพักขยับตัวระหว่างทำงานเลย มีระดับความปวดเมื่อยคอบ่าไหล่ประเมินด้วย NRS สูงถึง 6.0 คะแนน และสรีระท่านั่งผ่านเกณฑ์มาตรฐานเพียง 1 ใน 4 จุด (ร้อยละ 25.0) "
        "ดังแสดงสรุปในตารางที่ 4.1:"
    ))
    
    headers_4_1 = ["รายการประเมิน", "ค่าที่วัดได้จริง (Baseline D1)", "เกณฑ์มาตรฐาน", "สถานะการประเมิน"]
    data_4_1 = [
        ["1. ระดับความปวดคอบ่าไหล่ (NRS)", "6.0 คะแนน (ปวดปานกลาง-สูง)", "≤ 4.2 คะแนน (ลด ≥ 30%)", "พบปัญหาความปวดรบกวน"],
        ["2. มุมข้อศอก (Elbow Angle)", "138° (เอื้อมแขนลึกบนโต๊ะ)", "90° – 100°", "ไม่ผ่านเกณฑ์มาตรฐาน"],
        ["3. มุมสะโพก/ลำตัว (Hip Angle)", "125° (นั่งกึ่งนอนไม่พิงพนัก)", "90° – 100°", "ไม่ผ่านเกณฑ์มาตรฐาน"],
        ["4. มุมข้อเข่า (Knee Angle)", "65° (ยกขาพับบนเบาะนั่ง)", "85° – 100° (เท้าราบ)", "ไม่ผ่านเกณฑ์มาตรฐาน"],
        ["5. ระดับสายตา (Eye Level)", "กึ่งกลางจออยู่ในระดับสายตา", "0° ถึง -15° จากสายตา", "ผ่านเกณฑ์มาตรฐาน ✓"]
    ]
    col_w_4_1 = [Inches(1.35), Inches(1.4), Inches(1.4), Inches(1.6)]
    add_styled_table(doc, "4.1", "ข้อมูลฐานเปรียบเทียบก่อนการดำเนินโครงการในวันแรก (Baseline Day 1)", headers_4_1, data_4_1, col_w_4_1, font_size_pt=10.0, padding_twips=15)

    add_heading_1(doc, "4.2 ผลการเปลี่ยนแปลงของระดับความปวดบริเวณคอบ่าไหล่ (NRS D1–D14)", page_break_before=True)
    add_body_p(doc, (
        "จากการดำเนินโครงการแทรกแซงอย่างต่อเนื่องตั้งแต่วันที่ 1 ถึงวันที่ 14 (วันที่ 10–23 กันยายน 2569) "
        "ระดับความปวดก่อนการนวด (NRS Before) มีแนวโน้มลดลงอย่างต่อเนื่องจากระดับ 6.0 ในช่วงวันที่ 1–3 สู่ระดับ 5.0 ในช่วงวันที่ 4–7 "
        "และลดลงสู่ระดับ 4.0 อย่างมีเสถียรภาพตั้งแต่วันที่ 8 จนถึงวันที่ 14 ดังแสดงในตารางที่ 4.2 และรูปที่ 4.1:"
    ))

    # TABLE 4.2: EXACT 5.75 INCHES TOTAL WIDTH - ZERO OVERFLOW GUARANTEED
    headers_4_2 = ["ช่วงการประเมิน / จุดประเมิน", "ระยะเวลา", "นาทีจอเฉลี่ย", "NRS ก่อนนวด", "NRS หลังนวด", "% ลดทันที", "การประเมินและสถานะ"]
    data_4_2 = [
        ["ระยะเริ่มต้น (Baseline Phase)", "วันที่ 1–3 (D1–D3)", "291.0 นาที", "6.0 คะแนน", "5.0 คะแนน", "16.7%", "อาการปวดตึงระดับปานกลาง-สูง"],
        ["ระยะปรับตัว (Adaptation Phase)", "วันที่ 4–7 (D4–D7)", "390.0 นาที", "5.0 คะแนน", "4.0 คะแนน", "20.0%", "อาการปวดเริ่มลดลงอย่างมีนัยสำคัญ"],
        ["ระยะเริ่มทรงตัว (Stabilization)", "วันที่ 8–14 (D8–D14)", "359.0 นาที", "4.0 คะแนน", "3.0 คะแนน", "25.0%", "อาการปวดลดลงคงที่ที่ระดับ 4"],
        ["จุดประเมินครึ่งทาง (Midterm D14)", "วันที่ 14 (D14)", "480.0 นาที", "4.0 คะแนน", "3.0 คะแนน", "25.0%", "ผ่านเกณฑ์ Milestone 2"],
        ["สรุปผลระยะครึ่งทาง (D1 vs D12-14)", "14 วันแรก", "350.6 นาที", "ลดลง 2.0 คะแนน", "ลดลง 2.0 คะแนน", "ลดลง 33.3%", "บรรลุเป้าหมายโครงการ (≥30%) ✓"]
    ]
    col_w_4_2 = [Inches(1.15), Inches(0.85), Inches(0.65), Inches(0.65), Inches(0.65), Inches(0.6), Inches(1.2)]
    add_styled_table(doc, "4.2", "สรุปผลการประเมินระดับความปวดคอบ่าไหล่ (NRS 0–10) รายระยะ วันที่ 1 ถึง 14", headers_4_2, data_4_2, col_w_4_2, font_size_pt=9.5, padding_twips=14)
    add_body_p(doc, "(หมายเหตุ: รายละเอียดข้อมูลบันทึกดิบรายวันตลอด 14 วัน แสดงไว้ในภาคผนวก ก หน้า 34)", indent=False, font_size_pt=11.0, italic=True, space_after=3)

    add_figure(
        doc,
        os.path.join(ASSETS_DIR, 'chart_pain_progression.png'),
        "4.1",
        "กราฟแสดงการเปลี่ยนแปลงระดับความปวดบริเวณคอ บ่า ไหล่ (NRS Score) วันที่ 1 ถึง 14",
        width_inches=4.8,
        height_inches=2.3,
        source_text="วิเคราะห์จากฐานข้อมูลโครงการวิจัย Health30D, 2569"
    )

    add_body_p(doc, (
        "เมื่อคำนวณอัตราการลดลงของระดับความปวดตามสูตรของโครงการ โดยเปรียบเทียบคะแนนวันแรก D1 (6.0 คะแนน) "
        "กับค่าเฉลี่ย 3 วันสุดท้ายของระยะครึ่งทาง (D12–D14 ซึ่งมีคะแนนเฉลี่ย 4.0 คะแนน) พบว่า: "
        "%ลดปวด = (6.0 - 4.0) ÷ 6.0 × 100 = 33.3% ซึ่งสูงกว่าเกณฑ์เป้าหมายขั้นต่ำของโครงการที่กำหนดไว้ไม่น้อยกว่าร้อยละ 30 "
        "สะท้อนให้เห็นว่าการปรับท่าทางและการพักขยับตัวช่วยลดความตึงเครียดของกล้ามเนื้อ Upper Trapezius ได้อย่างมีประสิทธิภาพ"
    ))

    add_heading_1(doc, "4.3 ผลการปรับปรุงท่าทางและมุมข้อต่อสรีรศาสตร์ 90-90-90", page_break_before=True)
    add_body_p(doc, (
        "พัฒนาการของมุมข้อต่อสรีรศาสตร์ที่วัดจากภาพถ่ายด้านข้างด้วยแอปพลิเคชัน ImageMeter แสดงการเปลี่ยนแปลงอย่างชัดเจน "
        "โดยเฉพาะมุมข้อเข่าที่เข้าสู่เกณฑ์มาตรฐานตั้งแต่วันที่ 6 (86°) และรักษาระดับที่ 90°–92° ได้อย่างสม่ำเสมอจากการวางเท้าราบบนพื้น "
        "ในขณะที่มุมข้อศอกค่อยๆ ปรับลดลงจากการดึงคีย์บอร์ดและเมาส์เข้าใกล้ตัว จาก 138° ในวันแรก สู่ 115° ใน D7, 100° ใน D13 "
        "และเข้าสู่เกณฑ์มาตรฐานที่ 98° ใน D14 ส่วนมุมสะโพกดีขึ้นจาก 125° เป็น 108° ดังแสดงในตารางที่ 4.3 และรูปที่ 4.2:"
    ))

    # TABLE 4.3: EXACT 5.75 INCHES TOTAL WIDTH - ZERO OVERFLOW GUARANTEED
    headers_4_3 = ["จุดวัดมุมสรีระ", "เกณฑ์ 90-90-90", "วันแรก (D1)", "สัปดาห์ 1 (D7)", "ครึ่งทาง (D14)", "การเปลี่ยนแปลง", "สถานะการประเมิน"]
    data_4_3 = [
        ["1. มุมข้อศอก (Elbow Angle)", "90° – 100°", "138° (ไม่ผ่าน)", "115° (ไม่ผ่าน)", "98° (ผ่าน ✓)", "ลดลง 40° เข้าสู่เกณฑ์", "ผ่านเกณฑ์มาตรฐาน ✓"],
        ["2. มุมสะโพก/ลำตัว (Hip Angle)", "90° – 100°", "125° (ไม่ผ่าน)", "110° (ไม่ผ่าน)", "108° (ไม่ผ่าน)", "ลดลง 17° ใกล้เคียงเกณฑ์", "เตรียมเสริมหมอนรองหลัง"],
        ["3. มุมข้อเข่า (Knee Angle)", "85° – 100°", "65° (ไม่ผ่าน)", "90° (ผ่าน ✓)", "92° (ผ่าน ✓)", "เพิ่มขึ้น 27° เข้าสู่เกณฑ์", "ผ่านเกณฑ์มาตรฐาน ✓"],
        ["4. ระดับสายตา (Eye Level)", "กึ่งกลางจอ", "ผ่าน ✓", "ผ่าน ✓", "ผ่าน ✓", "คงที่ในระดับสายตา", "ผ่านเกณฑ์มาตรฐาน ✓"],
        ["สรุปภาพรวมท่าทางถูกต้อง", "ผ่าน ≥ 75.0%", "25.0% (1 ใน 4)", "50.0% (2 ใน 4)", "75.0% (3 ใน 4)", "เพิ่มขึ้น +50.0%", "บรรลุเป้าหมายโครงการ ✓"]
    ]
    col_w_4_3 = [Inches(1.2), Inches(0.75), Inches(0.65), Inches(0.7), Inches(0.7), Inches(0.85), Inches(0.9)]
    add_styled_table(doc, "4.3", "สรุปข้อมูลการวัดมุมข้อต่อสรีรศาสตร์เปรียบเทียบตามจุดประเมินสำคัญ (D1–D14)", headers_4_3, data_4_3, col_w_4_3, font_size_pt=9.0, padding_twips=12)
    add_body_p(doc, "(หมายเหตุ: รายละเอียดข้อมูลมุมข้อต่อดิบรายวันทั้ง 14 วัน แสดงไว้ในภาคผนวก ข หน้า 35)", indent=False, font_size_pt=11.0, italic=True, space_after=1)

    add_figure(
        doc,
        os.path.join(ASSETS_DIR, 'chart_joint_angles.png'),
        "4.2",
        "กราฟแสดงพัฒนาการของมุมข้อต่อสรีรศาสตร์เทียบกับแถบเกณฑ์มาตรฐาน 90-90-90 (D1–D14)",
        width_inches=4.0,
        height_inches=1.35,
        source_text="วิเคราะห์จากฐานข้อมูลโครงการวิจัย Health30D, 2569"
    )

    add_body_p(doc, (
        "จากตารางที่ 4.3 พบว่าในวันประเมินสำคัญระยะครึ่งทาง (Day 14) สรีระท่าทางการนั่งของผู้วิจัยสามารถผ่านเกณฑ์ได้ถึง 3 จาก 4 ข้อ "
        "คิดเป็นร้อยละ 75.0 (บรรลุเป้าหมายตัวชี้วัดที่ตั้งไว้ ≥ 75.0%) โดยมีมุมข้อศอก (98°), มุมข้อเข่า (92°) และระดับสายตาผ่านเกณฑ์อย่างสมบูรณ์ "
        "ในขณะทีมุมสะโพก (108°) มีแนวโน้มดีขึ้นอย่างมากเมื่อเทียบกับวันแรก (125°) โดยจะได้วางแผนเสริมหมอนรองหลัง (Lumbar Cushion) ในระยะที่ 2 ต่อไป"
    ), space_after=1, line_spacing=1.08)

    add_heading_1(doc, "4.4 ผลการปฏิบัติตามวินัยการพักขยับตัวทุก 45 นาที", page_break_before=True)
    add_body_p(doc, (
        "ตลอดระยะเวลา 14 วัน ผู้วิจัยมีเวลาใช้งานหน้าจอคอมพิวเตอร์รวมทั้งสิ้น 4,896 นาที (เฉลี่ย 349.7 นาที/วัน หรือประมาณ 5.8 ชั่วโมง/วัน) "
        "มีจำนวนรอบการพักที่ควรทำตามเกณฑ์ (D2 ถึง D14) รวมทั้งสิ้น 99 ครั้ง และสามารถลุกขึ้นพักขยับตัวได้จริง 88 ครั้ง "
        "คิดเป็นอัตราการปฏิบัติตามวินัยการพักขยับตัวเฉลี่ยรวมร้อยละ 89.2 ซึ่งสูงกว่าเกณฑ์เป้าหมายโครงการที่กำหนดไว้ไม่น้อยกว่าร้อยละ 80.0 ดังแสดงในตารางที่ 4.4 และรูปที่ 4.3:"
    ))

    # TABLE 4.4: EXACT 5.75 INCHES TOTAL WIDTH - ZERO OVERFLOW GUARANTEED
    headers_4_4 = ["ช่วงการประเมิน", "นาทีใช้จอเฉลี่ย/วัน", "รอบพักที่ควรทำ", "รอบพักจริง", "% การพักเฉลี่ย", "% ทำครบโปรแกรม", "ผลการประเมินตามเกณฑ์"]
    data_4_4 = [
        ["สัปดาห์ที่ 1 (D1–D7)", "347.6 นาที (5.8 ชม.)", "7.1 รอบ", "6.3 รอบ", "88.0%", "85.7% (6/7 วัน)", "ผ่านเกณฑ์เป้าหมาย"],
        ["สัปดาห์ที่ 2 (D8–D14)", "353.3 นาที (5.9 ชม.)", "7.4 รอบ", "6.7 รอบ", "90.3%", "85.7% (6/7 วัน)", "อัตราการพักพัฒนาขึ้นต่อเนื่อง"],
        ["วันงานด่วนช่วงดึก (D6 & D13)", "435.0 นาที (7.3 ชม.)", "9.0 รอบ", "7.5 รอบ", "83.3%", "0.0% (ขาดนวด/บริหาร)", "ได้รับบทเรียนปรับกลยุทธ์"],
        ["วันประเมินครึ่งทาง (D14)", "480.0 นาที (8.0 ชม.)", "10.0 รอบ", "9.0 รอบ", "90.0%", "100.0% (ทำครบทุกท่า)", "ผ่านเกณฑ์ Milestone 2"],
        ["เฉลี่ยรวม 14 วัน (D1–D14)", "350.5 นาที (5.8 ชม.)", "7.3 รอบ", "6.5 รอบ", "89.2%", "85.7% (12/14 วัน)", "บรรลุเกณฑ์เป้าหมายทั้ง 2 ด้าน ✓"]
    ]
    col_w_4_4 = [Inches(1.15), Inches(0.85), Inches(0.65), Inches(0.65), Inches(0.6), Inches(0.65), Inches(1.2)]
    add_styled_table(doc, "4.4", "สรุปเวลาใช้งานหน้าจอและอัตราการปฏิบัติตามวินัยการพักขยับตัวทุก 45 นาที (D1–D14)", headers_4_4, data_4_4, col_w_4_4, font_size_pt=9.5, padding_twips=14)

    add_figure(
        doc,
        os.path.join(ASSETS_DIR, 'chart_screen_and_breaks.png'),
        "4.3",
        "กราฟแสดงเวลาใช้งานหน้าจอและจำนวนรอบการพักขยับตัวรายวันเปรียบเทียบกับเกณฑ์ (D1–D14)",
        width_inches=4.4,
        height_inches=1.75,
        source_text="วิเคราะห์จากฐานข้อมูลโครงการวิจัย Health30D, 2569"
    )

    add_heading_1(doc, "4.5 ผลการประเมินความสม่ำเสมอในการปฏิบัติตามโปรแกรมกายบริหารและการนวด", page_break_before=True)
    add_body_p(doc, (
        "ความสม่ำเสมอในการปฏิบัติตามโปรแกรมนวด 6 นาที และกายบริหาร 6 ท่า ตลอด 14 วันแรก ผู้วิจัยสามารถปฏิบัติได้ครบถ้วนทั้งสิ้น 12 วัน "
        "คิดเป็นร้อยละ 85.7 (บรรลุเกณฑ์เป้าหมายที่กำหนดไว้ ≥ 85.0%) โดยมี 2 วันที่ขาดการปฏิบัติ ได้แก่ วันที่ 6 (D6) และวันที่ 13 (D13) "
        "เนื่องจากมีภารกิจส่งงานและอ่านหนังสือสอบช่วงดึกจนเกินเวลา 01:00 น. ทำให้เกิดความเหนื่อยล้าสะสมและละเลยการทำกิจกรรม "
        "อย่างไรก็ดี การแจ้งเตือนและบันทึกข้อมูลอย่างเคร่งครัดทำให้ผู้วิจัยตระหนักรู้และสามารถฟื้นฟูกลับมาทำครบ 100% ในวันถัดไปได้ทันที"
    ))

    add_heading_1(doc, "4.6 ผลสัมฤทธิ์ภาพรวมเปรียบเทียบกับตัวชี้วัดความสำเร็จ 5 ด้าน (Midterm Evaluation)")
    add_body_p(doc, "เมื่อนำผลการดำเนินงานตลอด 14 วันแรก มาประเมินเปรียบเทียบกับกรอบตัวชี้วัดความสำเร็จ 5 ด้านที่กำหนดไว้ในบทที่ 1 สรุปได้ดังตารางที่ 4.5:", indent=False, space_after=3)
    
    headers_4_5 = ["ตัวชี้วัดความสำเร็จ", "ค่าฐานตั้งต้น (Baseline D1)", "ผลลัพธ์จริงระยะครึ่งทาง (D14)", "เกณฑ์เป้าหมาย", "ผลการประเมินความสำเร็จ"]
    data_4_5 = [
        ["1. อัตราการลดระดับความปวด (KR 1)", "6.0 คะแนน", "4.0 คะแนน (ลดลง 33.3%)", "ลดลง ≥ 30.0%", "บรรลุเป้าหมาย ✓ (เกินเกณฑ์ 3.3%)"],
        ["2. อัตราความถูกต้องสรีระ (KR 2)", "25.0% (1 ใน 4 จุด)", "75.0% (3 ใน 4 จุด)", "ผ่านเกณฑ์ ≥ 75.0%", "บรรลุเป้าหมาย ✓ (ผ่าน 3 จุดหลัก)"],
        ["3. วินัยการพักขยับตัว (KR 3)", "0.0% (ไม่เคยพัก)", "89.2% (เฉลี่ย 6.5/7.3 รอบ)", "อัตราการพัก ≥ 80.0%", "บรรลุเป้าหมาย ✓ (เกินเกณฑ์ 9.2%)"],
        ["4. ความสม่ำเสมอทำโปรแกรม (KR 4)", "0.0% (ไม่เคยทำ)", "85.7% (ทำครบ 12 จาก 14 วัน)", "ความสม่ำเสมอ ≥ 85.0%", "บรรลุเป้าหมาย ✓ (ผ่านเกณฑ์พอดี)"],
        ["5. ประสิทธิผลทันทีจากการนวด (KR 5)", "0.0% (ไม่มีการนวด)", "ลดลงเฉลี่ย 21.6% ทันที", "ลดลงเฉลี่ย ≥ 20.0%", "บรรลุเป้าหมาย ✓ (ลดทันที 1.0 คะแนน)"]
    ]
    col_w_4_5 = [Inches(1.1), Inches(1.2), Inches(1.45), Inches(1.0), Inches(1.0)]
    add_styled_table(doc, "4.5", "สรุปผลสัมฤทธิ์เปรียบเทียบกับตัวชี้วัดความสำเร็จ 5 ด้าน (Midterm Evaluation)", headers_4_5, data_4_5, col_w_4_5, font_size_pt=9.5, padding_twips=14)

    # ------------------------------------------
    # CHAPTER 5: CONCLUSION & RECOMMENDATIONS
    # ------------------------------------------
    add_chapter_title(doc, "5", "สรุปผล อภิปรายผล และข้อเสนอแนะ", space_before=14, space_after=14)
    add_heading_1(doc, "5.1 สรุปผลการดำเนินโครงการระยะครึ่งทาง")
    add_body_p(doc, (
        "การดำเนินโครงงานวิจัยรายบุคคลเพื่อปรับเปลี่ยนพฤติกรรมสุขภาพและสภาพแวดล้อมโต๊ะทำงานตลอดระยะเวลา 14 วันแรก (Days 1 ถึง 14) "
        "สามารถสรุปผลการศึกษาเชื่อมโยงกับวัตถุประสงค์ทั้ง 3 ข้อของโครงงานได้ดังนี้:"
    ), space_after=3)
    add_numbered_item(doc, "1.", "ประสิทธิผลด้านการลดความปวดเมื่อย: การปรับสถานีงานตามหลัก 90-90-90 ร่วมกับการพักขยับตัวทุก 45 นาที และการบำบัดฟื้นฟูตนเอง สามารถลดระดับความปวดคอบ่าไหล่ (NRS) ได้ร้อยละ 33.3 (จาก 6.0 เหลือ 4.0 คะแนน) บรรลุตามวัตถุประสงค์ข้อที่ 1")
    add_numbered_item(doc, "2.", "ประสิทธิผลด้านการปรับปรุงสรีระท่าทาง: มุมข้อต่อสำคัญได้รับการฟื้นฟูเข้าสู่เกณฑ์มาตรฐานร้อยละ 75.0 (มุมข้อศอก 98°, มุมข้อเข่า 92° และระดับสายตาผ่านเกณฑ์) บรรลุตามวัตถุประสงค์ข้อที่ 2")
    add_numbered_item(doc, "3.", "ประสิทธิผลด้านระเบียบวิธีบันทึกข้อมูลเชิงประจักษ์: การใช้ Microsoft Excel บนคลาวด์ OneDrive ร่วมกับการจับเวลาบนสมาร์ตโฟน และแอป ImageMeter สามารถทำหน้าที่เป็นระบบบันทึกและประมวลผลหลักฐานเชิงประจักษ์ได้อย่างแม่นยำ บรรลุตามวัตถุประสงค์ข้อที่ 3")

    add_heading_1(doc, "5.2 การอภิปรายผลการวิจัย", page_break_before=True)
    add_body_p(doc, (
        "ผลการศึกษาที่ระดับความปวดลดลงอย่างมีนัยสำคัญร้อยละ 33.3 สอดคล้องกับทฤษฎีชีวกลศาสตร์และการยศาสตร์ของ NIOSH (1997) "
        "ซึ่งระบุว่าการขยับแป้นพิมพ์และเมาส์เข้ามาใกล้ลำตัว ช่วยลดโมเมนต์แขนกล (Moment Arm) ของข้อต่อหัวไหล่ลง ส่งผลให้กล้ามเนื้อ Upper Trapezius "
        "ไม่ต้องออกแรงพยุงแขนในลักษณะเกร็งค้างคงที่ (Static Load) อีกต่อไป นอกจากนี้ การลุกขึ้นยืนพักขยับตัวทุก 45 นาที ช่วยฟื้นฟูออกซิเจนและการไหลเวียนเลือดสู่กล้ามเนื้อคอ "
        "สอดคล้องกับงานวิจัยของ Hedge & Ray (2004) ที่ยืนยันว่า Active Microbreaks ช่วยลดความล้าของกล้ามเนื้อได้โดยไม่กระทบต่อประสิทธิภาพการทำงาน "
        "ส่วนการนวดกดจุด 6 นาทีและการบริหาร 6 ท่า ช่วยคลายจุด Trigger Points และเสริมสร้างความแข็งแรงของกล้ามเนื้อ Deep Cervical Flexors ป้องกันการกลับมาปวดซ้ำได้อย่างยั่งยืน"
    ))

    add_heading_1(doc, "5.3 ปัญหา อุปสรรค และแนวทางแก้ไขที่ได้เรียนรู้")
    add_body_p(doc, "ในการดำเนินงาน 14 วันแรก ผู้วิจัยได้พบบทเรียนและอุปสรรคสำคัญ 3 ประการ พร้อมแนวทางแก้ไข ดังนี้:", indent=False, space_after=3)
    add_numbered_item(doc, "1.", "ปัญหาการเผลอนั่งกึ่งนอนในวันที่เหนื่อยล้า: ในช่วงบ่ายบางวัน ผู้วิจัยมีแนวโน้มเลื่อนสะโพกมาข้างหน้า ทำให้มุมสะโพกเปิดกว้าง 108°–110° "
                          "แนวทางแก้ไขคือ เตรียมจัดหาหมอนรองหลัง (Lumbar Support Cushion) มาเสริมที่พนักพิงเก้าอี้เพื่อล็อกแนวกระดูกสันหลังส่วนเอวให้ตั้งตรงอัตโนมัติ")
    add_numbered_item(doc, "2.", "ปัญหาการลืมพักในวันที่มีงานเขียนโค้ดเร่งด่วน (D6 และ D13): เมื่อมีสมาธิจดจ่อสูง มักละเลยการลุกพัก "
                          "แนวทางแก้ไขคือ ปรับระดับเสียงเตือนของนาฬิกาสมาร์ตโฟนให้ดังขึ้น และวางโทรศัพท์ไว้ห่างจากโต๊ะทำงานเล็กน้อย เพื่อบังคับให้ต้องลุกขึ้นเดินไปปิดเสียง")
    add_numbered_item(doc, "3.", "ข้อจำกัดของเก้าอี้ที่ไม่มีที่วางแขนปรับระดับได้: ทำให้การรองรับน้ำหนักแขนต้องอาศัยการวางข้อศอกบนโต๊ะ "
                          "แนวทางแก้ไขคือ ปรับความลึกของโต๊ะและดึงคีย์บอร์ดให้อยู่ห่างจากขอบโต๊ะเพียง 10 เซนติเมตร เพื่อให้ท่อนแขนวางราบบนโต๊ะได้อย่างผ่อนคลาย")

    add_heading_1(doc, "5.4 ข้อเสนอแนะสำหรับการดำเนินงานระยะที่ 2 และการนำไปใช้ประโยชน์", page_break_before=True)
    add_body_p(doc, "ข้อเสนอแนะสำหรับการดำเนินโครงการระยะที่ 2 (Days 15 ถึง 30):", bold_prefix="1. ", indent=False, space_after=2)
    add_bullet_item(doc, "จัดหาหมอนรองหลังเพื่อแก้ไขมุมสะโพกให้ลดลงจาก 108° เข้าสู่เกณฑ์มาตรฐาน 90°–100° ให้สำเร็จภายใน Day 21")
    add_bullet_item(doc, "เพิ่มความเข้มข้นของการฝึกท่า Chin Tuck และ Wall Slide จาก 10 ครั้ง เป็น 12 ครั้งต่อรอบ เพื่อเพิ่มความทนทานของกล้ามเนื้อ")
    add_bullet_item(doc, "รักษาวินัยการบันทึกข้อมูลลงไฟล์ Microsoft Excel บนระบบ OneDrive ให้ครบ 100% จนสิ้นสุดโครงการ 30 วัน")
    
    add_body_p(doc, "ข้อเสนอแนะสำหรับการนำไปใช้ประโยชน์ในวงกว้าง:", bold_prefix="2. ", indent=False, space_before=4, space_after=2)
    add_bullet_item(doc, "สำหรับนักศึกษาสายคอมพิวเตอร์: ควรนำโมเดลการจัดโต๊ะทำงาน 90-90-90 ร่วมกับการตั้งเวลาเตือนพักในสมาร์ตโฟน ไปปรับใช้เป็นพฤติกรรมประจำวันเพื่อป้องกันโรคเรื้อรัง")
    add_bullet_item(doc, "สำหรับสถาบันการศึกษา: ควรส่งเสริมการนำโปรโตคอลการดูแลตนเองนี้ ไปใช้เป็นสื่อการสอนหรือกิจกรรมเสริมหลักสูตรในรายวิชาศึกษาทั่วไป หมวดสุขภาพเพื่อชีวิต")
    add_bullet_item(doc, "สำหรับงานวิจัยในอนาคต: ควรขยายขนาดกลุ่มตัวอย่างจาก N=1 สู่กลุ่มตัวอย่างขนาดใหญ่ เพื่อเปรียบเทียบผลลัพธ์เชิงสถิติในระดับประชากรต่อไป")

    # ==========================================
    # SECTION 7: BACK MATTER
    # ==========================================
    # REFERENCES
    sec_ref = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section_margins(sec_ref, top=1.5, bottom=1.0, left=1.5, right=1.0)
    clear_pg_num_types(sec_ref)
    add_header_page_field(sec_ref.header)
    
    add_front_matter_title(doc, "บรรณานุกรม")
    add_reference_item(doc, "กรมการแพทย์ กระทรวงสาธารณสุข. (2565). แนวทางการดูแลรักษาและป้องกันกลุ่มอาการออฟฟิศซินโดรมสำหรับประชาชน. กรุงเทพฯ: โรงพิมพ์ชุมนุมสหกรณ์การเกษตรแห่งประเทศไทย.")
    add_reference_item(doc, "ปาริฉัตร อารีรักษ์, และ สุรชัย ชัยทัศนีย์. (2563). ประสิทธิผลของการยืดเหยียดกล้ามเนื้อร่วมกับการนวดกดจุดในการลดอาการปวดกล้ามเนื้อคอบ่าในพนักงานสำนักงาน. วารสารกายภาพบำบัด, 42(2), 85-96.")
    add_reference_item(doc, "มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ. (2567). เกณฑ์มาตรฐานการจัดทำโครงงานวิจัยและปริญญานิพนธ์ระดับปริญญาบัณฑิต. กรุงเทพฯ: มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ.")
    add_reference_item(doc, "Hansraj, K. K. (2014). Assessment of stresses in the cervical spine caused by posture and position of the head. Surgical Technology International, 25, 277-279.")
    add_reference_item(doc, "Hedge, A., & Ray, E. J. (2004). Effects of an electronic ergonomics tutorial on seated posture and musculoskeletal discomfort in computer users. Proceedings of the Human Factors and Ergonomics Society Annual Meeting, 48(8), 1082-1086.")
    add_reference_item(doc, "Kazdin, A. E. (2011). Single-Case Research Designs: Methods for Clinical and Applied Settings (2nd ed.). Oxford University Press.")
    add_reference_item(doc, "National Institute for Occupational Safety and Health (NIOSH). (1997). Musculoskeletal Disorders and Workplace Factors: A Critical Review of Epidemiologic Evidence for Work-Related Musculoskeletal Disorders of the Neck, Upper Extremity, and Low Back. Cincinnati, OH: U.S. Department of Health and Human Services.")
    add_reference_item(doc, "Simons, D. G., Travell, J. G., & Simons, L. S. (1999). Travell & Simons' Myofascial Pain and Dysfunction: The Trigger Point Manual (2nd ed., Vol. 1). Baltimore: Williams & Wilkins.")

    # APPENDIX A: DAILY RAW DATA (Table ก.1 fits within 5.75 inches)
    sec_appa = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section_margins(sec_appa, top=1.5, bottom=1.0, left=1.5, right=1.0)
    clear_pg_num_types(sec_appa)
    add_header_page_field(sec_appa.header)
    
    add_front_matter_title(doc, "ภาคผนวก ก\\nตารางบันทึกข้อมูลประจำวันฉบับเต็ม Days 1–14\\n(Daily Health Logbook Raw Data)")
    add_body_p(doc, "ข้อมูลดิบการใช้งานคอมพิวเตอร์และการดำเนินกิจกรรมรายวัน บันทึกใน Microsoft Excel (6806021612037_Health30D.xlsx) ตลอด 14 วันแรก:", indent=False, space_after=3)
    
    headers_appa = ["วัน", "วันที่", "เริ่มจอ", "นาที", "ควรพัก", "พักจริง", "%พัก", "NRSก่อน", "NRSหลัง", "%ลด", "นวด", "บริหาร", "หมายเหตุการบันทึก"]
    data_appa = [
        ["D1", "10/09", "13:30", "275", "6", "0", "0.0%", "6.0", "6.0", "0.0%", "-", "-", "วันแรกเก็บค่าฐาน (Baseline) ไม่มีการแทรกแซง"],
        ["D2", "11/09", "09:15", "310", "6", "6", "100.0%", "6.0", "5.0", "16.7%", "ครบ", "ครบ", "ปรับโต๊ะวันแรก พักครบทุกรอบ"],
        ["D3", "12/09", "10:00", "288", "6", "5", "83.3%", "6.0", "5.0", "16.7%", "ครบ", "ครบ", "เริ่มคุ้นชินกับการเตือนพัก 45 นาที"],
        ["D4", "13/09", "09:00", "360", "8", "7", "87.5%", "5.0", "4.0", "20.0%", "ครบ", "ครบ", "อาการปวดเริ่มลดลงสู่ระดับ 5"],
        ["D5", "14/09", "13:00", "410", "9", "8", "88.9%", "5.0", "4.0", "20.0%", "ครบ", "ครบ", "ชั่วโมงใช้จอนานขึ้น แต่ยังคุมการพักได้ดี"],
        ["D6", "15/09", "08:30", "390", "8", "7", "87.5%", "5.0", "5.0", "0.0%", "ขาด", "ขาด", "งานดึกช่วงสอบ ละเลยนวดและบริหาร"],
        ["D7", "16/09", "10:30", "420", "9", "8", "88.9%", "5.0", "4.0", "20.0%", "ครบ", "ครบ", "ประเมิน Milestone 1 ถ่ายภาพ D7"],
        ["D8", "17/09", "09:00", "315", "7", "6", "85.7%", "4.0", "3.0", "25.0%", "ครบ", "ครบ", "อาการปวดลดลงสู่ระดับ 4 อย่างชัดเจน"],
        ["D9", "18/09", "11:00", "295", "6", "5", "83.3%", "4.0", "3.0", "25.0%", "ครบ", "ครบ", "กล้ามเนื้อบ่าผ่อนคลายขึ้นมาก"],
        ["D10", "19/09", "09:30", "340", "7", "7", "100.0%", "4.0", "3.0", "25.0%", "ครบ", "ครบ", "พักขยับตัวครบ 100% ท่าทางเริ่มนิ่ง"],
        ["D11", "20/09", "10:00", "355", "7", "6", "85.7%", "4.0", "3.0", "25.0%", "ครบ", "ครบ", "ความเมื่อยล้าสะสมระหว่างวันลดลง"],
        ["D12", "21/09", "13:00", "380", "8", "7", "87.5%", "4.0", "3.0", "25.0%", "ครบ", "ครบ", "ทำโปรแกรมกายบริหารคล่องแคล่ว"],
        ["D13", "22/09", "08:45", "480", "10", "8", "80.0%", "4.0", "4.0", "0.0%", "ขาด", "ขาด", "งานเขียนโค้ดเร่งด่วน ขาดนวดช่วงดึก"],
        ["D14", "23/09", "09:00", "480", "10", "9", "90.0%", "4.0", "3.0", "25.0%", "ครบ", "ครบ", "ประเมินครึ่งทาง D14 บรรลุเกณฑ์ทุกข้อ"]
    ]
    col_w_appa = [Inches(0.32), Inches(0.52), Inches(0.48), Inches(0.42), Inches(0.42), Inches(0.42), Inches(0.42), Inches(0.42), Inches(0.42), Inches(0.45), Inches(0.38), Inches(0.42), Inches(1.09)]
    add_styled_table(doc, "ก.1", "ตารางบันทึกข้อมูลประจำวันฉบับเต็ม Days 1–14 (Daily Raw Data)", headers_appa, data_appa, col_w_appa, font_size_pt=8.5, padding_twips=10)

    # APPENDIX B: JOINT ANGLE RAW DATA (Table ข.1 fits within 5.70 inches)
    sec_appb = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section_margins(sec_appb, top=1.5, bottom=1.0, left=1.5, right=1.0)
    clear_pg_num_types(sec_appb)
    add_header_page_field(sec_appb.header)
    
    add_front_matter_title(doc, "ภาคผนวก ข\\nข้อมูลการวัดมุมข้อต่อสรีรศาสตร์รายวันฉบับเต็ม Days 1–14\\n(ImageMeter Posture Angle Measurement)")
    add_body_p(doc, "ข้อมูลการวัดมุมข้อต่อสรีระ 3 ตำแหน่งและการประเมินระดับสายตาจากภาพถ่ายด้านข้างด้วย ImageMeter ตลอด 14 วัน:", indent=False, space_after=3)
    
    headers_appb = ["วัน", "วันที่", "สถานะภาพถ่าย", "มุมศอก", "เกณฑ์ศอก", "มุมสะโพก", "เกณฑ์สะโพก", "มุมเข่า", "เกณฑ์เข่า", "สายตา", "ผ่าน (4)", "ผลประเมิน"]
    data_appb = [
        ["D1", "10/09", "บันทึก Baseline", "138°", "ไม่ผ่าน", "125°", "ไม่ผ่าน", "65°", "ไม่ผ่าน", "ผ่าน", "1/4 (25%)", "ไม่ผ่านเกณฑ์"],
        ["D2", "11/09", "บันทึกติดตาม", "130°", "ไม่ผ่าน", "120°", "ไม่ผ่าน", "75°", "ไม่ผ่าน", "ผ่าน", "1/4 (25%)", "เริ่มปรับโต๊ะ"],
        ["D3", "12/09", "บันทึกติดตาม", "125°", "ไม่ผ่าน", "118°", "ไม่ผ่าน", "80°", "ไม่ผ่าน", "ผ่าน", "1/4 (25%)", "ดึงคีย์บอร์ดเข้าใกล้"],
        ["D4", "13/09", "บันทึกติดตาม", "120°", "ไม่ผ่าน", "115°", "ไม่ผ่าน", "82°", "ไม่ผ่าน", "ผ่าน", "1/4 (25%)", "ปรับเบาะนั่ง"],
        ["D5", "14/09", "บันทึกติดตาม", "118°", "ไม่ผ่าน", "112°", "ไม่ผ่าน", "85°", "ผ่าน ✓", "ผ่าน", "2/4 (50%)", "เข่าเริ่มผ่านเกณฑ์"],
        ["D6", "15/09", "บันทึกติดตาม", "116°", "ไม่ผ่าน", "112°", "ไม่ผ่าน", "86°", "ผ่าน ✓", "ผ่าน", "2/4 (50%)", "เท้าวางราบบนพื้น"],
        ["D7", "16/09", "Milestone 1", "115°", "ไม่ผ่าน", "110°", "ไม่ผ่าน", "90°", "ผ่าน ✓", "ผ่าน", "2/4 (50%)", "ข้อเข่า 90° สมบูรณ์"],
        ["D8", "17/09", "บันทึกติดตาม", "110°", "ไม่ผ่าน", "110°", "ไม่ผ่าน", "90°", "ผ่าน ✓", "ผ่าน", "2/4 (50%)", "ข้อศอกพัฒนาต่อเนื่อง"],
        ["D9", "18/09", "บันทึกติดตาม", "108°", "ไม่ผ่าน", "110°", "ไม่ผ่าน", "90°", "ผ่าน ✓", "ผ่าน", "2/4 (50%)", "หลังชิดพนักพิงมากขึ้น"],
        ["D10", "19/09", "บันทึกติดตาม", "105°", "ไม่ผ่าน", "108°", "ไม่ผ่าน", "90°", "ผ่าน ✓", "ผ่าน", "2/4 (50%)", "ศอกใกล้เคียงเกณฑ์"],
        ["D11", "20/09", "บันทึกติดตาม", "104°", "ไม่ผ่าน", "108°", "ไม่ผ่าน", "90°", "ผ่าน ✓", "ผ่าน", "2/4 (50%)", "ฝึกวางแขนให้เป็นนิสัย"],
        ["D12", "21/09", "บันทึกติดตาม", "102°", "ไม่ผ่าน", "108°", "ไม่ผ่าน", "92°", "ผ่าน ✓", "ผ่าน", "2/4 (50%)", "ศอกเกือบถึง 100°"],
        ["D13", "22/09", "บันทึกติดตาม", "100°", "ผ่าน ✓", "108°", "ไม่ผ่าน", "92°", "ผ่าน ✓", "ผ่าน", "3/4 (75%)", "ศอกผ่านเกณฑ์แล้ว"],
        ["D14", "23/09", "Midterm D14", "98°", "ผ่าน ✓", "108°", "ไม่ผ่าน", "92°", "ผ่าน ✓", "ผ่าน", "3/4 (75%)", "บรรลุเกณฑ์ KPI 2 ✓"]
    ]
    col_w_appb = [Inches(0.32), Inches(0.52), Inches(0.72), Inches(0.42), Inches(0.42), Inches(0.42), Inches(0.42), Inches(0.42), Inches(0.42), Inches(0.42), Inches(0.55), Inches(1.07)]
    add_styled_table(doc, "ข.1", "ข้อมูลการวัดมุมข้อต่อสรีรศาสตร์รายวันฉบับเต็ม Days 1–14", headers_appb, data_appb, col_w_appb, font_size_pt=8.5, padding_twips=10)

    # APPENDIX C: SAFETY LOG
    sec_appc = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section_margins(sec_appc, top=1.5, bottom=1.0, left=1.5, right=1.0)
    clear_pg_num_types(sec_appc)
    add_header_page_field(sec_appc.header)
    
    add_front_matter_title(doc, "ภาคผนวก ค\\nแบบบันทึกความปลอดภัยและเหตุการณ์ไม่พึงประสงค์\\n(Safety Monitoring & Adverse Event Log)")
    add_body_p(doc, "บันทึกการเฝ้าระวังอาการข้างเคียงหรือผลกระทบทางลบจากการดำเนินโครงการวิจัยตลอด 14 วันแรก:", indent=False, space_after=3)
    add_bullet_item(doc, "อาการทางระบบประสาท (Neurological Symptoms): ไม่พบอาการชา ปวดร้าวลงแขน หรือกล้ามเนื้อแขนอ่อนแรงตลอดระยะเวลาการศึกษา")
    add_bullet_item(doc, "อาการระบบไหลเวียนโลหิต (Circulatory Symptoms): ไม่พบอาการหน้ามืด วิงเวียนศีรษะ หรือความดันโลหิตผิดปกติจากการเปลี่ยนท่าทาง")
    add_bullet_item(doc, "อาการบวมช้ำหรือบาดเจ็บเฉียบพลัน (Tissue Trauma): ไม่พบรอยฟกช้ำหรือการอักเสบของผิวหนังบริเวณจุดนวดบ่าคอ")
    add_bullet_item(doc, "การหยุดกิจกรรมฉุกเฉิน (Stop Criteria Activation): ไม่มีการเปิดใช้เกณฑ์หยุดกิจกรรมฉุกเฉิน การดำเนินวิจัยเป็นไปอย่างราบรื่นและปลอดภัย 100%")

    # APPENDIX D: EXCEL STRUCTURE AND FORMULAS
    sec_appd = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section_margins(sec_appd, top=1.5, bottom=1.0, left=1.5, right=1.0)
    clear_pg_num_types(sec_appd)
    add_header_page_field(sec_appd.header)
    
    add_front_matter_title(doc, "ภาคผนวก ง\\nโครงสร้างไฟล์ Microsoft Excel และสูตรคำนวณที่ใช้ในการวิจัย\\n(Excel Workbook Architecture & Formulas)")
    add_body_p(doc, "โครงสร้างและการจัดวางสูตรคำนวณอัตโนมัติในไฟล์สมุดบันทึกสุขภาพ '6806021612037_Health30D.xlsx' บน Microsoft OneDrive ประกอบด้วย:", indent=False, space_after=3)
    add_bullet_item(doc, "โครงสร้างเวิร์กชีต Daily_Logbook: บันทึกข้อมูลคอลัมน์ A ถึง M ได้แก่ วัน (Day), วันที่ (Date), เวลาเริ่มใช้จอ, เวลาสิ้นสุด, นาทีใช้จอรวม, จำนวนรอบพักที่ควรทำ, จำนวนรอบพักจริง, อัตราส่วนร้อยละการพัก, คะแนนปวดก่อนนวด (NRS Pre), คะแนนปวดหลังนวด (NRS Post), ร้อยละการลดปวดทันที, การทำโปรแกรมนวดและบริหาร, และหมายเหตุ")
    add_bullet_item(doc, "สูตรคำนวณรอบการพักที่ควรทำ: =INT(E2/45) (นำเวลาใช้งานหน้าจอนาทีในคอลัมน์ E หารด้วย 45 และปัดเศษทศนิยมทิ้ง)")
    add_bullet_item(doc, "สูตรคำนวณร้อยละการพักขยับตัวรายวัน: =IF(F2>0, (G2/F2)*100, \"-\") (คำนวณสัดส่วนรอบพักจริงต่อรอบพักที่ควรทำ)")
    add_bullet_item(doc, "สูตรคำนวณร้อยละการลดความปวดทันทีจากการนวด: =IF(I2>0, ((I2-J2)/I2)*100, 0) (คำนวณอัตราลดลงของคะแนน NRS หลังนวดเทียบกับก่อนนวด)")
    add_bullet_item(doc, "สูตรคำนวณร้อยละการลดความปวดสุทธิตามเกณฑ์โครงการ (Key Result 1): =((I2-AVERAGE(I13:I15))/I2)*100 (เปรียบเทียบคะแนนวันแรก D1 กับค่าเฉลี่ย 3 วันสุดท้าย D12-D14)")
    add_bullet_item(doc, "สูตรคำนวณความสม่ำเสมอในการทำโปรแกรมครบถ้วน (Key Result 4): =(COUNTIF(L2:L15, \"ครบ\")/14)*100 (นับจำนวนวันที่ปฏิบัติตามโปรแกรมนวดและกายบริหารครบทั้ง 6 ท่า)")

    # BIOGRAPHY
    sec_bio = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section_margins(sec_bio, top=1.5, bottom=1.0, left=1.5, right=1.0)
    clear_pg_num_types(sec_bio)
    add_header_page_field(sec_bio.header)
    
    add_front_matter_title(doc, "ประวัติผู้จัดทำ")
    
    tbl_bio = doc.add_table(rows=6, cols=2)
    tbl_bio.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_bio.autofit = False
    
    tbl_borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="none"/><w:bottom w:val="none"/><w:left w:val="none"/>'
        f'<w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/>'
        f'</w:tblBorders>'
    )
    tbl_bio._tbl.tblPr.append(tbl_borders)
    col_w_bio = [Inches(1.8), Inches(3.95)]
    bio_data = [
        ["ชื่อ - นามสกุล:", "นายปภาวิน ธิติชุณหกุล (Mr. Paphavin Thitichunhakun)"],
        ["รหัสนักศึกษา:", "6806021612037"],
        ["การศึกษา:", "นักศึกษาชั้นปีที่ 1 แขนงวิชาเทคโนโลยีสารสนเทศ (IT)\\nภาควิชาเทคโนโลยีสารสนเทศ คณะเทคโนโลยีและการจัดการอุตสาหกรรม\\nมหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ วิทยาเขตปราจีนบุรี"],
        ["รายวิชา:", "080303609 สุขภาพเพื่อชีวิต (Healthy Life) กลุ่มเรียนที่ 2 (Sec 2)\\nภาคการศึกษาที่ 1 ปีการศึกษา 2569"],
        ["สถานที่ติดต่อ:", "คณะเทคโนโลยีและการจัดการอุตสาหกรรม มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ\\nวิทยาเขตปราจีนบุรี 129 หมู่ 21 ตำบลเนินหอม อำเภอเมือง จังหวัดปราจีนบุรี 25230"],
        ["ความสนใจทางวิชาการ:", "การยศาสตร์เชิงคอมพิวเตอร์ (Computer Ergonomics), การวิเคราะห์ข้อมูลสุขภาพเชิงประจักษ์ (Empirical Health Analytics), และการจัดการสารสนเทศส่วนบุคคล (Personal Information Management)"]
    ]
    
    for r_idx, (lbl, val) in enumerate(bio_data):
        cell_lbl = tbl_bio.cell(r_idx, 0)
        cell_val = tbl_bio.cell(r_idx, 1)
        cell_lbl.width = col_w_bio[0]
        cell_val.width = col_w_bio[1]
        
        p0 = cell_lbl.paragraphs[0]
        format_paragraph(p0, space_before=2, space_after=3, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.LEFT)
        p0.paragraph_format.first_line_indent = Inches(0)
        r0 = p0.add_run(thai_zwsp(lbl))
        set_run_font(r0, size_pt=15, bold=True)
        
        p1 = cell_val.paragraphs[0]
        format_paragraph(p1, space_before=2, space_after=3, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.THAI_JUSTIFY)
        p1.paragraph_format.first_line_indent = Inches(0)
        r1 = p1.add_run(thai_zwsp(val))
        set_run_font(r1, size_pt=15, bold=False)

    doc.save(output_path)
    print(f"Generated DOCX: {output_path}")

def convert_docx_to_pdf(docx_path, pdf_path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(os.path.abspath(docx_path))
    doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17) # 17 = wdExportFormatPDF
    doc.Close(False)
    word.Quit()
    print(f"Converted PDF: {pdf_path}")

def scan_pdf_pages(pdf_path):
    doc = fitz.open(pdf_path)
    page_map = {}
    
    FRONT_MATTER_KEYS = {
        'abstract_th': ['บทคัดย่อภาษาไทย', 'บทคัดย่อ'],
        'abstract_en': ['ABSTRACT'],
        'ack': ['กิตติกรรมประกาศ'],
        'toc': ['สารบัญ'],
        'lot': ['สารบัญตาราง'],
        'lof': ['สารบัญภาพ'],
    }
    
    BODY_PATTERNS = {
        'ch1': r'บทที่ 1\\s*\\n\\s*บทนำ',
        'ch1_2': r'1\\.2\\s*วัตถุประสงค์',
        'ch1_3': r'1\\.3\\s*ขอบเขต',
        'ch1_4': r'1\\.4\\s*สมมติฐาน',
        'ch1_5': r'1\\.5\\s*ประโยชน์',
        'ch1_6': r'1\\.6\\s*นิยามศัพท์',
        'ch2': r'บทที่ 2\\s*\\n\\s*วรรณกรรม',
        'ch2_2': r'2\\.2\\s*ทฤษฎีการยศาสตร์',
        'ch2_3': r'2\\.3\\s*สรีรวิทยาของการพัก',
        'ch2_4': r'2\\.4\\s*กลไกการบำบัด',
        'ch2_5': r'2\\.5\\s*ระเบียบวิธีวิจัยรายบุคคล',
        'ch2_6': r'2\\.6\\s*เครื่องมือและเทคโนโลยี',
        'ch2_7': r'2\\.7\\s*การสังเคราะห์วรรณกรรม',
        'ch3': r'บทที่ 3\\s*\\n\\s*วิธีดำเนินการ',
        'ch3_2': r'3\\.2\\s*กลุ่มตัวอย่าง',
        'ch3_3': r'3\\.3\\s*แผนปฏิบัติการ',
        'ch3_4': r'3\\.4\\s*เครื่องมือวิจัย',
        'ch3_5': r'3\\.5\\s*เกณฑ์และการวัด',
        'ch3_6': r'3\\.6\\s*มาตรการด้านความปลอดภัย',
        'ch4': r'บทที่ 4\\s*\\n\\s*ผลการดำเนินงาน',
        'ch4_2': r'4\\.2\\s*ผลการเปลี่ยนแปลงของระดับความปวด',
        'ch4_3': r'4\\.3\\s*ผลการปรับปรุงท่าทาง',
        'ch4_4': r'4\\.4\\s*ผลการปฏิบัติตามวินัย',
        'ch4_5': r'4\\.5\\s*ผลการประเมินความสม่ำเสมอ',
        'ch4_6': r'4\\.6\\s*ผลสัมฤทธิ์ภาพรวม',
        'ch5': r'บทที่ 5\\s*\\n\\s*สรุปผล',
        'ch5_2': r'5\\.2\\s*การอภิปรายผล',
        'ch5_3': r'5\\.3\\s*ปัญหา',
        'ch5_4': r'5\\.4\\s*ข้อเสนอแนะ',
        'ref': r'บรรณานุกรม',
        'app_a': r'ภาคผนวก ก',
        'app_b': r'ภาคผนวก ข',
        'app_c': r'ภาคผนวก ค',
        'app_d': r'ภาคผนวก ง',
        'bio': r'ประวัติผู้จัดทำ',
        'tbl_1_1': r'ตารางที่ 1\\.1',
        'tbl_2_1': r'ตารางที่ 2\\.1',
        'tbl_2_2': r'ตารางที่ 2\\.2',
        'tbl_3_1': r'ตารางที่ 3\\.1',
        'tbl_3_2': r'ตารางที่ 3\\.2',
        'tbl_4_1': r'ตารางที่ 4\\.1',
        'tbl_4_2': r'ตารางที่ 4\\.2',
        'tbl_4_3': r'ตารางที่ 4\\.3',
        'tbl_4_4': r'ตารางที่ 4\\.4',
        'tbl_4_5': r'ตารางที่ 4\\.5',
        'fig_1_1': r'รูปที่ 1\\.1',
        'fig_2_1': r'รูปที่ 2\\.1',
        'fig_3_1': r'รูปที่ 3\\.1',
        'fig_3_2': r'รูปที่ 3\\.2',
        'fig_3_3': r'รูปที่ 3\\.3',
        'fig_3_4': r'รูปที่ 3\\.4',
        'fig_4_1': r'รูปที่ 4\\.1',
        'fig_4_2': r'รูปที่ 4\\.2',
        'fig_4_3': r'รูปที่ 4\\.3',
    }
    
    arabic_start_idx = None
    for i, page in enumerate(doc):
        t = page.get_text()
        if 'บทที่ 1' in t and 'บทนำ' in t and '...' not in t and 'สารบัญ' not in t:
            arabic_start_idx = i
            break
            
    if arabic_start_idx is None:
        arabic_start_idx = 8
        
    thai_letters = ['ก', 'ข', 'ค', 'ง', 'จ', 'ฉ', 'ช', 'ซ', 'ฌ', 'ญ', 'ฎ', 'ฏ']
    
    for i, page in enumerate(doc):
        t = page.get_text()
        if i < arabic_start_idx:
            t_idx = i - 2
            th_char = thai_letters[t_idx] if 0 <= t_idx < len(thai_letters) else str(i)
            for k, kw_list in FRONT_MATTER_KEYS.items():
                if k not in page_map and any(kw in t for kw in kw_list):
                    page_map[k] = th_char
        else:
            arabic_page = str(i - arabic_start_idx + 1)
            for k, pat in BODY_PATTERNS.items():
                if k not in page_map and re.search(pat, t):
                    page_map[k] = arabic_page
                    
    print("Scanned Page Map:", page_map)
    return page_map

def run_two_pass_build():
    print("=== PASS 1: BUILDING DOCX & CONVERTING TO PDF ===")
    generate_report_docx(output_path=DOCX_OUTPUT)
    convert_docx_to_pdf(DOCX_OUTPUT, PDF_OUTPUT)
    
    print("=== SCANNING REAL PDF PAGE COORDINATES ===")
    real_pages = scan_pdf_pages(PDF_OUTPUT)
    
    print("=== PASS 2: REBUILDING DOCX WITH EXACT 100% SYNCHRONIZED TOC ===")
    generate_report_docx(toc_pages=real_pages, output_path=DOCX_OUTPUT)
    convert_docx_to_pdf(DOCX_OUTPUT, PDF_OUTPUT)
    print("=== BUILD COMPLETED SUCCESSFULLY ===")

if __name__ == '__main__':
    run_two_pass_build()
'''

with open('scripts/build_full_report.py', 'w', encoding='utf-8') as f:
    f.write(report_code)

print("Saved updated scripts/build_full_report.py successfully.")
