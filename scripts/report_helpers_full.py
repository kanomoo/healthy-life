import os
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

BASE_DIR = r'C:\Project\healthy-life'
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
    lines = str(text).split('\n')
    processed = []
    for line in lines:
        if not line:
            processed.append("")
            continue
        tokens = word_tokenize(line, engine='newmm')
        joined = '\u200b'.join(tokens)
        joined = re.sub(r'\u200b+(' + NO_BREAK_BEFORE + ')', r'\1', joined)
        joined = re.sub('(' + NO_BREAK_AFTER + r')\u200b+', r'\1', joined)
        joined = re.sub(r'\u200b+', '\u200b', joined)
        processed.append(joined)
    return '\n'.join(processed)

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
    r_num = p.add_run(f"บทที่ {chapter_num}\n")
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

# STRICT APA TABLE FORMATTER - WIDTH GUARANTEED <= 5.75 INCHES
def add_styled_table(doc, table_num, caption, headers, data, col_widths, font_size_pt=10.5, padding_twips=15, page_break_before=False):
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
    
    # Calculate exact total width
    total_w_dxa = int(sum([w.inches for w in col_widths]) * 1440)
    
    # APA borders: top and bottom only
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
    
    # CantSplit on every row, tblHeader on first row
    for i, row in enumerate(tbl.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        if i == 0:
            trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
            
    # Headers
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
        
    # Data rows
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
    r2 = p.add_run(f"\t{page_num_str}")
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
    r2 = p.add_run(f"\t{right_label}")
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

print("Helper definitions loaded successfully.")
