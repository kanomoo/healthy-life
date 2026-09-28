import docx
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import win32com.client, fitz, os, sys

sys.stdout.reconfigure(encoding='utf-8')

doc = docx.Document()
s1 = doc.sections[0]
s1.header_distance = docx.shared.Inches(0.5)
pnt1 = parse_xml(f'<w:pgNumType {nsdecls("w")} w:fmt="thaiLetters" w:start="1"/>')
s1._sectPr.append(pnt1)
p1 = doc.add_paragraph('Content 1')
doc.add_page_break()
p2 = doc.add_paragraph('Content 2')

s2 = doc.add_section(docx.enum.section.WD_SECTION.NEW_PAGE)
# If we omit pgNumType, does it continue thaiLetters? Let's check!
p3 = doc.add_paragraph('Content 3')

# add header ONLY to s1, s2 inherits it
h = s1.header.paragraphs[0]
r = h.add_run()
r._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w')))
r._r.append(parse_xml(r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls('w')))
r._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w')))
r._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w')))

doc.save('scratch/test_thai_pg2.docx')
word = win32com.client.Dispatch('Word.Application')
wdoc = word.Documents.Open(os.path.abspath('scratch/test_thai_pg2.docx'))
wdoc.SaveAs(os.path.abspath('scratch/test_thai_pg2.pdf'), FileFormat=17)
wdoc.Close(False)
word.Quit()

fdoc = fitz.open('scratch/test_thai_pg2.pdf')
for i, page in enumerate(fdoc):
    print(f'Page {i+1}: {repr(page.get_text().strip())}')
