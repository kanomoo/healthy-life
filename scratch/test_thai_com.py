import docx
import win32com.client, fitz, os, sys

sys.stdout.reconfigure(encoding='utf-8')

doc = docx.Document()
s1 = doc.sections[0]
p1 = doc.add_paragraph('Content 1')
doc.add_page_break()
p2 = doc.add_paragraph('Content 2')

s2 = doc.add_section(docx.enum.section.WD_SECTION.NEW_PAGE)
p3 = doc.add_paragraph('Content 3')

doc.save('scratch/test_thai_com.docx')
word = win32com.client.Dispatch('Word.Application')
wdoc = word.Documents.Open(os.path.abspath('scratch/test_thai_com.docx'))

# Set page numbering via Word COM
# wdPageNumberStyleThaiLetter = 54 (or let's check wdPageNumberStyle constants)
# Or let's inspect wdoc.Sections(1).Headers(1).PageNumbers
sec1 = wdoc.Sections(1)
pn1 = sec1.Headers(1).PageNumbers
print('Default NumberStyle:', pn1.NumberStyle)
pn1.NumberStyle = 54 # Thai letters
pn1.RestartNumberingAtSection = True
pn1.StartingNumber = 1
pn1.Add(PageNumberAlignment=2) # 2 = wdAlignPageNumberRight

sec2 = wdoc.Sections(2)
pn2 = sec2.Headers(1).PageNumbers
pn2.NumberStyle = 54
pn2.RestartNumberingAtSection = False
pn2.Add(PageNumberAlignment=2)

wdoc.SaveAs(os.path.abspath('scratch/test_thai_com.docx'))
wdoc.SaveAs(os.path.abspath('scratch/test_thai_com.pdf'), FileFormat=17)
wdoc.Close(False)
word.Quit()

fdoc = fitz.open('scratch/test_thai_com.pdf')
for i, page in enumerate(fdoc):
    print(f'Page {i+1}: {repr(page.get_text().strip())}')
