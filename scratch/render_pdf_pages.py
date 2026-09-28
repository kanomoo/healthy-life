import fitz # PyMuPDF
import os

pdf_path = r"00_ไฟล์ส่งงาน_Health30D_OfficeSyndrome/01_เล่มรายงาน_Health30D_OfficeSyndrome_ฉบับสมบูรณ์.pdf"
doc = fitz.open(pdf_path)

out_dir = r"scratch/rendered_pages"
os.makedirs(out_dir, exist_ok=True)

print(f"Total Pages in PDF: {len(doc)}")

# Render cover (page 0), table 3.1 (around page 13-15), table 4.2 (around page 20-22), table 4.3 (around page 22-24)
pages_to_render = [0, 1, 2, 3, 4, 5, 6, 13, 14, 15, 16, 20, 21, 22, 23, 24]

for p_num in pages_to_render:
    if p_num < len(doc):
        page = doc[p_num]
        pix = page.get_pixmap(dpi=150)
        img_path = os.path.join(out_dir, f"page_{p_num+1:02d}.png")
        pix.save(img_path)
        print(f"Saved: {img_path} (Page {p_num+1})")
