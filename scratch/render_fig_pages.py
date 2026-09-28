import fitz
import os

pdf_path = r"00_ไฟล์ส่งงาน_Health30D_OfficeSyndrome/01_เล่มรายงาน_Health30D_OfficeSyndrome_ฉบับสมบูรณ์.pdf"
doc = fitz.open(pdf_path)

out_dir = r"scratch/rendered_pages"
os.makedirs(out_dir, exist_ok=True)

for p_num in [24, 26]: # 0-indexed for pages 25 and 27
    page = doc[p_num]
    pix = page.get_pixmap(dpi=150)
    img_path = os.path.join(out_dir, f"page_{p_num+1:02d}.png")
    pix.save(img_path)
    print(f"Saved: {img_path} (Page {p_num+1})")
