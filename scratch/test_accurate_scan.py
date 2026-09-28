import fitz
import re
import sys

sys.stdout.reconfigure(encoding='utf-8')

pdf_path = r"00_ไฟล์ส่งงาน_Health30D_OfficeSyndrome/01_เล่มรายงาน_Health30D_OfficeSyndrome_ฉบับสมบูรณ์.pdf"
doc = fitz.open(pdf_path)

thai_letters = ['ก', 'ข', 'ค', 'ง', 'จ', 'ฉ', 'ช', 'ซ', 'ฌ', 'ญ', 'ฎ', 'ฏ']

# Arabic page 1 starts on PDF page 10 (index 9)
arabic_start_idx = 9

FRONT_MATTER_PATTERNS = {
    'abstract_th': r'^\s*บทคัดย่อภาษาไทย',
    'abstract_en': r'^\s*ABSTRACT',
    'ack': r'^\s*กิตติกรรมประกาศ',
    'toc': r'^\s*สารบัญ\s*$',
    'lot': r'^\s*สารบัญตาราง\s*$',
    'lof': r'^\s*สารบัญภาพ\s*$',
}

BODY_PATTERNS = {
    'ch1': r'บทที่ 1\s*\n\s*บทนำ',
    'ch1_2': r'1\.2\s*วัตถุประสงค์',
    'ch1_3': r'1\.3\s*ขอบเขต',
    'ch1_4': r'1\.4\s*สมมติฐาน',
    'ch1_5': r'1\.5\s*ประโยชน์',
    'ch1_6': r'1\.6\s*นิยามศัพท์',
    'ch2': r'บทที่ 2\s*\n\s*วรรณกรรม',
    'ch2_2': r'2\.2\s*ทฤษฎีการยศาสตร์',
    'ch2_3': r'2\.3\s*สรีรวิทยาของการพัก',
    'ch2_4': r'2\.4\s*กลไกการบำบัด',
    'ch2_5': r'2\.5\s*ระเบียบวิธีวิจัยรายบุคคล',
    'ch2_6': r'2\.6\s*เครื่องมือและเทคโนโลยี',
    'ch2_7': r'2\.7\s*การสังเคราะห์วรรณกรรม',
    'ch3': r'บทที่ 3\s*\n\s*วิธีดำเนินการ',
    'ch3_2': r'3\.2\s*กลุ่มตัวอย่าง',
    'ch3_3': r'3\.3\s*แผนปฏิบัติการ',
    'ch3_4': r'3\.4\s*เครื่องมือวิจัย',
    'ch3_5': r'3\.5\s*เกณฑ์และการวัด',
    'ch3_6': r'3\.6\s*มาตรการด้านความปลอดภัย',
    'ch4': r'บทที่ 4\s*\n\s*ผลการดำเนินงาน',
    'ch4_2': r'4\.2\s*ผลการเปลี่ยนแปลงของระดับความปวด',
    'ch4_3': r'4\.3\s*ผลการปรับปรุงท่าทาง',
    'ch4_4': r'4\.4\s*ผลการปฏิบัติตามวินัย',
    'ch4_5': r'4\.5\s*ผลการประเมินความสม่ำเสมอ',
    'ch4_6': r'4\.6\s*ผลสัมฤทธิ์ภาพรวม',
    'ch5': r'บทที่ 5\s*\n\s*สรุปผล',
    'ch5_2': r'5\.2\s*การอภิปรายผล',
    'ch5_3': r'5\.3\s*ปัญหา',
    'ch5_4': r'5\.4\s*ข้อเสนอแนะ',
    'ref': r'^\s*บรรณานุกรม\s*$',
    'app_a': r'^\s*ภาคผนวก ก',
    'app_b': r'^\s*ภาคผนวก ข',
    'app_c': r'^\s*ภาคผนวก ค',
    'app_d': r'^\s*ภาคผนวก ง',
    'bio': r'^\s*ประวัติผู้จัดทำ\s*$',
    'tbl_1_1': r'^\s*ตารางที่ 1\.1',
    'tbl_2_1': r'^\s*ตารางที่ 2\.1',
    'tbl_2_2': r'^\s*ตารางที่ 2\.2',
    'tbl_3_1': r'^\s*ตารางที่ 3\.1',
    'tbl_3_2': r'^\s*ตารางที่ 3\.2',
    'tbl_4_1': r'^\s*ตารางที่ 4\.1',
    'tbl_4_2': r'^\s*ตารางที่ 4\.2',
    'tbl_4_3': r'^\s*ตารางที่ 4\.3',
    'tbl_4_4': r'^\s*ตารางที่ 4\.4',
    'tbl_4_5': r'^\s*ตารางที่ 4\.5',
    'fig_1_1': r'^\s*รูปที่ 1\.1',
    'fig_2_1': r'^\s*รูปที่ 2\.1',
    'fig_3_1': r'^\s*รูปที่ 3\.1',
    'fig_3_2': r'^\s*รูปที่ 3\.2',
    'fig_3_3': r'^\s*รูปที่ 3\.3',
    'fig_3_4': r'^\s*รูปที่ 3\.4',
    'fig_4_1': r'^\s*รูปที่ 4\.1',
    'fig_4_2': r'^\s*รูปที่ 4\.2',
    'fig_4_3': r'^\s*รูปที่ 4\.3',
}

page_map = {}

for i, page in enumerate(doc):
    t = page.get_text()
    if i < arabic_start_idx:
        t_idx = i - 2
        th_char = thai_letters[t_idx] if 0 <= t_idx < len(thai_letters) else str(i)
        for k, pat in FRONT_MATTER_PATTERNS.items():
            if k not in page_map and re.search(pat, t, re.MULTILINE):
                page_map[k] = th_char
    else:
        arabic_page = str(i - arabic_start_idx + 1)
        for k, pat in BODY_PATTERNS.items():
            if k not in page_map and re.search(pat, t, re.MULTILINE):
                page_map[k] = arabic_page

print("Corrected Scanned Page Map:")
for k, v in page_map.items():
    print(f"  {k}: {v}")
