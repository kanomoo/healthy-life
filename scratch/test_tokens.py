from pythainlp import word_tokenize
import re, sys

sys.stdout.reconfigure(encoding='utf-8')

s1 = "โครงสร้างและการจัดวางสูตรคำนวณอัตโนมัติในไฟล์สมุดบันทึกสุขภาพ '6806021612037_Health30D.xlsx' บน Microsoft OneDrive ประกอบด้วย:"
print('Tokens for s1:', word_tokenize(s1, engine='newmm'))
s2 = "นักศึกษาชั้นปีที่ 1 แขนงวิชาเทคโนโลยีสารสนเทศ (IT)\nภาควิชาเทคโนโลยีสารสนเทศ คณะเทคโนโลยีและการจัดการอุตสาหกรรม\nมหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ วิทยาเขตปราจีนบุรี"
print('Tokens for s2 line 2:', word_tokenize(s2.split('\n')[1], engine='newmm'))
