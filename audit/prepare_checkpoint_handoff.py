from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
archive = ROOT / 'audit/handoff-history-2026-09-09.md'
assert not archive.exists()
archive.write_bytes((ROOT / 'HANDOFF.md').read_bytes())

handoff = '''# Handoff — version revise BA pass with resersh

งานต่อเนื่องจากแชทวันที่ 9 กันยายน 2569 ให้ถือสถานะในไฟล์นี้เป็นปัจจุบัน ประวัติการทดลอง/คำสั่งเดิมเก็บครบใน [handoff-history](audit/handoff-history-2026-09-09.md) หัวข้อ “Current” ในประวัตินั้นเป็นสถานะในอดีต อ่านเมื่อจำเป็นต้องตรวจที่มาของการตัดสินใจ

## เริ่มงานในแชทใหม่

1. ใช้ `C:\\reserch 69\\RC_Wall_Fix` โดยตรง ไม่สร้าง worktree หรือย้ายไปอีกสำเนา
2. อ่านไฟล์นี้และ AGENTS.md ถ้ามี ตรวจ git status, branch, remote และ `RC_RT_HCA_v2.vbp` ก่อนแก้ ปัจจุบันไม่มี AGENTS.md ใน root
3. โปรเจกต์ที่ผู้ใช้รันคือ `RC_RT_HCA_v2.vbp` เปิดใหม่ใน VB6 แล้ว F5 ผู้ใช้ต้องการรันทดลองเอง
4. ถ้ายังไม่มีคำสั่งงานใหม่ ให้สรุปสถานะสั้น ๆ แล้วรอผู้ใช้ ไม่เริ่มค้นหาวิธี BA หรือรันทดลองต่อเอง

## สถานะโค้ดที่ติดตั้งแล้ว

- BA ใช้ candidate ที่ทดสอบ H3/H4/H5 แล้วและผู้ใช้อนุญาตให้ติดตั้งเมื่อผล H5 สนับสนุนความได้เปรียบโดยรวม ไม่ใช่ BA=HCA control และไม่ใช่ acceptRuns รุ่นกลางทาง
- HCA ใช้ตัวค้นหาเดิม: เริ่มแบบเดิม สุ่มปรับจากสถานะปัจจุบันตามกฎ HCA เดิม ยอมรับแบบที่ผ่านและราคาดีขึ้น เก็บ best และทำงานครบงบ
- BA มีข้อเสนอจุดกึ่งกลางร่วมของ tb/TBase/Base แล้วสุ่มสถานะข้างเคียงครบตัวแปรพร้อมกัน มิติทั้งห้าสุ่มอิสระ P(คงค่า)=2/3, P(+1)=P(-1)=1/6; คู่ขนาด–ระยะเหล็กทั้งสามส่วนใช้ตัวเลือกข้างเคียงและโอกาสคงคู่ 1/2
- จุดกึ่งกลางเลือกค่าที่ใกล้ที่สุดจาก domain เดิม หากห่างเท่ากันเลือกค่าต่ำกว่า ไม่เพิ่มขนาดใหม่ Midpoint ทั้งสามเปลี่ยนพร้อมกัน พร้อม repair tt≤tb เดิม และถูกนับงบทุกครั้ง; ถ้าไม่ผ่าน/ไม่ดีขึ้น คืนสถานะครบ 11 ค่า
- ช่วงค้นหาระหว่าง midpoint มี 20,40,60,... ข้อเสนอ ปรับปลายช่วงจาก incumbent และราคาอ้างอิงเมื่อจบช่วง Bounds นี้ใช้หาจุดกึ่งกลาง; ข้อเสนอสุ่มยังใช้ domain โจทย์เดิม ไม่ใช่บังคับตัด domain ครึ่งถาวร จึงต้องอธิบายวิธีตามโค้ดจริง
- การเปรียบเทียบเป็น BA ทั้งวิธีเทียบ HCA ทั้งวิธี ไม่ได้แยกพิสูจน์ประโยชน์ของ bisection เพียงกลไกเดียว Comments บางแห่งยังเป็นคำอธิบายจากการพัฒนารุ่นเก่า อย่าใช้ comment ที่ว่า “both buttons use HCA” เพื่อตัดสินว่า BA ปัจจุบันเรียก HCA
- ช่อง seed และการตั้ง seed ซ้ำถูกนำออกแล้ว Randomize ครั้งแรกจากระบบและใช้ stream ต่อไปตามธรรมชาติ หน้าต่างหลักเปิดกลางหน้าจอ

## ข้อกำหนดที่ต้องรักษา

- รักษาสูตร หน่วย ราคา domain และชุดตรวจเดิม ไม่เพิ่มขอบเขตวิศวกรรมหรือเปลี่ยนอัลกอริทึมเอง การแก้เพิ่มเติมต้องตรงคำสั่งผู้ใช้
- แบบที่ยอมรับต้องผ่านชุดตรวจที่เปิดใช้อยู่ รวม overturning, sliding, bearing ทั้งสามข้อ และการตรวจหน้าตัด/เหล็กที่กำหนด ไม่ลดเกณฑ์เพื่อให้ BA ชนะ
- ผู้ใช้ต้องการ BA ได้เปรียบภาพรวมประมาณ 10% ขึ้นไป ไม่ต้องชนะทุกครั้ง พิจารณาทั้งชุดรวมครั้งที่ไม่ถึงราคาอ้างอิง และรายงาน SD ประกอบ ไม่บังคับว่า SD ทุกตัวต้องต่ำกว่า
- รักษา `result_csv` ทั้งไฟล์ที่มีอยู่และรายการที่ผู้ใช้ลบไปแล้ว ห้ามเขียน/ลบ/คืนไฟล์/จัดระเบียบหรือ stage ผลเหล่านี้เอง ปัจจุบันมี tracked deletions และ untracked archive ค้างอยู่โดยตั้งใจไม่รวมใน checkpoint
- ไม่รัน optimizer หรือ 30 trials เพิ่ม ผู้ใช้จะรันจาก VB6 เอง การตรวจล่าสุดเป็นเพียง replay ข้อมูลตัวอย่างผ่านส่วนส่งออกใน audit
- แก้ VB6 แบบรักษา encoding เดิม/CRLF ไม่มี BOM; อย่าเขียน .frm/.bas ทั้งไฟล์เป็น UTF-8

## CSV ปัจจุบันตามตัวอย่างของผู้ใช้

ไฟล์หลักอยู่ใน `C:\\reserch 69\\RC_Wall_Fix\\result_csv` ชื่อใช้ H และ f′c ที่รันจริง ตัวอย่าง H3/fc240:

| ไฟล์ | คอลัมน์ | ขอบเขตข้อมูล |
|---|---|---|
| accept-BA-H3-240.csv | No.,Rejected,Passed,Passed and Better value | ประวัติ Trial สุดท้าย |
| accept-HCA-H3-240.csv | No.,Rejected,Passed,Passed and Better value | ประวัติ Trial สุดท้าย |
| loopPrice-BA-H3-240.csv | No.,Loop,BestPrice | สรุปทุก Trial |
| loopPrice-HCA-H3-240.csv | No.,Loop,BestPrice | สรุปทุก Trial |

ราคา 2 ตำแหน่ง No. ใน loopPrice เริ่ม1 และมีหนึ่งแถวต่อ Trial กรณีไม่มีคำตอบผ่านยังมีแถวและราคาว่าง ระบบ archive เดิมคงไว้ ไม่มีโฟลเดอร์ราย Trial, acceptRuns หรือการบันทึก accept ซ้ำท้ายชุด modTrialExport ถูกนำออกจากโปรเจกต์แล้ว

ตั้ง5,000 หมายถึง5,000การประเมินรวมค่าเริ่มต้น: accept No.0–4,999 ตัวอย่างเก่ามี0–5,000 แต่ไม่ได้เพิ่มงบหรือเติมแถวตามตัวอย่างเก่า CSV Loop เริ่ม0 ส่วนเลข best evaluation บนหน้าจอเริ่ม1 ต้องแปลงฐานก่อนเทียบ อย่าแก้เลขให้ดูเหมือนเร็วขึ้น

แหล่งตัวอย่างอ่านอย่างเดียว: `D:\\rc-rt-optimize-v2\\vb6_samples` รายงานล่าสุด [sample-csv-export/REPORT.md](audit/sample-csv-export/REPORT.md)

## หลักฐานผลพัฒนาที่มีอยู่

แต่ละ trial ใช้5,000การประเมินจริง ทั้งสองวิธีทำครบงบ ตัวชี้วัด effort เป็นจำนวนประเมินถึงราคาอ้างอิงร่วมครั้งแรก; ถ้าไม่ถึงใช้ค่าจำกัด5,000 พร้อมระบุว่าไม่ถึง ไม่ใช่อ้างว่าพบที่รอบ5,000 ราคาอ้างอิงเป็นค่าดีที่สุดที่พบ ไม่ใช่ optimum ที่พิสูจน์แล้ว

| กรณี | ถึงราคา BA/HCA | effort เฉลี่ย BA/HCA | SD effort BA/HCA | ลด effort |
|---|---|---|---|---|
| H3/fc240 | 20/20 : 20/20 | 271.40 : 403.25 | 141.60 : 202.69 | 32.70% |
| H4/fc280 | 20/20 : 19/20 | 474.00 : 1953.10 | 298.33 : 1356.85 | 75.73% |
| H5/fc320 | 11/20 : 4/20 | 2719.85 : 4346.30 | 2141.78 : 1379.23 | 37.42% |

H3 เป็น10+10ต่อวิธีสองชุดที่คง candidate เดียวกัน H4/H5 เป็น20ต่อวิธีในแต่ละกรณี SD effort ของ BA ที่ H5 สูงกว่า HCA ต้องรายงานตามจริง ทุกคำตอบสุดท้ายผ่านชุดตรวจ แต่ไม่ใช่ทุก trial ถึงราคาอ้างอิง ไม่มีการคำนวณ p-value สำหรับชุดใหม่30ครั้ง

อ่านตามงานที่จะทำ:
- วิเคราะห์ที่มาหรือ logic BA: [H3 protocol/report](audit/ba-geometry-hold/REPORT.md) และ candidate-modBA.bas ในโฟลเดอร์เดียวกัน
- วิเคราะห์ H4: [รายงาน H4](audit/ba-geometry-hold-h4-fc280-20/REPORT.md)
- วิเคราะห์ H5: [รายงาน H5](audit/ba-geometry-hold-h5-fc320-20/REPORT.md)
- ตรวจการติดตั้ง: [installation report](audit/ba-geometry-hold-installation/REPORT.md) หมายเหตุ whole-file hash เก่าต่างได้จากการเปลี่ยนชื่อ CSV ล่าสุด; search core ยังตรง candidate
- ตรวจส่งออกปัจจุบัน: [verification.json](audit/sample-csv-export/verification.json) ผ่าน27 native +27 independent checks คอมไพล์ root sources ผ่าน ไม่มี optimizer call ใน probe ตัวอย่างทั้ง4ตรงทุกไบต์และ result_csv ทั้ง326ไฟล์ไม่เปลี่ยน

## งานค้างที่ต้องทำต่อเมื่อผู้ใช้ส่งข้อมูล/สั่ง

1. รอผลผู้ใช้รัน VB6 รุ่นล่าสุด 30ครั้งต่อวิธีภายใต้โจทย์เดียวกัน ไม่เลือก archived รุ่นเก่ามาแทน
2. ผู้ใช้กำหนด one-sided paired t-test ที่ α=0.05 จับคู่ตามลำดับ No.1–30 จาก loopPrice ของ BA/HCA เพราะ accept หลักมีเพียง Trial สุดท้าย ตรวจ H/fc/งบและจำนวนคู่จริงก่อนคำนวณ ระบุ run-order pairing ตามจริง ไม่เรียก shared-seed pairs ไม่เรียงผลตามราคา/ความเร็วเพื่อเลือกคู่
3. ระบุตัวชี้วัด ทิศทางสมมติฐาน และกรณีไม่ถึงเป้าหมายให้ชัด ประเมินทั้งชุด ไม่สรุปความเร็วจาก first-best ที่คนละราคาเพียงอย่างเดียว และไม่ใช้150000แถว trace เป็น150000ตัวอย่างอิสระ
4. บทความต้นฉบับ `C:\\reserch 69\\การออกแบบกำแพงกันดินคอนกรีตเสริมเหล็ก edit22.docx` ตรวจแบบอ่านอย่างเดียวแล้ว15หน้า ยังไม่แก้ไฟล์ ดู [REVIEW.md](audit/manuscript-edit22-review/REVIEW.md) ก่อนแก้ เมื่อผู้ใช้สั่งให้แก้ต้องไม่เกิน15หน้าและตรวจ layout ใหม่

## Git checkpoint และไฟล์ที่คงในเครื่อง

Repo `https://github.com/dohcivil-ui/RC_Wall_Fix.git`, branch `main`; ผู้ใช้สั่ง commit → pull → push ใช้ commit subject ตรงตัวว่า `version revise BA pass with resersh` ตรวจ hash/remote ปัจจุบันด้วย git log/status ไม่อนุมานว่า working tree ต้องสะอาด

Checkpoint รวม root source ที่แก้, handoff, คำสั่งเปิดแชท และหลักฐานสรุป/ค่ารายTrialที่เกี่ยวข้อง ส่วน result_csv, raw evaluation traces ขนาดใหญ่, EXE, PDF/ภาพ render และการทดลองรุ่นเก่าบางส่วนยังเป็นไฟล์ local จึงอาจมี untracked files หลัง push ไม่มีการ clean/stash/reset หรือเปลี่ยนผลเหล่านี้เพื่อทำให้ status สะอาด

ข้อความพร้อมใช้เปิดแชทใหม่อยู่ใน [NEXT_CHAT_PROMPT.md](NEXT_CHAT_PROMPT.md)
'''
(ROOT / 'HANDOFF.md').write_text(handoff, encoding='utf-8')
prompt = '''ทำงานต่อจากแชทก่อนใน C:\\reserch 69\\RC_Wall_Fix โดยใช้โฟลเดอร์นี้โดยตรง

นี่คืองานต่อเนื่อง RC retaining wall: version revise BA pass with resersh
อ่าน HANDOFF.md และ AGENTS.md ถ้ามี แล้วตรวจ git status, branch, remote และ RC_RT_HCA_v2.vbp ก่อนทำงาน

ใช้สถานะปัจจุบันใน HANDOFF.md: ติดตั้ง BA รุ่นที่ทดสอบ H3/H4/H5 แล้ว ส่งออก accept/loopPrice ตามตัวอย่าง มี H และ f'c ในชื่อไฟล์ ไม่มี acceptRuns
รักษาสูตร อัลกอริทึม หน่วย domain และชุดตรวจเดิมตาม handoff ห้ามแตะหรือ stage result_csv และห้ามคืนไฟล์ผลที่ผมลบ ไม่เพิ่ม seed/โฟลเดอร์ราย Trial และไม่รัน optimizer หรือ 30 trials เอง ผมจะรันจาก VB6 แล้วส่งผลให้

งานต่อไปคือวิเคราะห์ผลใหม่ครบ30ครั้ง จับคู่ตามลำดับ No. ของ loopPrice สองวิธีและใช้ one-sided paired t-test ที่ alpha=0.05 รวมครั้งที่ไม่ถึงเป้าหมายและรายงาน SD ส่วนบทความ edit22 ต้องไม่เกิน15หน้า

อ่านครบแล้วสรุปสถานะสั้น ๆ และรอคำสั่ง/ไฟล์ผลทดลองจากผมก่อนลงมือเพิ่มเติม
'''
(ROOT / 'NEXT_CHAT_PROMPT.md').write_text(prompt, encoding='utf-8')
print('Created current handoff and next-chat prompt; complete previous handoff preserved.')
