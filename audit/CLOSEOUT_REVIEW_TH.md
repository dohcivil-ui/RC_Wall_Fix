# รายงานตรวจปิดงาน — 9 กันยายน 2026

ตรวจโปรเจกต์สำเนา `C:\reserch 69\RC_Wall_Fix` ตามไฟล์อ้างอิงใน `RC_RT_HCA_v2.vbp` ทั้ง 11 ไฟล์ เก็บ SHA256 ของซอร์สและ EXE ที่คอมไพล์ล่าสุดใน [closeout-manifest.json](closeout-manifest.json) ไม่พบ AGENTS.md หรือ JIT_code_audit.md ตามรายการตรวจตั้งต้น

ผลนี้เป็นการตรวจซอฟต์แวร์และรายการคำนวณที่ระบุ ไม่ใช่การพิสูจน์ว่าโค้ดปราศจากข้อผิดพลาดทุกกรณี หรือรับรองครบทุกข้อของ วสท. 2562/ACI ฐานเกณฑ์ปัจจุบันคือ [PROJECT_WSD_ACI99_V3_MAIN_ONLY](PROJECT_CHECKS_IMPLEMENTATION_TH.md)

## ขอบเขตที่ตรวจ

| ไฟล์ active | สิ่งที่ตรวจ |
|---|---|
| Form1.frm | รับข้อมูลและหน่วย, seed, ปุ่ม BA/HCA, เก็บผู้ชนะทั้ง session, Compare, NO_SOLUTION, เรียก popup หลังจบ loop ครั้งเดียว |
| frmBestDesign.frm | มิติจากคำตอบผู้ชนะ, ตำแหน่งเหล็กหลักและ DB/spacing, ราคา/วิธี/trial/evaluation, ไม่เปลี่ยนค่าคำนวณหรือสถานะสุ่ม |
| modShared.bas | เรขาคณิต น้ำหนัก centroid แรง active/passive สมดุลฐาน หน่วย การส่งเข้าตัวตรวจ ราคา seed evaluation และไฟล์ผล |
| modProjectChecks.bas | เสถียรภาพทั้งสามและ full contact, โมเมนต์/เฉือน toe-heel, หน่วยแรงตลอด stem, เหล็กจริง cover/spacing/min/max, ปริมาณและรายงานจากตัวตรวจเดียวกัน |
| modWSD.bas | n=Es/Ec, k/j/R ของหน้าตัดสมดุล, cracked-section ตาม As จริง, effective depth และแยกเกณฑ์ทดสอบเก่าออกจาก production |
| modBA.bas | บีบเฉพาะ tb/TBase/Base, recovery, initial/reset/neighbor ทุกทางเข้าผ่านตัวนับและ global best |
| modHillClimbing.bas | ใช้งบและตัวนับเดียวกับ BA, seed และเก็บคำตอบที่ผ่าน |
| modDataStructures.bas | Design/material และตัวแปรร่วม; helper ราคาดั้งเดิมที่ไม่มี caller ไม่ใช่แหล่งราคาที่ production ใช้ |
| modGraphing.bas | ใช้ประวัติคำตอบผู้ชนะและแกน evaluation |
| modEnhancedValidation.bas | ปุ่มจริงใช้ Form1.ValidateInputs; เอาข้อความ helper เก่าที่อ้างว่าผ่านทั้งที่ไม่ได้ตรวจออก |
| modBatch.bas | อ่านเส้นทางเรียก optimizer/CSV; แก้ Timer ติดลบเมื่อข้ามเที่ยงคืนหนึ่งครั้ง แต่ไม่ได้รัน matrix วิจัยหรือรับรองผลสถิติของ batch |

modPDFReport, modExport, modSettings, modFontHelper และไฟล์สำรองไม่อยู่ใน `.vbp` นี้ จึงไม่อ้างว่ารายงานหรือเส้นทางของไฟล์เหล่านั้นได้รับการทดสอบแล้ว

## สิ่งที่แก้ในรอบปิดงาน

- Compare เดิมตัดสินจาก evaluation อย่างเดียว แม้คำตอบแพงกว่า: แก้เป็นราคาต่ำสุดก่อน เมื่อราคาเท่ากันจึงใช้ evaluation แรกที่พบราคานั้น แสดงเป็นผลของ session ที่สังเกตได้
- เก็บชุดข้อมูล/วัสดุ/seed/จำนวน trials/งบของแต่ละ session และงดจัดอันดับถ้าเงื่อนไขไม่ตรงกัน ไม่แสดง sentinel เป็นราคาเมื่อ NO_SOLUTION
- ล้างกราฟก่อนวาดผล session ใหม่ เพื่อไม่คงรูปคำตอบเก่าเมื่อรอบใหม่ไม่พบคำตอบ
- อ่าน seed ตั้งต้นเพียงครั้งเดียวต่อการกดปุ่ม ป้องกันการแก้ textbox ระหว่าง DoEvents เปลี่ยน seed ของ trials ที่เหลือ เปลี่ยนป้ายแกนกราฟเป็น Evaluation
- helper validation เก่าที่ไม่มี caller จะบอกตามจริงว่าไม่ได้ตรวจข้อมูล; ไม่แสดง All inputs are valid โดยไม่มีการตรวจ
- Timer ของ batch รองรับการข้ามเที่ยงคืนหนึ่งครั้ง (ระยะรันน้อยกว่า 24 ชั่วโมง); ไม่ได้ทดสอบ matrix ข้ามวันจริง ไม่แก้สูตรหรือราคาเพื่อรักษาคำตอบเก่า
- ติดป้ายรายงานตรวจตั้งต้นให้เป็นประวัติ แยกจากเกณฑ์ production ปัจจุบัน

## หลักฐาน

- [คอมไพล์โปรเจกต์หลักล่าสุด](closeout-checks/20260909/compile-main.log): VB6 compiler รายงาน succeeded หลังแก้ไฟล์ active ทั้งหมดในรอบนี้
- [Native regression](checks/20260909-000004-601/native-regression.txt): 267 checks, failures=0 เป็นการตรวจสูตร/อัลกอริทึม รวมโหมดเก่าและข้อมูลจำลองที่ระบุชัด
- [GUI regression](checks/20260909-000004-601/gui-regression.txt): failures=0; กดปุ่มบน Form1 ที่คอมไพล์จริง ทดสอบ Compare เพิ่ม 7 กรณี ได้แก่ราคาถูกกว่าแต่ช้ากว่า, ราคาเท่ากัน, evaluation เท่ากัน และ no-solution รวมทั้งปุ่ม Compare เมื่อทั้งสองไม่พบคำตอบ และเมื่อ seed ต่างกัน
- [Production regression](project-checks/20260909-000004-335/project-regression.txt): 47 checks, failures=0; H=3/4/5 ทั้งกรณีรับและปฏิเสธ พร้อม BA/HCA งบสั้นและปุ่มจริง
- [คำนวณอิสระ](project-checks/20260909-000004-335/independent-project.json): 300 ค่าตรงกับผล VB6 การคำนวณอิสระใช้ตรวจเลข ไม่ได้ใช้แทนหลักฐาน VB6
- [Popup regression](sketch-checks/20260909-000043-512/sketch-regression.txt): 19 checks, failures=0; ทดลอง BA หนึ่ง trial 64 evaluations ภาพ H3/H4/H5/no-solution และ [ภาพจาก VB6 จริง](sketch-checks/20260909-000043-512/popup-BA.png) ตรวจด้วยตาแล้ว ป้ายมิติ/เหล็ก/ราคาครบ
- หลักฐานก่อนหน้านี้ [25 popup checks](sketch-checks/20260908-234556-842/sketch-regression.txt) ตรวจทั้ง BA/HCA แบบ 2 trials ว่าหยิบผู้ชนะ trial แรกมาแสดง แม้ trial สุดท้ายเป็น trial 2; รอบปิดงานไม่ได้เปลี่ยนตรรกะเลือกผู้ชนะหรือสูตรวาดภาพนั้น
- [Source/trace checks](source-checks.txt): neighbor เดิมไม่เปลี่ยน, BA มี bounds เพียง 3 ตัว, ไฟล์ active อยู่ในสำเนา, ไฟล์ Downloads เดิมไม่เปลี่ยน, CRLF/encoding ของ VB6 คงเดิม และ 526 traces มีเลข evaluation ต่อเนื่อง ครบงบและ prefix-best ตรงกับคำตอบที่ผ่านทุกทางเข้า

Native suites ข้างต้นรันหลังแก้ Form1/กราฟ/helper validation แล้ว การแก้สุดท้ายหลัง suites มีเฉพาะการชดเชย Timer และข้อความ comment ใน modBatch ซึ่งอยู่ในโปรเจกต์หลักที่คอมไพล์รอบสุดท้าย ไม่ได้อ้างว่ารัน batch แล้ว

กรณีอ้างอิงยืนยัน hs=4.70 ม., active moment=10.3823 และ full-passive net moment=9.7262 tf·m/m หน้าตัด stem .20 ม., clear cover .075 ม., DB12@.25 ถูกปฏิเสธ

ตัวอย่าง BA smoke หนึ่ง trial พบราคา 10,542.98 บาท/ม. ที่ evaluation 61: tt=.25, tb=.35, TBase=.45, B=3.00, toe=.80, heel=1.85 ม.; stem DB25@.20, toe DB28@.25, heel DB28@.20 เป็นเพียงคำตอบที่พบในงบสั้น ไม่ใช่ราคาต่ำสุดสากลหรือผลทดลอง 30 ครั้ง

## งานที่ยังไม่ได้รันและขอบเขตยืนยัน

ไม่ได้รันชุด 30 trials ของรุ่นปัจจุบัน หรือ batch วิจัย ผู้ใช้รันเองตาม [วิธีรัน](PROJECT_CHECKS_IMPLEMENTATION_TH.md#ผู้ใช้รัน-30-ครั้ง) เมื่อครบทั้ง session จะแสดงภาพผู้ชนะเพียงครั้งเดียว

ราคาเป็นคอนกรีตโครงสร้างทั้ง stem และฐาน รวมเฉพาะเหล็กหลักสามส่วนตาม DB/spacing ที่เลือก ไม่รวมเหล็กประกอบ ฝังยึด ทาบ หรือไม้แบบตามขอบเขตที่สั่ง ไม่ใช่ BOQ ครบสำหรับก่อสร้าง ยังไม่ได้ตรวจข้อกำหนดการโก่งตัว การทรุดตัว น้ำ แผ่นดินไหว รอยต่อ/ปลายกำแพง หรือรับรองมาตรฐานฉบับเต็ม ต้องยืนยันส่วนเหล่านี้แยกก่อนนำไปใช้เป็นแบบก่อสร้าง
