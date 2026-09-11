ทำงานต่อใน C:\reserch 69\RC_Wall_Fix โดยใช้โฟลเดอร์นี้โดยตรง

Checkpoint โค้ดที่พร้อมให้ผู้ใช้รัน: 5cd5aca (5cd5acab710ff0586ad911a431cc17d7c520ed0b) บน main และ push ไป origin แล้ว เอกสาร handoff อาจมี commit ใหม่กว่านี้ ให้ตรวจ git log ก่อนงาน รุ่น d84d218 / e73c024 / cd31cc8 เป็นประวัติเก่า ห้าม checkout หรือคืนไฟล์จากรุ่นเหล่านั้นเพื่อเริ่มแชทใหม่

อัปเดตการส่งออกล่าสุด: ผู้ใช้ต้องการเพียง 6 ไฟล์ต่อกรณี H/fc เมื่อรัน BA และ HCA: accept-BA/HCA-H{H}-{fc}.csv, loopPrice-BA/HCA-H{H}-{fc}.csv และ design-BA/HCA-H{H}-{fc}.bmp ใช้ชื่อเดิมทับได้ ไม่มี selectedTrial CSV หรือ convergence BMP เพิ่ม กราฟยังแสดงในโปรแกรม BMP เป็นภาพหน้าตัดของ Trial ที่เลือกราคาต่ำสุดก่อน/รอบน้อยสุดเมื่อราคาเท่ากัน แสดง Trial, ราคา และ Loop=bestEvaluation-1 ให้ตรงกับ CSV ที่เริ่ม No.0 ไม่เปลี่ยนตัวนับภายในหรือวิธีเลือก Trial

ทดสอบเฉพาะการส่งออกด้วย fixtures และภาพทั้ง H3/H4/H5 ผ่าน54ข้อ รวมเลือก accept ที่ดีที่สุดซึ่งไม่ใช่ Trial สุดท้าย เก็บ loopPrice ทุก Trial และรันซ้ำเขียนทับ CSV/BMP ชื่อเดิม ตรวจได้4CSV+2BMPต่อกรณี คอมไพล์GUIผ่าน ไม่มี optimizer runs ผลผู้ใช้เดิมไม่เปลี่ยน ไฟล์ selectedTrial/convergence เก่าถ้ามีไม่ได้ลบ ดู audit/minimal-result-export-20260911/verified/runtime.txt และ verified-gui/compile.log

สถานะล่าสุด 11 กันยายน 2569: ผู้ใช้เปลี่ยนคำสั่งเป็น “แบบเริ่มต้นร่วมที่ตรวจผ่านจริง” เพื่อให้กราฟ BA/HCA เริ่มที่ No.0 ราคาเดียวกัน คำสั่งนี้แทนการคง max/Reject ใน cd31cc8 อย่าคืน max หรือเติมเส้น/ราคา Reject ลงกราฟ

ค่าเริ่มต้นร่วมทุกความสูง: tt=.40, tb=.60, TBase=.50, B=3.00, toe=1.10, heel=1.30 เมตร; stem/toe DB20@.125, heel DB16@.15 อยู่ใน catalogue เดิมทั้งหมด ใช้ SetCommonInitialIndices ใน modShared.bas ทั้ง BA และ HCA ไม่มีการสุ่มหรือค้นหาเพื่อเตรียมจุดเริ่มต้น

RequireFeasibleInitial ตรวจแบบอ้างอิงนี้กับ inputs ปัจจุบันผ่านชุดตรวจเดิมก่อน EvaluateCandidate(initial) ถ้าไม่ผ่านให้แจ้งเหตุและหยุดก่อน No.0 โดยไม่เขียนทับผลเดิม ถ้าผ่าน EvaluateCandidate จะตรวจและบันทึกราคาจริงใน Passed and Better value ที่ No.0 ซึ่งนับเป็น evaluation 1 จากงบเดิม ไม่บังคับสถานะ Passed และไม่แก้สูตร หน่วย ขอบเขต กลไกค้นหา/ยอมรับ หรือ RNG

ทดสอบ H3/fc240, H4/fc280, H5/fc320 ด้วย H1=1.2, gamma_soil=1.8, gamma_concrete=2.4, phi=30, mu=.6, qa=30, cover=.075, SD40: แบบผ่านจริงทั้งสาม ราคาเริ่ม BA=HCA คือ 8,626.89 / 10,621.25 / 12,705.85 บาท/ม. ตรวจซ้ำข้ามวิธีและ Trial ด้วยงบ 1 รวม 12 initial-only calls ไม่มี neighbor/midpoint; รวม regression 109 ข้อผ่านและคอมไพล์ GUI ผ่าน ค่า input อื่นต้องตรวจจริง ไม่อ้างว่าผ่านทุกเงื่อนไข

ผู้ใช้เปิด RC_RT_HCA_v2.vbp ใหม่แล้ว F5 เอง ยังไม่ได้รัน 30 trials จากจุดตั้งต้นใหม่ และไม่เปลี่ยน EXE เดิม ผลชนะเก่า H3=26/30 H4=24/30 H5=21/30 เป็นรุ่นเริ่ม max ไม่ใช่ผลยืนยัน initializer ใหม่นี้

BA กลไก adaptive-coupling เดิม เปลี่ยนเฉพาะการตั้งต้นและเพิ่ม guard; SHA256 ปัจจุบัน cea19b7801309e8370b5fe943b31ac6c743085860be8341ef3df4361f4977a8b
กราฟ Form1.frm/modGraphing.bas ตาม cd31cc8 อ่าน 4 คอลัมน์เดิม ใช้ No. จริง เฉพาะ Passed and Better value อัปเดต best เส้นขั้นบันไดตัดแกน X, BA น้ำเงิน/HCA ส้ม ไม่มี initial ประดิษฐ์หรือเลื่อนเลขรอบ ผลเก่าไม่ได้วาดใหม่

หลักฐานสั้น: audit/shared-feasible-start-20260911/verified/runtime.txt และ verified-gui/compile.log; driver verify.py และ fixtures.bas
อ่าน HANDOFF.md สถานะบนสุด, AGENTS.md ถ้ามี และตรวจ git status/branch/remote/RC_RT_HCA_v2.vbp ก่อนทำงาน ตรวจ git log เพื่อทราบ commit ล่าสุด
ห้ามเขียน ลบ คืนไฟล์ จัดระเบียบ หรือ stage result_csv และไฟล์ภาพที่ผู้ใช้ลบ ณ ตอนส่งต่อพบผลเดิม896ไฟล์และคงไฟล์เหล่านี้ตามจริง คำสั่งทับCSV/BMPใช้เมื่อผู้ใช้รันVB6เท่านั้น ส่งออกเฉพาะ accept/loopPrice/design BMP ชื่อเดิมตามข้อกำหนด6ไฟล์ด้านบน ไม่เพิ่มคอลัมน์
เลือก Trial แยกวิธีจากราคาต่ำสุดก่อน แล้วรอบน้อยสุดเมื่อราคาเท่ากัน ไม่เลือก seed เพื่อให้ชนะ รักษา VB6 legacy encoding/CRLF/no BOM

งานต่อในแชทใหม่: อ่าน NEXT_CHAT_PROMPT.md, สถานะบนสุดของ HANDOFF.md และ AGENTS.md ถ้ามี ตรวจ git status/branch/remote และ RC_RT_HCA_v2.vbp แล้วสรุปสถานะสั้น ๆ รอคำสั่งหรือผลรันจากผู้ใช้ ยังไม่เริ่ม optimizer/30 trials/แก้สูตรหรืออัลกอริทึมเพิ่ม ผู้ใช้จะเปิด VBP ใหม่แล้ว F5 เอง

เมื่อผู้ใช้ส่งผลใหม่ ให้ตรวจ No.0 ของ accept ว่า BA/HCA ผ่านและราคาเท่ากันเมื่อ inputs เดียวกัน ตรวจ loopPrice ครบทุก Trial และ BMP แสดง Trial/ราคา/Loop ของ accept ที่เลือก อย่าใช้ตัวเลขผลเก่าหรือ fixtures ใน audit เป็นผลของชุดรันใหม่ เป้าหมายการทดลองเดิมคือ BA ราคาไม่สูงกว่า HCA และพบราคาสุดท้ายของตนก่อน โดยต้องการชนะอย่างน้อย60%ใน30ครั้งครบ H3/H4/H5 แต่ยังไม่ยืนยันผลเป้าหมายนี้หลังเปลี่ยนจุดเริ่มต้น

GUI ยัง Randomize จากเวลาครั้งแรกและใช้ stream ต่อเนื่อง ไม่ได้จับคู่ seed อัตโนมัติ หากจะวิเคราะห์ paired t-test ต้องตรวจรูปแบบการจับคู่ของชุดทดลองจริงก่อน และห้ามอ้างว่าใช้ seed เดียวกันโดยไม่มีหลักฐาน

มีรายการลบค้างใน git โดยตั้งใจ เช่น 3m hca.jpg และ result_csv พร้อมไฟล์ untracked จำนวนมาก ให้รักษาไว้ ไม่ใช้ git add -A / reset / clean / restore ทั้งโฟลเดอร์เพื่อทำให้ status สะอาด งานทดสอบถ้าผู้ใช้อนุญาตภายหลังให้แยกใน audit; เขียนผลจริงใน result_csv เฉพาะตามคำสั่งผู้ใช้
