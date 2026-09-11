ทำงานต่อใน C:\reserch 69\RC_Wall_Fix โดยใช้โฟลเดอร์นี้โดยตรง

สถานะล่าสุด 11 กันยายน 2569: ผู้ใช้เปลี่ยนคำสั่งเป็น “แบบเริ่มต้นร่วมที่ตรวจผ่านจริง” เพื่อให้กราฟ BA/HCA เริ่มที่ No.0 ราคาเดียวกัน คำสั่งนี้แทนการคง max/Reject ใน cd31cc8 อย่าคืน max หรือเติมเส้น/ราคา Reject ลงกราฟ

ค่าเริ่มต้นร่วมทุกความสูง: tt=.40, tb=.60, TBase=.50, B=3.00, toe=1.10, heel=1.30 เมตร; stem/toe DB20@.125, heel DB16@.15 อยู่ใน catalogue เดิมทั้งหมด ใช้ SetCommonInitialIndices ใน modShared.bas ทั้ง BA และ HCA ไม่มีการสุ่มหรือค้นหาเพื่อเตรียมจุดเริ่มต้น

RequireFeasibleInitial ตรวจแบบอ้างอิงนี้กับ inputs ปัจจุบันผ่านชุดตรวจเดิมก่อน EvaluateCandidate(initial) ถ้าไม่ผ่านให้แจ้งเหตุและหยุดก่อน No.0 โดยไม่เขียนทับผลเดิม ถ้าผ่าน EvaluateCandidate จะตรวจและบันทึกราคาจริงใน Passed and Better value ที่ No.0 ซึ่งนับเป็น evaluation 1 จากงบเดิม ไม่บังคับสถานะ Passed และไม่แก้สูตร หน่วย ขอบเขต กลไกค้นหา/ยอมรับ หรือ RNG

ทดสอบ H3/fc240, H4/fc280, H5/fc320 ด้วย H1=1.2, gamma_soil=1.8, gamma_concrete=2.4, phi=30, mu=.6, qa=30, cover=.075, SD40: แบบผ่านจริงทั้งสาม ราคาเริ่ม BA=HCA คือ 8,626.89 / 10,621.25 / 12,705.85 บาท/ม. ตรวจซ้ำข้ามวิธีและ Trial ด้วยงบ 1 รวม 12 initial-only calls ไม่มี neighbor/midpoint; รวม regression 109 ข้อผ่านและคอมไพล์ GUI ผ่าน ค่า input อื่นต้องตรวจจริง ไม่อ้างว่าผ่านทุกเงื่อนไข

ผู้ใช้เปิด RC_RT_HCA_v2.vbp ใหม่แล้ว F5 เอง ยังไม่ได้รัน 30 trials จากจุดตั้งต้นใหม่ และไม่เปลี่ยน EXE เดิม ผลชนะเก่า H3=26/30 H4=24/30 H5=21/30 เป็นรุ่นเริ่ม max ไม่ใช่ผลยืนยัน initializer ใหม่นี้

BA กลไก adaptive-coupling เดิม เปลี่ยนเฉพาะการตั้งต้นและเพิ่ม guard; SHA256 ปัจจุบัน e7010f8573f69b0aa5c989d8a90d88d1d507f71c9fa2a2289185d9f5339f8c5c
กราฟ Form1.frm/modGraphing.bas ตาม cd31cc8 อ่าน 4 คอลัมน์เดิม ใช้ No. จริง เฉพาะ Passed and Better value อัปเดต best เส้นขั้นบันไดตัดแกน X, BA น้ำเงิน/HCA ส้ม ไม่มี initial ประดิษฐ์หรือเลื่อนเลขรอบ ผลเก่าไม่ได้วาดใหม่

หลักฐานสั้น: audit/shared-feasible-start-20260911/verified/runtime.txt และ verified-gui/compile.log; driver verify.py และ fixtures.bas
อ่าน HANDOFF.md สถานะบนสุด, AGENTS.md ถ้ามี และตรวจ git status/branch/remote/RC_RT_HCA_v2.vbp ก่อนทำงาน ตรวจ git log เพื่อทราบ commit ล่าสุด
ห้ามเขียน ลบ คืนไฟล์ จัดระเบียบ หรือ stage result_csv และไฟล์ภาพที่ผู้ใช้ลบ ผลเดิม901ไฟล์ตรง snapshot ก่อนงาน คำสั่งทับCSV/BMPใช้เมื่อผู้ใช้รันVB6เท่านั้น คง accept/loopPrice/selectedTrial/BMP ชื่อเดิม ไม่เพิ่มไฟล์หรือคอลัมน์ส่งออก
เลือก Trial แยกวิธีจากราคาต่ำสุดก่อน แล้วรอบน้อยสุดเมื่อราคาเท่ากัน ไม่เลือก seed เพื่อให้ชนะ รักษา VB6 legacy encoding/CRLF/no BOM
