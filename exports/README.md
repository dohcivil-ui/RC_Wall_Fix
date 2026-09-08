# ภาพเวกเตอร์จากผล VB6

- `best_BA_trial1_7x6cm.svg`: เวกเตอร์จริง (เส้น วงกลม และข้อความ ไม่มี bitmap ฝังใน SVG) ขนาดเอกสาร 7 × 6 ซม. SVG ไม่ขึ้นกับ DPI
- `best_BA_trial1_7x6cm_600dpi.png`: เรนเดอร์จาก SVG โดยตรง ขนาด 1,654 × 1,417 พิกเซล พร้อม metadata 600 DPI

ข้อมูลมาจาก `audit/results/BA-25690909-000943-seed20260908/run.txt`: ผลทดสอบ VB6 เดิม BA 1 trial, 64 evaluations ไม่ได้รัน optimization ใหม่ ไม่ใช่ผลทดลอง 30 trials วงกลมเป็นสัญลักษณ์ของชั้นเหล็กหลัก ไม่ใช้จำนวนวงกลมถอดปริมาณ ขนาดและระยะที่ป้ายมาจากรายงาน

ตัวส่งออกอยู่ใน `audit/export_best_svg.py`; ตัวเรนเดอร์ PNG คือ `audit/render_svg_png.cjs` ใช้ Node.js และ sharp รายการตรวจขนาดไฟล์ ความละเอียด และ SHA256 อยู่ใน `export-verification.json` ไม่มีการเปลี่ยนซอร์ส VB6 ในงานส่งออกไฟล์นี้
