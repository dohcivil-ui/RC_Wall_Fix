# RC_Wall_Fix

โปรเจกต์ VB6 ออกแบบกำแพงกันดิน พร้อม BA ที่บีบช่วงเฉพาะ `tb`, `TBase`, `Base` และ HCA

รุ่นปัจจุบันเปิดตัวตรวจ `PROJECT_WSD_ACI99_V3_MAIN_ONLY`: WSD จากเอกสารโครงการร่วมกับข้อเสริม ACI 318-99 ใช้ active และ passive เต็มค่า ตรวจเสถียรภาพทั้งสาม หน่วยแรงตลอด stem และฐาน ก่อนเทียบราคา **คิดเฉพาะคอนกรีตและเหล็กหลัก stem/toe/heel ไม่คิดเหล็กแนวนอนหรือเหล็กประกอบ ไม่ตรวจ/รวมระยะฝังยึดหรือทาบ และไม่คิดราคาไม้แบบ ตามขอบเขตผู้ใช้** ปริมาณใช้ความยาวช่วงสมาชิก ไม่มีค่าเผื่อ .40 ม. ยังไม่รับรองครบทุกข้อของ วสท. 2562 หรือ ACI

คอมไพล์และทดสอบด้วย VB6 จริง: regression เดิม 267 ข้อ และ production ใหม่ 47 ข้อ failures=0; เทียบรายการคำนวณอิสระ 300 ค่า รุ่นนี้ยังไม่รัน 30 trials — ผู้ใช้จะรันเอง ผล NO_SOLUTION ทั้ง 30 ครั้งในรายงานเก่าเป็นผลก่อนเปิดตัวตรวจรุ่นนี้

- [รายละเอียดเกณฑ์ ผลทดสอบ และวิธีรัน 30 ครั้ง](audit/PROJECT_CHECKS_IMPLEMENTATION_TH.md)
- [หลักฐาน VB6 ของตัวตรวจปัจจุบันและปุ่ม BA](audit/project-checks/20260908-232727-197/project-regression.txt)

- [กรณีแรงโครงการ: active และ passive เต็มค่า](audit/PROJECT_LOAD_CASE_TH.md)
- [ประเมินผลกระทบ n และใช้ Es/Ec ตามหลักฐาน](audit/MODULAR_RATIO_DECISION_TH.md)
- [ผล BA 30 ครั้งในอดีต ก่อนเปิดตัวตรวจปัจจุบัน](audit/TRIAL30_RESULT_TH.md)
- [ข้อ ACI ที่ตรวจพบและผลตรวจเสริมด้วย VB6](audit/ACI_SUPPLEMENT_RESULT_TH.md)
- เปิดโปรเจกต์หลัก: `RC_RT_HCA_v2.vbp`
- [รายการแก้ สมมติฐาน สูตร และสิ่งที่ต้องยืนยัน](audit/CALCULATION_AUDIT.md)
- [ผลตรวจด้วย VB6 จริง: 267 checks, failures=0](audit/native-regression.txt)
- [ผลทดสอบปุ่ม BA/HCA บนฟอร์มจริง](audit/gui-regression.txt)
- [หลักฐานคอมไพล์โปรเจกต์หลัก](audit/final-compile-RC_RT_HCA_v2.log)
- [หลักฐานก่อนแก้และสำเนาซอร์สตั้งต้น](audit/baseline)

ทดสอบซ้ำบน Windows ที่ติดตั้ง VB6 และ dependencies ของโปรเจกต์:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File audit\run_project_checks.ps1
```

สคริปต์ทดสอบยังอ้างตำแหน่ง VB6 และไฟล์ตั้งต้นของเครื่องที่ตรวจไว้ ต้องปรับตำแหน่งเมื่อใช้เครื่องอื่น ส่วน `audit/patch_*.py`, `finalize_fixes.py`, `prepare_baseline.py`, `install_project_checks.py` และ `finish_project_checks.py` เป็นบันทึกเครื่องมือที่ใช้แก้ครั้งนี้ ไม่ควรรันซ้ำบนซอร์สที่แก้แล้ว

เก็บซอร์ส ไฟล์สำรอง และรายงานทดสอบใน Git; ไม่เก็บ EXE หรือสถานะ editor ที่สร้างเฉพาะเครื่อง `.gitattributes` ปิดการแปลง line endings เพื่อรักษาไฟล์ VB6 เดิม
