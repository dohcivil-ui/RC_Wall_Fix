# RC_Wall_Fix

โปรเจกต์ VB6 ออกแบบกำแพงกันดิน พร้อม BA ที่บีบช่วงเฉพาะ `tb`, `TBase`, `Base` และ HCA

**ยังไม่รับรองแบบตาม วสท. 2562 และยังไม่ได้รัน batch วิจัย** เกณฑ์ที่ยังไม่มีข้ออ้างอิงครบแสดงสถานะ `WSD_CRITERIA_UNVERIFIED` ค่าที่ทำให้แบบผ่านใน test harness เป็นข้อมูลจำลองสำหรับตรวจซอฟต์แวร์เท่านั้น

- [ตรวจ ref: n, k, j, R และผลคำนวณใหม่ล่าสุด](audit/WSD_REFERENCE_PARAMETERS_TH.md)
- [ผลคำนวณ H5 รอบก่อนแก้ n: qa ตาม VB6, แรงทั้งกำแพง, เหล็กและราคา](audit/H5_REVIEWED_CALCULATION_TH.md)
- เปิดโปรเจกต์หลัก: `RC_RT_HCA_v2.vbp`
- [รายการแก้ สมมติฐาน สูตร และสิ่งที่ต้องยืนยัน](audit/CALCULATION_AUDIT.md)
- [ผลตรวจด้วย VB6 จริง: 256 checks, failures=0](audit/native-regression.txt)
- [ผลทดสอบปุ่ม BA/HCA บนฟอร์มจริง](audit/gui-regression.txt)
- [หลักฐานคอมไพล์โปรเจกต์หลัก](audit/final-compile-RC_RT_HCA_v2.log)
- [หลักฐานก่อนแก้และสำเนาซอร์สตั้งต้น](audit/baseline)

ทดสอบซ้ำบน Windows ที่ติดตั้ง VB6 และ dependencies ของโปรเจกต์:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File audit\run_checks.ps1
```

สคริปต์ทดสอบยังอ้างตำแหน่ง VB6 และไฟล์ตั้งต้นของเครื่องที่ตรวจไว้ ต้องปรับตำแหน่งเมื่อใช้เครื่องอื่น ส่วน `audit/patch_*.py`, `finalize_fixes.py` และ `prepare_baseline.py` เป็นบันทึกเครื่องมือที่ใช้แก้ครั้งนี้ ไม่ควรรันซ้ำบนซอร์สที่แก้แล้ว

เก็บซอร์ส ไฟล์สำรอง และรายงานทดสอบใน Git; ไม่เก็บ EXE หรือสถานะ editor ที่สร้างเฉพาะเครื่อง `.gitattributes` ปิดการแปลง line endings เพื่อรักษาไฟล์ VB6 เดิม
