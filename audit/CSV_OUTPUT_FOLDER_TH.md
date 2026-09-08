# โฟลเดอร์ CSV ของโปรเจกต์

เปลี่ยนเมื่อ 9 กันยายน 2026 ตามคำขอ: CSV ทุกประเภทที่โปรเจกต์หลักส่งออกใช้ราก `C:\reserch 69\RC_Wall_Fix\result_csv\` โปรแกรมสร้างโฟลเดอร์นี้อัตโนมัติเมื่อเริ่มส่งออก ตำแหน่งไม่เปลี่ยนตามที่เก็บ EXE หรือ working directory

กำหนดไว้จุดเดียวที่ `modShared.RESULT_CSV_ROOT` และใช้ `ResultCsvRoot()` สำหรับสร้าง/คืนตำแหน่ง `BeginSearch` กับ `modBatch.BatchOutputPath` เปลี่ยนมาใช้ฟังก์ชันเดียวกัน ส่วน Form1 แสดงตำแหน่งนี้ให้ผู้ใช้เห็น

- ไฟล์หลัก `accept-BA-H...csv`, `accept-HCA-H...csv`, `loopPrice-BA-H...csv` และ `loopPrice-HCA-H...csv` บันทึกที่ราก `result_csv` โดยตรงอัตโนมัติเมื่อใช้ปุ่ม BA/HCA
- `accept` แยกแต่ละ trial โดยใส่ trial/seed ในชื่อไฟล์ เช่น `accept-BA-H5-trial1-seed12345.csv` และเติม suffix เมื่อชื่อซ้ำ มีหนึ่งแถวต่อ evaluation รวม initial/reset/neighbor
- `loopPrice` รวมทุก trial ของการกดปุ่มหนึ่ง session เช่น `loopPrice-BA-H5.csv` ถ้าชื่อซ้ำจะเป็น `loopPrice-BA-H5-1.csv` เป็นต้น เมื่อตั้ง 30 trials จะมี 30 แถวในไฟล์นี้
- `evaluations.csv` และ `run.txt` อยู่ในโฟลเดอร์ย่อย BA/HCA แยกตามวันที่ เวลา seed และ suffix เมื่อชื่อซ้ำ
- `trial-summary.csv` อยู่ในโฟลเดอร์ trial แรกของ session ตามเดิม
- `batch_step3*.csv` อยู่ใต้ราก `result_csv` โดยตรง พร้อมวันที่ เวลา และ suffix เมื่อชื่อซ้ำ

จึงยังแยกแต่ละ trial ไม่ทับกัน ไม่ได้ย้ายหรือลบผลรันเก่าใน `results`/`audit/results` โมดูลสำรองที่ไม่อยู่ใน `.vbp` และ CSV fixtures ของชุดตรวจใน `audit` ไม่ใช่ตัวส่งออกของโปรแกรมหลักและไม่ได้เปลี่ยน

## คืนการส่งออกไฟล์หลัก

พบว่า SaveAcceptCSV เดิมเป็น stub และปุ่ม Form1 ไม่ได้เรียก SaveLoopPriceCSV หลังจบ loop จึงแก้ให้ BeginSearch เริ่ม buffer accept, EvaluateCandidate บันทึกการจัดประเภทจากผลตรวจจริง และ FinishSearch บันทึก accept แต่ละ trial ส่วน Form1 บันทึก loopPrice ครั้งเดียวเมื่อจบทุก trial พร้อมแสดงชื่อไฟล์ที่ส่งออก

`accept` คงคอลัมน์ `No.,Rejected,Passed,Passed and Better value`: No. คือ evaluation; Rejected ใส่ `INVALID` เพราะไม่มีราคาของคำตอบที่ไม่ผ่าน; Passed คือราคาของคำตอบที่ผ่านแต่ไม่ลด global best; Passed and Better value คือราคาที่ปรับปรุง global best ของ trial นั้น

`loopPrice` คงสามคอลัมน์แรก `No.,Loop,BestPrice` และเพิ่ม `Status`: No. คือ trial, Loop คือ evaluation แรกที่พบราคาดีที่สุดของ trial, BestPrice ว่างและ Loop=0 เมื่อ NO_SOLUTION ไม่ใช้ sentinel เป็นราคา ค่าตัวเลข CSV ใช้จุดทศนิยมเพื่อไม่ชนตัวคั่นคอลัมน์

หลักฐานล่าสุด: [VB6 47 checks, failures=0](csv-path-checks/20260909-003550-400/native-csv-paths.txt), [ตรวจเนื้อหา CSV กับ native logs](csv-path-checks/20260909-003550-400/primary-csv-verification.json) ครบทั้ง Rejected 95 แถว, Passed 19 แถว, Better 16 แถวจาก 4 sessions ตรวจ loopPrice แบบ 1 และ 2 trials รวมกรณี NO_SOLUTION; [trace BA/HCA 64 evaluations เท่ากับก่อนแก้ byte-for-byte](csv-path-checks/20260909-003550-400/trace-preservation.txt) ไม่เปลี่ยนคำตอบหรือขั้นตอนค้นหา ไม่ได้รัน 30 trials หรือ batch matrix

## หลักฐานรอบย้ายโฟลเดอร์ก่อนคืนไฟล์หลัก

คอมไพล์โปรเจกต์หลักและ harness VB6 จริงสำเร็จ ทดสอบ [19 ข้อ failures=0](csv-path-checks/20260909-002834-089/native-csv-paths.txt) ครอบคลุมการสร้างโฟลเดอร์, เรียก export ก่อนเริ่ม trial, ปุ่ม BA/HCA จริงอย่างละ 1 trial × 8 evaluations, CSV ทุกชนิดของสองปุ่ม, เส้นทางที่หน้าจอแสดง และการบันทึก loop ซ้ำไม่ทับไฟล์ ตรวจ [ทุกเส้นทาง CSV ในไฟล์ active](csv-path-checks/20260909-002834-089/routes.json) ประกอบ

โฟลเดอร์ `BA-25690909-002837-seed20260909` และ `HCA-25690909-002837-seed20260909` ใน `result_csv` เป็นข้อมูล smoke test นี้ ไม่ใช่ผลวิจัย ไม่ได้รันชุด 30 trials หรือ batch matrix และไม่ได้เปลี่ยนสูตร ราคา หรือขั้นตอนค้นหา BA/HCA
