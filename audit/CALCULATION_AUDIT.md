# RC_Wall_Fix — รายการแก้และตรวจยืนยัน 2026-09-08

สถานะ: แก้ส่วนที่ตรวจสอบได้และรัน VB6 จริงแล้ว **ยังไม่รับรองการออกแบบตาม วสท. 2562** เพราะยังไม่มีเนื้อหาข้อกำหนดฉบับที่ใช้อ้างอิงครบ ค่าทดสอบที่ทำให้แบบผ่านใน harness เป็นข้อมูลจำลองสำหรับทดสอบซอฟต์แวร์ ไม่ใช่เกณฑ์ วสท. และไม่ถูกเปิดใช้ในโปรแกรมหลัก

## ขอบเขตไฟล์และหลักฐานตั้งต้น

- ไม่พบ AGENTS.md ในโปรเจกต์หรือโฟลเดอร์บรรพบุรุษที่ตรวจ และไม่พบ JIT_code_audit.md ใต้ `C:\reserch 69` ไม่มี Git repository ในสำเนานี้
- `.vbp` ตั้งต้นใช้ Form1 และโมดูล modEnhancedValidation, modGraphing, modWSD, modDataStructures, modShared, modBatch, modBA, modHillClimbing
- modShared ที่ใช้งานจริงตั้งต้นอยู่ที่ `C:\Users\moosu\Downloads\modShared.bas` ไม่ใช่ไฟล์ในสำเนา ความต่างจากไฟล์ local คือประกาศ BatchMode และข้อความ comment หนึ่งจุด เก็บไฟล์ active ไว้ที่ [baseline/modShared.ACTIVE-external.bas](baseline/modShared.ACTIVE-external.bas) และ [diff](baseline/active-vs-local.diff)
- เปลี่ยน `.vbp` ให้ใช้ modShared ในสำเนาและเติม BatchMode ที่ขาด ไม่แก้ไฟล์ Downloads; ตรวจ byte comparison หลังแก้แล้ว
- สำรองไฟล์ต้นฉบับใน [baseline](baseline) ก่อนแก้ ตรวจ BA สำรองทุกไฟล์ที่พบ พร้อม hash และตัวแปรช่วงใน [BA-versions.txt](baseline/BA-versions.txt) ทุกรุ่นที่พบมีช่วงเพียง tb/TBase/Base ใช้รุ่นปัจจุบันเป็นฐาน ไม่ย้อนนำ patch ของรุ่นอื่นมาปะปน
- modPDFReport, modExport, modSettings และ modFontHelper ไม่ได้อยู่ใน `.vbp` จึงไม่แก้และไม่อ้างว่ารายงานจากโมดูลเหล่านี้ถูกตรวจแล้ว รายงานที่ใช้งานจริงคือ FormatResults และรายงานต่อ trial จาก FinishSearch
- คอมไพล์โปรเจกต์ตั้งต้นด้วย VB6 สำเร็จ: [compile.log](baseline/compile.log) เปิด EXE ตั้งต้นและบันทึกสถานะกระบวนการ: [gui-launch.txt](baseline/gui-launch.txt)
- รัน harness ที่ลิงก์ **modShared active ตั้งต้น** ด้วย VB6: [baseline-runtime.txt](baseline/baseline-runtime.txt) ได้ Mstem=0.1728 และ CheckSteelOK รับ DB12@0.25 เป็น True ทั้งที่กรณีอ้างอิงควรไม่ผ่าน

## สมมติฐานและสูตรที่แก้

ใช้หน่วยเมตร ตันแรง และหน้าตัดกว้าง 1 ม. ค่า f'c/fy/fs/fc เป็น kgf/cm² ไม่ใช่ MPa

1. H และ H1 เป็นระดับจาก **ท้องฐาน**; H เป็นดินด้านหลัง H1 เป็นดินด้านหน้า ไม่ใช่ความสูงดินด้านหลังอีกตัวหนึ่ง ผิวหลัง stem ตั้งตรง ผิวหน้าลาดตาม tt/tb ดินแห้งไม่มี cohesion ผิวดินราบ ไม่มี surcharge น้ำใต้ดินหรือแรงแผ่นดินไหวในโมเดลนี้
2. hs=H−TBase และ hp=H1−TBase สำหรับการตรวจ stem เท่านั้น; Ka=(1−sinφ)/(1+sinφ), Kp=1/Ka
3. Mstem=γ(Ka·hs³−ηKp·hp³)/6. CalculateMomentStem รับ Design ของคำตอบแต่ละตัวแล้ว ไม่ใช้ H1³ แทน hs³ และไม่เปลี่ยน H เหมารวมในแรงเสถียรภาพ
4. เสถียรภาพทั้งกำแพงยังใช้ Pa=γKaH²/2, Pp=ηγKpH1²/2 และแขน H/3, H1/3. FSot=(ΣWx+PpH1/3)/(PaH/3), FSsl=(μΣW+Pp)/Pa. เก็บเกณฑ์โครงการเดิม FSot≥2 และ FSsl≥1.5 โดยยังต้องยืนยันความเหมาะสมสำหรับงานนี้
5. η=PassiveFactor ใช้ **1 ตามเงื่อนไขโครงการที่ผู้ใช้กำหนด** รวม active ด้าน H และ passive เต็มค่าด้าน H1; InitializeArrays ตั้งค่านี้เมื่อโหลดฟอร์มและเริ่ม BA/HCA ดู [กรณีแรงปัจจุบัน](PROJECT_LOAD_CASE_TH.md) กรณี hp<0 ถูกปฏิเสธว่าอยู่นอกโมเดล ไม่บังคับความลึกให้เป็นบวก
6. cover คือระยะใสถึงผิว **เหล็กรับแรงดึงหลัก** ที่เลือก; หนึ่งชั้น d=t−cover−db/2. ถ้ามีเหล็กขวางอยู่ใกล้ผิวกว่าต้องทบทวนระยะถึงเหล็กหลัก ยังไม่รองรับการจัดสองชั้นและไม่เพิ่มเหล็กสองชั้นอัตโนมัติ ค่า d≤0 และ index ผิดถูกปฏิเสธ/แจ้ง error
7. น้ำหนักดินสามเหลี่ยมบนผิวลาดหน้ารวม centroid จริงกับส่วนสี่เหลี่ยม แทนใช้ x=LToe/2 สำหรับน้ำหนักทั้งหมด ตรวจ centroid stem และดินอิสระด้วย polygon area/centroid
8. กำหนด x จากปลาย toe ไป heel และ e=B/2−(ΣWx−Mnet)/ΣW. e บวกไปทาง toe, qtoe=W/B(1+6e/B), qheel=W/B(1−6e/B). ไม่ใช้ Abs(e) ก่อนกำหนดทิศแรงดัน
9. รองรับเฉพาะการสัมผัสฐานเต็มพื้นที่และไม่มีแรงดึง (|e|≤B/6) กรณี partial contact ถูกปฏิเสธ; เอาการประมาณสามเหลี่ยมแล้วลากเส้นตลอด B เดิมออก เงื่อนไข middle third สอดคล้องกับสมดุลของฐานสัมผัสเต็มพื้นที่ และดูคำอธิบายประกอบได้ใน [FHWA, external stability/bearing discussion](https://www.fhwa.dot.gov/clas/ctip/rockery_design_construction_guidelines/ch_4_recommended.aspx) ไม่ใช้เกณฑ์ออกแบบ rockery แทนเกณฑ์กำแพง ค.ส.ล.
10. qa ตีความเป็น **ค่ายอมให้** ตามป้ายเดิม จึงใช้ qa/qmax≥1 แทนหารเผื่อซ้ำด้วย 2. ถ้าค่าจากงานวิจัยจริงเป็นกำลังวิบัติ ต้องยืนยันก่อนใช้งานและเปลี่ยนป้าย/เกณฑ์ทั้งชุดอย่างชัดเจน
11. โมเมนต์ toe/heel อินทิเกรตแรงดันเส้นตรงกับแขนถึงหน้าตัด stem; toe หักทั้งน้ำหนักฐานและดินหน้าเหนือ toe, heel หักแรงดันจากน้ำหนักฐานและดินหลัง บวกของ toe หมายถึงเหล็กล่างรับแรงดึง บวกของ heel หมายถึงเหล็กบนรับแรงดึง โมเมนต์กลับทิศถูกปฏิเสธเพราะไม่มีข้อมูลเหล็กอีกหน้า ไม่ใช้ Abs เพื่อซ่อนการกลับหน้า

## หน้าตัด WSD และสิ่งที่ยังไม่มีข้อยืนยัน

สมการหน้าตัดแตกร้าวได้จากสมดุลแรงและ strain compatibility (ละคอนกรีตส่วนรับแรงดึง):

```
rho = As/(b*d)
k = sqrt((n*rho)^2 + 2*n*rho) - n*rho
j = 1 - k/3
fs_actual = |M|/(As*j*d)
fc_actual = 2*|M|/(b*k*j*d^2)
```

ใช้ As ของ DB/spacing จริงในการตรวจ **ทั้ง**คอนกรีตและเหล็ก; ไม่พึ่งการเช็ก As อย่างเดียวหรือ j สมดุลค่าคงที่ รายงานแสดงหน่วยแรงและ d จากเครื่องคำนวณชุดเดียวกับตัวตรวจ

ปัจจุบันใช้ **`n=9` ตามสมมติฐานที่ผู้ใช้ตั้งใจเลือกสำหรับโปรเจกต์นี้** ไม่เปลี่ยนเป็น Es/Ec อัตโนมัติ (ดู [ข้อชี้แจงและพารามิเตอร์](WSD_REFERENCE_PARAMETERS_TH.md)); ส่วน `fc_allow=0.45f'c`, `fs_allow=1500/1700` ยังมีไว้เป็น **legacy screening parameters** ไม่อ้างว่าเป็นข้อกำหนดที่ยืนยันแล้วจาก วสท. 2562. แก้คำอธิบาย n เป็น Es/Ec และ ρbalanced จากสมดุลมีตัวคูณ 1/2. เอาการใช้ 0.75ρbalanced เป็นเกณฑ์สูงสุดอัตโนมัติออก เพราะยังไม่มีข้ออ้างอิง WSD ที่ตรวจสอบได้

หน้ามาตรฐานของ [วสท. รายการหนังสือหน้า 80](https://eit.or.th/showcase/EIT/issue2_68/files/basic-html/page80.html) ยืนยันชื่อมาตรฐาน WSD รหัส **011007-19**, ISBN 978-616-396-023-8 ได้ แต่ไม่ได้ให้เนื้อหาข้อกำหนด จึง **ยังยืนยันไม่ได้**: modular ratio, หน่วยแรงยอมให้แยกวัสดุ/สมาชิก, หน่วยแรงเฉือนและหน้าตัดวิกฤต, เหล็กขั้นต่ำของ stem/ฐานและทิศทาง, เหล็กหดตัว/อุณหภูมิ, ระยะเรียง/ระยะใส, development/anchorage และรายละเอียดต่อเนื่อง

- มีตัวตรวจแรงเฉือน: stem ใช้ค่าสูงสุดของ nominal |V(y)|/[b*d(y)] ตลอดความสูง (แก้ข้อสันนิษฐานเดิมที่ว่าแรงเฉือนโคน conservative เมื่อมี passive), toe/heel ใช้แรงนอกระยะ d จากหน้ารองรับ; ค่ารายงานเป็น nominal V/(b*d). **ต้องยืนยันนิยามหน่วยแรง (รวมประเด็น j), หน้าตัดวิกฤต และค่าที่ยอมให้ตามฉบับจริงก่อนเปิดใช้**
- มีตัวตรวจ As_min แยก MinStemRatio/MinBaseRatio โดยใช้พื้นที่คอนกรีตรวม. ต้องยืนยันว่าข้อกำหนดแต่ละสมาชิกใช้ฐานใดและต้องมีเหล็กทิศอื่นเท่าใด; ไม่ใส่ 0.0015 เดิมเป็นข้อกำหนดทั่วทุกสมาชิก
- `AllowableShear`, `MinStemRatio`, `MinBaseRatio`, `WSDSource`, `WSDReviewed` เริ่มเป็นค่าว่าง/0/False. ตัวตรวจให้เหตุผล `WSD_CRITERIA_UNVERIFIED` และไม่รับคำตอบจนข้อมูลเกณฑ์พร้อม ไม่มีปุ่มในโปรแกรมหลักที่อ้างว่ารับรองเกณฑ์เหล่านี้แล้ว
- harness ตั้ง shear=8 kgf/cm², ratios=0.0015 และ source ที่ระบุ **SYNTHETIC TEST FIXTURE ONLY** เพื่อพิสูจน์ branch ผ่าน/ไม่ผ่านเท่านั้น ค่าเหล่านี้ไม่ถูกบันทึกเป็นค่ามาตรฐานของโปรแกรมหลัก
- ยังไม่เพิ่มรายละเอียดเหล็กทิศอื่น/anchor หรือต้นทุนที่เกี่ยวข้อง จึงไม่อ้างว่าราคาที่ค้นพบเป็นราคาของแบบพร้อมก่อสร้าง สูตรราคา/ปริมาณเดิมไม่ได้ปรับเพื่อรักษาราคาเดิม

## BA/HCA และการเก็บผล

- initializer และ routine GenerateNeighbor ของทั้งสองวิธียังคงเดิม (ตรวจเทียบ source แล้ว) BA บีบช่วงเฉพาะ tb, TBase, Base; tt, LToe และเหล็กยังสุ่มตาม routine เดิม
- ทุกคำตอบผ่าน EvaluateCandidate: initial, reset, neighbor. นับหนึ่งคำขอประเมินเป็นหนึ่ง evaluation แม้รูปทรงไม่ผ่าน อัปเดต RunBest จากทุกคำตอบที่ผ่านโดยไม่ขึ้นกับการยอมรับเป็น current
- งบรวม MaxIterations เดิมเปลี่ยนความหมายชัดเจนเป็นจำนวน evaluations รวมทั้งหมด. HCA ไม่มี initial แถมฟรี; BA ไม่มี reset แถมฟรี. กราฟและ BestCostIteration ใช้เลข evaluation
- BA ไม่บีบต่อจากรอบที่ไม่พบคำตอบผ่าน และเปิดคืนช่วงเดิมเมื่อทั้งสามช่วงยุบตัว ใช้ reset ที่สุ่มเฉพาะสามตัวในช่วงเดิม การ reopen ถูกนับใน RunRecoveryCount และการประเมิน reset ใช้งบปกติ
- ราคากับความเป็นไปได้ไม่ได้เป็นฟังก์ชัน monotone ของสามมิติเมื่อเหล็กและมิติอื่นเปลี่ยนพร้อมกัน BA นี้จึงเป็น heuristic; การกู้ช่วงแก้การติดช่วงไม่ผ่าน แต่ **ไม่รับประกัน global optimum** และไม่ใช่หลักฐานว่า BA ดีกว่า HCA
- seed เริ่มต้น 12345 แก้ได้ในฟอร์ม trial n ใช้ seed+n−1; ใช้ Rnd(-1) ก่อน Randomize เพื่อให้ทำซ้ำจริง BA/HCA ใช้ตาราง seed เดียวกัน รวมโหมด Random f'c ที่ใช้ seed ชัดเจน
- no-solution: IsValid=False, BestEvaluation=0, TotalCost=NO_SOLUTION_COST; ไม่ทำให้แบบว่างกลายเป็นราคาศูนย์และชนะ global best หน้าจอไม่วาดกราฟ best ที่ไม่มี
- ทุก trial สร้างโฟลเดอร์ใหม่ `App.Path\results\algorithm-timestamp-seed[-suffix]` พร้อม `evaluations.csv` และ `run.txt` (algorithm, trial, seed, budget, inputs, criteria, status, design/report). การตั้งปีในชื่อโฟลเดอร์ขึ้นกับ locale VB6 ของเครื่องนี้ (แสดง พ.ศ.) ไม่มีการทับ trial แม้เวลา/seed ตรงกัน
- ปิดปุ่มอัลกอริทึมอื่นระหว่างประมวลผล ป้องกัน DoEvents เรียกการค้นหาอีกวิธีมาทับตัวแปรร่วม
- batch มี guard เมื่อเกณฑ์ WSD ยังไม่ยืนยัน และ **ไม่ได้รัน batch วิจัยในงานนี้**

## ผลตรวจและวิธีรันซ้ำ

เรียกจากสำเนาโปรเจกต์:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File audit\run_checks.ps1
```

สคริปต์คอมไพล์โปรเจกต์จริง, native regression และ GUI regression ด้วย `VB6.EXE` ของเครื่องนี้ เก็บ compile log แยกใหม่ทุกครั้ง (VB6 /out เป็น append จึงไม่ใช้ log เก่าตัดสิน) รัน EXE จริงและตรวจข้อความผล ไม่ใช้ Python แทนหลักฐาน VB6

- [ผลสูตร/อัลกอริทึม VB6](native-regression.txt): **257 checks, failures=0**
- [ผลฟอร์มจริง](gui-regression.txt): **GUI failures=0**; Load/Show Form1 แล้วเรียก Click ของปุ่ม BA/HCA จริง ทั้งแบบ no-solution 2 trials และเกณฑ์จำลองที่ได้คำตอบ พร้อมรายงาน/กราฟและการคืนสถานะปุ่ม
- [ผล compile โปรเจกต์](final-compile-RC_RT_HCA_v2.log), [native harness](final-compile-Regression.log), [GUI harness](final-compile-GuiRegression.log); สำเนาหลักฐานแต่ละรอบและ hash EXE อยู่ใต้ `checks`
- [รายการอิสระ](independent-results.json) ใช้ Simpson integration ของแรง/แขน, polygon centroid และแก้ neutral axis โดยตรงเพื่อสร้าง expected values จากนั้นเทียบกับ VB6. มี H=3,4,5 ทั้งหน้าตัดหนาและบาง, e บวก/ลบ, โมเมนต์และแรงเฉือน toe/heel, concrete-only failure, ข้อมูลผิด, reversal และ passive ทั้ง 0/1
- กรณีอ้างอิง H5: hs=4.7000, Ma=10.3823, Mnet(η=1)=9.7262 ตันแรง·ม./ม. DB12@0.25 ใน t=0.20, cover=0.075 ได้ d=0.119 ม.; สำหรับ active-only, fc_actual=691.5640 และ fs_actual=20883.9863 kgf/cm² จึงถูกปฏิเสธอย่างชัดเจน
- ผ่านการตรวจ global best initial/reset/neighbor, initial-only budget=1, exact budget BA/HCA, same seed ทั้ง trace เหมือนกัน, no-solution, rejected-round recovery และไฟล์ trial ไม่ทับกัน
- [source-checks.txt](source-checks.txt) ตรวจช่วง BA มีแค่สามตัว, neighbor routine เดิมไม่เปลี่ยน, สูตรราคาคงเดิม, `.vbp` อ้างไฟล์ในสำเนา, ไฟล์ Downloads ไม่เปลี่ยน และตรวจ best prefix/evaluation budget ของ trace ที่เก็บไว้ทั้งหมด
- [changed-files.json](changed-files.json) เป็น hash ไฟล์ที่แก้; ไฟล์ VB6 คง byte encoding เดิมโดยแก้ข้อความใหม่เป็น ASCII และคง CRLF ไม่แปลงเป็น UTF-8/BOM

**ยังไม่ได้รัน/ยังไม่สรุป:** batch วิจัย, การเปรียบเทียบสถิติหลาย seed/งบใหญ่, exhaustive optimum, แบบสองชั้นหรือ partial-contact, การรับรอง วสท. 2562 และรายละเอียดเหล็กครบเพื่อก่อสร้าง ต้องยืนยันเอกสารมาตรฐานจริง, qa เป็น allowable/ultimate, passive ที่ใช้ได้ และตำแหน่งเหล็ก/cover ก่อนทำงานวิจัยต่อ

## Per-check report update (2026-09-08)

See [H5 result in Thai](H5_CHECK_RESULT_TH.md) and [actual VB6 tables](h5-vb6-checks.md). The new report continues independent checks after a failure, distinguishes unset criteria, and rejects the weak reference using a necessary elastic yield bound. The prior BA candidate remains INDETERMINATE_WSD. Actual VB6: 140 checks, failures=0; GUI failures=0; independent report comparisons: 44, mismatches=0. No production WSD criteria or BA search behavior changed.

## Feasibility before cost (2026-09-08)

See [selection order and native evidence](FEASIBLE_COST_ORDER.md). Search pricing now occurs only after all three stability checks AND the remaining structural criteria succeed. Rejected candidates retain the no-solution sentinel, never compete on price, and still consume one evaluation. Native regression: 150 checks, failures=0; GUI failures=0. BA and original pricing formulas remain unchanged.

## H5 independent-member recalculation (2026-09-08)

See [new active/passive report](H5_RECALCULATION_TH.md). Native BA retains full validation and reports no solution; a separate explicitly relaxed grid sizes the three members independently. Two complete 520200-row grids agree with independent calculation after fixing binary-roundoff handling of heel=toe through shared CheckHeelLayout. No EIT criteria were invented or enabled. Native regression: 152 checks, failures=0; GUI failures=0.


## Latest whole-wall recalculation and stem shear correction

See [latest Thai report](H5_REVIEWED_CALCULATION_TH.md). This run reads qa=30 allowable and the 5000-evaluation budget from Form1 defaults, retaining the requested H=5/fc=320 study. Same-geometry native tests verify active/passive loads, their one-third-height arms, all three stability checks, bearing reactions and toe/heel moments against independent equilibrium. Fixed a real stem shear defect: passive cancellation and taper can make an interior nominal shear stress exceed the face value (1.6475 versus 1.4006 kgf/cm2 for the selected geometry). The validator, screen and report now share the envelope. Native regression: 211 checks, failures=0; GUI failures=0. Both complete 520200-row diagnostic grids and 120 selected-geometry steel alternatives agree independently. The screening minimum remains 8273.14368 baht/m, but production BA reports NO_SOLUTION in both cases because complete WSD criteria remain unverified. No research batch was run. Earlier update paragraphs above describe historical runs.


## Prior automatic Es/Ec change (superseded by user clarification)

The user supplied the Pongnathee Chapter 10 example. Its n=9 is an approximation for f'c=210, not a constant for all strengths. CalculateWSDParameters now uses the unrounded Es/Ec equation; balanced k/j/R are explicitly labelled, while actual-bar checks retain their own neutral axis. See [reference audit and current recalculation](WSD_REFERENCE_PARAMETERS_TH.md). Native regression: 256 checks, failures=0; GUI failures=0. Both full grids were rerun; the selected prices are unchanged but stresses and screened counts changed. This establishes the teaching-reference model, not complete EIT 011007-19 compliance.


## User-selected modular ratio restored (current)

The user clarified that n=9 was an intentional project assumption. The earlier classification of this choice as a software bug was incorrect. Production CalculateWSDParameters now uses PROJECT_MODULAR_RATIO=9, with no automatic replacement by Es/Ec. Balanced k/j/R still respond to the material allowables; actual-bar k/j still come from As and d. See [current parameter report](WSD_REFERENCE_PARAMETERS_TH.md). Earlier Es/Ec calculations are historical comparisons, not the active project configuration.


## Current project load case: active plus full passive

The user specifies active loading from H and full passive from front H1. InitializeArrays now sets PROJECT_PASSIVE_FACTOR=1 for Form1 and both optimizers. The current sizing study contains one project case; earlier comparison runs are historical. See [project load case](PROJECT_LOAD_CASE_TH.md). Unit-level isolated force checks are not alternative project designs.
