"""Validate all 30 persisted native runs, including every candidate evaluation."""
from pathlib import Path
import sys, json, csv, re, hashlib
from collections import Counter

P = Path(__file__).resolve().parent
E = Path(sys.argv[1])
before = set(json.loads((E/'prior-folders.json').read_text(encoding='utf-8-sig')))
folders = [path for path in (P/'results').iterdir() if path.is_dir() and path.name not in before]
assert len(folders) == 30, len(folders)
summary = []
all_reasons = Counter()
for path in folders:
    run = (path/'run.txt').read_text()
    def number(key):
        return float(re.search(r'(?:^|; )'+re.escape(key)+r'=([^;\r\n]+)',run,re.M)[1])
    trial, seed, evaluations, budget = [int(number(k)) for k in ['Trial','Seed','Evaluations','Budget']]
    assert 'Algorithm=BA;' in run
    assert evaluations == budget == 5000 and seed == 12344+trial and 1 <= trial <= 30
    for key,value in dict(H=5,H1=1.2,gamma_soil=1.8,gamma_concrete=2.4,phi=30,mu=.6,qa_allowable=30,clear_cover=.075,fc_prime=320,fy=4000,passive_fraction=1).items():
        assert abs(number(key)-value) < 1e-10, (path.name,key)
    assert 'WSDReviewed=False; source=' in run
    assert number('AllowableShear') == number('MinStemRatio') == number('MinBaseRatio') == 0
    rows = list(csv.DictReader((path/'evaluations.csv').open()))
    assert len(rows) == 5000
    best = 999999999.0
    reasons, entries = Counter(), Counter()
    valid_count = 0
    for index,row in enumerate(rows,1):
        assert int(row['evaluation']) == index
        reasons[row['reason']] += 1
        entries[row['entry']] += 1
        if row['valid'] == 'True':
            valid_count += 1
            best = min(best,float(row['cost']))
        assert abs(float(row['best_cost'])-best)<1e-7
    status = re.search(r'^Status=(.+)$',run,re.M)[1]
    assert status == 'NO_SOLUTION' and valid_count == 0 and number('BestEvaluation') == 0
    all_reasons.update(reasons)
    summary.append(dict(trial=trial,seed=seed,evaluations=evaluations,status=status,valid_candidates=valid_count,
                        best_cost_baht_per_m=None,folder=str(path.relative_to(P)),reasons=dict(reasons),entries=dict(entries),
                        evaluations_sha256=hashlib.sha256((path/'evaluations.csv').read_bytes()).hexdigest()))
summary.sort(key=lambda row:row['trial'])
assert [row['trial'] for row in summary] == list(range(1,31))
assert len({row['evaluations_sha256'] for row in summary}) == 30, 'Unexpected identical search traces'
with (E/'trial-summary.csv').open('w',newline='',encoding='ascii') as f:
    writer=csv.DictWriter(f,fieldnames=['trial','seed','evaluations','status','valid_candidates','best_cost_baht_per_m','folder','evaluations_sha256'])
    writer.writeheader()
    writer.writerows({k:row[k] for k in writer.fieldnames} for row in summary)
result=dict(algorithm='BA',trials=30,total_evaluations=sum(row['evaluations'] for row in summary),accepted_trials=0,
            status='NO_SOLUTION_ALL_TRIALS',best_cost=None,cost_statistics=None,rejection_counts=dict(all_reasons),runs=summary,
            note='Actual Form1 BA button executed. Missing criteria were not bypassed. No cost distribution or optimum claim can be made.')
(E/'trial-details.json').write_text(json.dumps(result,indent=2),encoding='ascii')
(P/'trial30-latest.json').write_text(json.dumps(dict(evidence=str(E.relative_to(P)),trials=30,total_evaluations=150000,status=result['status'],best_cost=None),indent=2),encoding='ascii')
elapsed = float(re.search(r'NativeElapsedSeconds=([0-9.]+)',(E/'trial30-native.txt').read_text())[1])
report = f'''# ผลทดลอง BA 30 ครั้งด้วย VB6 จริง

รันครบ **30/30 ครั้ง รวม 150,000 evaluations** ผ่านปุ่ม BA ของ Form1 จริง ใช้เวลารวมที่ VB6 วัดได้ **{elapsed/60:.2f} นาที** ผลทั้ง 30 ครั้งเป็น **NO_SOLUTION** และไม่มีคำตอบที่ผ่านตัวตรวจครบ จึงไม่มีราคาต่ำสุด ค่าเฉลี่ย หรือส่วนเบี่ยงเบนมาตรฐานของราคาจากการทดลองนี้

## เงื่อนไข

H=5.00, H1=1.20 ม., γดิน=1.8, γคอนกรีต=2.4 ตันแรง/ม.³, φ=30°, μ=.60, qa=30 ตันแรง/ม.² (allowable), clear cover=.075 ม., f′c=320 และ fy=4000 กก.แรง/ซม.²

รวม active และ passive เต็มค่า ใช้ n=Es/Ec=7.552282573 และ BA บีบเพียง tb/TBase/Base ทุกครั้งใช้งบ 5,000 evaluations ตั้ง seed 12345–12374 ตามลำดับ trial 1–30 การทดลองนี้เป็น BA อย่างเดียว ไม่ใช่ผลเปรียบเทียบ BA/HCA

ไม่ได้เปิดเกณฑ์ WSD จำลองหรือปิดตัวตรวจเพื่อให้ได้ราคา WSDReviewed=False, AllowableShear/MinStemRatio/MinBaseRatio ยังไม่ตั้งค่าจากมาตรฐานครบ จำนวนต่อไปนี้เป็นเหตุผลแรกที่ตัวตรวจหยุดในแต่ละ evaluation ไม่ใช่จำนวนข้อบกพร่องทั้งหมดของคำตอบนั้น

| เหตุผลที่ปฏิเสธ | จำนวน evaluations |
| --- | ---: |
'''
for reason,count in sorted(all_reasons.items()):
    report += f'| {reason} | {count:,} |\n'
report += f'''
## หลักฐาน

- [ผลหน้าจอจาก VB6 หลังรันครบ]({E.relative_to(P).as_posix()}/trial30-native.txt)
- [สรุปครบทั้ง 30 trials]({E.relative_to(P).as_posix()}/trial-summary.csv): seed, งบ, สถานะ และตำแหน่งรายงานแต่ละรอบ ช่องราคาว่างหมายถึงไม่มีคำตอบที่ผ่าน ไม่ใช่ราคาศูนย์
- [รายละเอียดและจำนวนเหตุผลต่อ trial]({E.relative_to(P).as_posix()}/trial-details.json)
- [หลักฐานคอมไพล์และ SHA256]({E.relative_to(P).as_posix()}/): ตรวจทุกไฟล์ evaluations.csv แล้ว ลำดับ 1–5000 ครบ และ best_cost สอดคล้องกับคำตอบที่ตัวตรวจรับ

ผล NO_SOLUTION ทั้งหมดในสภาพที่เกณฑ์ WSD ยังไม่ครบ **ไม่พิสูจน์ว่าไม่มีแบบที่ออกแบบได้ และใช้สรุปความสามารถหา optimum ของ BA ไม่ได้** ตัวเลข sentinel ภายในโปรแกรมไม่ถูกนำมาคำนวณเป็นราคาหรือสถิติราคา

## ทดลองเอง

เปิดโปรเจกต์หลัก `RC_RT_HCA_v2.vbp` ใน VB6 แล้วกด F5 ตั้งค่าตามเงื่อนไขข้างต้น เลือกกำลังคอนกรีต 320, Trials=30, MaxIter=5000, Seed=12345 แล้วกดปุ่ม BA ผลแต่ละรอบอยู่ในโฟลเดอร์ results ใต้ตำแหน่ง EXE

หรือรันซ้ำผ่านสคริปต์ที่ตั้งค่าฟอร์มและเรียกปุ่มเดียวกัน:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File audit/run_trial30.ps1
```

สคริปต์สร้างโฟลเดอร์หลักฐานใหม่ทุกครั้ง ใช้ compiler VB6 ที่ติดตั้งในเครื่องนี้ การตรวจ Python ใช้ตรวจแฟ้มผลที่ VB6 สร้าง ไม่ใช้แทนหลักฐานรัน VB6
'''
(P/'TRIAL30_RESULT_TH.md').write_text(report,encoding='utf-8')
print('Verified actual VB6: BA 30/30 trials, 150000 sequential evaluations, fixed settings and distinct seeds, no accepted designs.')
print('Rejection counts:',dict(all_reasons))
