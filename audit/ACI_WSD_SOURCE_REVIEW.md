# ACI sources for the project WSD review

Reviewed 2026-09-08. Source review only; no production criteria or acceptance flags changed. See [native supplementary calculations](ACI_SUPPLEMENT_RESULT_TH.md).

ACI can supplement the project basis, but the edition and the purpose of each adopted provision must be explicit. The existing source comment “ACI 318-19” does not establish that its WSD constants are provisions of that edition.

## Historical WSD: ACI 318-99

The [official ACI catalog](https://www.concrete.org/store/productdetail/itemid/31899.aspx?Format=PROTECTED_PDF&ItemID=31899&Language=English&Units=US_Units) confirms Appendix A is the Alternate Design Method. Full original pages were inspected in an [Internet Archive mirror of the ACI publication](https://ia800601.us.archive.org/6/items/1.ACI31899/1.%20ACI%20-318-99.pdf). This is original ACI content, **not an ACI-hosted or authenticated current copy**. The same URL is cited by a [university thesis, bibliography item 9](https://dspace.ups.edu.ec/bitstream/123456789/25530/1/TTS1388.pdf).

Compact clause map from the original (printed page; PDF page = printed page + 1):

| Clause | Verified provision / location |
|---|---|
| A.1.1–A.2.1 | Service loads; unit load/reduction factors. A.1.4 retains 10.4–10.7 and deflection control. pp358–359 |
| A.2.3 | Beneficial dead load reduced to 85 percent in member design combinations. p359 |
| A.3.1–A.3.2 | Flexural concrete limit 0.45f′c; steel 20,000 psi for Grade40/50, 24,000 psi for Grade60+. pp359–360 |
| A.5.4 | n=Es/Ec; nearest integer, minimum6, permitted. p361 |
| A.7.1; A.7.4.1 | v=V/(bwd); normal-weight flexure/shear-only vc=1.1√f′c, psi. pp361–362 |
| A.7.5.5.1 | Minimum shear reinforcement above vc/2; exceptions include slabs/footings. p364 |
| 14.1.2 | Cantilever retaining walls: Chapter10 flexure; horizontal minimum14.3.3. p230 |
| 14.3.3 | Horizontal minimum .0020Ag for bars≤No5 and fy≥60ksi; otherwise .0025Ag. p231 |
| 10.5.1 | Flexural minimum max(3√f′c/fy,200/fy)bwd, psi. p112 |
| 10.5.3–10.5.4 | One-third-extra analysis exception; uniform slabs/footings use7.12 instead. p113 |
| 14.3.4 | Two-layer exception says basement walls, not retaining walls. p231 |

Implementation inference: the tapered stem cannot automatically use the **uniform-thickness** slab exception in 10.5.4. A bounded conservative check can use 10.5.1 without invoking the 10.5.3 exception. Uniform toe/heel footing minimum follows 10.5.4→7.12; its exact material-grade branch must be recorded. Do not substitute the vertical wall ratio for this flexural requirement merely because the member is called a wall.

Independent unit conversion: 1 psi = 0.0703069579639 kgf/cm² gives 1.1√f′c in psi = **0.291670052√f′c** in kgf/cm². The textbook coefficient .29 is a slightly lower rounded value. fy=4000 kgf/cm² is about56,893psi, so it does **not** establish Grade60 eligibility. Retaining textbook fs=1700 requires attribution to that textbook/project basis, not relabeling it an exact ACI allowable.

## ACI 318-19: different scope and detailing

Inspected the [original ACI 318-19 publication hosted at UC Berkeley](https://www.ocf.berkeley.edu/~chiep/wp-content/uploads/2024/01/CE-123-ACI-318-19.pdf), first printing June2019. It is a university-hosted copy, not the official errata-updated subscription.

| Clause | Verified distinction (printed page; PDF page = printed page + 2) |
|---|---|
| 11.1.4 | Cantilever retaining walls refer to Chapter13. p165 |
| 13.3.6.1 | Stem designed as a one-way slab using applicable Chapter7 provisions. p198 |
| 13.3.6.3 | Uniform stem critical section at footing interface; tapered/variable stem shear and moment investigated throughout height. p198 |
| 7.6.1.1 / R7.6.1.1 | Minimum flexural steel .0018Ag, near tension face; distinct from distributing temperature steel between faces. p92 |
| 11.7.2.3 | Cantilever retaining walls explicitly excepted from the general two-layer rule. p172 |

The [PCI article by members involved in the ACI code change](https://www.pci.org/PCI_Docs/Publications/PCI%20Journal/2001/May-June/Significant%20Changes%20in%20the%202002%20ACI%20Code.pdf) records removal of the Alternate Design Method from the 2002 edition. Thus neither Chapter11 shearwall equations nor 2019 strength-design shear capacity should be pasted into a service-load WSD check without a defined compatible design basis.

## Consequences for current work

- Source verification now supports a supplementary **ACI318-99 service shear and flexural-minimum check**. It does not establish full EIT2562 or full ACI318-19 compliance.
- A single `MinStemRatio` cannot express flexural minimum, horizontal distribution, material-grade branches, and any opposite-face requirements. The three optimized main-bar pairs are not a complete reinforcement schedule.
- Full acceptance still needs an explicit basis for the chosen edition, compatible material allowables, stem checks along its taper, anchorage/development, spacing/crack control, and whatever detailing the project elects to adopt. Price must include those provided bars before claiming a complete design price.
- Both active and full passive remain the user-authorized project load case. This review introduces no alternate passive load case.

Downloaded standards and extracted full-text working files are local review cache only under `audit/aci-source-cache/`; **do not commit or redistribute these copyrighted full documents**. Citations and the compact clause map above are the deliverable.
