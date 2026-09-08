# EIT WSD primary-source search

Update: the user subsequently authorized ACI as a supplementary source. Original ACI 318-99 and 318-19 pages have now been inspected; see [ACI clause review](ACI_WSD_SOURCE_REVIEW.md) and [native supplementary results](ACI_SUPPLEMENT_RESULT_TH.md). The EIT-only search outcome below is historical and does not mean that no usable ACI clauses are available.

Search date: 2026-09-08. Scope: public EIT/COE/institutional sources for EIT 011007-19, with emphasis on actual clauses or official errata. No VB6 source or criteria flags were changed by this search.

## Verified publication identity

- The publisher's [EIT Engineering Journal 78(4), catalogue p. 90](https://eit.or.th/showcase/EIT/issue4_68/files/basic-html/page90.html) lists the reinforced-concrete working-stress standard as **011007-19**, ISBN **978-616-396-023-8**. This verifies book identity, not the numerical requirements inside it.
- The [Council of Engineers announcement 109/2568, PDF p. 4, row 22](https://coe.or.th/wp-content/uploads/2026/01/109_2568-รายชื่อมาตรฐานการให้บริการวิชาชีพที่สภาวิศวกรให้การรับรอง.pdf) identifies the WSD standard as **second revision, November 2562**, with a recognition period of 3 April 2566 through 2 April 2571. The web tool returned its text; screenshot retrieval timed out. This is a recognition listing, not a source of design clauses.
- The university's [KMUTNB reinforced-concrete course specification, PDF p. 5](https://cit.kmutnb.ac.th/wp-content/uploads/2022/04/TQF3-RC-Design_NMT25มิถุนายน2564.pdf) recommends 011007-19, second revision, 2562. This supports edition identification only.

## Actual-clause verification status

No public publisher-issued text or official errata containing the following requirements was located in this bounded search:

| Requirement | Status against actual 011007-19 clauses |
| --- | --- |
| Concrete allowable flexural compression and applicable conditions | Unverified |
| Steel tensile allowable stresses and grade-specific caps | Unverified |
| Concrete modulus, steel modulus, modular ratio and rounding | Unverified |
| Shear stress definition: V/(bd) versus V/(bjd), allowable stress and critical section | Unverified |
| Minimum vertical/horizontal stem reinforcement and applicability to a retaining wall | Unverified |
| Toe/heel minimum flexural and shrinkage reinforcement, spacing and detailing | Unverified |
| Cover and effective-depth detailing requirements | Unverified |

This outcome means that this search has not verified those clauses. It does not establish that the standard is unavailable elsewhere or that the formulas used in any existing program are wrong. Publication metadata, a citation to the standard, and passing software tests cannot establish compliance with its unexamined clauses.

## Useful leads, with authority limits

- A [publisher-hosted textbook preview](https://bundanjai-static.reeeed.com/book/ckabwe78dfhb7078916f86s9w/preview/9786160838721_PDF.pdf?supportedpurview=project), printed p. 29, explicitly attributes a concrete-cover table to 011007-19 and describes cover as measured to the outermost steel surface. It is an educational interpretation, not the standard publisher's clause text; search text has Thai font extraction problems. No numerical requirement was promoted to verified status from this preview.
- The original [KMUTNB cantilever-stair optimization paper](https://ojs.kmutnb.ac.th/index.php/joindtech/article/view/7592) explicitly cites 011007-19. Its implementation can be studied as research evidence, but it cannot by itself verify the standard's retaining-wall clauses or their complete applicability.
- Followed the publisher link to the [18-page article PDF](https://ojs.kmutnb.ac.th/index.php/joindtech/article/download/7592/5498), DOI 10.14416/j.ind.tech.2025.04.006. Printed p. 74, section 2.2 was inspected as text and a rendered page: it states checks for concrete moment, steel moment, concrete shear and cantilever deflection (equations 1–4), citing references 14–16. That section does not supply the missing numerical shear/minimum-steel rules for this retaining wall. A paper citing the code therefore does not resolve these missing implementation requirements.
- The textbook preview ends before its substantive WSD allowable-stress chapter (table of contents points to printed p. 38); publicly visible introductory pages do not provide a complete replacement for the standard.
- Unauthenticated uploads of standards, lecture notes, and software manuals appeared in search results. They were not accepted as publisher-authoritative evidence, and requirements from different editions or strength-design standards were not substituted.

## Search coverage and next concrete step

Queries combined 011007-19, the Thai WSD title, 2562, November/second revision, shear, clause and correction/errata terms, including site restrictions to eit.or.th and coe.or.th. The public EIT catalogue provides a route to identify the exact edition for library consultation or obtaining the publisher's copy. No purchase, account request, or external message was made.

When authoritative pages are available, record each applicable clause/table/page, unit system, material grade and applicability before changing production verification status. Until then, distinguish a research-source calculation from a design claimed to be fully verified to EIT 011007-19.

## Practical outcome for this project

The user does not have a standard file. The search was performed independently; no further upload is assumed. A physical publisher's copy or library consultation would also permit clause verification—PDF ownership is not required. This search did not obtain a complete numerical design specification, so it made no change to WSDReviewed, engineering formulas, or the compiled program and did not run a research batch. Temporary downloaded research-paper copies and renders were used for inspection only; they are not republished in this repository.
