# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## MFU ARIC Proposal Workspace — read first

**MFU-ARIC** = Mae Fah Luang University — AI & Robotics Innovation Center.
A new center being proposed for Thai government funding (FY 2570–2574 / 2027–2031).
This directory is the working area for the **200M proposal submission** under the
**MDES "200M+" computer-system procurement template** (mandatory form, effective 9 Dec 2025;
download at https://200m.mdes.go.th).

The proposal must be authored in Thai, follows the MDES section numbering (ข้อ 1–17 + ภาคผนวก),
and is reviewed by both MDES and the Ministry of Higher Education, Science, Research and
Innovation (อว.). Submitted in the name of **มหาวิทยาลัยแม่ฟ้าหลวง (MFU)** as
the proposing agency (หน่วยงานผู้เสนอโครงการ), with **สำนักวิชาเทคโนโลยีดิจิทัล
ประยุกต์ (ADT)** — School of Applied Digital Technology — designated by the
University as the principal operating unit (หน่วยปฏิบัติงานหลัก) that runs the
centre day-to-day (confirmed 2026-05-18; earlier framing that named ADT as the
sole proposer is stale, as are older drafts that name "สำนักวิชาเทคโนโลยี
สารสนเทศ" — both must be reworded when reused).

**Active template (confirmed 2026-05-15):** `ARIC 200M-2568_V3.docx` — revision V3 dated
พ.ศ. ๒๕๖๘. This **supersedes** the earlier `โครงการ 200M-ARIC.docx` referenced in prior
work. V3 extends the body from ข้อ 1–10 to **ข้อ 1–17** and reshuffles the annex labels
(see "V3 template structure" below).

## Four active workstreams

This workspace now serves **four related but separately-governed deliverables**. Identify
which one the request belongs to before touching files — they have different audiences,
different tone, and different source files.

1. **The MDES "200M+" procurement form** (main track, since 2026-05) — the formal
   ข้อ ๑–๑๗ + ภาคผนวก submission described throughout this file. Audience: MDES and อว.
   reviewers. Files: `sections/`, `annexes/`, `data/`, `document/section_NN.docx`.

2. **The University-Council / President executive report** (added 2026-06/07) — a Thai
   executive report summarising the consultations with **๔ ผู้ทรงคุณวุฒิ** (รศ.ยืน ภู่วรวรรณ ·
   ศ.ดร.จำรัส ลิ้มตระกูล · ศ.ดร.วรภัทร โตธนะเกษม · ดร.ทวินโชค ตันธุวนิตย์, มิ.ย. ๒๕๖๙) and asking
   the Council for approval in principle. Plan: `plan-executive-report.md`. Source of the
   advisors' comments: `_reviews/advisor-comments.docx` + `_sources/MFU_ARIC.pptx`.
   Delivered as `document/ARIC_council_report.html` → `.pdf` (Thai-only, TH Sarabun, A4,
   ~8 pages, no signature block). **The HTML was hand-built, not script-generated** — edit
   it directly and re-export the PDF.
   Direction agreed in this track: **flagship domain = Health & สังคมสูงวัย**, positioning =
   application layer of the Physical AI stack ("สร้างสมอง ไม่สร้างมือกล"), plus short-term
   wins / revenue model / risk answers. `sections/section_01-04.md` and `section_07.md`
   already reflect the health flagship; other sections may not.
   **`document/memo_president_ARIC.docx` no longer belongs to this track**: on 2026-09-03 it was
   rewritten (`document/scripts/rewrite_memo_president.py`) into the cover memo for track 3,
   citing the President's instruction of ๒๕ ส.ค. ๒๕๖๙ to submit an establishment project, and
   it now names two attachments — the establishment form and `document/ARIC_concept_plan.pdf`.

3. **The MFU internal "establishment" form** (โครงการจัดตั้งหน่วยงานใหม่, added 2026-08-31) —
   the University's own **แบบฟอร์มเสนอโครงการจัดตั้งโครงการใหม่/หน่วยงานใหม่ มหาวิทยาลัยแม่ฟ้าหลวง**,
   filled for ARIC as an *organisational* unit. Audience: the President and University
   Council, not MDES. Content source: **`phasing.md`** (two-phase setup plan — Phase 1 launch
   under ADT in Years 1–2, Phase 2 spin-off under มูลนิธิมหาวิทยาลัยแม่ฟ้าหลวง from Year 3;
   มูลนิธิ takes 15 % equity in each spin-out). Scope confirmed 2026-08-31: **only Phase 1 is
   budgeted (ปีงบประมาณ ๒๕๗๐–๒๕๗๑), Phase 2 is described conceptually with no budget, and no
   salary is set for the acting head.** Live artifact:
   `document/ข้อเสนอโครงการจัดตั้งหน่วยงาน ARIC (ฉบับกรอก).docx`, with attachments
   `document/ARIC_concept_plan.pdf` and `document/ARIC_org_structure_A4.pdf`
   (org chart source: `figures/org_chart_establishment.svg`). This budget is **separate from
   and much smaller than** the ฿575M MDES ask (~฿6M over two years) — never merge the two.
   See "Establishment-form toolchain" under Toolchain for why the docx is edited in place.
   A 10-slide Thai deck of `phasing.md` is published at
   https://cnacha-mfu.github.io/aric-plan-deck/ (separate repo `cnacha-mfu/aric-plan-deck`);
   when a number changes in `phasing.md`, that deck's `index.html` must be updated too.

4. **The building space plan / renovation brief** (added 2026-08-26) — `space-plan.md`
   (English, v1.4, design brief for architect / QS / MEP) + `space-plan-th.md` (Thai
   translation for MFU readers — **keep the two in sync**) + `docs/index.html` (a static
   Thai web page presenting the plan, with `docs/*.jpg|png` photos). Programme: a 4-storey
   building, ~400 sqm/floor (≈1,600 sqm gross): F1 "Launchpad" (collaboration/showcase),
   F2 "Academy" (six 46-sqm training rooms, combinable to 60/120 seats), F3 "Engine"
   (specialized labs + machine room), F4 "Bridge" (office/partners). **Renovation cost is
   deliberately outside the ฿575M MDES submission** (equipment-only) and must be funded
   separately, complete before Q2 FY ๒๕๗๑. Open decisions D-01–D-06 live in §2.1 of
   `space-plan.md`; **D-02 records that `sections/section_09.md` and
   `annexes/annex_khor_installation.md` still follow the older DIC/S7A layout and contradict
   this brief** — see gotcha 8.

## Active working plan

**`PLAN_sections_5-10.md`** is the authoring plan for ข้อ ๕–๑๐. Read it before starting
any section-๕–๑๐ work — it lists decisions D1–D9, per-subsection deliverables, the 13-step
execution order, and the 12-figure plan (F1–F12) with the SVG/Mermaid authoring approach.
**Note:** PLAN does NOT cover ข้อ ๑๑–๑๗ added in V3 — those need a new authoring sweep
(task #25).

## Folder layout (reorganized 2026-05-18)

```
G:\My Drive\School\IIC\
├── CLAUDE.md ··················· this file
├── ARIC 200M-2568_V3.docx ····· active MDES template
├── PLAN_sections_5-10.md ······ authoring plan (covers ๕–๑๐ only)
├── plan-executive-report.md ··· plan for the Council/President executive report (track 2)
├── phasing.md ················· two-phase organisational setup plan — source for track 3
├── space-plan.md / space-plan-th.md ··· building renovation brief EN/TH (track 4)
├── docs/ ······················ static Thai web page of the space plan (index.html + photos)
│
├── sections/ ·················· section_01-04.md … section_17.md (active body drafts)
├── annexes/ ··················· annex_a1/a2/a3/khor1/khor/ngor + supporting_personnel_jd /
│                                supporting_rental_terms
├── data/ ······················ equipment_master.csv, asset_rollup.{csv,md},
│                                price_list_clause_mapping.md, price_verification_log.md,
│                                cost-summary.md, vendor_outreach_plan.md
├── scripts/ ··················· all build_*.py, fill_template.py, verify_draft.py,
│                                render_*.py, convert_svgs.py, asset_rollup.py,
│                                make_simple_figures.py, docx_thai_editor.py,
│                                build_/fill_establishment_form.py (track 3)
│
├── figures/ ··················· F01–F12 (+ F12a/b/c/d, F_AI_LLM_dev, F_Robot_dev) SVG/Mermaid
│                                sources + rendered PNGs; org_chart_establishment.svg (track 3);
│                                logo/ (ARIC logo mark + lockup, SVG + PNG)
├── document/ ·················· LIVE artifacts, not throwaway output — see gotcha 6
│   ├── section_05..17.docx ···· per-section docx (amended in place during June rounds)
│   ├── ภาคภนวก ก.docx ········· annex ก (HW+SW spec)
│   ├── ภาคภนวก ง.xlsx ········· annex ง / equipment master — **source of truth for prices**
│   ├── ใบเสนอราคา/ ············ real vendor quotations received (feeds ภาคผนวก ข-๑)
│   ├── ARIC_council_report.{html,pdf} ··· track-2 deliverable
│   ├── ข้อเสนอโครงการจัดตั้งหน่วยงาน ARIC (ฉบับกรอก).docx ··· track-3 live artifact (edited in place)
│   ├── memo_president_ARIC.docx, ARIC_concept_plan.pdf, ARIC_org_structure_A4.pdf ··· track-3 cover memo + attachments
│   ├── scripts/ ··············· ~45 one-off amendment scripts (see Toolchain)
│   ├── _reviews/ ·············· price-audit multi-agent review loop (review.md + iterations)
│   └── _backup/ ··············· automatic docx backups written by docx_thai_editor.py
├── sample/ ···················· reference proposal example (WUH Network 2018)
│
├── _sources/ ·················· external read-only inputs (pptx/docx/xlsx/pdf)
├── _reviews/ ·················· reviewer-author audit trail (rounds 01-03 + revision logs
│                                + advisor-comments.docx for track 2)
└── _archive/ ·················· historical drafts (proposal*, blueprint, executive_summary,
                                 infographic, agent-editor, one-off scripts)
```

In-prose references in the sections/*.md / PLAN files use the `_sources/` / `_archive/`
prefixes. If you encounter an unprefixed legacy reference (e.g., raw `proposal_mdes.md`)
in a future paste-in, normalize it to `_archive/proposal_mdes.md` before saving. Likewise,
references to `section_05.md`, `annex_a1_hw_spec.md`, or `equipment_master.csv` should be
normalized to `sections/section_05.md`, `annexes/annex_a1_hw_spec.md`, or
`data/equipment_master.csv` respectively.

**Naming conventions in use:**
- `_`-prefixed files are **frozen snapshots taken before a large rewrite**
  (`sections/_section_08.before_annexd.md`, `data/_equipment_master.before_annexd.csv`).
  Read them to see what a change replaced; never edit them, and never let them become the
  file a build script reads.
- `*_DEPRECATED.svg` figures (F01, F10, F11) were superseded by split versions
  (F01a/b, F10a/b, F11a/b). Keep the split ones; don't re-render the deprecated files.
- The generated annex files carry **Thai filenames** (`ภาคภนวก ก.docx`, `ภาคภนวก ง.xlsx`) —
  note the spelling "ภาคภนวก" as it appears on disk. Older script constants still point at
  the retired English names (`annex_a.docx`, `annex_d.docx`, `price_audit.xlsx`); fix the
  constant rather than recreating the old file.

## Source documents in this directory

| File | What it is | Use for |
|---|---|---|
| `ARIC 200M-2568_V3.docx` | **The blank MDES template, revision V3** (Thai, พ.ศ. ๒๕๖๘). Defines required structure for ข้อ 1–17 + ภาคผนวก ก-๑/ก-๒/ก-๓/ข-๑/ค/ง. | Authoritative form layout; do not deviate from heading numbering. |
| ~~`โครงการ 200M-ARIC.docx`~~ | **Stale** — earlier template revision. | Do not use; replaced by V3. |
| `PLAN_sections_5-10.md` | Authoring plan for ข้อ ๕–๑๐. Decisions D1–D9, per-subsection deliverables, 13-step execution order, 12-figure plan (F1–F12). | Read first when picking up section ๕–๑๐ work. **Does NOT cover ข้อ ๑๑–๑๗ added in V3** — those need a new authoring sweep. |
| `_sources/MFU_ARIC_Equipment_Specifications_Budget.docx` | Equipment spec sheet, **150M CapEx**, 3 หมวด / 67 รายการ (Network+Cyber 15M, AI Computing 67M, Robotics 68M). Year-by-year purchase plan included. | Source for ข้อ 8 (รายการที่จะจัดหา) and ภาคผนวก ก-๑/ก-๒. **Note budget gap below.** |
| `_sources/MFU_ARIC_Government_Funding_Proposal_2569.docx` | Older narrative version of the proposal. | Background prose, vision, ยุทธศาสตร์ alignment. |
| `_sources/MFU_ARIC.pptx` | Pitch / overview deck for the center. | Visual cues, talking points; not a content source for the form. |
| `_archive/proposal_mdes.md` | A previous **200M MDES-style proposal draft** (full sections 1–4 + room/lab budget breakdown summing to 200M). Different equipment cut from the 150M docx (organized by lab, not by category). | Cross-reference for narrative wording, อัตรากำลัง 25 ตำแหน่ง, KPIs, ผังเครือข่าย, แผน 3 ปี. **Resolve which equipment list is authoritative before writing ข้อ 8.** |
| `_archive/proposal.md`, `_archive/proposal_old.md` | Earlier full drafts. | Historical reference only. |
| `_archive/blueprint.md` | Strategic blueprint (English). Global benchmarks, humanoid-robotics rationale, China/CLMV partnerships, research focus areas. | Source for ข้อ 7.1 (Business Architecture) — ยุทธศาสตร์, พันธกิจ, ผู้ใช้งาน. |
| `_archive/executive_summary_IIH.md` | One-pager. | Quick context. |
| `_sources/ข้อเสนอโครงการ MFU ARIC.pdf`, `_sources/งบเจรจาศูนย์ AI.pdf` | Prior submitted versions / budget-negotiation deck. | Read-only reference. |
| `_sources/2025-ITschool-Asset-Checking-System.xlsx` | **Current asset inventory** of สำนักวิชาเทคโนโลยีดิจิทัลประยุกต์ (ADT) — 2,313 รายการ, ฿46M total. 918 IT-relevant items. Locations: **S7A (1,079), S7B (592), E3A (617)**. Includes Lenovo M72E, Mac mini, Apple iMac, Huawei AirEngine WiFi, CISCO routers, Meta Quest 3S, Oculus Quest 2, IBM/HP servers, Huion/Wacom tablets. | Source for ข้อ 5.1/5.3 (สถานภาพระบบ + อุปกรณ์ที่มีอยู่). Filter to IT only — skip ตู้/โต๊ะ/เก้าอี้. |
| `../Lab/Digital Innovation Lab - School of Applied Digital Technology.pptx` | Existing lab plan deck (37 slides). Shows current floor layout S7A-1F, S7A-2F, S7B-3F, planned upgrades (IoT, Data Analytics, AI Solution, UI-UX, Smart Studio labs), 2-phase model (Utilize → Upgrade), budgets. | Source for ข้อ 5 (current spaces + issues like humidity, leakage), ข้อ 9 (สถานที่ติดตั้ง). |
| `../Lab/สำรวจความต้องการครุภัณฑ์ ประจำปีงบประมาณ 2569 - ADT (1).xlsx` | Equipment needs survey for FY 2569. | Cross-check equipment list. |

## V3 template structure (read this before editing any section/annex)

The V3 form runs ข้อ ๑–๑๗ in the body, plus ๖ annexes:

**Body (ส่วน ก = ข้อ ๑–๔ admin · ส่วน ข = ข้อ ๑–๑๗ project · ส่วน ค = signature page):**
- ข้อ ๑ ชื่อโครงการ
- ข้อ ๒ ส่วนราชการ (incl. **DCIO + MCIO** signature blocks — V3 adds explicit MCIO line)
- ข้อ ๓ วงเงิน · ข้อ ๔ สัดส่วนงบประมาณ (% breakdown)
- ข้อ ๕ การพิจารณาของคณะกรรมการบริหารและจัดหาฯ ของกระทรวง (note: not ส่วน ข๕)
- ข้อ ๖ ความสอดคล้องกับ **๖ ยุทธศาสตร์ของแผนพัฒนาดิจิทัล** (checkboxes)
- ส่วน ข — ข้อ ๑–๑๗:
  - ๑ หลักการและเหตุผล · ๒ วัตถุประสงค์ · ๓ เป้าหมาย
  - ๔ จัดหาใหม่ / ทดแทน / ขยาย-ปรับปรุง (radio)
  - ๕ สภาพปัจจุบัน (๕.๑/๕.๒/๕.๓) · ๖ ระบบงานและปริมาณงาน (๖.๑–๖.๕ incl. buy-vs-rent)
  - ๗ EA (๗.๑ Business / ๗.๒ Application / ๗.๓ Data / ๗.๔ Technology + network/system/security diagrams = รูป ๑–๔; รูป ๕ = equipment-linkage diagram)
  - ๘ รายการที่จัดหา (๘.๑ HW/SW · ๘.๒ บุคลากร · ๘.๓ หน้าที่ความรับผิดชอบของบุคลากร) — ราคากลาง footnote now lists **6 ranked sources** per พ.ร.บ. ๒๕๖๐
  - ๙ สถานที่ติดตั้ง · ๑๐ การฝึกอบรม
  - **NEW in V3 / not in PLAN_sections_5-10.md:**
    - ๑๑ ระยะเวลาดำเนินงาน (๑๑.๑ ระยะเวลา · ๑๑.๒ แผนตลอดโครงการ · ๑๑.๓ Migration Plan · ๑๑.๔ การบำรุงรักษาหลังหมดประกัน)
    - ๑๒ ผลผลิต (เชิงปริมาณ/คุณภาพ)
    - ๑๓ ตัวชี้วัดสัมฤทธิ์ผลด้าน IT (เชิงปริมาณ/คุณภาพ)
    - ๑๔ ความพร้อมของหน่วยงาน (๑๔.๑ บุคลากร ICT ปัจจุบัน · ๑๔.๒ ความพร้อมด้านอื่น)
    - ๑๕ ความเสี่ยงและแนวทางการแก้ไข
    - ๑๖ กฎหมายที่เกี่ยวข้อง
    - ๑๗ ประโยชน์ที่จะได้รับ

**Annexes (V3 labels — corrected 2026-05-15):**
- **ก-๑** รายละเอียดคุณลักษณะของ **ฮาร์ดแวร์** → `annexes/annex_a1_hw_spec.md` ✓
- **ก-๒** รายละเอียดคุณลักษณะของ **ซอฟต์แวร์สำเร็จรูป** → `annexes/annex_a2_sw_spec.md` ✓
- **ก-๓** รายละเอียดคุณลักษณะของ **ซอฟต์แวร์ประยุกต์ (พัฒนาขึ้นเอง)** → `annexes/annex_a3_in_house_apps.md` (status: ไม่ใช้ — all ARIC software is COTS)
- **ข-๑** ใบเสนอราคา (≥ 3 vendors, 3 products) → `annexes/annex_khor1_quotations.md` (placeholder — populated at RFQ phase; current price-discovery evidence lives in `annexes/annex_ngor_3vendor.md` + `data/price_verification_log.md`)
- **ค** ตารางการติดตั้งอุปกรณ์ → `annexes/annex_khor_installation.md` (all 67 lines mapped to Zone/Floor/Room) ✓
- **ง** ตารางเปรียบเทียบราคา ๓ ราย ๓ ผลิตภัณฑ์ (VAT-incl.) → `annexes/annex_ngor_3vendor.md` ✓

**Supporting (non-annex) reference files:**
- `annexes/supporting_personnel_jd.md` — JD detail backing §8.2/§8.3 (formerly mislabeled `annex_a3_personnel.md`)
- `annexes/supporting_rental_terms.md` — multi-year commitment detail backing §6.5 + §8.1 lines ๗-๘ (formerly mislabeled `annex_kor_rental.md`)

**Authoring status (last verified 2026-08-26):**
- ข้อ ๑–๔: drafted (`sections/section_01-04.md`) ✓ — 5-yr form total ฿575M (Y1 ฿235M + Y2–Y5 ฿85M × 4).
- ข้อ ๕–๑๐: drafted (`sections/section_05.md` … `sections/section_10.md`) ✓
- ข้อ ๑๑–๑๗: drafted (`sections/section_11.md` … `sections/section_17.md`) ✓ — total ~13,300 words. Open user-confirmations: post-warranty maint projection (§11.4), risk P/I scores (§15), KPI audit targets (§13), Council resolution timing (§14.2 D / §16.1 C), CCTV exclusion language (§16.4 L).
- Annexes: all 6 V3 annexes present ✓; ก-๓ and ข-๑ are placeholders by design.
- Generated docx for ข้อ ๐๕–๑๗ + ภาคผนวก ก/ง exist under `document/` and carry June amendments
  that the markdown does not always reflect — see gotcha 6.

## Key constraints and gotchas

1. **Confirmed scope (2026-05-15, 5-year expansion):**
   - **Project window:** ปีงบประมาณ ๒๕๗๑–๒๕๗๕ (FY 2028–2032, 1 Oct 2027 – 30 Sep 2032).
   - **5-year form total = ฿575M** = ฿150M Year-1 CapEx + ฿85M/year OpEx × 5 years.
   - **Year 1 (๒๕๗๑)**: ฿235M = ฿150M CapEx (all 46 equipment lines) + ฿85M Year-1 OpEx.
   - **Years 2–5 (๒๕๗๒–๒๕๗๕)**: ฿85M/year OpEx only — **no new equipment**.
   - Multi-year commitment requires งบผูกพันข้ามปี authorisation per พ.ร.บ. วินัยการเงินการคลังของรัฐ พ.ศ. ๒๕๖๑ มาตรา ๒๖ + University Council resolution + Cabinet approval.
   - **Do not** inflate the Year-1 equipment list above ฿150M, and do not add equipment to Years 2–5, unless re-authorized.
   - Prior single-year framing (฿235M / ฿235M) is superseded by this 5-year ฿575M scope.

2. **MDES form splits IT vs Non-IT and ในเกณฑ์/นอกเกณฑ์ราคากลาง** — ข้อ 8.1 table has 4 sub-tables
   (HW-in / SW-in / HW-out / SW-out) plus Non-IT. **Current cut (2026-06-19): 46 lines in 9 หมวด
   totalling exactly ฿150,000,000 — HW-out 35 · SW-out 6 · Non-IT 5 · HW-in/SW-in = 0.**
   Every remaining line's spec exceeds the MDES central-price-list ceiling, so nothing sits
   ในเกณฑ์ any more (§8 carries an explicit note saying so). The older "67 รายการ with some
   ในเกณฑ์ PCs/laptops/APs" description is stale — do not restore it, and if you add a line
   that *is* within the price list, re-open the HW-in sub-table deliberately rather than
   filing it under HW-out.

3. **Equipment procurement = Year 1 only (FY ๒๕๗๑):** all 150M of CapEx ครุภัณฑ์ is
   procured in **ปีงบประมาณ ๒๕๗๑** (single year for the equipment). Years 2–5 carry **only
   operating expense** (personnel + training + ops cloud/license + maintenance) — **no new
   equipment in Years 2–5**. Section 8 has the Year-1 procurement tables plus a Y2–Y5 OpEx
   table within the same form (multi-year filing per the MDES V3 template footnote: "หาก
   งบประมาณในข้อ ๓ มีมากกว่า ๑ ปี ให้จัดทำตารางรวม และตารางแยกรายปี ตามจำนวนปีในข้อ ๓").

4. **Thai numerals** — the form uses Thai numerals (๑–๙) for section numbering and dates.
   Body content can be Arabic numerals (the existing equipment docx uses Arabic). Keep
   section headers Thai-numeral, table contents Arabic.

5. **ราคากลาง source must be specified per line.** V3 footnote (per พ.ร.บ. ๒๕๖๐) lists
   **6 ranked sources** the team must pick from:
   1. ราคาที่คำนวณตามหลักเกณฑ์ที่คณะกรรมการราคากลางกำหนด
   2. ฐานข้อมูลราคาอ้างอิงของกรมบัญชีกลาง
   3. ราคามาตรฐานของสำนักงบประมาณ / หน่วยงานกลางอื่น
   4. ราคาที่ได้จากการสืบราคาจากท้องตลาด
   5. ราคาที่เคยซื้อ/จ้างครั้งหลังสุดภายใน ๒ ปีงบประมาณ
   6. ราคาอื่นใดตามหลักเกณฑ์ของหน่วยงานนั้น ๆ
   Most ARIC out-of-spec items will use source (4) สืบราคา + ภาคผนวก ข-๑ (≥ 3 quotations) and
   ภาคผนวก ง (3-vendor comparison). Annex citation in `data/price_list_clause_mapping.md` already
   identifies MDES central price-list edition พ.ค. ๒๕๖๙ for in-spec items.

6. **Know which file is authoritative before you edit anything.** The project moved from
   "markdown is the source, docx is generated" to "docx/xlsx is the live artifact" during the
   June price-audit rounds, and the two layers have drifted:
   - **Equipment list / prices → `document/ภาคภนวก ง.xlsx`, sheet `Clean Items`.**
     `build_section_08.py` states it plainly: *"source of truth (CSV retired)"*.
     `data/equipment_master.csv` is a frozen earlier snapshot (still what `fill_template.py`
     reads) — treat it as historical unless you deliberately re-sync it from the xlsx.
   - **Section body text → `document/section_NN.docx`** for ๐๕–๑๗. These were amended in place
     by the one-off scripts in `document/scripts/`; the matching `sections/section_NN.md` files
     were **not** always updated (compare mtimes before trusting either).
   - **`ARIC 200M-2568_V3_draft.docx` (2026-05-18) is stale** — it predates the June equipment
     re-cut. Re-running `scripts/fill_template.py` regenerates it from the *retired* CSV, so
     re-sync the CSV from the xlsx first or the draft will silently revert to the old cut.
   - When a section's md and docx disagree, ask the user which one they want to carry forward
     rather than guessing; state the drift explicitly.

7. **Existing assets live across three buildings (S7A, S7B, E3A)** — not just one lab.
   The new ARIC should integrate with, not duplicate, existing networking (Huawei AirEngine
   WiFi 6, CloudEngine S3710H switches) and existing compute (Mac mini cluster, Apple iMac
   labs). Mention this in ข้อ 5.1 to show context.

8. **Installation site (confirmed 2026-05-14):** **อาคาร S7A ชั้น 1–4** organized into **3 zones**
   (per https://dic-website-853767948689.us-central1.run.app/ — the official DIC master plan).
   1,500 sqm total across 4 floors:
   - **Zone 1: Collaboration & Startup Ecosystem** (500 sqm) — Pitch Arena (151 sqm, 100–150 seats),
     co-working space (157 sqm), innovation café, 5 meeting rooms, reception, community pantry.
   - **Zone 2: Specialized Labs** (500 sqm) — 5 labs × 100 sqm each: Software Intelligence,
     Vision & Neural, **Immersive Game Studio** (note: replaces "Generative AI Lab" from
     _archive/proposal_mdes.md), Autonomous Systems, Industrial AI & IoT.
   - **Zone 3: Office & Training** (500 sqm) — Talent pool office (150 sqm), 3 training rooms
     (60 seats each, 90 sqm each), 1 meeting room (20–30 seats).
   GPU Server Room and Data Center are not explicitly listed by the master plan — likely
   tucked inside Zone 2 or a service corridor. Resolve location for the GPU cluster + rack
   before ข้อ 9 finalization.
   **Superseding programme (2026-08-26, not yet propagated):** `space-plan.md` v1.4 lays out a
   different programme — 4 floors × ~400 sqm, training on F2, specialized labs + machine room
   on F3 (decision D-01, conditional on a structural/lift survey). Its decision D-02 says the
   submission should adopt the brief and **re-issue ข้อ ๙ and ภาคผนวก ค**, which still follow
   the 3-zone layout above. Until the user confirms D-02, treat ข้อ ๙ / ภาคผนวก ค / F12a–d
   as the *submitted* layout and `space-plan.md` as the *intended* one, and say which you used.

9. **The lab pptx shows real pain points** — S7A humidity/leakage from neighboring wet lab,
   undersized rooms, no weekend access, old furniture. Use these verbatim in ข้อ 5.2
   (สภาพปัญหา) — they are documented student/faculty feedback.

## File outputs the user will eventually need

- A `.docx` filled into the MDES template (`ARIC 200M-2568_V3.docx`) — the final submission.
- ภาคผนวก **ก-๑** (HW spec), **ก-๒** (packaged SW spec), **ก-๓** (in-house app spec — if any),
  **ข-๑** (≥ 3 vendor quotations), **ค** (installation-location table), **ง** (3-vendor price comparison).
- `figures/` subdirectory holding F1–F12 source files (Mermaid `.mmd` + hand-authored SVG)
  and rendered PNGs to embed in the docx. See `PLAN_sections_5-10.md` for the list.
- SVG → PNG rendering is done by `convert_svgs.py` (Playwright/headless Chromium); the
  pipeline `python make_simple_figures.py && python convert_svgs.py` regenerates the
  รูปที่ ๑–๕ image set.
- Working drafts in Markdown are fine for review iterations before transferring into the docx.
- Track 2 (Council/President): `document/ARIC_council_report.html` + `.pdf` — hand-authored,
  Thai-only, no signature block.
- Track 3 (establishment form): `document/ข้อเสนอโครงการจัดตั้งหน่วยงาน ARIC (ฉบับกรอก).docx` +
  cover memo `document/memo_president_ARIC.docx` + attachments `ARIC_concept_plan.pdf`,
  `ARIC_org_structure_A4.pdf`.
- Track 4 (space plan): `space-plan.md` / `space-plan-th.md` and the `docs/index.html` page;
  eventually excerpted into ข้อ ๙ and ภาคผนวก ค once D-02 is settled.

## Toolchain — Python scripts in this repo

All scripts now live under `scripts/`. They are standalone, invoked from the **repo root**
(`python scripts/<name>.py`), and force UTF-8 stdout (Windows PowerShell mangles Thai
output otherwise). Scripts that use relative paths self-locate via
`os.chdir(Path(__file__).resolve().parent.parent)` so they keep working regardless of CWD.
Dependencies: `python-docx`, `openpyxl`, `python-pptx`, `playwright` (with
`playwright install chromium`).

| Script (under `scripts/`) | Purpose | Inputs → Outputs |
|---|---|---|
| `asset_rollup.py` | Roll up the ADT asset register into the ข้อ ๕.๓ table; filters to IT-relevant categories. | `_sources/2025-ITschool-Asset-Checking-System.xlsx` → `data/asset_rollup.csv` + `data/asset_rollup.md` |
| `make_simple_figures.py` | Generate simplified SVGs for รูปที่ ๑–๕ (large Thai labels for A4 embed). | (none) → `figures/simple_fig{1..5}_*.svg` |
| `convert_svgs.py` | Render the simple SVGs (and other selected figures) to high-DPI PNGs via headless Chromium. | `figures/*.svg` → `figures/*.png` |
| `render_mermaid.py` | Render the F04/F06/F07 Mermaid swimlane/flow diagrams (incl. F06a–d) via Playwright + mermaid CDN. | `figures/F0*.mmd` → `figures/F0*.png` |
| `render_section_05_figs.py` | Render the 3 Section-๕ SVGs (F01a current logical, F01b current compute, F02 floor plan) to PNG. | `figures/F01a/F01b/F02*.svg` → `figures/*.png` |
| `render_f03.py` | Render the F03 old/new integration SVG to PNG. | `figures/F03_old_new_integration.svg` → `figures/F03*.png` |
| `fill_template.py` | Populate the V3 docx template with Thai content, equipment/personnel tables, and embedded figure PNGs. **Authoritative draft generator.** | `ARIC 200M-2568_V3.docx` + `data/equipment_master.csv` + `figures/*.png` → `ARIC 200M-2568_V3_draft.docx` |
| `verify_draft.py` | Sanity-check the generated draft (table row counts, header text, image count, CapEx/OpEx totals). Run after every `fill_template.py`. | `ARIC 200M-2568_V3_draft.docx` → stdout report |
| `build_section_05.py … build_section_17.py` | Section-by-section standalone docx generators (formal Thai, A4). | `data/*.csv` + embedded prose → `document/section_NN.docx` |
| `build_annex_a.py` | Combined ภาคผนวก ก docx from annex_a1/a2 markdown. | `annexes/annex_a1_hw_spec.md` + `annexes/annex_a2_sw_spec.md` → `document/annex_a.docx` (on disk today: `ภาคภนวก ก.docx`) |
| `build_annex_d.py` | ภาคผนวก ง 3-vendor comparison docx. | `annexes/annex_ngor_3vendor.md` + `data/equipment_master.csv` → `document/annex_d.docx` (superseded by `ภาคภนวก ง.xlsx`) |
| `build_price_audit_xlsx.py` | Team price-audit workbook — **the workbook it created was renamed to `document/ภาคภนวก ง.xlsx`; re-running it recreates `price_audit.xlsx` from the retired CSV and will NOT reflect the current cut.** | `data/equipment_master.csv` + `annexes/annex_ngor_3vendor.md` → `document/price_audit.xlsx` |
| `docx_thai_editor.py` | In-place Thai-typography fix-up agent (9 spacing/paragraph rules) on any docx. | `<docx_path>` CLI arg → overwrites file (backup in `document/_backup/`) |
| `fill_establishment_form.py` | **Track 3.** Converts the University's blank `.doc` establishment form to `.docx` via Word COM (`pywin32`, retries on "Call was rejected by callee") and fills it from `phasing.md` Phase-1 numbers. | `ข้อเสนอโครงการจัดตั้งหน่วยงานARIC.doc` (repo root — **currently missing**, see below) → `document/ข้อเสนอโครงการจัดตั้งหน่วยงาน ARIC (ฉบับกรอก).docx` |
| `build_establishment_form.py` | Superseded from-scratch generator of the same form (structure extracted from the `.doc` via olefile). Imports the `thai-docx` skill module from `C:\Users\DELL\.claude\skills\thai-docx`. | (none) → `document/ข้อเสนอโครงการจัดตั้งหน่วยงาน ARIC.docx` (not on disk today) |

### Establishment-form toolchain (track 3) — edit the docx in place

- The blank University `.doc` template that `fill_establishment_form.py` reads **is no longer in
  the workspace** (noted in `document/scripts/apply_traction_and_board.py`). Do **not** re-run
  `fill_establishment_form.py` unless the user restores the `.doc`; every change since
  2026-08-31 has been applied to the live docx by one-off scripts in `document/scripts/`
  (`apply_traction_and_board.py`, `apply_budget_recut.py`, `apply_supplies_taper.py`,
  `apply_traction_10_8_3.py`, `apply_roi_valuation.py`), each mirrored into the fill script
  "in case the template comes back". Follow the same pattern for new amendments.
- `fill_establishment_form.py` also hard-codes a `SCRATCH` path from an earlier session; point
  it at the current scratchpad before running.
- **Budget figures drift between `phasing.md` and the docx.** `phasing.md` (last edited
  2026-09-17) shows Phase-1 Y1 ฿3,065,250 / Y2 ฿2,485,250; the fill script's docstring shows
  ฿3,278,500 / ฿2,778,500 (total ฿6,057,000) after the 2026-08-31 recut. Check which figure
  the user wants before quoting a Phase-1 total, and keep `phasing.md`, the docx, and the
  published deck in step.

Normal regeneration order (run from repo root):
```
python scripts/asset_rollup.py            # only when the source xlsx changes
python scripts/make_simple_figures.py     # only when figure design changes
python scripts/convert_svgs.py            # after any SVG edit
python scripts/fill_template.py           # after any sections/*.md or data/*.csv edit
python scripts/verify_draft.py            # always, to confirm the draft is intact
```

If Word has `ARIC 200M-2568_V3_draft.docx` open, `fill_template.py` will fail to write
(lock file `~$IC 200M-2568_V3_draft.docx` will be present). Close Word first.

**Caution: this regeneration order reflects the pre-June "markdown → docx" model.** Since the
June equipment re-cut, `data/equipment_master.csv` is retired in favour of
`document/ภาคภนวก ง.xlsx` (gotcha 6), so `fill_template.py` / `build_price_audit_xlsx.py`
will regenerate the *old* 150M cut unless the CSV is re-synced from the xlsx first. Confirm
with the user before running either against the live artifacts.

### `document/scripts/` — one-off amendment scripts

A second, differently-governed script family (~45 files) that edits the **live docx/xlsx in
place** rather than regenerating from markdown. They are single-use, dated by intent, and
document their reasoning in the module docstring — read that docstring before assuming what
a script did. Recurring families:

- `apply_iteration_N.py` / `apply_*_findings.py` / `apply_procurement_decisions.py` — apply one
  round of the price-audit review loop (add/drop/reprice items, keeping the ฿150M cap).
- `expand_to_150M.py`, `adjust_to_150M.py`, `drop_*.py`, `merge_thai_vendors.py`,
  `rebuild_annex_d_*.py` — equipment-list surgery.
- `reconcile_*.py`, `align_section_09_to_08.py`, `fix_*.py` — cross-section consistency fixes
  (headcount, floor plan, table 12.1, stale figure refs).
- `refresh_embedded_images.py` / `refresh_all_embedded_images.py` / `resize_embedded_images.py` /
  `insert_infographics_into_docx.py` — re-embed the latest `figures/*.png` into the section
  docx by matching the **Thai figure caption** in the following paragraph.
- `check_urls.py` / `cleanup_and_fix_urls.py` / `final_url_cleanup.py` / `strip_comments.py` —
  vendor-URL liveness auditing; `strip_comments.py` exists because URL annotation left orphan
  Word comment references that made Word report "unreadable content".
- `make_aric_infographics*.py` (v1→v4) — the ARIC capability infographics; **v4 is the
  vendor-neutral version** (all brand and Thai-distributor names stripped) and is the one to
  reuse. Earlier versions name vendors and must not go into the submission.
- `apply_traction_and_board.py` / `apply_budget_recut.py` / `apply_supplies_taper.py` /
  `apply_traction_10_8_3.py` / `apply_roi_valuation.py` — **track 3** amendments to the
  establishment-form docx (traction targets, budget lines, ROI / equity-valuation prose).
- `rewrite_memo_president.py` / `memo_add_attachments.py` — the 2026-09-03 rewrite of the
  President memo into the track-3 cover letter, and the two-attachment wording.

When writing a new amendment, follow the same pattern: a docstring stating *why* + the exact
deltas, hard-coded absolute paths, UTF-8 stdout, and a backup before overwriting.

## Review / revision-log workflow

The `_reviews/` folder is the audit trail between reviewer and author. Each round is two
files: `review_round_NN.md` (reviewer findings) and `revision_log_round_NN.md` (author's
response — what changed, what was deferred, what was rejected with reasoning).
`revision_log_specs.md` tracks spec-level changes to the annex ก-๑/ก-๒ documents
independent of round numbering. When closing a round, append entries here rather than
editing the original review file. `_reviews/advisor-comments.docx` belongs to track 2 (the
Council report), not to these rounds.

A **second, separate loop** lives in `document/_reviews/`: the equipment price-audit run over
`ภาคภนวก ง.xlsx`, orchestrated as four parallel expert personas (Innovation / Procurement /
Government-funding / Technical) with explicit convergence criteria in `review.md` — Year-1
CapEx within ฿0.5M of ฿150M, zero dead vendor URLs, ≥ 2 vendors per row, ≥ 70 % of rows 🟢
(price visible at a public URL), max 5 iterations. Findings from iteration N are applied by
`document/scripts/apply_iteration_N.py` and logged in `review_iteration_N.md`. If you re-open
this loop, keep the tier legend (🟢 public list price · 🟡 manufacturer page + industry
report · 🔴 sole-source / contact-sales) and the FX convention (฿35/USD + ~10 % Thai dealer
markup + 7 % VAT) used by `data/price_verification_log.md`.

## Tooling notes (Windows / PowerShell)

- This is a Google Drive mirror — paths contain spaces and Thai characters.
  Quote everything. When calling python, force UTF-8: `sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')`.
- **Drive-mirror I/O is slow.** A recursive `find .` over the workspace times out. Use
  targeted `ls <dir>` / `Glob` / `Grep` instead of whole-tree walks.
- `python-docx` reads the templates fine. `openpyxl` for the xlsx.
- `python-pptx` for the lab deck.
- `~$…` files mean Word/Excel has that document open — writing to the real filename will
  fail. Close the app first (`~$MFU_ARIC.pptx` and
  `~$U_ARIC_Equipment_Specifications_Budget.docx` have both appeared).
- Fonts: **TH Sarabun New / TH SarabunPSK**, body 14–16 pt. `scripts/build_section_*.py`
  writes `w:rFonts` for `ascii`/`hAnsi`/**`cs`** — the `cs` (complex-script) attribute is what
  actually controls Thai glyphs; omitting it produces the classic broken-vowel rendering.
- Relevant skills for this workspace: **`thai-docx`** (authoring/repairing Thai .docx — use it
  instead of hand-rolling font and justification fixes) and **`adt-school-management`**
  (ADT school data, official MFU forms, EdPEx/SAR context).
- `.claude/settings.local.json` pre-allows `python`, `pip`, `git add/commit`, `WebSearch`, and
  `WebFetch` on the MDES/อว./DIC domains used for price and policy verification.

## How to author Thai content for this project

- Project name (Thai): **โครงการจัดตั้งศูนย์นวัตกรรมด้านปัญญาประดิษฐ์และหุ่นยนต์ มหาวิทยาลัยแม่ฟ้าหลวง**
- Project name (English): **MFU AI & Robotics Innovation Center (MFU-ARIC)**.
  (Earlier drafts also use **Intelligence Innovation Hub / IIH** — current canonical short name
  is **MFU-ARIC**. Don't mix.)
- Address: 333 หมู่ 1 ต.ท่าสุด อ.เมือง จ.เชียงราย 57100.
- Responsible unit (confirmed 2026-05-14): **สำนักวิชาเทคโนโลยีดิจิทัลประยุกต์ (ADT)** —
  School of Applied Digital Technology. (Not the IT school named in older _archive/proposal_mdes.md
  drafts — those need to be reworded when reused.)
- Tone: formal, official Thai administrative language. Avoid translation of English buzzwords
  where a standard Thai term exists; keep English-only for vendor/model names.
