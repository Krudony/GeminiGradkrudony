# Excel Grade XML Mapping (Safe Method)

This reference documents the specific XML files and cell locations for editing ปพ.5.

## Primary Level (ประถม)
- **Main/Front Page:** `xl/worksheets/sheet1.xml` (rId1)
- **Attendance (Semester 1):** `xl/worksheets/sheet5.xml` (rId5)
- **Attendance (Semester 2):** `xl/worksheets/sheet6.xml` (rId6)
- **Scores (Both Semesters):** `xl/worksheets/sheet8.xml` (rId8)
  - Semester 1 Columns: `I–BI`
  - Semester 2 Columns: `BJ–DV`
  - Midterm Score: `BH` (Semester 1), `DI` (Semester 2)
  - **Critical Rule:** Scores in Semester 2 (`BJ–BQ`) should not be 0 (recommended ≥ 5).
  - **Critical Rule:** Row 7 is the header for full scores. Do not delete. Set `BJ7:BQ7 = 10` for a total base of 200.
- **Characteristics (คุณลักษณะ):** `xl/worksheets/sheet9.xml` (rId9)
- **Reading/Thinking (อ่านคิดวิเคราะห์):** `xl/worksheets/sheet10.xml` (rId10)
- **Competencies (สมรรถนะ):** `xl/worksheets/sheet11.xml` (rId11)
  - Semester 1 Columns: `H, L, P, T, X`
  - Semester 2 Columns: `I, M, Q, U, Y`

## Secondary Level (มัธยม)
- **Main/Front Page:** `xl/worksheets/sheet1.xml` (rId1)
- **Scores (คะแนน1):** `xl/worksheets/sheet6.xml` (rId6)
- **Characteristics (คุณลักษณะ):** `xl/worksheets/sheet7.xml` (rId7)
- **Reading/Thinking (อ่านคิดวิเคราะห์):** `xl/worksheets/sheet8.xml` (rId8)
- **Competencies (สมรรถนะ):** `xl/worksheets/sheet9.xml` (rId9)

## Sheet "Scores" Columns

### Primary (sheet8.xml)
#### Semester 1 (I–BI)
| Excel | Name | Type | Notes |
|---|---|---|---|
| I–P | Indicators 1–8 | INPUT | max=10 |
| BG | Sum in-between | FORMULA | SUMIF(I:BF,"<>-1") |
| BH | Midterm exam | INPUT | max=20 |
| BI | Total sem 1 | FORMULA | SUMIF(BG:BH,"<>-1") |

#### Semester 2 (BJ–DV)
| Excel | Name | Type | Notes |
|---|---|---|---|
| BJ–BQ | Indicators 1–8 | INPUT | max=10 (Ensure elements are created) |
| DH | Sum in-between | FORMULA | SUMIF(BJ:DG,"<>-1") |
| DI | Final exam | INPUT | max=20 |
| DJ | Final proportion | FORMULA | ROUND(DI8*DJ$7/DI$7,0) |
| DN | Sem 2 total | FORMULA | DH8+DM8 |
| DQ | Year total | FORMULA | DO8+DP8 |
| DR | 100% Score | FORMULA | ROUND(DQ8/DQ$7*DR$7,0) |

### Secondary (sheet6.xml)
| Excel | Name | Type | Notes |
|---|---|---|---|
| I–O | Indicators 1–7 | INPUT | max=10 |
| BG | Midterm exam | INPUT | max=10 |
| BH | Sum in-between | FORMULA | SUMIF(I:BG,"<>-1") |
| BI | Final raw score | INPUT | max=40 |
| BJ | Final calculated | FORMULA | ROUND(BI*BJ$7/BI$7,0) |
| BN | Total score | FORMULA | SUM(BH,BM) |
| BO | 100% Score | FORMULA | ROUND(BN/BN$7*BO$7,0) |

## Core Rules for All Levels
1. **Backup First:** Always create a `[filename]_backup.xlsx` before any modification.
2. **Direct XML Editing:** Unzip, edit XML, then Zip back. Avoid standard Excel libraries that re-save the entire file and strip elements.
3. **Text Values:** Use `<c r="..." t="str"><v>VALUE</v></c>` for text values to prevent Excel repair errors.
4. **Formula Re-calculation:** Delete `xl/calcChain.xml` and its reference in `[Content_Types].xml` to force Excel to recalculate upon opening.
5. **Auto-detection:** Detect student count from Column C (Student ID) which contains formulas. Do not hardcode student counts.
6. **Cached Values Update:** Always update the `<v>` inside `<c>` for formulas if modifying dependent inputs so the view is correct before saving.
