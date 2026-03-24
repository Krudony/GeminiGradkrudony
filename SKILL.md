---
name: excel-grade-krudony
description: Safe Excel grade and score adjustment for ปพ.5 files (Krudony's method). Preserves drawings, charts, and XML integrity.
---

# Excel Grade Krudony (Safe Method)

This skill implements the "Safe Method" for editing school grading files (ปพ.5), ensuring that direct XML manipulation is used to preserve the integrity of the workbook. 

The traditional `openpyxl.save()` method strips out images, drawings, and charts from ปพ.5 files. This method extracts the XLSX, edits the internal XML, and repacks it safely.

## Core Workflow

1. **Pre-edit Audit (Mandatory):** Before making any changes to the Main Sheet (Sheet1), the agent MUST read and report ALL existing key values (School, Subject, Teacher, Year, etc.) to the user for verification. Do NOT skip any fields.
2. **Identify Level:** Determine if the file is for Primary (ประถม) or Secondary (มัธยม).
   - **Primary:** คะแนน 1 (Scores) has *both* Semester 1 and 2 in `sheet8.xml`. Attendance is split into `sheet5.xml` (Sem 1) and `sheet6.xml` (Sem 2).
   - **Secondary:** คะแนน 1 (Scores) is usually in `sheet6.xml`.
2. **Execute Safe Edit Script:** Use the bundled Python script functions to handle the Zip/Unzip and XML cleanup.
   ```bash
   python .gemini/skills/excel-grade-krudony/scripts/xlsx_safe_edit.py <action> ...
   ```
3. **Manual Verification:** If needed, verify the XML structure using `references/mapping.md`.

## Critical Safety Rules

1. **Backup First:** Always create a `_backup.xlsx` file before starting.
2. **No Standard Save:** Do NOT use standard libraries like `openpyxl`'s `.save()` method.
3. **XML Cleanup:** Always delete `xl/calcChain.xml` and its reference in `[Content_Types].xml` to force Excel to recalculate formulas upon opening.
4. **Encoding:** Ensure all XML edits use UTF-8. 
5. **Text Values:** When setting text strings in XML `<v>`, ensure the cell `<c>` has `t="str"` (or `t="s"` if referencing shared strings) to prevent Excel repair errors.
6. **Do NOT Modify Row 7 (Header):** Row 7 contains the maximum score values. Modifying it can corrupt formula scaling. For Semester 2 (Primary), ensure `BJ7:BQ7` are set to `10`.
7. **Cached Values:** When changing inputs that affect formulas (like BG, BH, DI, etc.), you *must* also update the formula's cached value in the XML, otherwise Excel might miscalculate initially.

## Features

- Update Semester 1 & 2 Scores and calculate grade boundaries.
- Update Attributes (คุณลักษณะ), Reading/Thinking (อ่านคิดวิเคราะห์), and Competencies (สมรรถนะ).
- Fill automated Attendance (เวลาเรียน 1 & 2) considering Thai public holidays.
- Auto-detect student rows from column C formulas.

## Triggers
- "Update grades in [file]"
- "Change scores for student [name]"
- "Fix attendance in ปพ.5"
- "Safe edit excel"

## Expert Knowledge: Por Por 5 Attendance (Sheet 6)
- **Month Selectors:** Row 4 (H4:EU4) contains List Box selectors for months. Formulas in Row 6 (Dates) depend on these.
- **Weekly Mapping (Sem 2):** 
  - Week 1-4 (H, N, T, Z) -> 'พฤศจิกายน'
  - Week 5-8 (AF, AL, AR, AX) -> 'ธันวาคม'
  - Week 9-13 (BD, BJ, BP, BV, CB) -> 'มกราคม'
  - Week 14-17 (CH, CN, CT, CZ) -> 'กุมภาพันธ์'
  - Week 18-22 (DF, DL, DR, DX, ED) -> 'มีนาคม'
- **Execution:** Always write the Thai month name as a string directly to these selector cells to trigger automatic date calculations.

## Advanced Surgical Editing (Safety First)
- **Rule:** Never recreate the entire <sheetData> tag for complex sheets (Sheet 5, 6, 8, etc.). 
- **Method:** Locate existing <row> and <c> tags and update only their <v> (value) or <f> (formula). This preserves merged cells, conditional formatting, and complex XML structures.
- **Recalculation:** Always delete xl/calcChain.xml and its reference in [Content_Types].xml to force Excel to update formulas on next open.

## Standard Cell Mapping (Primary Por Por 5)
### Main Sheet (Sheet 1):
- **School Name:** E5
- **Subject Code:** E12
- **Subject Name:** I12
- **Semester:** E13
- **Academic Year:** I13
- **Teacher:** E15

### Attendance Sheets (Sheet 5/6):
- **Month Selectors (List Box):** Row 4 (e.g., L4, AJ4, BH4, CL4, DJ4)
- **Weekday Labels:** Row 5 (ID 107 = 'ศ', ID 104 = 'อ', etc.)
- **Day Numbers:** Row 6
- **Period Count:** Row 7
- **Student Attendance:** Rows 8, 9, 10... (ID 141 = '/')

## Holiday Logic
- **Constraint:** Always cross-reference Thai public holidays for the specific academic year before filling attendance. 
- **Action:** Skip marking student attendance (Row 8+) and periods (Row 7) for holiday columns, but ensure Month and Day headers are still filled for visual consistency.

## Advanced Attendance Logic: Full Week Sequencing
- **Date Consistency (Row 6):** When updating attendance dates, you MUST calculate and populate the dates for the entire 5-day week (Monday through Friday) within each weekly block, regardless of the actual teaching day. This prevents visual anomalies where old data mixes with new data. 
- **Targeted Marking:** While the dates (Row 6) are filled for all 5 days, the attendance marks (Row 8+) and period counts (Row 7) should ONLY be applied to the specific teaching day(s) (e.g., Friday).
- **Tab Verification:** Never assume XML file names correspond directly to Semesters. Always verify the 
Id mapping. For this standard Por Por 5 format:
  - **เวลาเรียน1 (Term 1):** Maps to sheet5.xml
  - **เวลาเรียน2 (Term 2):** Maps to sheet6.xml
