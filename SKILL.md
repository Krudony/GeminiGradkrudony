---
name: excel-grade-krudony
description: Safe Excel grade and score adjustment for ปพ.5 files (Krudony's method). Preserves drawings, charts, and XML integrity.
---

# Excel Grade Krudony (Safe Method)

This skill implements the "Safe Method" for editing school grading files (ปพ.5), ensuring that direct XML manipulation is used to preserve the integrity of the workbook. 

The traditional `openpyxl.save()` method strips out images, drawings, and charts from ปพ.5 files. This method extracts the XLSX, edits the internal XML, and repacks it safely.

## Core Workflow

1. **Identify Level:** Determine if the file is for Primary (ประถม) or Secondary (มัธยม).
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
