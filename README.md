# GeminiGradkrudony 📊

Skill สำหรับ Gemini CLI ในการจัดการแก้ไขไฟล์ ปพ.5 (Excel) ด้วยวิธี **Safe Method** (Direct XML Editing) เพื่อรักษาความสมบูรณ์ของไฟล์ (รูปภาพ, กราฟ, สูตร) ตามมาตรฐานของครูดอน (Krudony)

## ✨ ฟีเจอร์หลัก
- **Safe Edit**: แก้ไขไฟล์ XML ภายใน XLSX โดยตรง ไม่ผ่าน Library ที่ทำให้ไฟล์พัง
- **Auto Detect**: ตรวจหานักเรียนในห้องอัตโนมัติ
- **Smart Scoring**: คำนวณและกระจายคะแนนให้สอดคล้องกับเกรดเป้าหมาย
- **Attendance**: จัดการเวลาเรียน ภาค 1 และ 2 พร้อมเช็ควันหยุดไทย
- **Assessment**: ให้คะแนนคุณลักษณะ, อ่านคิดวิเคราะห์ และสมรรถนะ แบบ Batch

## 📂 โครงสร้าง
- `SKILL.md`: คู่มือการใช้งานสำหรับ Gemini CLI
- `scripts/xlsx_safe_edit.py`: Master script สำหรับจัดการ XML
- `references/mapping.md`: ตารางตำแหน่ง Cell อ้างอิง (ประถม/มัธยม)

## 🚀 วิธีติดตั้ง
Copy โฟลเดอร์นี้ไปไว้ที่ `.gemini/skills/excel-grade-krudony` ในเครื่องของคุณ

---
*สร้างโดย Si-Som (Gemini CLI Partner) เพื่อสนับสนุนงานวิชาการครูไทย*
