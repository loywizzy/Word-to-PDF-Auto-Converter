# 📄 Word to PDF Converter

แปลงไฟล์ Word (.docx) เป็น PDF อัตโนมัติทั้งโฟลเดอร์ พร้อม GUI สวยงาม

![Windows](https://img.shields.io/badge/Windows-10%2F11-blue?logo=windows)
![Python](https://img.shields.io/badge/Python-3.8+-green?logo=python)
![License](https://img.shields.io/badge/License-MIT-yellow)

---

## ✨ Features

- 🚀 แปลงไฟล์ Word เป็น PDF อัตโนมัติทั้งโฟลเดอร์
- 🎨 GUI สวยงาม พร้อม Dark theme
- 📊 แสดง Progress bar แบบ real-time
- 📋 Log แสดงสถานะการแปลงแต่ละไฟล์
- 💾 ไม่ต้องติดตั้ง - ใช้งานได้ทันที (.exe)

---

## 📥 Download

👉 **[ดาวน์โหลด word_to_pdf_gui.exe](../../releases/latest)**

---

## 📋 Requirements

- **Windows 10/11**
- **Microsoft Word** 2010 ขึ้นไป (จำเป็นต้องติดตั้ง)

---

## 🚀 Quick Start

1. ดาวน์โหลด `word_to_pdf_gui.exe` จาก [Releases](../../releases)
2. ดับเบิ้ลคลิกเพื่อเปิดโปรแกรม
3. เลือกโฟลเดอร์ที่มีไฟล์ .docx
4. กดปุ่ม **"🚀 เริ่มแปลงไฟล์"**
5. ไฟล์ PDF จะถูกสร้างในโฟลเดอร์เดียวกัน

---

## 🛠️ Run from Source

```bash
# ติดตั้ง dependencies
pip install pywin32

# รันโปรแกรม
python word_to_pdf_gui.py
```

---

## 📦 Build .exe

```bash
pip install pyinstaller
python -m PyInstaller --onefile --windowed word_to_pdf_gui.py
```

ไฟล์ .exe จะอยู่ในโฟลเดอร์ `dist/`

---

## 📁 Project Structure

```
├── word_to_pdf_gui.py        # โปรแกรมหลัก (GUI)
├── word_to_pdf_converter.py  # Command-line version
├── document.md               # เอกสารการใช้งาน
└── RELEASE_NOTES.md          # Release notes
```

---

## ⚠️ Limitations

- ต้องมี Microsoft Word ติดตั้งในเครื่อง
- รองรับเฉพาะ Windows เท่านั้น

---

## 📝 License

MIT License - ใช้งานได้อิสระ
