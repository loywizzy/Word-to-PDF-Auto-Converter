"""
Word to PDF Auto Converter (Improved Version)
แปลงไฟล์ .docx ทั้งหมดในโฟลเดอร์เป็น PDF อัตโนมัติ
ใช้ win32com เพื่อควบคุม Microsoft Word โดยตรง
"""

import os
import sys
import time
from pathlib import Path

try:
    import win32com.client
    from win32com.client import constants
except ImportError:
    print("กรุณาติดตั้ง pywin32 ก่อน: pip install pywin32")
    print("หมายเหตุ: โปรแกรมนี้ต้องใช้ Microsoft Word ที่ติดตั้งในเครื่อง")
    sys.exit(1)


def convert_all_word_to_pdf(folder_path: str) -> None:
    """
    แปลงไฟล์ .docx ทั้งหมดในโฟลเดอร์ที่ระบุเป็น PDF
    
    Args:
        folder_path: พาธของโฟลเดอร์ที่มีไฟล์ .docx
    """
    folder = Path(folder_path)
    
    # ตรวจสอบว่าโฟลเดอร์มีอยู่จริง
    if not folder.exists():
        print(f"❌ ไม่พบโฟลเดอร์: {folder_path}")
        return
    
    if not folder.is_dir():
        print(f"❌ พาธที่ระบุไม่ใช่โฟลเดอร์: {folder_path}")
        return
    
    # หาไฟล์ .docx ทั้งหมด
    docx_files = sorted(folder.glob("*.docx"))
    
    if not docx_files:
        print(f"⚠️ ไม่พบไฟล์ .docx ในโฟลเดอร์: {folder_path}")
        return
    
    print(f"📁 พบไฟล์ .docx จำนวน {len(docx_files)} ไฟล์")
    print("-" * 50)
    
    success_count = 0
    error_count = 0
    
    # เปิด Word Application
    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # ซ่อน Word
        word.DisplayAlerts = False  # ปิด Alert
        
        wdFormatPDF = 17  # PDF format constant
        
        for i, docx_file in enumerate(docx_files, 1):
            pdf_file = docx_file.with_suffix(".pdf")
            print(f"🔄 [{i}/{len(docx_files)}] กำลังแปลง: {docx_file.name}")
            
            doc = None
            try:
                # เปิดไฟล์แบบ Read Only
                doc = word.Documents.Open(
                    str(docx_file),
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    Visible=False
                )
                
                # รอให้เปิดเสร็จ
                time.sleep(0.5)
                
                # Export เป็น PDF
                doc.ExportAsFixedFormat(
                    str(pdf_file),
                    wdFormatPDF,
                    OpenAfterExport=False,
                    OptimizeFor=0  # wdExportOptimizeForPrint
                )
                
                print(f"   ✅ สำเร็จ: {pdf_file.name}")
                success_count += 1
                
            except Exception as e:
                print(f"   ❌ ผิดพลาด: {e}")
                error_count += 1
            finally:
                # ปิดเอกสาร
                if doc:
                    try:
                        doc.Close(SaveChanges=False)
                    except:
                        pass
                
                # รอเล็กน้อยก่อนแปลงไฟล์ถัดไป
                time.sleep(0.3)
        
    except Exception as e:
        print(f"❌ ไม่สามารถเปิด Microsoft Word ได้: {e}")
        return
    finally:
        # ปิด Word Application
        if word:
            try:
                word.Quit()
            except:
                pass
    
    print("-" * 50)
    print(f"📊 สรุปผล: สำเร็จ {success_count} ไฟล์, ผิดพลาด {error_count} ไฟล์")
    print("✨ Export PDF เสร็จสิ้น!")


def main():
    # กำหนดโฟลเดอร์เริ่มต้น (สามารถเปลี่ยนได้)
    default_folder = r"C:\Users\soraw\OneDrive\Desktop\Document Project\Fixing\now-fix"
    
    print("=" * 50)
    print("   Word to PDF Auto Converter")
    print("   แปลงไฟล์ Word เป็น PDF อัตโนมัติ")
    print("=" * 50)
    print()
    
    # ถามผู้ใช้ว่าจะใช้โฟลเดอร์ไหน
    print(f"โฟลเดอร์เริ่มต้น: {default_folder}")
    user_input = input("กด Enter เพื่อใช้โฟลเดอร์เริ่มต้น หรือพิมพ์พาธใหม่: ").strip()
    
    if user_input:
        folder_path = user_input
    else:
        folder_path = default_folder
    
    print()
    convert_all_word_to_pdf(folder_path)
    print()
    input("กด Enter เพื่อปิดโปรแกรม...")


if __name__ == "__main__":
    main()
