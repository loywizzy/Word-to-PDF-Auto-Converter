"""
Word to PDF Auto Converter - GUI Version
แปลงไฟล์ .docx ทั้งหมดในโฟลเดอร์เป็น PDF อัตโนมัติ
พร้อม GUI สวยงาม
"""

import os
import sys
import time
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import win32com.client
    import pythoncom
except ImportError:
    messagebox.showerror("Error", "กรุณาติดตั้ง pywin32 ก่อน:\npip install pywin32")
    sys.exit(1)


class WordToPDFConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Word to PDF Converter")
        self.root.geometry("650x700")
        self.root.minsize(600, 650)
        self.root.resizable(True, True)
        self.root.configure(bg="#1a1a2e")
        
        # ตั้งค่า icon (ถ้ามี)
        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass
        
        self.folder_path = tk.StringVar()
        self.folder_path.set(r"C:\Users\soraw\OneDrive\Desktop\Document Project\Fixing\now-fix")
        
        self.is_converting = False
        
        self.create_widgets()
        self.center_window()
    
    def center_window(self):
        """จัดหน้าต่างให้อยู่กลางจอ"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
    
    def create_widgets(self):
        """สร้าง UI ทั้งหมด"""
        # สไตล์
        style = ttk.Style()
        style.theme_use('clam')
        
        # Header
        header_frame = tk.Frame(self.root, bg="#16213e", pady=20)
        header_frame.pack(fill="x")
        
        title_label = tk.Label(
            header_frame,
            text="📄 Word to PDF Converter",
            font=("Segoe UI", 24, "bold"),
            fg="#e94560",
            bg="#16213e"
        )
        title_label.pack()
        
        subtitle_label = tk.Label(
            header_frame,
            text="แปลงไฟล์ Word เป็น PDF อัตโนมัติ",
            font=("Segoe UI", 12),
            fg="#a0a0a0",
            bg="#16213e"
        )
        subtitle_label.pack(pady=(5, 0))
        
        # Main content
        main_frame = tk.Frame(self.root, bg="#1a1a2e", padx=30, pady=20)
        main_frame.pack(fill="both", expand=True)
        
        # Folder selection
        folder_label = tk.Label(
            main_frame,
            text="📁 เลือกโฟลเดอร์:",
            font=("Segoe UI", 11),
            fg="#ffffff",
            bg="#1a1a2e"
        )
        folder_label.pack(anchor="w")
        
        folder_frame = tk.Frame(main_frame, bg="#1a1a2e")
        folder_frame.pack(fill="x", pady=(5, 15))
        
        self.folder_entry = tk.Entry(
            folder_frame,
            textvariable=self.folder_path,
            font=("Segoe UI", 10),
            bg="#0f3460",
            fg="#ffffff",
            insertbackground="#ffffff",
            relief="flat",
            bd=0
        )
        self.folder_entry.pack(side="left", fill="x", expand=True, ipady=10, ipadx=10)
        
        browse_btn = tk.Button(
            folder_frame,
            text="📂 Browse",
            font=("Segoe UI", 10),
            bg="#e94560",
            fg="#ffffff",
            activebackground="#ff6b6b",
            activeforeground="#ffffff",
            relief="flat",
            cursor="hand2",
            command=self.browse_folder
        )
        browse_btn.pack(side="right", padx=(10, 0), ipady=8, ipadx=15)
        
        # Convert button (ย้ายมาไว้ด้านบน)
        self.convert_btn = tk.Button(
            main_frame,
            text="🚀 เริ่มแปลงไฟล์",
            font=("Segoe UI", 14, "bold"),
            bg="#e94560",
            fg="#ffffff",
            activebackground="#ff6b6b",
            activeforeground="#ffffff",
            relief="flat",
            cursor="hand2",
            command=self.start_conversion
        )
        self.convert_btn.pack(fill="x", ipady=12, pady=(0, 15))
        
        # Progress frame
        progress_frame = tk.Frame(main_frame, bg="#1a1a2e")
        progress_frame.pack(fill="x", pady=(10, 5))
        
        self.progress_label = tk.Label(
            progress_frame,
            text="พร้อมแปลงไฟล์",
            font=("Segoe UI", 10),
            fg="#a0a0a0",
            bg="#1a1a2e"
        )
        self.progress_label.pack(anchor="w")
        
        # Progress bar
        style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor="#0f3460",
            background="#e94560",
            thickness=8
        )
        
        self.progress_bar = ttk.Progressbar(
            main_frame,
            style="Custom.Horizontal.TProgressbar",
            orient="horizontal",
            mode="determinate"
        )
        self.progress_bar.pack(fill="x", pady=(5, 15))
        
        # Log area
        log_label = tk.Label(
            main_frame,
            text="📋 Log:",
            font=("Segoe UI", 11),
            fg="#ffffff",
            bg="#1a1a2e"
        )
        log_label.pack(anchor="w")
        
        log_frame = tk.Frame(main_frame, bg="#0f3460")
        log_frame.pack(fill="both", expand=True, pady=(5, 15))
        
        self.log_text = tk.Text(
            log_frame,
            font=("Consolas", 10),
            bg="#0f3460",
            fg="#ffffff",
            insertbackground="#ffffff",
            relief="flat",
            wrap="word",
            state="disabled"
        )
        self.log_text.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        # Configure tags for colored text
        self.log_text.tag_configure("success", foreground="#4ade80")
        self.log_text.tag_configure("error", foreground="#f87171")
        self.log_text.tag_configure("info", foreground="#60a5fa")
        self.log_text.tag_configure("warning", foreground="#fbbf24")
        
        # Footer
        footer_label = tk.Label(
            self.root,
            text="Made with ❤️ for Word to PDF conversion",
            font=("Segoe UI", 9),
            fg="#666666",
            bg="#1a1a2e"
        )
        footer_label.pack(pady=10)
    
    def browse_folder(self):
        """เปิด dialog เลือกโฟลเดอร์"""
        folder = filedialog.askdirectory(
            title="เลือกโฟลเดอร์ที่มีไฟล์ Word",
            initialdir=self.folder_path.get()
        )
        if folder:
            self.folder_path.set(folder)
    
    def log(self, message, tag=None):
        """เพิ่มข้อความใน log"""
        self.log_text.configure(state="normal")
        if tag:
            self.log_text.insert("end", message + "\n", tag)
        else:
            self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
    
    def clear_log(self):
        """ล้าง log"""
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
    
    def start_conversion(self):
        """เริ่มการแปลงไฟล์ใน thread แยก"""
        if self.is_converting:
            return
        
        folder = self.folder_path.get().strip()
        if not folder:
            messagebox.showwarning("Warning", "กรุณาเลือกโฟลเดอร์ก่อน")
            return
        
        if not os.path.isdir(folder):
            messagebox.showerror("Error", f"ไม่พบโฟลเดอร์:\n{folder}")
            return
        
        # รันใน thread แยก
        thread = threading.Thread(target=self.convert_files, args=(folder,))
        thread.daemon = True
        thread.start()
    
    def convert_files(self, folder_path):
        """แปลงไฟล์ทั้งหมดในโฟลเดอร์"""
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        
        self.is_converting = True
        self.convert_btn.configure(state="disabled", bg="#666666")
        self.clear_log()
        
        folder = Path(folder_path)
        docx_files = sorted(folder.glob("*.docx"))
        
        if not docx_files:
            self.log("⚠️ ไม่พบไฟล์ .docx ในโฟลเดอร์", "warning")
            self.reset_ui()
            pythoncom.CoUninitialize()
            return
        
        total_files = len(docx_files)
        self.log(f"📁 พบไฟล์ .docx จำนวน {total_files} ไฟล์", "info")
        self.log("-" * 40)
        
        success_count = 0
        error_count = 0
        
        word = None
        try:
            self.update_progress(0, "กำลังเปิด Microsoft Word...")
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            
            wdFormatPDF = 17
            
            for i, docx_file in enumerate(docx_files, 1):
                pdf_file = docx_file.with_suffix(".pdf")
                progress = int((i / total_files) * 100)
                
                self.update_progress(progress, f"[{i}/{total_files}] กำลังแปลง: {docx_file.name}")
                self.log(f"🔄 [{i}/{total_files}] กำลังแปลง: {docx_file.name}")
                
                doc = None
                try:
                    doc = word.Documents.Open(
                        str(docx_file),
                        ReadOnly=True,
                        AddToRecentFiles=False,
                        Visible=False
                    )
                    
                    time.sleep(0.5)
                    
                    doc.ExportAsFixedFormat(
                        str(pdf_file),
                        wdFormatPDF,
                        OpenAfterExport=False,
                        OptimizeFor=0
                    )
                    
                    self.log(f"   ✅ สำเร็จ: {pdf_file.name}", "success")
                    success_count += 1
                    
                except Exception as e:
                    self.log(f"   ❌ ผิดพลาด: {e}", "error")
                    error_count += 1
                finally:
                    if doc:
                        try:
                            doc.Close(SaveChanges=False)
                        except:
                            pass
                    time.sleep(0.3)
        
        except Exception as e:
            self.log(f"❌ ไม่สามารถเปิด Microsoft Word ได้: {e}", "error")
        finally:
            if word:
                try:
                    word.Quit()
                except:
                    pass
            # Uninitialize COM
            pythoncom.CoUninitialize()
        
        self.log("-" * 40)
        self.log(f"📊 สรุปผล: สำเร็จ {success_count} ไฟล์, ผิดพลาด {error_count} ไฟล์", "info")
        self.log("✨ Export PDF เสร็จสิ้น!", "success")
        
        self.update_progress(100, f"เสร็จสิ้น! สำเร็จ {success_count}/{total_files} ไฟล์")
        
        if error_count == 0:
            messagebox.showinfo("สำเร็จ", f"แปลงไฟล์เสร็จสิ้น!\n\nสำเร็จ: {success_count} ไฟล์")
        else:
            messagebox.showwarning("เสร็จสิ้น", f"แปลงไฟล์เสร็จสิ้น!\n\nสำเร็จ: {success_count} ไฟล์\nผิดพลาด: {error_count} ไฟล์")
        
        self.reset_ui()
    
    def update_progress(self, value, text):
        """อัพเดท progress bar และ label"""
        self.root.after(0, lambda: self.progress_bar.configure(value=value))
        self.root.after(0, lambda: self.progress_label.configure(text=text))
    
    def reset_ui(self):
        """รีเซ็ต UI หลังแปลงเสร็จ"""
        self.is_converting = False
        self.root.after(0, lambda: self.convert_btn.configure(state="normal", bg="#e94560"))


def main():
    root = tk.Tk()
    app = WordToPDFConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()
