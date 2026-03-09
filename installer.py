#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Installer Tu Dong
"""

import os
import shutil
import subprocess
import tkinter as tk
from tkinter import messagebox, ttk


class DocTienInstaller:
    def __init__(self):
        self.app_name = "DocTien Excel Add-in"
        self.version = "1.0.0"

    def get_excel_startup_path(self):
        try:
            appdata = os.environ.get("APPDATA")
            xlstart_path = os.path.join(appdata, "Microsoft", "Excel", "XLSTART")
            os.makedirs(xlstart_path, exist_ok=True)
            return xlstart_path
        except Exception:
            return None

    def get_excel_addins_path(self):
        try:
            appdata = os.environ.get("APPDATA")
            addins_path = os.path.join(appdata, "Microsoft", "AddIns")
            os.makedirs(addins_path, exist_ok=True)
            return addins_path
        except Exception:
            return None

    def create_auto_open_file(self, destination_path):
        try:
            bas_file = os.path.join(os.path.dirname(__file__), "DocTien.bas")
            if not os.path.exists(bas_file):
                return None
            dest_bas = os.path.join(destination_path, "DocTien.bas")
            shutil.copy2(bas_file, dest_bas)
            return dest_bas
        except Exception:
            return None

    def create_vbs_installer(self, bas_file_path, xlstart_path):
        vbs_content = f'''
On Error Resume Next

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

Set objWorkbook = objExcel.Workbooks.Add
objWorkbook.VBProject.VBComponents.Import "{bas_file_path.replace(chr(92), chr(92) + chr(92))}"

If Err.Number <> 0 Then
    WScript.Echo "ERROR: " & Err.Description
    WScript.Quit 1
End If

strSavePath = "{xlstart_path.replace(chr(92), chr(92) + chr(92))}\\DocTien.xlam"
objWorkbook.SaveAs strSavePath, 55

If Err.Number <> 0 Then
    WScript.Echo "ERROR: " & Err.Description
    WScript.Quit 1
End If

objWorkbook.Close False
objExcel.Quit

Set objWorkbook = Nothing
Set objExcel = Nothing

WScript.Echo "SUCCESS"
'''
        vbs_path = os.path.join(os.path.dirname(bas_file_path), "install_doctien.vbs")
        with open(vbs_path, "w", encoding="utf-8") as f:
            f.write(vbs_content)
        return vbs_path

    def install(self, progress_callback=None):
        try:
            if progress_callback:
                progress_callback(10, "\u0110ang t\u00ecm th\u01b0 m\u1ee5c Excel...")

            xlstart_path = self.get_excel_startup_path()
            addins_path = self.get_excel_addins_path()
            if not xlstart_path or not addins_path:
                return False, None, "Kh\u00f4ng t\u00ecm th\u1ea5y th\u01b0 m\u1ee5c Excel"

            if progress_callback:
                progress_callback(30, "\u0110ang chu\u1ea9n b\u1ecb file...")

            bas_path = self.create_auto_open_file(addins_path)
            if not bas_path:
                return False, None, "Kh\u00f4ng t\u00ecm th\u1ea5y file DocTien.bas"

            if progress_callback:
                progress_callback(50, "\u0110ang t\u1ea1o Add-in t\u1ef1 \u0111\u1ed9ng...")

            vbs_path = self.create_vbs_installer(bas_path, xlstart_path)

            if progress_callback:
                progress_callback(70, "\u0110ang c\u00e0i \u0111\u1eb7t v\u00e0o Excel...")

            result = subprocess.run(
                ["cscript", "//Nologo", vbs_path],
                capture_output=True,
                text=True,
                timeout=30,
            )

            if result.returncode == 0 and "SUCCESS" in result.stdout:
                if progress_callback:
                    progress_callback(100, "C\u00e0i \u0111\u1eb7t ho\u00e0n t\u1ea5t!")
                return True, xlstart_path, None

            error_msg = (result.stderr or result.stdout or "").strip()
            manual_message = (
                "Kh\u00f4ng th\u1ec3 k\u00edch ho\u1ea1t t\u1ef1 \u0111\u1ed9ng trong Excel.\n\n"
                "File DocTien.bas \u0111\u00e3 \u0111\u01b0\u1ee3c copy \u0111\u1ebfn:\n"
                f"{bas_path}\n\n"
                "H\u00e3y import th\u1ee7 c\u00f4ng trong Excel (Alt + F11 -> File -> Import File)."
            )
            if error_msg:
                manual_message = f"{manual_message}\n\nChi ti\u1ebft l\u1ed7i: {error_msg[:300]}"

            if progress_callback:
                progress_callback(100, "Kh\u00f4ng th\u1ec3 c\u00e0i t\u1ef1 \u0111\u1ed9ng")
            return False, bas_path, manual_message
        except Exception as e:
            return False, None, f"L\u1ed7i: {str(e)}"


class InstallerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("C\u00e0i \u0111\u1eb7t DocTien - Excel Add-in")
        self.root.geometry("550x500")
        self.root.resizable(False, False)

        self.bg_color = "#f0f0f0"
        self.primary_color = "#2563eb"
        self.success_color = "#10b981"

        self.root.configure(bg=self.bg_color)
        self.installer = DocTienInstaller()
        self.create_widgets()

    def create_widgets(self):
        header_frame = tk.Frame(self.root, bg=self.primary_color, height=80)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)

        title_label = tk.Label(
            header_frame,
            text="DocTien Excel Add-in",
            font=("Segoe UI", 20, "bold"),
            bg=self.primary_color,
            fg="white",
        )
        title_label.pack(expand=True)

        content_frame = tk.Frame(self.root, bg=self.bg_color)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)

        desc_text = (
            "C\u00c0I \u0110\u1eb6T T\u1ef0 \u0110\u1ed8NG - KH\u00d4NG C\u1ea6N TH\u1ee6 C\u00d4NG!\n\n"
            "Sau khi c\u00e0i \u0111\u1eb7t, b\u1ea1n c\u00f3 th\u1ec3:\n"
            "- S\u1eed d\u1ee5ng c\u00f4ng th\u1ee9c: =DocTien(A1)\n"
            "- Chuy\u1ec3n s\u1ed1 th\u00e0nh ch\u1eef ti\u1ebfng Vi\u1ec7t t\u1ef1 \u0111\u1ed9ng\n"
            "- S\u1eed d\u1ee5ng trong t\u1ea5t c\u1ea3 file Excel\n\n"
            "V\u00ed d\u1ee5:\n"
            "- 12345 -> M\u01b0\u1eddi hai ngh\u00ecn ba tr\u0103m b\u1ed1n m\u01b0\u01a1i l\u0103m \u0111\u1ed3ng\n"
            "- 1000000 -> M\u1ed9t tri\u1ec7u \u0111\u1ed3ng\n\n"
            "Nh\u1ea5n 'C\u00e0i \u0111\u1eb7t' \u0111\u1ec3 b\u1eaft \u0111\u1ea7u..."
        )

        desc_label = tk.Label(
            content_frame,
            text=desc_text,
            font=("Segoe UI", 10),
            bg=self.bg_color,
            fg="#333",
            justify=tk.LEFT,
        )
        desc_label.pack(pady=10)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            content_frame,
            variable=self.progress_var,
            maximum=100,
            mode="determinate",
            length=400,
        )
        self.progress_bar.pack(pady=15)

        self.status_label = tk.Label(
            content_frame,
            text="Nh\u1ea5n 'C\u00e0i \u0111\u1eb7t' \u0111\u1ec3 b\u1eaft \u0111\u1ea7u",
            font=("Segoe UI", 9),
            bg=self.bg_color,
            fg="#666",
        )
        self.status_label.pack()

        button_frame = tk.Frame(content_frame, bg=self.bg_color)
        button_frame.pack(side=tk.BOTTOM, pady=20, anchor=tk.CENTER)

        self.install_button = tk.Button(
            button_frame,
            text="C\u00e0i \u0111\u1eb7t",
            font=("Segoe UI", 11, "bold"),
            bg=self.primary_color,
            fg="white",
            padx=40,
            pady=10,
            command=self.install,
            cursor="hand2",
            relief=tk.FLAT,
        )
        self.install_button.pack(side=tk.LEFT, padx=5)

        self.close_button = tk.Button(
            button_frame,
            text="\u0110\u00f3ng",
            font=("Segoe UI", 11),
            bg="#6b7280",
            fg="white",
            padx=40,
            pady=10,
            command=self.root.quit,
            cursor="hand2",
            relief=tk.FLAT,
        )
        self.close_button.pack(side=tk.LEFT, padx=5)

    def update_progress(self, value, message):
        self.progress_var.set(value)
        self.status_label.config(text=message)
        self.root.update()

    def install(self):
        self.install_button.config(state=tk.DISABLED)
        success, location, message = self.installer.install(self.update_progress)

        if success:
            messagebox.showinfo(
                "Th\u00e0nh c\u00f4ng",
                "C\u00c0I \u0110\u1eb6T HO\u00c0N T\u1ea4T!\n\n"
                "C\u00c1CH S\u1eec D\u1ee4NG:\n\n"
                "B\u01af\u1edaC 1: M\u1edf Microsoft Excel (ho\u1eb7c kh\u1edfi \u0111\u1ed9ng l\u1ea1i n\u1ebfu \u0111ang m\u1edf)\n\n"
                "B\u01af\u1edaC 2: N\u1ebfu xu\u1ea5t hi\u1ec7n c\u1ea3nh b\u00e1o b\u1ea3o m\u1eadt:\n"
                "   - Nh\u1ea5n 'Enable Content' ho\u1eb7c 'Enable Macros'\n\n"
                "B\u01af\u1edaC 3: S\u1eed d\u1ee5ng c\u00f4ng th\u1ee9c:\n"
                "   =DocTien(A1)\n\n"
                "V\u00ed d\u1ee5:\n"
                "   =DocTien(12345)\n"
                "   -> M\u01b0\u1eddi hai ngh\u00ecn ba tr\u0103m b\u1ed1n m\u01b0\u01a1i l\u0103m \u0111\u1ed3ng\n\n"
                "N\u1ebfu h\u00e0m DocTien() ch\u01b0a ho\u1ea1t \u0111\u1ed9ng:\n"
                "   File -> Options -> Add-ins -> Go...\n"
                "   T\u00edch ch\u1ecdn 'DocTien' -> OK\n\n"
                f"File \u0111\u00e3 \u0111\u01b0\u1ee3c c\u00e0i t\u1ea1i:\n{location}"
            )
            self.status_label.config(text="C\u00e0i \u0111\u1eb7t th\u00e0nh c\u00f4ng!", fg=self.success_color)
        else:
            self.install_button.config(state=tk.NORMAL)
            self.status_label.config(text="C\u00e0i \u0111\u1eb7t th\u1ea5t b\u1ea1i", fg="red")
            messagebox.showwarning("C\u1ea7n thao t\u00e1c th\u1ee7 c\u00f4ng", message or "C\u00e0i \u0111\u1eb7t th\u1ea5t b\u1ea1i")


def main():
    root = tk.Tk()
    app = InstallerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
