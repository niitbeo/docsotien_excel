#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DocTien Auto Installer - Fully Automatic
Runs PowerShell script to install VBA module automatically
"""

import os
import sys
import subprocess
import tempfile
import tkinter as tk
from tkinter import messagebox, ttk
import threading


class DocTienAutoInstaller:
    def __init__(self, root):
        self.root = root
        self.root.title("DocTien Auto Installer")
        self.root.geometry("650x600")
        self.root.resizable(False, False)

        self.bg_color = "#f0f0f0"
        self.primary_color = "#2563eb"
        self.success_color = "#10b981"

        self.root.configure(bg=self.bg_color)
        self.create_widgets()

    def create_widgets(self):
        header_frame = tk.Frame(self.root, bg=self.primary_color, height=80)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)

        title_label = tk.Label(
            header_frame,
            text="DocTien Auto Installer",
            font=("Segoe UI", 24, "bold"),
            bg=self.primary_color,
            fg="white"
        )
        title_label.pack(expand=True)

        content_frame = tk.Frame(self.root, bg=self.bg_color)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=40, pady=20)

        desc_text = (
            "C\u00e0i \u0111\u1eb7t T\u1ef0 \u0110\u1ed8NG h\u00e0m DocTien() v\u00e0o Excel\n\n"
            "Sau khi c\u00e0i \u0111\u1eb7t:\n"
            "- M\u1edf Excel v\u00e0 g\u00f5: =DocTien(A1)\n"
            "- Chuy\u1ec3n s\u1ed1 th\u00e0nh ch\u1eef ti\u1ebfng Vi\u1ec7t t\u1ef1 \u0111\u1ed9ng\n"
            "- D\u00f9ng \u0111\u01b0\u1ee3c trong m\u1ecdi file Excel\n\n"
            "V\u00ed d\u1ee5: 12345 -> \"M\u01b0\u1eddi hai ngh\u00ecn ba tr\u0103m b\u1ed1n m\u01b0\u01a1i l\u0103m \u0111\u1ed3ng\""
        )

        desc_label = tk.Label(
            content_frame,
            text=desc_text,
            font=("Segoe UI", 11),
            bg=self.bg_color,
            fg="#333",
            justify=tk.LEFT
        )
        desc_label.pack(pady=10)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            content_frame,
            variable=self.progress_var,
            maximum=100,
            mode="indeterminate",
            length=500
        )
        self.progress_bar.pack(pady=15)

        self.status_label = tk.Label(
            content_frame,
            text="Nh\u1ea5n 'C\u00e0i \u0111\u1eb7t t\u1ef1 \u0111\u1ed9ng' \u0111\u1ec3 b\u1eaft \u0111\u1ea7u",
            font=("Segoe UI", 10),
            bg=self.bg_color,
            fg="#666"
        )
        self.status_label.pack(pady=10)

        button_frame = tk.Frame(content_frame, bg=self.bg_color)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=15)

        self.install_button = tk.Button(
            button_frame,
            text="C\u00e0i \u0111\u1eb7t t\u1ef1 \u0111\u1ed9ng",
            font=("Segoe UI", 12, "bold"),
            bg=self.primary_color,
            fg="white",
            padx=50,
            pady=12,
            command=self.install,
            cursor="hand2",
            relief=tk.FLAT
        )
        self.install_button.pack(side=tk.LEFT, padx=10, expand=True)

        self.close_button = tk.Button(
            button_frame,
            text="\u0110\u00f3ng",
            font=("Segoe UI", 12),
            bg="#6b7280",
            fg="white",
            padx=50,
            pady=12,
            command=self.root.quit,
            cursor="hand2",
            relief=tk.FLAT
        )
        self.close_button.pack(side=tk.LEFT, padx=10, expand=True)

        footer_frame = tk.Frame(self.root, bg="#1e293b", height=50)
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X)
        footer_frame.pack_propagate(False)

        author_label = tk.Label(
            footer_frame,
            text="\u00a9 2024 Viber Code - Nguy\u1ec5n L\u00ea Tr\u01b0\u1eddng - 0888849413",
            font=("Segoe UI", 9),
            bg="#1e293b",
            fg="#94a3b8"
        )
        author_label.pack(expand=True)

    def get_resource_path(self, filename):
        base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, filename)

    def load_bas_content(self):
        bas_path = self.get_resource_path("DocTien.bas")
        with open(bas_path, "r", encoding="utf-8-sig") as f:
            return f.read()

    def run_powershell_script(self):
        bas_file = None
        try:
            doctien_bas_content = self.load_bas_content()

            with tempfile.NamedTemporaryFile(delete=False, suffix=".bas", mode="w", encoding="utf-8") as f:
                f.write(doctien_bas_content)
                bas_file = f.name

            ps_script = f"""
$basFile = '{bas_file.replace(chr(92), chr(92) * 2)}'

Write-Host '[PROGRESS] Starting Excel...'

try {{
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $xlStartPath = Join-Path $env:APPDATA 'Microsoft\\Excel\\XLSTART'

    if (-not (Test-Path $xlStartPath)) {{
        New-Item -ItemType Directory -Path $xlStartPath -Force | Out-Null
    }}

    $personalFile = Join-Path $xlStartPath 'PERSONAL.XLSB'

    if (Test-Path $personalFile) {{
        $workbook = $excel.Workbooks.Open($personalFile)
    }} else {{
        $workbook = $excel.Workbooks.Add()
    }}

    foreach ($component in $workbook.VBProject.VBComponents) {{
        if ($component.Name -eq 'DocTienModule') {{
            $workbook.VBProject.VBComponents.Remove($component)
            break
        }}
    }}

    $workbook.VBProject.VBComponents.Import($basFile) | Out-Null

    if (Test-Path $personalFile) {{
        $workbook.Save()
    }} else {{
        $workbook.SaveAs($personalFile, 50)
    }}

    $workbook.Close($false)
    $excel.Quit()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Host 'SUCCESS'

}} catch {{
    Write-Host "ERROR: $_"
    exit 1
}}
"""

            result = subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
                capture_output=True,
                text=True,
                timeout=60
            )

            if result.returncode == 0 and "SUCCESS" in result.stdout:
                return True, "C\u00e0i \u0111\u1eb7t th\u00e0nh c\u00f4ng!"

            error_msg = result.stderr if result.stderr else result.stdout
            if "Trust access" in error_msg or "0x800A03EC" in error_msg:
                return False, (
                    "C\u1ea7n b\u1eadt 'Trust access to VBA project'\n\n"
                    "H\u01b0\u1edbng d\u1eabn:\n"
                    "1. M\u1edf Excel\n"
                    "2. File -> Options -> Trust Center -> Trust Center Settings\n"
                    "3. Macro Settings -> t\u00edch 'Trust access to the VBA project object model'\n"
                    "4. OK v\u00e0 ch\u1ea1y l\u1ea1i installer"
                )
            return False, f"L\u1ed7i: {error_msg[:200]}"
        except FileNotFoundError:
            return False, "Kh\u00f4ng t\u00ecm th\u1ea5y file DocTien.bas trong b\u1ed9 c\u00e0i. H\u00e3y build l\u1ea1i installer k\u00e8m resource n\u00e0y."
        except subprocess.TimeoutExpired:
            return False, "Timeout - qu\u00e1 tr\u00ecnh c\u00e0i \u0111\u1eb7t m\u1ea5t qu\u00e1 nhi\u1ec1u th\u1eddi gian"
        except Exception as e:
            return False, f"L\u1ed7i: {str(e)}"
        finally:
            if bas_file and os.path.exists(bas_file):
                try:
                    os.remove(bas_file)
                except OSError:
                    pass

    def install_thread(self):
        self.progress_bar.start(10)
        self.status_label.config(text="\u0110ang c\u00e0i \u0111\u1eb7t... Vui l\u00f2ng \u0111\u1ee3i...")

        success, message = self.run_powershell_script()

        self.progress_bar.stop()
        self.progress_var.set(100 if success else 0)

        if success:
            self.status_label.config(text="C\u00e0i \u0111\u1eb7t th\u00e0nh c\u00f4ng!", fg=self.success_color)
            messagebox.showinfo(
                "Th\u00e0nh c\u00f4ng",
                "C\u00c0I \u0110\u1eb6T HO\u00c0N T\u1ea4T!\n\n"
                "C\u00e1ch s\u1eed d\u1ee5ng:\n"
                "1. M\u1edf Excel (ho\u1eb7c kh\u1edfi \u0111\u1ed9ng l\u1ea1i n\u1ebfu \u0111ang m\u1edf)\n"
                "2. G\u00f5 c\u00f4ng th\u1ee9c: =DocTien(A1)\n\n"
                "V\u00ed d\u1ee5:\n"
                "   =DocTien(12345)\n"
                "   -> M\u01b0\u1eddi hai ngh\u00ecn ba tr\u0103m b\u1ed1n m\u01b0\u01a1i l\u0103m \u0111\u1ed3ng\n\n"
                "L\u01b0u \u00fd: N\u1ebfu c\u00f3 c\u1ea3nh b\u00e1o b\u1ea3o m\u1eadt, nh\u1ea5n 'Enable Content'"
            )
        else:
            self.status_label.config(text="C\u00e0i \u0111\u1eb7t th\u1ea5t b\u1ea1i", fg="red")
            messagebox.showerror("L\u1ed7i", message)

        self.install_button.config(state=tk.NORMAL)

    def install(self):
        self.install_button.config(state=tk.DISABLED)
        thread = threading.Thread(target=self.install_thread)
        thread.daemon = True
        thread.start()


def main():
    root = tk.Tk()
    app = DocTienAutoInstaller(root)
    root.mainloop()


if __name__ == "__main__":
    main()
