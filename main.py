"""全格式文档转换器 — GUI 主界面（美化版）。"""

import os
import sys
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from utils import detect_format, is_valid_file, get_output_path, ALLOWED_EXTENSIONS
from engines.word_engine import docx_to_pdf, pdf_to_docx

CONVERSION_TARGETS = {
    "docx": {"PDF": ("pdf", docx_to_pdf), "Markdown": ("md", "md_to_docx")},
    "pdf":  {"Word": ("docx", pdf_to_docx)},
    "md":   {"Word": ("docx", "md_to_docx"), "PDF": ("pdf", "md_to_pdf")},
}

FORMAT_NAMES = {"docx": "Word 文档 (.docx)", "pdf": "PDF 文档 (.pdf)", "md": "Markdown (.md)"}
FORMAT_ICONS = {"docx": "📝", "pdf": "📕", "md": "📋"}
TARGET_COLORS = {"pdf": "#E74C3C", "docx": "#2E86C1", "md": "#27AE60"}


class FormatConverterApp:
    def __init__(self, root):
        self.root = root
        self.input_path = None
        self.input_format = None
        self.output_path = None
        self.target_ext = None
        self.converted = False

        root.title("FormatForge — 全格式文档转换器")
        root.geometry("560x530")
        root.resizable(False, False)
        root.configure(bg="#F5F6FA")

        self.colors = {
            "bg": "#F5F6FA", "card": "#FFFFFF",
            "primary": "#4A6CF7", "primary_hover": "#3B5DE7",
            "text": "#2D3436", "subtext": "#636E72",
            "success": "#00B894", "error": "#E17055",
            "border": "#DFE6E9", "drop_bg": "#F0F3FF",
            "drop_border": "#B2BEC3",
        }
        self._build_ui()

    def _build_ui(self):
        bg = self.colors["bg"]; card = self.colors["card"]
        r = self.root

        # === 标题 ===
        h = tk.Frame(r, bg=bg)
        h.pack(fill="x", padx=30, pady=(20, 8))
        tk.Label(h, text="FormatForge", font=("Segoe UI", 20, "bold"),
                 fg=self.colors["primary"], bg=bg).pack(anchor="w")
        tk.Label(h, text="拖拽或选择文件，一键转换格式",
                 font=("Microsoft YaHei", 9), fg=self.colors["subtext"], bg=bg).pack(anchor="w")

        # === 投放区 ===
        self.drop_card = tk.Frame(r, bg=card, bd=0)
        self.drop_card.pack(padx=30, pady=(5, 10), fill="x", ipady=22)

        di = tk.Frame(self.drop_card, bg=self.colors["drop_bg"], bd=0,
                      highlightthickness=1, highlightbackground=self.colors["drop_border"])
        di.pack(padx=18, pady=16, fill="both", ipady=16)

        self.drop_icon = tk.Label(di, text="📂", font=("Segoe UI", 24), bg=self.colors["drop_bg"])
        self.drop_icon.pack()
        self.drop_label = tk.Label(di, text="拖拽文件到此处，或点击下方按钮选择",
                                   font=("Microsoft YaHei", 10),
                                   fg=self.colors["subtext"], bg=self.colors["drop_bg"])
        self.drop_label.pack(pady=(4, 8))
        self.browse_btn = tk.Button(di, text="＋ 选择文件",
                                    font=("Microsoft YaHei", 10, "bold"),
                                    bg=self.colors["primary"], fg="white",
                                    bd=0, padx=22, pady=5,
                                    activebackground=self.colors["primary_hover"],
                                    activeforeground="white",
                                    cursor="hand2", relief="flat",
                                    command=self._browse_file)
        self.browse_btn.pack()

        # === 文件信息（隐藏） ===
        self.info_card = tk.Frame(r, bg=card, bd=0)
        ii = tk.Frame(self.info_card, bg=card)
        ii.pack(padx=16, pady=8, fill="x")
        self.file_icon_var = tk.StringVar()
        self.file_info_var = tk.StringVar()
        self.file_size_var = tk.StringVar()
        tk.Label(ii, textvariable=self.file_icon_var, font=("Segoe UI", 18), bg=card).pack(side="left", padx=(0, 8))
        tk.Label(ii, textvariable=self.file_info_var, font=("Microsoft YaHei", 10, "bold"),
                 fg=self.colors["text"], bg=card).pack(side="left")
        tk.Label(ii, textvariable=self.file_size_var, font=("Microsoft YaHei", 9),
                 fg=self.colors["subtext"], bg=card).pack(side="left", padx=(8, 0))

        # === 目标格式 + 转换按钮（同一卡片） ===
        self.target_card = tk.Frame(r, bg=card, bd=0)
        self.ti = tk.Frame(self.target_card, bg=card)
        self.ti.pack(padx=16, pady=10, fill="x")

        tk.Label(self.ti, text="转换为", font=("Microsoft YaHei", 10, "bold"),
                 fg=self.colors["text"], bg=card).pack(side="left", padx=(0, 10))
        self.target_buttons = {}

        # 开始转换按钮（放在目标格式同一行右侧）
        self.convert_btn = tk.Button(self.ti, text="▶ 开始转换",
                                     font=("Microsoft YaHei", 10, "bold"),
                                     bg=self.colors["primary"], fg="white",
                                     bd=0, padx=18, pady=5,
                                     activebackground=self.colors["primary_hover"],
                                     activeforeground="white",
                                     cursor="hand2", relief="flat",
                                     state="disabled",
                                     command=self._start_conversion)
        self.convert_btn.pack(side="right", padx=(10, 0))

        # === 输出路径 ===
        self.output_card = tk.Frame(r, bg=card, bd=0)
        oi = tk.Frame(self.output_card, bg=card)
        oi.pack(padx=16, pady=8, fill="x")
        tk.Label(oi, text="输出到", font=("Microsoft YaHei", 10, "bold"),
                 fg=self.colors["text"], bg=card).pack(side="left", padx=(0, 8))
        self.output_var = tk.StringVar()
        tk.Label(oi, textvariable=self.output_var, font=("Microsoft YaHei", 9),
                 fg="#0984E3", bg=card, wraplength=300).pack(side="left")

        bf = tk.Frame(oi, bg=card)
        bf.pack(side="right")
        tk.Button(bf, text="📂", font=("Segoe UI", 10), bd=0, bg=card,
                  cursor="hand2", command=self._open_folder,
                  activebackground="#DFE6E9").pack(side="left", padx=1)
        tk.Button(bf, text="✏️", font=("Segoe UI", 10), bd=0, bg=card,
                  cursor="hand2", command=self._change_output,
                  activebackground="#DFE6E9").pack(side="left", padx=1)

        # === 进度条 ===
        self.progress = ttk.Progressbar(r, mode="indeterminate", length=500)

        # === 状态 ===
        self.status_var = tk.StringVar()
        self.status_label = tk.Label(r, textvariable=self.status_var,
                                     font=("Microsoft YaHei", 9),
                                     fg=self.colors["subtext"], bg=bg)
        self.status_label.pack(pady=(8, 16))

    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="选择文档",
            filetypes=[
                ("所有支持的格式", "*.docx *.doc *.pdf *.md"),
                ("Word 文档", "*.docx *.doc"), ("PDF 文档", "*.pdf"), ("Markdown", "*.md"),
            ]
        )
        if not path or not os.path.exists(path): return
        if not is_valid_file(path):
            messagebox.showwarning("不支持的格式", "支持：.docx / .doc / .pdf / .md")
            return

        self.input_path = path
        self.input_format = detect_format(path)
        size_kb = os.path.getsize(path) // 1024
        size_str = f"{size_kb} KB" if size_kb < 1024 else f"{size_kb // 1024} MB"

        self.file_icon_var.set(FORMAT_ICONS.get(self.input_format, "📄"))
        self.file_info_var.set(os.path.basename(path))
        self.file_size_var.set(size_str)
        self.drop_icon.config(text="✅")
        self.drop_label.config(text="已选择文件，可重新选择")

        self.info_card.pack(padx=30, pady=(0, 6), fill="x", ipady=6, after=self.drop_card)
        self._update_targets()

    def _update_targets(self):
        for btn in self.target_buttons.values():
            btn.destroy()
        self.target_buttons.clear()

        targets = CONVERSION_TARGETS.get(self.input_format, {})
        if not targets:
            self.target_card.pack_forget()
            self.output_card.pack_forget()
            return

        self.target_card.pack(padx=30, pady=(0, 6), fill="x", ipady=6, after=self.info_card)

        first = True
        for label, (ext, _) in targets.items():
            color = TARGET_COLORS.get(ext, self.colors["primary"])
            btn = tk.Button(self.ti, text=label, width=9,
                            font=("Microsoft YaHei", 9, "bold"),
                            bg=color, fg="white", bd=0, padx=10, pady=4,
                            activebackground=color, activeforeground="white",
                            cursor="hand2", relief="flat",
                            command=lambda e=ext, l=label, c=color: self._select_target(e, l, c))
            # pack after the "转换为" label but before convert_btn
            btn.pack(side="left", padx=2, before=self.convert_btn)
            self.target_buttons[label] = btn
            if first:
                self._select_target(ext, label, color)
                first = False

    def _select_target(self, ext, label, color):
        self.target_ext = ext
        self.output_path = get_output_path(self.input_path, ext)
        self.output_var.set(self.output_path)
        self.convert_btn.config(state="normal", bg=color, activebackground=color)
        self.status_var.set("")
        self.status_label.config(fg=self.colors["subtext"])
        self.output_card.pack(padx=30, pady=(0, 6), fill="x", ipady=2, after=self.target_card)
        for lbl, btn in self.target_buttons.items():
            btn.config(relief="flat" if lbl != label else "sunken", bd=2 if lbl == label else 0)

    def _change_output(self):
        if not self.input_path or not self.target_ext: return
        default_name = os.path.splitext(os.path.basename(self.input_path))[0]
        path = filedialog.asksaveasfilename(
            title="选择输出位置",
            initialfile=f"{default_name}.{self.target_ext}",
            defaultextension=f".{self.target_ext}",
            filetypes=[(f"*.{self.target_ext}", f"*.{self.target_ext}")]
        )
        if path:
            self.output_path = path
            self.output_var.set(self.output_path)

    def _open_folder(self):
        if self.output_path and os.path.exists(os.path.dirname(self.output_path)):
            subprocess.Popen(["explorer", os.path.dirname(self.output_path)])

    def _start_conversion(self):
        if not self.input_path or not self.output_path: return

        self.convert_btn.config(state="disabled", text="⏳ 转换中...", bg=self.colors["subtext"])
        self.progress.pack(padx=30, pady=(4, 4), fill="x")
        self.progress.start()
        self.status_var.set("正在转换，请稍候...")
        self.status_label.config(fg=self.colors["primary"])
        self.root.update()

        targets = CONVERSION_TARGETS.get(self.input_format, {})
        engine_func = None
        for lbl, (ext, func) in targets.items():
            if ext == self.target_ext:
                engine_func = func
                break

        def run():
            if engine_func is None:
                result = (False, "不支持的转换方向")
            elif engine_func == "md_to_docx":
                from engines.markdown_engine import md_to_docx
                result = md_to_docx(self.input_path, self.output_path)
            elif engine_func == "md_to_pdf":
                from engines.markdown_engine import md_to_docx
                tmp = os.path.splitext(self.output_path)[0] + "_tmp.docx"
                ok, msg = md_to_docx(self.input_path, tmp)
                if ok:
                    ok2, msg2 = docx_to_pdf(tmp, self.output_path)
                    try: os.remove(tmp)
                    except: pass
                    result = (ok2, msg2)
                else:
                    result = (ok, msg)
            else:
                result = engine_func(self.input_path, self.output_path)
            self.root.after(0, lambda: self._on_done(result))

        threading.Thread(target=run, daemon=True).start()

    def _on_done(self, result):
        self.progress.stop(); self.progress.pack_forget()
        success, message = result
        if success:
            self.status_var.set(f"✅ {message}")
            self.status_label.config(fg=self.colors["success"])
            if self.output_path and os.path.exists(self.output_path):
                sz = os.path.getsize(self.output_path) // 1024
                szs = f"{sz} KB" if sz < 1024 else f"{sz//1024} MB"
                self.output_var.set(f"{self.output_path}  ({szs}) ✓")
            self.convert_btn.config(text="✓ 完成", bg=self.colors["success"], state="normal")
            self.converted = True
        else:
            self.status_var.set(f"❌ {message}")
            self.status_label.config(fg=self.colors["error"])
            self.convert_btn.config(text="重试", state="normal", bg=self.colors["error"])
        self.convert_btn.config(command=self._reset_or_convert)

    def _reset_or_convert(self):
        if self.converted: self._reset()
        else: self._start_conversion()

    def _reset(self):
        self.input_path = self.input_format = self.output_path = self.target_ext = None
        self.converted = False
        self.file_icon_var.set(""); self.file_info_var.set(""); self.file_size_var.set("")
        self.output_var.set(""); self.status_var.set("")
        self.info_card.pack_forget(); self.target_card.pack_forget()
        self.output_card.pack_forget()
        self.drop_icon.config(text="📂")
        self.drop_label.config(text="拖拽文件到此处，或点击下方按钮选择")
        for btn in self.target_buttons.values(): btn.destroy()
        self.target_buttons.clear()
        self.convert_btn.config(text="▶ 开始转换", state="disabled",
                                bg=self.colors["primary"], command=self._start_conversion)


def main():
    root = tk.Tk()
    FormatConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
