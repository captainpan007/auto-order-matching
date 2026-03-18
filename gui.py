# -*- coding: utf-8 -*-
"""
超市采购对账系统 - GUI 启动界面
双击运行即可启动
"""

import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import threading
import subprocess
import sys
import os
import json
from pathlib import Path


# ─── 默认配置 ───
DEFAULT_BASE_DIR = r"F:\claude开发项目\atutoordermatching\3.16-对账汇总"
DEFAULT_OUTPUT_DIR = r"F:\claude开发项目\atutoordermatching\对账结果"


class ReconcileApp:
    def __init__(self, root):
        self.root = root
        self.root.title("超市采购对账系统")
        self.root.geometry("800x620")
        self.root.resizable(True, True)
        self._center_window()

        self.base_dir = tk.StringVar(value=DEFAULT_BASE_DIR)
        self.no_cache = tk.BooleanVar(value=False)
        self.running = False
        self.supplier_vars = {}

        self._build_ui()
        self._scan_suppliers()

    def _center_window(self):
        self.root.update_idletasks()
        w, h = 800, 620
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        # 标题
        title = tk.Label(self.root, text="超市采购对账系统", font=("Microsoft YaHei", 18, "bold"))
        title.pack(pady=(10, 5))

        # 数据目录
        dir_frame = tk.Frame(self.root)
        dir_frame.pack(fill="x", padx=15, pady=5)
        tk.Label(dir_frame, text="数据目录:", font=("Microsoft YaHei", 10)).pack(side="left")
        tk.Entry(dir_frame, textvariable=self.base_dir, font=("Consolas", 9), width=55).pack(side="left", padx=5, fill="x", expand=True)
        tk.Button(dir_frame, text="浏览", command=self._browse_dir, width=6).pack(side="left")

        # 供应商选择 + 选项（左右布局）
        mid_frame = tk.Frame(self.root)
        mid_frame.pack(fill="x", padx=15, pady=5)

        # 左侧：供应商列表
        supplier_frame = tk.LabelFrame(mid_frame, text="供应商（勾选参与对账）", font=("Microsoft YaHei", 9))
        supplier_frame.pack(side="left", fill="both", expand=True)

        self.supplier_inner = tk.Frame(supplier_frame)
        self.supplier_inner.pack(fill="both", padx=5, pady=3)

        # 全选/取消按钮
        btn_row = tk.Frame(supplier_frame)
        btn_row.pack(fill="x", padx=5, pady=2)
        tk.Button(btn_row, text="全选", command=self._select_all, width=6).pack(side="left", padx=2)
        tk.Button(btn_row, text="取消全选", command=self._deselect_all, width=8).pack(side="left", padx=2)

        # 右侧：选项
        opt_frame = tk.LabelFrame(mid_frame, text="选项", font=("Microsoft YaHei", 9))
        opt_frame.pack(side="left", fill="y", padx=(10, 0))

        tk.Checkbutton(opt_frame, text="强制重新OCR\n(清除缓存)", variable=self.no_cache,
                       font=("Microsoft YaHei", 9), justify="left").pack(anchor="w", padx=10, pady=10)

        # 启动按钮
        self.start_btn = tk.Button(self.root, text="开始对账", font=("Microsoft YaHei", 14, "bold"),
                                   bg="#4472C4", fg="white", activebackground="#365fa1",
                                   command=self._start, height=1, width=20, cursor="hand2")
        self.start_btn.pack(pady=8)

        # 进度区域
        progress_frame = tk.Frame(self.root)
        progress_frame.pack(fill="x", padx=15)

        self.status_label = tk.Label(progress_frame, text="就绪", font=("Microsoft YaHei", 10),
                                     anchor="w", fg="#333")
        self.status_label.pack(fill="x")

        self.progress = ttk.Progressbar(progress_frame, mode="determinate", length=750)
        self.progress.pack(fill="x", pady=3)

        # 日志窗口
        self.log_text = scrolledtext.ScrolledText(self.root, height=12, font=("Consolas", 9),
                                                   bg="#1e1e1e", fg="#d4d4d4",
                                                   insertbackground="white", wrap="word")
        self.log_text.pack(fill="both", expand=True, padx=15, pady=(5, 5))

        # 底部按钮
        bottom_frame = tk.Frame(self.root)
        bottom_frame.pack(fill="x", padx=15, pady=(0, 10))

        self.open_btn = tk.Button(bottom_frame, text="打开报告文件夹", command=self._open_output,
                                  state="disabled", width=18)
        self.open_btn.pack(side="right")

        self.summary_label = tk.Label(bottom_frame, text="", font=("Microsoft YaHei", 9),
                                       anchor="w", fg="#666")
        self.summary_label.pack(side="left", fill="x", expand=True)

    def _scan_suppliers(self):
        """扫描数据目录下的供应商文件夹"""
        for w in self.supplier_inner.winfo_children():
            w.destroy()
        self.supplier_vars.clear()

        base = self.base_dir.get()
        if not os.path.isdir(base):
            return

        suppliers = sorted([d.name for d in Path(base).iterdir()
                           if d.is_dir() and not d.name.startswith("_")])

        # 每行放5个
        cols = 5
        for i, name in enumerate(suppliers):
            var = tk.BooleanVar(value=True)
            self.supplier_vars[name] = var
            cb = tk.Checkbutton(self.supplier_inner, text=name, variable=var,
                               font=("Microsoft YaHei", 9))
            cb.grid(row=i // cols, column=i % cols, sticky="w", padx=3)

    def _browse_dir(self):
        d = filedialog.askdirectory(initialdir=self.base_dir.get())
        if d:
            self.base_dir.set(d)
            self._scan_suppliers()

    def _select_all(self):
        for var in self.supplier_vars.values():
            var.set(True)

    def _deselect_all(self):
        for var in self.supplier_vars.values():
            var.set(False)

    def _log(self, text):
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")

    def _open_output(self):
        output_dir = DEFAULT_OUTPUT_DIR
        if os.path.isdir(output_dir):
            os.startfile(output_dir)

    def _start(self):
        if self.running:
            return

        selected = [name for name, var in self.supplier_vars.items() if var.get()]
        if not selected:
            self._log("请至少选择一个供应商")
            return

        self.running = True
        self.start_btn.config(state="disabled", text="对账中...", bg="#999")
        self.open_btn.config(state="disabled")
        self.log_text.delete("1.0", "end")
        self.progress["value"] = 0
        self.summary_label.config(text="")

        thread = threading.Thread(target=self._run_reconcile, args=(selected,), daemon=True)
        thread.start()

    def _run_reconcile(self, selected_suppliers):
        try:
            base_dir = self.base_dir.get()
            total = len(selected_suppliers)

            self.root.after(0, lambda: self.progress.config(maximum=total))
            self.root.after(0, lambda: self._log(f"开始对账: {total} 家供应商\n"))

            # 构建命令
            cmd = [sys.executable, "-X", "utf8", "-u",
                   os.path.join(os.path.dirname(__file__), "main.py")]
            if self.no_cache.get():
                cmd.append("--no-cache")

            # 设置环境变量传递选中的供应商
            env = os.environ.copy()
            env["RECONCILE_SUPPLIERS"] = ",".join(selected_suppliers)
            env["PYTHONIOENCODING"] = "utf-8"

            # 启动子进程
            process = subprocess.Popen(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, encoding="utf-8", errors="replace",
                env=env, cwd=os.path.dirname(__file__)
            )

            current_supplier = 0
            for line in process.stdout:
                line = line.rstrip()
                if not line:
                    continue

                # 更新进度
                if "处理供应商:" in line:
                    current_supplier += 1
                    supplier_name = line.split("处理供应商:")[-1].strip()
                    self.root.after(0, lambda s=supplier_name, n=current_supplier:
                        self._update_progress(s, n, total))

                # 输出到日志
                self.root.after(0, lambda l=line: self._log(l))

            process.wait()

            # 完成
            self.root.after(0, lambda: self._on_complete(total))

        except Exception as e:
            self.root.after(0, lambda: self._log(f"\n错误: {e}"))
            self.root.after(0, self._reset_ui)

    def _update_progress(self, supplier_name, current, total):
        self.status_label.config(text=f"正在处理: {supplier_name} ({current}/{total})")
        self.progress["value"] = current

    def _on_complete(self, total):
        self.progress["value"] = total
        self.status_label.config(text="对账完成", fg="#2e7d32")
        self.start_btn.config(state="normal", text="开始对账", bg="#4472C4")
        self.open_btn.config(state="normal")
        self.running = False

        self._log("\n" + "=" * 50)
        self._log("  对账完成！报告已生成。")
        self._log("=" * 50)

        self.summary_label.config(text=f"已完成 {total} 家供应商对账，点击右侧按钮查看报告")

    def _reset_ui(self):
        self.start_btn.config(state="normal", text="开始对账", bg="#4472C4")
        self.running = False


def main():
    root = tk.Tk()
    app = ReconcileApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
