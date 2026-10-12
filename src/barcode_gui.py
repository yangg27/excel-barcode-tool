#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
条形码批量生成工具 - 图形界面版

双击运行，选 Excel 文件，点一个按钮就出结果。无需命令行操作。

用法:
  双击 "条形码工具.bat"
  或直接双击本文件（.pyw 关联）
  也支持把 Excel 文件拖到 .bat 图标上直接打开处理
"""
from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import traceback

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except ImportError:
    import ctypes
    ctypes.windll.user32.MessageBoxW(
        0, "当前 Python 缺少图形界面组件(tkinter)，无法启动窗口。\n请改用命令行版 excel_barcode.py。",
        "条形码工具 - 启动失败", 0x10)
    sys.exit(1)

try:
    import excel_barcode as core
except Exception as _e:  # noqa: BLE001
    import ctypes
    ctypes.windll.user32.MessageBoxW(
        0, f"核心模块加载失败：\n{_e}\n\n请确认 excel_barcode.py 与本文件在同一目录。",
        "条形码工具 - 启动失败", 0x10)
    sys.exit(1)


PROMPT_FONT = ("Microsoft YaHei UI", 9)
TITLE_FONT = ("Microsoft YaHei UI", 13, "bold")
SMALL_FONT = ("Microsoft YaHei UI", 8)


class App:
    def __init__(self, root: tk.Tk, initial_file: str = ""):
        self.root = root
        self.msg_q: queue.Queue = queue.Queue()
        self.result: dict | None = None
        self.running = False
        self._build_ui()
        if initial_file and os.path.isfile(initial_file):
            self.file_var.set(initial_file)
            self.root.after(200, self.on_file_chosen)

    # ------------------------------------------------------------ 界面 --
    def _build_ui(self):
        r = self.root
        r.title("条形码批量生成工具")
        r.geometry("560x520")
        r.minsize(520, 480)
        try:
            r.iconbitmap(default="")
        except Exception:  # noqa: BLE001
            pass

        style = ttk.Style()
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure(".", font=PROMPT_FONT)
        style.configure("Title.TLabel", font=TITLE_FONT)
        style.configure("Hint.TLabel", font=SMALL_FONT, foreground="#777777")
        style.configure("Go.TButton", font=("Microsoft YaHei UI", 11, "bold"))

        outer = ttk.Frame(r, padding=(14, 12, 14, 10))
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="条形码批量生成工具", style="Title.TLabel").pack(anchor="w")
        ttk.Label(outer, text="Excel / CSV 里的编码  →  一键生成 Code128 条形码",
                  style="Hint.TLabel").pack(anchor="w", pady=(2, 12))

        # --- 步骤1: 文件
        f1 = ttk.LabelFrame(outer, text=" 第 1 步　选择表格文件 ", padding=(10, 8))
        f1.pack(fill="x")
        row = ttk.Frame(f1)
        row.pack(fill="x")
        self.file_var = tk.StringVar()
        entry = ttk.Entry(row, textvariable=self.file_var)
        entry.pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="浏览…", width=9, command=self.pick_file).pack(side="left", padx=(6, 0))
        ttk.Label(f1, text="支持 .xlsx / .xlsm / .xls / .csv　（也可把文件拖到本窗口）",
                  style="Hint.TLabel").pack(anchor="w", pady=(5, 0))

        # --- 步骤2: 编码列
        f2 = ttk.LabelFrame(outer, text=" 第 2 步　确认编码所在列 ", padding=(10, 8))
        f2.pack(fill="x", pady=(10, 0))
        row2 = ttk.Frame(f2)
        row2.pack(fill="x")
        ttk.Label(row2, text="编码列：").pack(side="left")
        self.col_var = tk.StringVar(value="自动识别")
        self.col_box = ttk.Combobox(row2, textvariable=self.col_var, state="readonly", width=26)
        self.col_box["values"] = ("自动识别",)
        self.col_box.pack(side="left", fill="x", expand=True)
        self.sheet_var = tk.StringVar()
        ttk.Label(row2, text="  工作表：").pack(side="left")
        self.sheet_box = ttk.Combobox(row2, textvariable=self.sheet_var, state="readonly", width=14)
        self.sheet_box.pack(side="left")
        self.sheet_box.bind("<<ComboboxSelected>>", lambda e: self.on_file_chosen(silent=True))
        ttk.Label(f2, text="打开表格后自动读取表头；识别不准时在此手动指定列",
                  style="Hint.TLabel").pack(anchor="w", pady=(5, 0))

        # --- 步骤3: 输出
        f3 = ttk.LabelFrame(outer, text=" 第 3 步　输出设置 ", padding=(10, 8))
        f3.pack(fill="x", pady=(10, 0))
        row3 = ttk.Frame(f3)
        row3.pack(fill="x")
        ttk.Label(row3, text="输出到：").pack(side="left")
        self.out_var = tk.StringVar()
        ttk.Entry(row3, textvariable=self.out_var).pack(side="left", fill="x", expand=True)
        ttk.Button(row3, text="浏览…", width=9, command=self.pick_out).pack(side="left", padx=(6, 0))

        row4 = ttk.Frame(f3)
        row4.pack(fill="x", pady=(8, 0))
        self.cols_var = tk.StringVar(value="3")
        self.bw_var = tk.StringVar(value="2")
        self.bh_var = tk.StringVar(value="60")
        ttk.Label(row4, text="每行标签数").pack(side="left")
        ttk.Spinbox(row4, from_=1, to=8, width=4, textvariable=self.cols_var).pack(side="left", padx=(4, 14))
        ttk.Label(row4, text="条码宽度").pack(side="left")
        ttk.Spinbox(row4, from_=1, to=6, width=4, textvariable=self.bw_var).pack(side="left", padx=(4, 14))
        ttk.Label(row4, text="高度").pack(side="left")
        ttk.Spinbox(row4, from_=20, to=200, increment=10, width=5,
                    textvariable=self.bh_var).pack(side="left", padx=(4, 0))

        row5 = ttk.Frame(f3)
        row5.pack(fill="x", pady=(8, 0))
        self.opt_text = tk.BooleanVar(value=True)
        self.opt_png = tk.BooleanVar(value=True)
        self.opt_dedupe = tk.BooleanVar(value=False)
        ttk.Checkbutton(row5, text="显示编码文字", variable=self.opt_text).pack(side="left")
        ttk.Checkbutton(row5, text="同时输出 PNG 图片", variable=self.opt_png).pack(side="left", padx=(12, 0))
        ttk.Checkbutton(row5, text="重复编码去重", variable=self.opt_dedupe).pack(side="left", padx=(12, 0))

        # --- 主按钮
        self.go_btn = ttk.Button(outer, text="开始生成条形码", style="Go.TButton", command=self.run)
        self.go_btn.pack(fill="x", pady=(14, 10), ipady=6)

        # --- 状态区
        f4 = ttk.LabelFrame(outer, text=" 运行状态 ", padding=(8, 6))
        f4.pack(fill="both", expand=True)
        self.log = tk.Text(f4, height=7, wrap="word", relief="flat", bg="#F7F8FA",
                           font=("Microsoft YaHei UI", 9), state="disabled")
        self.log.pack(fill="both", expand=True)
        self.log.tag_config("ok", foreground="#0A7A2F")
        self.log.tag_config("err", foreground="#C00000")
        self.log.tag_config("dim", foreground="#777777")

        # --- 完成后动作
        act = ttk.Frame(outer)
        act.pack(fill="x", pady=(8, 0))
        self.btn_html = ttk.Button(act, text="打开打印页", state="disabled",
                                   command=lambda: self.open_path("html"))
        self.btn_html.pack(side="left")
        self.btn_dir = ttk.Button(act, text="打开输出文件夹", state="disabled",
                                  command=lambda: self.open_path("dir"))
        self.btn_dir.pack(side="left", padx=(8, 0))

        self._log("请选择要处理的 Excel 或 CSV 文件。", "dim")
        self._poll()
        self._enable_drop()

    # ------------------------------------------------------ 拖放支持 --
    def _enable_drop(self):
        """尝试启用文件拖放（需 tkinterdnd2，缺失则静默跳过）。"""
        try:
            import tkinterdnd2  # type: ignore
        except ImportError:
            return
        try:
            self.root.drop_target_register("DND_Files")  # type: ignore
            self.root.dnd_bind("<<Drop>>", self._on_drop)  # type: ignore
        except Exception:  # noqa: BLE001
            pass

    def _on_drop(self, event):
        raw = event.data.strip()
        if raw.startswith("{") and raw.endswith("}"):
            raw = raw[1:-1]
        path = raw.split("} {")[0].strip("{}")
        if os.path.isfile(path):
            self.file_var.set(path)
            self.on_file_chosen()

    # ---------------------------------------------------------- 选择 --
    def pick_file(self):
        path = filedialog.askopenfilename(
            title="选择编码表格",
            filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.xls *.csv *.txt"),
                       ("Excel 工作簿", "*.xlsx *.xlsm *.xls"),
                       ("CSV 文件", "*.csv *.txt"),
                       ("所有文件", "*.*")])
        if path:
            self.file_var.set(path)
            self.on_file_chosen()

    def pick_out(self):
        d = filedialog.askdirectory(title="选择输出文件夹")
        if d:
            self.out_var.set(d)

    # ---------------------------------------------------- 读取表头 --
    def on_file_chosen(self, silent: bool = False):
        path = self.file_var.get().strip()
        if not path or not os.path.isfile(path):
            return
        # 换了文件则重置工作表选择, 避免沿用旧文件的工作表名
        if path != getattr(self, "_cur_path", None):
            self._cur_path = path
            self.sheet_var.set("")
        try:
            sheet = self.sheet_var.get().strip() or None
            sheets = self._list_sheets(path)
            if sheets:
                self.sheet_box["values"] = sheets
                if not self.sheet_var.get():
                    self.sheet_var.set(sheets[0])
                    sheet = None
            else:
                self.sheet_box["values"] = ()
            rows = core.read_table(path, sheet)
        except Exception as e:  # noqa: BLE001
            if not silent:
                self._log(f"读取失败：{e}", "err")
            return

        headers: list[str] = []
        if rows:
            for cell in rows[0]:
                c = cell.strip()
                headers.append(c if c else "(空列)")
        values = ["自动识别"] + [f"{chr(65 + i) if i < 26 else i + 1} 列 · {h}"
                                 for i, h in enumerate(headers)]
        self.col_box["values"] = values
        if self.col_var.get() not in values:
            self.col_var.set("自动识别")

        # 自动推荐输出目录
        if not self.out_var.get().strip():
            base = os.path.dirname(os.path.abspath(path))
            self.out_var.set(os.path.join(base, "条码输出"))

        if not silent:
            n = max(0, len(rows) - 1)
            self._log(f"已读取：{os.path.basename(path)}　共约 {n} 行数据", "dim")
            self._log("　　　 表头：" + " | ".join(headers[:8]), "dim")

    @staticmethod
    def _list_sheets(path: str) -> list[str]:
        ext = os.path.splitext(path)[1].lower()
        if ext not in (".xlsx", ".xlsm", ".xltx", ".xls"):
            return []
        try:
            import zipfile
            if ext == ".xls":
                import openpyxl  # type: ignore
                return list(openpyxl.load_workbook(path, read_only=True).sheetnames)
            with zipfile.ZipFile(path) as zf:
                return [name for name, _ in core._xlsx_sheet_map(zf)]
        except Exception:  # noqa: BLE001
            return []

    # -------------------------------------------------------- 执行 --
    def run(self):
        if self.running:
            return
        path = self.file_var.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showwarning("还没选文件", "请先选择要处理的 Excel 或 CSV 文件。")
            return

        try:
            cols = int(self.cols_var.get())
            bw = int(self.bw_var.get())
            bh = int(self.bh_var.get())
            if not (1 <= cols <= 8 and 1 <= bw <= 6 and 20 <= bh <= 200):
                raise ValueError
        except ValueError:
            messagebox.showwarning("参数不对", "请检查每行标签数(1-8)、条码宽度(1-6)、高度(20-200)。")
            return

        # 解析列选择
        column = None
        cv = self.col_var.get().strip()
        if cv and cv != "自动识别":
            letter = cv.split(" 列")[0].strip()
            column = letter or None
            # 若是 "A 列 · 资产编码" 形式, 优先用列标, 更稳

        out_dir = self.out_var.get().strip() or None
        sheet = self.sheet_var.get().strip() or None

        self.running = True
        self.go_btn.config(state="disabled", text="正在生成，请稍候…")
        self.btn_html.config(state="disabled")
        self.btn_dir.config(state="disabled")
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")

        # 界面变量必须在主线程读取，再交给后台线程使用
        opts = {
            "text": self.opt_text.get(),
            "png": self.opt_png.get(),
            "dedupe": self.opt_dedupe.get(),
        }

        t = threading.Thread(target=self._worker,
                             args=(path, column, sheet, out_dir, cols, bw, bh, opts), daemon=True)
        t.start()

    def _worker(self, path, column, sheet, out_dir, cols, bw, bh, opts):
        try:
            rows = core.read_table(path, sheet)
            col_idx, header_idx = core.detect_column(rows, column, None)
            codes, meta = core.build(rows, col_idx, header_idx, opts["dedupe"])
            if not codes:
                raise core.BarcodeError(f"第 {col_idx + 1} 列没有读到任何编码，请确认第 2 步的列选择")

            if not out_dir:
                out_dir = os.path.join(os.path.dirname(os.path.abspath(path)), "条码输出")
            os.makedirs(out_dir, exist_ok=True)

            class _A:
                pass
            a = _A()
            a.bar_width = bw
            a.barcode_height = bh
            a.no_text = not opts["text"]
            a.font_size = 14
            a.png = opts["png"]
            a.cols = cols

            res = core.write_outputs(codes, out_dir, a)
            name = (rows[header_idx][col_idx]
                    if header_idx is not None and col_idx < len(rows[header_idx])
                    else f"第{col_idx + 1}列")
            self.result = {
                "out_dir": out_dir, "html": res["html"], "csv": res["csv"],
                "svg_dir": res["svg_dir"], "count": len(res["items"]),
                "failed": len(res["errors"]), "errors": res["errors"],
                "column_name": name, "col_index": col_idx + 1,
                "dup": meta["duplicates"] if opts["dedupe"] else 0,
                "png_count": res["png_count"],
            }
        except Exception as e:  # noqa: BLE001
            self.result = {"error": f"{e}", "trace": traceback.format_exc()}
        finally:
            self.msg_q.put("done")

    # ------------------------------------------------------ 状态刷新 --
    def _poll(self):
        try:
            while True:
                self.msg_q.get_nowait()
                self._on_done()
        except queue.Empty:
            pass
        self.root.after(120, self._poll)

    def _on_done(self):
        self.running = False
        self.go_btn.config(state="normal", text="开始生成条形码")
        r = self.result
        if not r:
            return
        if "error" in r:
            self._log("生成失败：" + r["error"], "err")
            return
        self._log(f"完成！共生成 {r['count']} 个条形码（Code128）", "ok")
        self._log(f"　　编码列：第 {r['col_index']} 列 · {r['column_name']}", "dim")
        if r["dup"]:
            self._log(f"　　重复编码已跳过 {r['dup']} 个", "dim")
        if r["failed"]:
            self._log(f"　　有 {r['failed']} 个编码无法生成（含中文等非 ASCII 字符会被跳过）", "err")
        if r["png_count"]:
            self._log(f"　　另输出 {r['png_count']} 张 PNG 图片", "dim")
        self._log(f"　　输出位置：{r['out_dir']}", "dim")
        self._log("", "dim")
        self._log("点击下方「打开打印页」，在浏览器里按 Ctrl+P 即可打印贴标。", "dim")
        self.btn_html.config(state="normal")
        self.btn_dir.config(state="normal")

    def _log(self, text: str, tag: str = ""):
        self.log.config(state="normal")
        self.log.insert("end", text + "\n", tag)
        self.log.see("end")
        self.log.config(state="disabled")

    def open_path(self, what: str):
        if not self.result:
            return
        if what == "html":
            target = self.result.get("html")
        else:
            target = self.result.get("out_dir")
        if not target or not os.path.exists(target):
            messagebox.showinfo("找不到", "输出文件不存在，请重新生成。")
            return
        try:
            os.startfile(target)  # type: ignore[attr-defined]
        except Exception:  # noqa: BLE001
            subprocess.Popen(["explorer", "/select,", target])


def main():
    initial = sys.argv[1] if len(sys.argv) > 1 else ""
    root = tk.Tk()
    App(root, initial)
    root.mainloop()


if __name__ == "__main__":
    main()