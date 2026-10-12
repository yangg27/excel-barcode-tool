#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
excel_barcode.py - Excel/CSV 编码批量生成条形码

功能:
  读取 Excel(.xlsx/.xls/.xlsm) 或 CSV 文件中的一列编码 -> 批量生成 Code128 条形码
  输出: 每码一张 SVG + 一份 A4 打印页 HTML + 编码映射 CSV

零依赖兜底: 内置 XLSX(OOXML) 解析器, 无 openpyxl 也能读 .xlsx; 内置 Code128 编码器。
可选加速: 若环境中有 openpyxl, 自动优先使用(支持 .xls/.xlsm 更稳)。

用法:
  python excel_barcode.py --input 编码表.xlsx
  python excel_barcode.py --input 编码表.xlsx --column 资产编码 --out ./输出
  python excel_barcode.py --input 编码表.csv --column A --cols 3 --json

参数:
  --input / -i    必填, Excel 或 CSV 文件路径
  --column / -c   编码所在列: 列名(表头文字) 或 字母(A/B/C) 或 序号(1/2/3); 默认自动识别
  --sheet / -s    工作表名或序号(1起); 默认第一个工作表
  --out / -o      输出目录; 默认 <输入文件同目录>/barcodes_<时间戳>
  --cols          打印页每行标签数(默认 3)
  --barcode-height 条码高度 px(默认 60)
  --bar-width     单模块宽度 px(默认 2)
  --font-size     文字字号 px(默认 14)
  --no-text       不在条码下方显示编码文字
  --dedupe        去重(相同编码只生成一次)
  --header        强制首行为表头 / --no-header 强制首行非表头
  --json          以 JSON 输出执行摘要
"""
from __future__ import annotations

import argparse
import csv
import io
import json
import os
import re
import sys
import time
import zipfile
from datetime import datetime
from xml.etree import ElementTree as ET

# ---------------------------------------------------------------- Code128 ---
# 每个 code 值对应 6 个元素宽度(条,空,条,空,条,空), STOP 为 7 个元素
_CODE128_PATTERNS = [
    "212222", "222122", "222221", "121223", "121322", "131222", "122213", "122312", "132212", "221213",
    "221312", "231212", "112232", "122132", "122231", "113222", "123122", "123221", "223211", "221132",
    "221231", "213212", "223112", "312131", "311222", "321122", "321221", "312212", "322112", "322211",
    "212123", "212321", "232121", "111323", "131123", "131321", "112313", "132113", "132311", "211313",
    "231113", "231311", "112133", "112331", "132131", "113123", "113321", "133121", "313121", "211331",
    "231131", "213113", "213311", "213131", "311123", "311321", "331121", "312113", "312311", "332111",
    "314111", "221411", "431111", "111224", "111422", "121124", "121421", "141122", "141221", "112214",
    "112412", "122114", "122411", "142112", "142211", "241211", "221114", "413111", "241112", "134111",
    "111242", "121142", "121241", "114212", "124112", "124211", "411212", "421112", "421211", "212141",
    "214121", "412121", "111143", "111341", "131141", "114113", "114311", "411113", "411311", "113141",
    "114131", "311141", "411131", "211412", "211214", "211232", "2331112",
]
_START_A, _START_B, _START_C, _STOP = 103, 104, 105, 106


class BarcodeError(ValueError):
    pass


_START_CODE = {0: _START_A, 1: _START_B, 2: _START_C}
_SWITCH_CODE = {(0, 1): 100, (0, 2): 99,
                (1, 0): 101, (1, 2): 99,
                (2, 0): 101, (2, 1): 100}
_IN_SET_A = lambda o: 0 <= o <= 95    # noqa: E731
_IN_SET_B = lambda o: 32 <= o <= 127  # noqa: E731


def _code128_codewords(text: str) -> list[int]:
    """ISO/IEC 15417 最优编码: 动态规划自动在 A/B/C 字符集间切换并使用数字压缩。

    码字数最少 <=> 条码最短, 与行业库(python-barcode / zint)行为一致。
    """
    text = text.strip()
    if not text:
        raise BarcodeError("编码为空")
    n = len(text)
    for ch in text:
        if not (0 <= ord(ch) <= 127):
            raise BarcodeError(f"Code128 无法编码字符 {ch!r}(仅支持 ASCII 字符)")

    INF = 10 ** 9
    dp = [[INF] * 3 for _ in range(n + 2)]      # dp[i][c]: 从 i 起、当前集 c 的最少码字数
    base = [[INF] * 3 for _ in range(n + 1)]    # 仅含"直接编码"的代价
    dp[n] = [0, 0, 0]
    dp[n + 1] = [0, 0, 0]

    for i in range(n - 1, -1, -1):
        o = ord(text[i])
        b = [INF, INF, INF]
        if _IN_SET_A(o):
            b[0] = 1 + dp[i + 1][0]
        if _IN_SET_B(o):
            b[1] = 1 + dp[i + 1][1]
        if o > 95:                              # 仅属 B -> 在 A 集用 SHIFT(98)
            b[0] = min(b[0], 1 + dp[i + 1][0])
        if o <= 31:                             # 仅属 A -> 在 B 集用 SHIFT(98)
            b[1] = min(b[1], 1 + dp[i + 1][1])
        if i + 1 < n and text[i].isdigit() and text[i + 1].isdigit():
            b[2] = 1 + dp[i + 2][2]             # C 集: 两位数字 = 1 码字
        base[i] = b
        val = b[:]
        for _ in range(2):                      # 集间切换闭包(3 节点, 2 轮收敛)
            nv = val[:]
            for c in range(3):
                for c2 in range(3):
                    if c != c2 and val[c2] + 1 < nv[c]:
                        nv[c] = val[c2] + 1
            val = nv
        dp[i] = val

    # 选起始集(代价相同时优先 C, 其次 B)
    c = min(range(3), key=lambda x: (dp[0][x], {2: 0, 1: 1, 0: 2}[x]))
    codes: list[int] = [_START_CODE[c]]
    i = 0
    while i < n:
        o = ord(text[i])
        if dp[i][c] == base[i][c]:              # 直接编码(含 SHIFT)
            if c == 2:
                codes.append(int(text[i:i + 2]))
                i += 2
            elif c == 0:
                if o <= 95:
                    codes.append(o - 32 if o >= 32 else o + 64)
                else:
                    codes.extend((98, o - 32))  # SHIFT -> A 集临时用 B
                i += 1
            else:
                if o >= 32:
                    codes.append(o - 32)
                else:
                    codes.extend((98, o + 64))  # SHIFT -> B 集临时用 A
                i += 1
        else:                                   # 切换字符集
            c2 = min((x for x in range(3) if x != c), key=lambda x: base[i][x])
            codes.append(_SWITCH_CODE[(c, c2)])
            c = c2

    checksum = codes[0] + sum(cv * idx for idx, cv in enumerate(codes[1:], start=1))
    codes.append(checksum % 103)
    codes.append(_STOP)
    return codes


def code128_modules(text: str) -> str:
    """返回 '1'/'0' 模块串(不含静区, 静区由渲染层添加)。"""
    bits: list[str] = []
    for cw in _code128_codewords(text):
        pattern = _CODE128_PATTERNS[cw]
        is_bar = True
        for w in pattern:
            bits.append(("1" if is_bar else "0") * int(w))
            is_bar = not is_bar
    return "".join(bits)


_PATTERN_TO_CODE = {p: i for i, p in enumerate(_CODE128_PATTERNS)}


def code128_decode(modules: str) -> str:
    """独立实现的 Code128 解码器(仅用于自检/往返验证)。"""
    runs: list[int] = []
    prev, cnt = modules[0], 1
    for b in modules[1:]:
        if b == prev:
            cnt += 1
        else:
            runs.append(cnt)
            prev, cnt = b, 1
    runs.append(cnt)

    codes: list[int] = []
    i, total = 0, len(runs)
    while i < total:
        take = 7 if total - i == 7 else 6
        pat = "".join(str(w) for w in runs[i:i + take])
        if take == 7:
            if pat != "2331112":
                raise BarcodeError("STOP 图案非法")
            codes.append(_STOP)
        else:
            if pat not in _PATTERN_TO_CODE:
                raise BarcodeError(f"未知图案 {pat}")
            codes.append(_PATTERN_TO_CODE[pat])
        i += take

    if not codes or codes[-1] != _STOP:
        raise BarcodeError("缺少 STOP")
    body = codes[:-1]
    if len(body) < 2:
        raise BarcodeError("码字过短")
    data, chk = body[:-1], body[-1]
    expect = (data[0] + sum(cv * idx for idx, cv in enumerate(data[1:], start=1))) % 103
    if chk != expect:
        raise BarcodeError(f"校验位不符: 期望 {expect} 实际 {chk}")

    charset = {103: "A", 104: "B", 105: "C"}.get(data[0])
    if charset is None:
        raise BarcodeError("起始符非法")
    out: list[str] = []
    shift = None
    for cv in data[1:]:
        cs = shift or charset
        shift = None
        if cs == "C":
            if cv <= 99:
                out.append(f"{cv:02d}")
            elif cv == 100:
                charset = "B"
            elif cv == 101:
                charset = "A"
            else:
                raise BarcodeError(f"C 集非法码字 {cv}")
        elif cs == "A":
            if cv <= 63:
                out.append(chr(cv + 32))
            elif cv <= 95:
                out.append(chr(cv - 64))
            elif cv == 98:
                shift = "B"
            elif cv == 99:
                charset = "C"
            elif cv == 100:
                charset = "B"
            else:
                raise BarcodeError(f"A 集非法码字 {cv}")
        else:
            if cv <= 95:
                out.append(chr(cv + 32))
            elif cv == 98:
                shift = "A"
            elif cv == 99:
                charset = "C"
            elif cv == 101:
                charset = "A"
            else:
                raise BarcodeError(f"B 集非法码字 {cv}")
    return "".join(out)


def code128_svg(text: str, bar_width: int = 2, height: int = 60,
                show_text: bool = True, font_size: int = 14,
                quiet: int = 10) -> str:
    """生成一张自包含 SVG 条形码。"""
    modules = code128_modules(text)
    total_modules = len(modules) + quiet * 2
    width = total_modules * bar_width
    text_h = int(font_size * 1.8) if show_text else 0
    total_h = height + text_h
    # 合并连续黑条, 减少 rect 数量
    rects: list[str] = []
    i, n = 0, len(modules)
    while i < n:
        if modules[i] == "1":
            j = i
            while j < n and modules[j] == "1":
                j += 1
            x = (quiet + i) * bar_width
            w = (j - i) * bar_width
            rects.append(f'<rect x="{x}" y="0" width="{w}" height="{height}" fill="#000"/>')
            i = j
        else:
            i += 1
    esc = (text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
    label = ""
    if show_text:
        label = (f'<text x="{width / 2:.1f}" y="{height + font_size + 2}" text-anchor="middle" '
                 f'font-family="Consolas,\'Courier New\',monospace" font-size="{font_size}" '
                 f'fill="#000">{esc}</text>')
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{total_h}" '
            f'viewBox="0 0 {width} {total_h}" role="img" aria-label="{esc}">'
            f'<rect width="{width}" height="{total_h}" fill="#fff"/>{label}'
            f'{"".join(rects)}</svg>')


# ------------------------------------------------------------- 输入读取 ---
def _col_letter_to_index(letters: str) -> int:
    idx = 0
    for ch in letters.strip().upper():
        if not ("A" <= ch <= "Z"):
            raise BarcodeError(f"非法列标: {letters}")
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1


_XLSX_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PKG_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def _xlsx_sheet_map(zf: zipfile.ZipFile) -> list[tuple[str, str]]:
    """返回 [(sheet名, zip内路径), ...]"""
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rid2target = {}
    for rel in rels.iter(f"{_PKG_REL_NS}Relationship"):
        rid2target[rel.get("Id")] = rel.get("Target")
    out = []
    for sh in wb.iter(f"{_XLSX_NS}sheet"):
        name = sh.get("name") or ""
        target = rid2target.get(sh.get(f"{_REL_NS}id"), "")
        if not target:
            continue
        if target.startswith("/"):
            path = target.lstrip("/")
        else:
            path = "xl/" + target
        out.append((name, path.replace("/./", "/")))
    return out


def _xlsx_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    strings = []
    for si in root.iter(f"{_XLSX_NS}si"):
        strings.append("".join(t.text or "" for t in si.iter(f"{_XLSX_NS}t")))
    return strings


def _num_to_str(v: str) -> str:
    try:
        f = float(v)
    except (TypeError, ValueError):
        return v
    return str(int(f)) if f.is_integer() else repr(f)


def read_xlsx_stdlib(path: str, sheet: str | None = None) -> list[list[str]]:
    with zipfile.ZipFile(path) as zf:
        sheets = _xlsx_sheet_map(zf)
        if not sheets:
            raise BarcodeError("工作簿中没有工作表")
        target_path = None
        if sheet:
            if sheet.strip().isdigit():
                idx = int(sheet) - 1
                if not (0 <= idx < len(sheets)):
                    raise BarcodeError(f"工作表序号超范围(共 {len(sheets)} 个)")
                target_path = sheets[idx][1]
            else:
                for name, p in sheets:
                    if name.strip() == sheet.strip():
                        target_path = p
                        break
                if target_path is None:
                    raise BarcodeError(f"未找到工作表 {sheet!r}, 可用: {[n for n, _ in sheets]}")
        else:
            target_path = sheets[0][1]

        shared = _xlsx_shared_strings(zf)
        try:
            data = zf.read(target_path)
        except KeyError:
            # 少数文件 rels 目标名不一致, 退回 sheet1.xml
            data = zf.read("xl/worksheets/sheet1.xml")
        root = ET.fromstring(data)
        rows: list[list[str]] = []
        for row in root.iter(f"{_XLSX_NS}row"):
            cells: dict[int, str] = {}
            auto_idx = 0
            for c in row.iter(f"{_XLSX_NS}c"):
                ref = c.get("r") or ""
                m = re.match(r"^([A-Za-z]+)", ref)
                ci = _col_letter_to_index(m.group(1)) if m else auto_idx
                auto_idx = ci + 1
                ctype = c.get("t")
                val = ""
                if ctype == "inlineStr":
                    val = "".join(t.text or "" for t in c.iter(f"{_XLSX_NS}t"))
                else:
                    v = c.find(f"{_XLSX_NS}v")
                    if v is not None and v.text is not None:
                        if ctype == "s":
                            try:
                                val = shared[int(v.text)]
                            except (ValueError, IndexError):
                                val = v.text
                        elif ctype == "b":
                            val = "TRUE" if v.text.strip() == "1" else "FALSE"
                        else:
                            val = _num_to_str(v.text.strip())
                    else:
                        is_el = c.find(f"{_XLSX_NS}is")
                        if is_el is not None:
                            val = "".join(t.text or "" for t in is_el.iter(f"{_XLSX_NS}t"))
                cells[ci] = val
            width = (max(cells) + 1) if cells else 0
            rows.append([cells.get(i, "") for i in range(width)])
        while rows and not any(x.strip() for x in rows[-1]):
            rows.pop()
        return rows


def read_xlsx_openpyxl(path: str, sheet: str | None = None) -> list[list[str]]:
    import openpyxl  # type: ignore

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if sheet:
        if sheet.strip().isdigit():
            ws = wb.worksheets[int(sheet) - 1]
        else:
            ws = wb[sheet.strip()]
    else:
        ws = wb.worksheets[0]
    rows: list[list[str]] = []
    for r in ws.iter_rows(values_only=True):
        rows.append(["" if v is None else (str(int(v)) if isinstance(v, float) and v.is_integer()
                                           else str(v)) for v in r])
    while rows and not any(x.strip() for x in rows[-1]):
        rows.pop()
    return rows


def read_csv_rows(path: str) -> list[list[str]]:
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "utf-8", "gb18030", "gbk", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise BarcodeError("CSV 编码无法识别")
    sample = text[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    return [row for row in csv.reader(io.StringIO(text), dialect)]


def read_table(path: str, sheet: str | None = None) -> list[list[str]]:
    if not os.path.isfile(path):
        raise BarcodeError(f"文件不存在: {path}")
    ext = os.path.splitext(path)[1].lower()
    if ext in (".csv", ".txt", ".tsv"):
        rows = read_csv_rows(path) if ext != ".txt" else [
            [ln.strip()] for ln in open(path, encoding="utf-8", errors="replace").read().splitlines()]
        return rows
    if ext in (".xlsx", ".xlsm", ".xltx", ".xls"):
        try:
            return read_xlsx_openpyxl(path, sheet)
        except ImportError:
            if ext == ".xls":
                raise BarcodeError(".xls 旧格式需安装 openpyxl/pandas 支持, 请另存为 .xlsx 或 .csv")
            return read_xlsx_stdlib(path, sheet)
    raise BarcodeError(f"不支持的文件类型: {ext}")


# --------------------------------------------------------- 列自动识别 ----
# 分级识别: 分越高越像"编码列", 避免"序号"抢占"资产编码"
_HEADER_TIERS = [
    (3, ("编码", "条码", "条形码", "barcode", "code128", "二维码")),
    (2, ("编号", "物料", "资产", "设备", "sku", "sn", "序列号", "货号", "品号")),
    (1, ("序号", "code", "id", "no", "号码")),
]


def detect_column(rows: list[list[str]], column: str | None,
                  header: bool | None) -> tuple[int, int | None]:
    """返回 (列下标, 表头行下标或 None)"""
    if not rows:
        raise BarcodeError("表格无数据")
    first = rows[0]
    # 显式列名 / 列标 / 序号
    if column:
        col = column.strip()
        if col.isdigit():
            return int(col) - 1, (0 if header else None)
        if re.fullmatch(r"[A-Za-z]{1,3}", col):
            return _col_letter_to_index(col), (0 if header else None)
        for i, cell in enumerate(first):
            if cell.strip().lower() == col.lower():
                return i, 0
        raise BarcodeError(f"未找到列名 {column!r}, 表头为: {first}")

    # 自动识别表头关键字（分级打分, 同分取最左）
    header_idx = None
    col_idx = 0
    best_tier = 0
    for i, cell in enumerate(first):
        low = cell.strip().lower()
        if not low:
            continue
        for tier, words in _HEADER_TIERS:
            if any(w in low for w in words):
                if tier > best_tier:
                    best_tier, header_idx, col_idx = tier, 0, i
                break
    if header is False:
        header_idx = None
    if header is True and header_idx is None:
        header_idx = 0
    if header_idx is None:
        # 无表头: 选第一个"看起来像编码"的列
        best, best_score = 0, -1
        for i in range(len(first)):
            vals = [r[i] for r in rows[:60] if i < len(r) and r[i].strip()]
            if not vals:
                continue
            score = sum(1 for v in vals if len(v) >= 3) - sum(1 for v in vals if len(v) > 60)
            if len(set(vals)) > best_score and score > 0:
                best, best_score = i, len(set(vals))
        col_idx = best
    return col_idx, header_idx


# --------------------------------------------------------------- 输出 ----
def _safe_name(text: str, idx: int) -> str:
    s = re.sub(r'[\\/:*?"<>|\s]+', "_", text.strip())[:60] or f"code{idx}"
    return s


def build(rows: list[list[str]], col_idx: int, header_idx: int | None,
          dedupe: bool) -> tuple[list[str], dict]:
    start = (header_idx + 1) if header_idx is not None else 0
    codes: list[str] = []
    seen = set()
    dup = 0
    for r in rows[start:]:
        if col_idx >= len(r):
            continue
        v = r[col_idx].strip()
        if not v:
            continue
        if v in seen:
            dup += 1
            if dedupe:
                continue
        seen.add(v)
        codes.append(v)
    return codes, {"duplicates": dup}


def _try_png(svg: str, path: str) -> bool:
    """可选 PNG 输出: 依赖 Pillow(非标准库), 缺失时静默跳过。"""
    try:
        from PIL import Image, ImageDraw  # type: ignore
    except ImportError:
        return False
    try:
        text = re.search(r'aria-label="([^"]*)"', svg)
        text = (text.group(1).replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
                if text else "")
        w = int(float(re.search(r'viewBox="0 0 ([\d.]+)', svg).group(1)))
        h = int(float(re.search(r'viewBox="0 0 [\d.]+ ([\d.]+)', svg).group(1)))
        scale = 3
        img = Image.new("L", (w, h), 255)
        draw = ImageDraw.Draw(img)
        for m in re.finditer(r'<rect x="([\d.]+)" y="([\d.]+)" width="([\d.]+)" height="([\d.]+)"',
                             svg):
            x, y, rw, rh = (float(m.group(i)) for i in range(1, 5))
            draw.rectangle([x, y, x + rw - 1, y + rh - 1], fill=0)
        img = img.resize((w * scale, h * scale), Image.NEAREST)
        if text:
            img = img.convert("RGB")
            d2 = ImageDraw.Draw(img)
            try:
                from PIL import ImageFont  # type: ignore
                font = ImageFont.load_default(size=14 * scale)
            except Exception:  # noqa: BLE001
                font = None
            tw = d2.textlength(text, font=font) if font else len(text) * 6 * scale
            d2.text(((img.width - tw) / 2, (h - 18) * scale), text, fill=0, font=font)
        img.save(path, "PNG")
        return True
    except Exception:  # noqa: BLE001
        return False


def write_outputs(codes: list[str], out_dir: str, args) -> dict:
    svg_dir = os.path.join(out_dir, "svg")
    os.makedirs(svg_dir, exist_ok=True)
    png_dir = os.path.join(out_dir, "png")
    if args.png:
        os.makedirs(png_dir, exist_ok=True)
    used: dict[str, int] = {}
    items: list[dict] = []
    errors: list[dict] = []
    png_count = 0
    for i, code in enumerate(codes, 1):
        try:
            svg = code128_svg(code, bar_width=args.bar_width, height=args.barcode_height,
                              show_text=not args.no_text, font_size=args.font_size)
        except BarcodeError as e:
            errors.append({"code": code, "error": str(e)})
            continue
        base = _safe_name(code, i)
        used[base] = used.get(base, 0) + 1
        fname = f"{base}.svg" if used[base] == 1 else f"{base}_{used[base]}.svg"
        with open(os.path.join(svg_dir, fname), "w", encoding="utf-8") as f:
            f.write(svg)
        if args.png and _try_png(svg, os.path.join(png_dir, fname[:-4] + ".png")):
            png_count += 1
        items.append({"code": code, "svg": fname, "svg_inline": svg})

    # 映射 CSV
    csv_path = os.path.join(out_dir, "mapping.csv")
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["序号", "编码", "SVG文件", "条码类型"])
        for i, it in enumerate(items, 1):
            w.writerow([i, it["code"], f"svg/{it['svg']}", "Code128"])

    # 打印页 HTML
    html_path = os.path.join(out_dir, "print.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(render_html(items, args))
    return {"svg_dir": svg_dir, "csv": csv_path, "html": html_path,
            "items": items, "errors": errors, "png_count": png_count,
            "png_dir": png_dir if args.png else None}


def render_html(items: list[dict], args) -> str:
    cells = []
    for it in items:
        # 内嵌到 HTML 时由 CSS 控制尺寸: 剥离 width/height, 保留 viewBox 维持比例
        svg = re.sub(r'\s(width|height)="[\d.]+"', "", it["svg_inline"], count=2)
        cells.append(f'<div class="label">{svg}</div>')
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    return f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<title>条形码打印页 - {len(items)} 个</title>
<style>
  * {{ box-sizing: border-box; }}
  body {{ margin: 0; padding: 10mm; background: #f5f6f8; font-family: "Microsoft YaHei", Arial, sans-serif; }}
  .bar {{ margin: 0 0 6mm; padding: 8px 12px; background: #fff; border-radius: 6px;
         display: flex; justify-content: space-between; font-size: 12px; color: #555; }}
  .sheet {{ background: #fff; padding: 6mm; }}
  .grid {{ display: grid; grid-template-columns: repeat({args.cols}, 1fr); gap: 3mm; }}
  .label {{ border: 1px dashed #c8ccd4; padding: 3mm 2mm; text-align: center;
           break-inside: avoid; page-break-inside: avoid; }}
  .label svg {{ width: 100%; height: auto; display: block; }}
  @media print {{
    body {{ background: #fff; padding: 0; }}
    .bar {{ display: none; }}
    .sheet {{ padding: 0; }}
    .label {{ border: none; }}
    @page {{ size: A4 portrait; margin: 8mm; }}
  }}
</style></head>
<body>
<div class="bar"><span>条形码打印页 · 共 {len(items)} 个编码 · 每行 {args.cols} 个</span><span>生成时间 {stamp}</span></div>
<div class="sheet"><div class="grid">
{chr(10).join(cells)}
</div></div>
</body></html>"""


# ---------------------------------------------------------------- CLI ----
def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        description="Excel/CSV 编码批量生成 Code128 条形码",
        formatter_class=argparse.RawTextHelpFormatter)
    ap.add_argument("--input", "-i", required=True, help="Excel(.xlsx/.xls) 或 CSV 文件路径")
    ap.add_argument("--column", "-c", help="编码列: 列名 / 列标(A) / 序号(1); 默认自动识别")
    ap.add_argument("--sheet", "-s", help="工作表名或序号(1起); 默认第一个")
    ap.add_argument("--out", "-o", help="输出目录; 默认 输入目录/barcodes_<时间戳>")
    ap.add_argument("--cols", type=int, default=3, help="打印页每行标签数(默认 3)")
    ap.add_argument("--barcode-height", type=int, default=60, help="条码高度 px(默认 60)")
    ap.add_argument("--bar-width", type=int, default=2, help="单模块宽度 px(默认 2)")
    ap.add_argument("--font-size", type=int, default=14, help="文字字号 px(默认 14)")
    ap.add_argument("--no-text", action="store_true", help="条码下方不显示编码文字")
    ap.add_argument("--png", action="store_true", help="同时输出 PNG 位图(需 Pillow, 缺失则跳过)")
    ap.add_argument("--dedupe", action="store_true", help="相同编码去重")
    ap.add_argument("--header", dest="header", action="store_true", default=None, help="强制首行为表头")
    ap.add_argument("--no-header", dest="header", action="store_false", help="强制首行非表头")
    ap.add_argument("--json", action="store_true", help="以 JSON 输出摘要")
    args = ap.parse_args(argv)

    t0 = time.time()
    try:
        rows = read_table(args.input, args.sheet)
        col_idx, header_idx = detect_column(rows, args.column, args.header)
        if col_idx < 0:
            raise BarcodeError("列下标非法")
        codes, meta = build(rows, col_idx, header_idx, args.dedupe)
        if not codes:
            raise BarcodeError(f"第 {col_idx + 1} 列未读到任何编码")
        out_dir = args.out or os.path.join(
            os.path.dirname(os.path.abspath(args.input)) or ".",
            "barcodes_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
        os.makedirs(out_dir, exist_ok=True)
        res = write_outputs(codes, out_dir, args)
    except BarcodeError as e:
        print(f"[错误] {e}", file=sys.stderr)
        if args.json:
            print(json.dumps({"ok": False, "error": str(e)}, ensure_ascii=False))
        return 2
    except Exception as e:  # noqa: BLE001
        print(f"[错误] {type(e).__name__}: {e}", file=sys.stderr)
        if args.json:
            print(json.dumps({"ok": False, "error": f"{type(e).__name__}: {e}"}, ensure_ascii=False))
        return 3

    summary = {
        "ok": True,
        "input": os.path.abspath(args.input),
        "column_index": col_idx + 1,
        "column_name": (rows[header_idx][col_idx] if header_idx is not None and col_idx < len(rows[header_idx])
                        else f"第{col_idx + 1}列"),
        "total": len(res["items"]),
        "duplicates_skipped": meta["duplicates"] if args.dedupe else 0,
        "failed": len(res["errors"]),
        "errors": res["errors"][:20],
        "output_dir": os.path.abspath(out_dir),
        "svg_dir": os.path.abspath(res["svg_dir"]),
        "png_dir": res["png_dir"] and os.path.abspath(res["png_dir"]),
        "png_count": res["png_count"],
        "mapping_csv": os.path.abspath(res["csv"]),
        "print_html": os.path.abspath(res["html"]),
        "elapsed_sec": round(time.time() - t0, 2),
    }
    if args.json:
        print(json.dumps(summary, ensure_ascii=False, indent=2))
    else:
        print(f"[完成] 共 {summary['total']} 个条形码 (Code128)")
        print(f"       编码列: 第 {summary['column_index']} 列 {summary['column_name']}")
        if summary["failed"]:
            print(f"       失败 {summary['failed']} 个: {res['errors'][:3]}")
        print(f"       输出目录: {summary['output_dir']}")
        print(f"       SVG 目录: {summary['svg_dir']}")
        if summary["png_count"]:
            print(f"       PNG 目录: {summary['png_dir']} ({summary['png_count']} 张)")
        print(f"       映射表:   {summary['mapping_csv']}")
        print(f"       打印页:   {summary['print_html']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())