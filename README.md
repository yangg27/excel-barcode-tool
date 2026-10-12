# Excel 条形码批量生成工具

> Excel / CSV 一列编码 → 一键批量生成 **Code128 条形码**，直接打印贴标。
> 图形界面 + 命令行双模式，**零依赖可跑**，支持免安装 exe。

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Code128](https://img.shields.io/badge/Barcode-Code128-orange)

---

## 特性

- **双击即用** —— 图形界面三步操作，无需命令行
- **零依赖** —— 纯标准库实现 XLSX 解析与 Code128 编解码，不装任何第三方库也能跑
- **标准最优编码** —— 采用 ISO/IEC 15417 动态规划算法，自动切换 A/B/C 字符集并使用数字压缩，生成最短条码
- **四种产出** —— SVG 矢量图 + PNG 位图 + A4 打印页 HTML + 编码映射 CSV
- **智能识别编码列** —— 分级打分自动定位编码列，避免误选"序号"列
- **宽格式兼容** —— `.xlsx` / `.xlsm` / `.xls` / `.csv`，CSV 自动识别 UTF-8 / GBK / GB18030

## 效果预览

`print.html` 生成的 A4 打印页（浏览器 `Ctrl+P` 直接打印贴标）：

```
┌─────────────┬─────────────┬─────────────┐
│ ▮▯▮▯▮▯▮▯▮▯▮ │ ▮▯▮▯▮▯▮▯▮▯▮ │ ▮▯▮▯▮▯▮▯▮▯▮ │
│ SB-ZC-0001  │ SB-ZC-0002  │ WZ2026-0088 │
└─────────────┴─────────────┴─────────────┘
```

## 快速开始

### 方式 1：图形界面（推荐）

```bash
# 双击运行
条形码工具.bat

# 或直接运行界面程序
python src/barcode_gui.py
```

三步操作：**选文件 → 确认编码列 → 点「开始生成条形码」**

### 方式 2：命令行

```bash
# 最简用法（自动识别编码列）
python src/excel_barcode.py -i "台账.xlsx"

# 指定列 + 输出目录 + 同时出 PNG
python src/excel_barcode.py -i "台账.xlsx" --column 资产编码 --out ./输出 --png

# 打印页每行 4 个，条码加大
python src/excel_barcode.py -i "台账.xlsx" --cols 4 --bar-width 3 --barcode-height 80
```

## 产出物

| 文件 | 说明 |
|---|---|
| `svg/` | 每码一张矢量图，打印放大不糊，可贴进 Word / PPT |
| `png/` | 位图版，方便插进 Office 文档 |
| `print.html` | **A4 打印页**，浏览器打开 `Ctrl+P` 直接打印贴标 |
| `mapping.csv` | 序号 ↔ 编码 ↔ 文件 对照表 |

## 命令行参数

| 参数 | 说明 | 默认 |
|---|---|---|
| `-i, --input` | Excel / CSV 路径（**必填**） | — |
| `-c, --column` | 编码列：列名 / 列标 `B` / 序号 `2` | 自动识别 |
| `-s, --sheet` | 工作表名或序号 | 第一个 |
| `-o, --out` | 输出目录 | 输入文件旁 `barcodes_时间戳` |
| `--cols` | 打印页每行标签数 | 3 |
| `--bar-width` | 单模块宽度 px | 2 |
| `--barcode-height` | 条码高度 px | 60 |
| `--font-size` | 编码文字字号 px | 14 |
| `--no-text` | 不显示编码文字 | 显示 |
| `--png` | 额外输出 PNG | 否 |
| `--dedupe` | 重复编码去重 | 不去重 |
| `--header` / `--no-header` | 强制首行是/不是表头 | 自动 |
| `--json` | 输出 JSON 摘要 | 否 |

## 编码列自动识别规则

按优先级分级打分，命中即选（分高者优先，同分取最左列）：

| 优先级 | 关键词 |
|---|---|
| 最高 | `编码`、`条码`、`条形码`、`barcode`、`code128` |
| 次之 | `编号`、`物料`、`资产`、`设备`、`SKU`、`序列号`、`货号` |
| 最低 | `序号`、`code`、`id`、`no`、`号码` |

这样设计是为了避免误选"序号"列。识别不准时用 `--column` 手动指定。

## 条形码规格

- **类型**：Code 128（ISO/IEC 15417）
- **字符集**：动态规划自动在 A/B/C 间切换 + 数字压缩，生成最短编码
- **字符范围**：ASCII 0–127（数字 / 字母 / 常见符号）
- **校验**：Modulo 103 自动校验位
- **静区**：两侧各留 10 模块，符合标准

## 验证情况

| 测试项 | 结果 |
|---|---|
| 真实扫码引擎 ZXing | 16/16 识别成功 |
| 1000 行台账压力测试 | 3.56 秒完成，零失败（约 280 条/秒） |
| 大批量抽样扫码 | 60/60 通过 |
| 图形界面产出扫码 | 6/6 通过 |
| 内置编解码器往返 | 18/18 通过 |

> 附注：业界常用的 `python-barcode` 库在超长奇数位纯数字（如 19 位）上会丢字符，本工具无此问题。

## 环境依赖

| 使用方式 | 依赖 |
|---|---|
| 图形界面 | Python 3 自带 `tkinter`，**零额外依赖** |
| 命令行读 `.xlsx` + 出 SVG | 纯标准库，**零依赖** |
| 读 `.xls` / 输出 PNG | 可选 `openpyxl` / `Pillow` |

**打包免安装 exe**（可选）：

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "条形码工具" \
  --hidden-import openpyxl --hidden-import PIL.ImageDraw --hidden-import PIL.ImageFont \
  src/barcode_gui.py
```

## 项目结构

```
excel-barcode-tool/
├── src/
│   ├── excel_barcode.py    # 核心引擎 + 命令行版（740 行）
│   └── barcode_gui.py      # 图形界面（430 行）
├── examples/
│   ├── 示例-编码清单.xlsx   # 示例输入
│   └── 示例输出/            # 示例产出（SVG/PNG/HTML/CSV）
├── 条形码工具.bat           # Windows 双击入口
├── 使用说明.md              # 详细中文使用文档
├── LICENSE
└── README.md
```

## 注意事项

- **编码含中文会跳过** —— Code128 规范不支持中文，含中文的行会提示失败数（中文场景请改用 QR 码）
- **条码太密扫不出** —— 编码很长时把「条码宽度」从 2 调到 3 或 4
- **打印建议** —— 浏览器打印时取消页眉页脚、选「实际大小」效果最佳

## License

[MIT](LICENSE) © 2026 yangg27