# FormatForge — 全格式文档转换器

一键在 **Word ↔ PDF ↔ Markdown** 之间互转，支持 LaTeX 数学公式渲染。

桌面 GUI 工具，基于 tkinter，可打包成单文件 exe，**使用者无需安装 Python**。

> ⚠️ **仅支持 Windows**。Word ↔ PDF 的双向转换通过 `pywin32` 调用 Word COM 接口实现，
> 依赖本机安装的 Microsoft Word（2013 及以上）。Markdown 相关转换无此限制。

---

## 功能

| 转换方向 | 状态 | 依赖 |
|----------|------|------|
| Word → PDF | ✅ | Microsoft Word |
| PDF → Word | ✅ | Microsoft Word 2013+ |
| Markdown → Word | ✅ | 无 |
| Word → Markdown | ✅ | 无 |
| Markdown → PDF | ✅（Word→PDF 两步自动完成） | Microsoft Word |
| 数学公式（LaTeX） | ✅ | 无 |

---

## 快速开始

### 方式一：直接运行（不需要 Python）

到 [**Releases**](../../releases) 页面下载 `FormatForge.exe`，双击即可。

> 首次运行 Windows Defender 可能弹窗，点「更多信息」→「仍要运行」。

### 方式二：从源码运行

```bash
pip install -r requirements.txt
python main.py
```

---

## 界面

```
┌──────────────────────────────────────┐
│   FormatForge                        │
│   ┌──────────────────────────────┐   │
│   │   📂 拖拽文件或点击选择       │   │
│   └──────────────────────────────┘   │
│   输入: report.docx (42KB)           │
│   转换为: [PDF] [Markdown]  [▶ 开始]  │
│   输出到: C:\...\report.pdf          │
└──────────────────────────────────────┘
```

---

## 技术栈

| 组件 | 技术 |
|------|------|
| GUI | tkinter（Python 内置） |
| Word 引擎 | pywin32 → Word COM 自动化 |
| Markdown 引擎 | mistune + python-docx |
| 数学公式 | latex2mathml |
| 打包 | PyInstaller（单文件 exe） |

---

## 项目结构

```
format-converter/
├── main.py                  # GUI 主界面与转换流程编排
├── utils.py                 # 格式检测、路径处理
├── engines/
│   ├── word_engine.py       # docx ↔ PDF（Word COM）
│   └── markdown_engine.py   # MD ↔ docx（mistune + python-docx + latex2mathml）
├── requirements.txt
├── FormatForge.spec         # PyInstaller 打包配置
├── FormatForge.bat          # 一键打包脚本
├── DEVLOG.md                # 开发日志
└── LICENSE
```

---

## 打包

```bash
pip install pyinstaller
pyinstaller FormatForge.spec
```

产物输出到 `dist/FormatForge.exe`。构建产物已在 `.gitignore` 中排除，
发布版本请上传到 [Releases](../../releases) 而非提交进仓库。

---

## 实现要点

**格式探测驱动转换路径**：`utils.py` 依据输入文件的扩展名判定源格式，再结合界面上
选择的目标格式，路由到对应引擎；Markdown → PDF 这类没有直连路径的组合会自动拆成
`Markdown → Word → PDF` 两步执行。

**Word 交给 COM，纯文本自己处理**：凡是需要精确排版保真的转换（docx ↔ PDF）都委托给
本机 Word，避免自行解析 OOXML 带来的格式丢失；而 Markdown 与 docx 之间的结构化转换
用 `mistune` 解析语法树、`python-docx` 构建文档，数学公式经 `latex2mathml` 转为
MathML 后写入，全程不依赖 Office。

---

## 已知限制

- 仅 Windows 可用（Word COM 依赖）。
- PDF → Word 的还原质量取决于 Word 自身的解析能力，复杂排版可能出现偏差。
- 单文件转换，暂无批量队列功能。

---

## License

[MIT](LICENSE)
