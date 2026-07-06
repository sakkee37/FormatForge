# FormatForge — 全格式文档转换器

一键在 Word ↔ PDF ↔ Markdown 之间互转，支持数学公式渲染。

## 功能

| 转换方向 | 状态 |
|----------|------|
| Word → PDF | ✅ |
| PDF → Word | ✅（需 Word 2013+） |
| Markdown → Word | ✅ |
| Word → Markdown | ✅ |
| Markdown → PDF | ✅（两步自动完成） |
| 数学公式（LaTeX） | ✅ |

## 快速开始

### 方式一：直接运行（不需要 Python）

下载 `dist/FormatForge.exe`，双击即可。

> 第一次运行 Windows Defender 可能弹窗，点「更多信息」→「仍要运行」。

### 方式二：从源码运行

```bash
pip install -r requirements.txt
python main.py
```

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

## 技术栈

| 组件 | 技术 |
|------|------|
| GUI | tkinter（Python 内置） |
| Word 引擎 | pywin32 → Word COM |
| Markdown 引擎 | mistune + python-docx |
| 数学公式 | latex2mathml |
| 打包 | PyInstaller（单文件 exe） |

## 项目结构

```
format-converter/
├── main.py              # GUI 主界面
├── utils.py             # 格式检测、路径处理
├── engines/
│   ├── word_engine.py   # docx ↔ PDF
│   └── markdown_engine.py # MD ↔ docx
├── requirements.txt
├── DEVLOG.md            # 开发日志
└── dist/
    └── FormatForge.exe  # 打包好的程序
```

## 许可

MIT
