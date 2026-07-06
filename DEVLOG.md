# 开发日志 — FormatForge

## [2026-07-07 00:45] — 修复转换按钮不可见问题

**改动者:** AI (reasonix, Superpowers)

**改动内容:**
- 修改 format-converter/main.py
- 将「开始转换」按钮从独立底部行移到目标格式选择同一行
- 按钮现在与 PDF/Markdown 目标按钮并排显示
- 窗口高度从 620px 降至 530px
- 重新打包 FormatForge.exe

**原因:**
- 选择文件后所有卡片展开，底部「开始转换」按钮被挤出窗口不可见
- 用户两次反馈看不到按钮

**影响范围:**
- 仅 GUI 布局变更，引擎层不受影响

**Git commit:** `(4e0e045)`

---

## [2026-07-07 00:22] — 全格式文档转换器初始版本

**改动者:** AI (reasonix, Superpowers: brainstorming → writing-plans → TDD)

**改动内容:**
- 新建 format-converter/ 项目：
  - main.py — tkinter 美化 GUI
  - utils.py — 格式检测、路径处理
  - engines/word_engine.py — Word COM: docx→PDF, PDF→docx
  - engines/markdown_engine.py — mistune + python-docx: MD?docx, 含 latex2mathml 数学公式
  - requirements.txt — pywin32, mistune, python-docx, latex2mathml
- 新建 tests/test_format_converter.py — 15 个单元测试，全部通过
- 设计文档: docs/superpowers/specs/2026-07-07-format-converter-design.md

**支持方向:**
- docx→PDF ? | PDF→docx ??(需 Word 365) | MD→docx ? | docx→MD ? | MD→PDF ?(两步)

**原因:**
- 用户需求：全格式文档互转工具，最初从 docx→pdf 示例扩展而来
- 设计阶段确认了 Word?PDF?Markdown 互转范围

**影响范围:**
- 新项目，无破坏性变更
- 与旧 docx2pdf/ 目录并存（可后续清理）

**Git commit:** `d458aa4 - docs: add format converter design spec`

---

## [2026-07-07 00:06] — Superpowers 技能包装配

**改动者:** AI (reasonix)

**改动内容:**
- 安装 skills: brainstorming, writing-plans, tdd, superpowers, devlog
- 存储记忆: superpowers-default-workflow
- 安装依赖: mistune, python-docx, latex2mathml, pyinstaller

**原因:**
- 用户要求从 GitHub 安装 obra/superpowers 技能包

**影响范围:**
- 全局开发流程变更：所有新开发任务将遵循 brainstorming → plans → TDD → devlog
