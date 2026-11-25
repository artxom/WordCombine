# WordCombine — Word 文档高保真拼接与目录刷新工具

WordCombine 是一个跨平台（Windows + macOS）桌面工具，可将选中的多个 `.docx` 文档按文件名自然数字顺序首尾相连拼接成一个完整文件，并在拼接后后台静默刷新文档目录（TOC），同时提供严格的合并前后内容一致性校验。

## 核心能力

- **自然文件名排序**：按数值大小排序文件名中的数字段（例如 `3.1 3.2` < `3.3` < ... < `3.9` < `3.10` < `4.`），自动过滤 `~$*.docx` 临时锁文件。
- **高保真 OOXML 拼接**：直接操作 OOXML 结构，完整保留原文段落、表格、样式、独立章节编号、内嵌图片、Excel OLE 对象、批注、脚注/尾注、超链接与节页眉/页脚隔离，不丢失或篡改原内容。
- **后台静默刷新目录**：
  - **Windows**：通过 COM 后台调用本机 **Microsoft Word** 或 **WPS Office** 静默更新目录与页码（未安装时自动优雅跳过并保留已合并文档）。
  - **macOS**：通过 AppleScript 在沙盒可访问目录静默调用 **Microsoft Word** 刷新目录，无授权弹窗。
- **实时可视化界面**：显示待拼接文件列表、每个文档的当前处理状态（待处理 / 合并中 / 已完成）、阶段进度条、百分比、已用时与详细日志。
- **严格一致性校验**：自动校验 ZIP/XML 完整性、ID/样式唯一性、正文段落序列、表格单元格、图片/OLE SHA-256 哈希与页眉节隔离。

## 下载与使用独立可执行版本（无需 Python 环境）

前往仓库的 **Releases** 页面下载对应平台安装包：

- **Windows（免解压秒开推荐）**：下载 `WordCombine-v1.0.1-Windows-x64-FastStart.zip`，解压后双击文件夹内的 `WordCombine.exe`，免去单文件解包等待，几乎秒开。
- **Windows（单文件便携版）**：下载 `WordCombine-v1.0.1-Windows-x64.exe`，双击即可运行；内建原生启动画面与进度提示，避免冷启动及后台引擎初始化期间让用户误以为卡死。
- **macOS**：下载 `WordCombine-v1.0.1-macOS-arm64.zip`，解压后打开 `WordCombine.app`。

> **提示**：如 Windows SmartScreen 或 macOS 首次打开未签名应用提示安全警告，选择“仍要运行 / 仍要打开”即可。

## 启动与交互性能优化

- **原生启动反馈 & 秒开打包模式**：Windows 单文件版禁用 UPX 避免杀毒引擎长时扫描，并在 Python/Tk 加载前由引导器优先展示启动卡片；同时提供免解压 `onedir` (`FastStart.zip`) 打包。
- **界面秒开与延迟加载**：将 `lxml` 解析引擎、目录检测和校验模块延迟到窗口显示后于后台线程预热，主窗口无白屏闪烁。
- **异步目录扫描**：添加多份大型 `.docx` 文件时，表格立即呈现（目录列显示 `检测中...`），后台采用字节级快速预过滤配合异步线程刷新目录状态，杜绝主线程阻塞导致的 `(未响应)`。

## 本地打包选项

```bash
# Windows 单文件版（默认，含启动画面）
pyinstaller packaging/word_combine.spec --noconfirm --clean

# Windows 免解压秒开版 (onedir)
WORDCOMBINE_WIN_MODE=onedir pyinstaller packaging/word_combine.spec --noconfirm --clean
```

## 本地开发与命令行运行

```bash
pip install -r requirements-dev.txt

# 启动图形界面
python word_combine.py

# 命令行批量拼接并校验
python word_combine.py --dir ./docs -o ./output/合并文档.docx --verify

# 在文档开头自动插入目录并静默刷新
python word_combine.py --dir ./docs -o ./output/合并文档_含目录.docx --insert-toc --verify
```

运行单元测试：

```bash
python -m unittest discover -s tests -v
```

## 通过 GitHub Actions 打包并发布

本仓库已配置 `.github/workflows/release.yml`，推送形如 `v1.0.1` 的版本标签即可自动在云端构建 Windows 单文件 `.exe`、Windows 免解压 `FastStart.zip` 与 macOS `.app` 压缩包并创建 GitHub Release：

```bash
git remote add origin https://github.com/<你的用户名>/<你的仓库名>.git
git push -u origin main
git tag v1.0.1
git push origin v1.0.1
```
