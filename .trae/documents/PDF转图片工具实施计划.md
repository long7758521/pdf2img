# PDF 转图片 CLI 工具 — 需求分析与实施计划

## 一、需求分析总结

从 `pdf转图片需求` 文档提炼出的核心诉求：

> 开发一个**零依赖、单文件 exe** 的 PDF 转图片命令行工具，供 Word VBA 宏等脚本通过命令行调用，实现"标书制作"场景下的无人值守批量转换（PDF 每页 → 图片 → 插入 Word）。

**核心功能清单：**
1. 多页 PDF 逐页渲染为 PNG/JPG 图片
2. 可配置输出目录（-o）、分辨率 DPI（-d，默认 150）、图片格式（-f，默认 png）
3. `--get-page-count`：只向 stdout 打印总页数，不转换，供 VBA 循环控制
4. `--clean`：转换完成后自动删除生成的临时图片
5. 通过 stderr 日志 + 进程退出码（0~6）向上层精准反馈错误
6. `--quiet` 静默模式、`--log-file` 日志文件、`--version` 版本号

**退出码规范（VBA 依赖此判断成败）：**

| 码 | 含义 |
|---|---|
| 0 | 成功 |
| 1 | 输入参数错误 |
| 2 | PDF 文件不存在/无法访问 |
| 3 | PDF 损坏、加密或无法解析 |
| 4 | 某页渲染失败 |
| 5 | 无法创建输出目录/磁盘空间不足 |
| 6 | 图片写入失败 |

**非功能性目标：** 150DPI 单页 <1s；纯本地离线；模块化可扩展（未来可加 TIFF/水印等）。

## 二、现状分析

- 项目目录 `d:\小玩具\PDF转图片工具` 为空目录，仅含需求文档，属全新项目
- 本机已装 Python 3.13.5 + pip 25.2；**PyMuPDF、PyInstaller 均未安装**（实施时需 pip 安装）
- 已与用户确认两个决策：
  - 目标系统**仅 Win10/11**（Python 3.13 已不支持 Win7，放弃 Win7 兼容）
  - 交付物中**附带 Word VBA 示例宏**（.bas 文件）

## 三、技术方案

### 3.1 技术选型（沿用需求文档）
Python 3.13 + PyMuPDF (fitz) + argparse + logging + PyInstaller（--onefile --console 打包为 pdf2img.exe）。

### 3.2 分层架构与文件结构

```
d:\小玩具\PDF转图片工具\
├── src\pdf2img\
│   ├── __init__.py      # __version__ = "1.0.0"
│   ├── errors.py        # 自定义异常类 → 退出码映射
│   ├── processor.py     # 业务逻辑层：打开/校验PDF、页数、逐页渲染
│   ├── saver.py         # 文件操作层：建目录、写图片、命名、路径跟踪、清理
│   ├── cli.py           # CLI交互层：argparse 解析、日志初始化、路由、异常→退出码
│   └── __main__.py      # 支持 python -m pdf2img
├── pdf2img.py           # PyInstaller 打包入口（根目录薄封装）
├── vba\
│   └── PDF转图片插入Word示例.bas   # Word VBA 示例宏
├── build.ps1            # 一键构建脚本（装依赖 + 打包 exe）
├── requirements.txt     # PyMuPDF、pyinstaller
├── README.md            # 使用说明 + VBA 集成说明 + 退出码表（交付文档）
└── tests\
    └── make_test_pdf.py # 生成多页测试 PDF，供反复验证用
```

### 3.3 各模块设计要点

**errors.py** —— 6 个异常与退出码一一对应：`ParamError(1)`、`PdfNotFoundError(2)`、`PdfParseError(3)`、`RenderError(4)`、`OutputDirError(5)`、`SaveError(6)`。

**processor.py**（业务逻辑层 `PdfProcessor`）
- 打开前先校验文件存在，不存在抛 `PdfNotFoundError`
- `fitz.open` 失败（损坏/加密/非法）抛 `PdfParseError`
- `page_count` 属性返回总页数；`render_page(i, dpi, fmt)` 按 `zoom = dpi/72` 缩放矩阵渲染
- 统一 `alpha=False`（PNG/JPG 均白底，避免 JPG 黑底问题）
- 某页渲染失败即抛 `RenderError`（fail-fast，日志注明页码）

**saver.py**（文件操作层 `ImageSaver`）
- 构造时创建输出目录（`mkdir` 失败抛 `OutputDirError`，捕获磁盘满等 OSError）
- 命名规则：`{PDF文件名主干}_{页码}.{扩展名}`，页码从 1 开始
- 保存失败（权限、磁盘满）抛 `SaveError`
- 维护 `saved_paths` 列表；`cleanup()` 尽力删除，删除失败仅记 warning 日志不影响退出码（按需求 6.3 节）

**cli.py**（CLI 交互层）
- 参数与需求 5.1 节完全一致：位置参数 `pdf_file`；`-o/--output`(默认 ./output)；`-d/--dpi`(默认150，校验为正整数)；`-f/--format`(choices: png/jpg/jpeg，jpeg 归一为 jpg)；`--get-page-count`；`--clean`；`--quiet`；`--log-file`；`--version`
- 自定义 ArgumentParser 子类重写 `error()`：参数错误走 `ParamError` → 退出码 **1**（默认值是 2，不符合需求）
- 日志：一律输出到 **stderr**（stdout 留给页数纯数字输出）；`--quiet` 屏蔽 WARNING 以下；`--log-file` 追加 FileHandler
- 路由逻辑：
  - `--get-page-count`：打开 PDF → stdout 仅打印页数 → 退出 0，不生成任何文件
  - 标准转换：打开 → 逐页渲染保存 → 成功退出 0；`--clean` 时转换后执行 cleanup
- `main()` 顶层 try/except 捕获全部业务异常 → 对应退出码；全局兜底 `Exception` → 退出码 7 并打印堆栈日志（防御性兜底，不破坏 0~6 规范）

### 3.4 VBA 示例宏要点（vba\PDF转图片插入Word示例.bas）
完整演示需求第 7 节流程，顶部用常量配置工具路径/PDF路径/输出目录/DPI：
1. `WScript.Shell.Exec` 运行 `pdf2img.exe "...pdf" --get-page-count`，读 `StdOut` 得页数 N
2. `Shell.Run` 带参数完整转换，`bWaitOnReturn:=True` 获取退出码，`Select Case` 按 0~6 弹窗提示错误
3. 循环 1..N，`InlineShapes.AddPicture` 按 `文件名_i.ext` 逐张插入光标处
4. 插入完成后用 `Kill` 删除临时图片（按需求推荐，比 --clean 更可控）

### 3.5 构建方案
- `requirements.txt`：`PyMuPDF>=1.24`、`pyinstaller>=6.0`
- `build.ps1`：装依赖 → `pyinstaller --onefile --console --clean --noconfirm --name pdf2img pdf2img.py`
- 产物：`dist\pdf2img.exe`（约 15~20MB，与需求预估一致）

## 四、实施步骤

1. 创建 `src\pdf2img` 包：errors.py → processor.py → saver.py → cli.py → `__init__.py`/`__main__.py`
2. 创建根入口 `pdf2img.py`（PyInstaller 打包入口）
3. 创建 `requirements.txt`、`build.ps1`
4. 创建 `tests\make_test_pdf.py`，生成 3 页测试 PDF
5. 编写 `vba\PDF转图片插入Word示例.bas`
6. pip 安装依赖，按"验证方案"完成源码级自测
7. 执行 build.ps1 打包 exe，对 exe 复跑核心验证
8. 编写 `README.md`（用法、退出码表、VBA 集成说明、部署建议）

## 五、关键假设与决策

| 项 | 决策 |
|---|---|
| 目标系统 | 仅 Win10/11，Python 3.13 构建（用户确认，放弃 Win7） |
| VBA 宏 | 附带 .bas 示例（用户确认） |
| 页面失败策略 | fail-fast：首页渲染失败立即中止，退出码 4 |
| 图片背景 | 统一白底（alpha=False），PNG/JPG 一致 |
| 文件命名 | `{PDF文件名}_{页码}.{ext}`，页码起始 1 |
| stdout 纯净性 | 日志全走 stderr，stdout 仅输出页数 |
| 参数错误退出码 | 重写 argparse error 使参数错误返回 1（非默认 2） |
| --clean 失败 | 仅 warning，不影响退出码（需求 6.3 明确规定） |
| 版本号 | 1.0.0 |

## 六、验证方案

源码级（`python pdf2img.py ...`）：
1. 生成测试 PDF → 转换：确认输出 3 张图片名 `测试_1.png`~`测试_3.png`，退出码 0
2. `--get-page-count`：stdout 仅输出 `3`，目录无新文件
3. `-f jpg -d 300`：输出 jpg 格式且分辨率符合 300DPI
4. `--clean`：转换后图片被删除，退出码仍 0
5. `--quiet`、`--log-file`：日志行为符合预期
6. 异常路径：不存在的 PDF→码2；伪造损坏文件→码3；非法格式参数→码1；只读输出目录→码5/6
7. `python -m pdf2img --version` 正常

exe 级：`dist\pdf2img.exe` 复跑第 1、2、6 项；`--version` 输出 1.0.0。

VBA 级（需用户在 Word 中手动验证）：导入 .bas → 运行宏 → 图片插入、退出码提示、Kill 清理均正常。