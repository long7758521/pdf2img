# PDF 转图片 CLI 工具（pdf2img）

零依赖、单文件可执行（exe）的 PDF 转图片命令行工具，专为 **Word VBA 宏等脚本无人值守调用**设计：快速获取 PDF 页数、逐页渲染为 PNG/JPG、规范化的退出码与 stderr 日志反馈，支持转换后自动清理临时图片。

## 功能特性

- **多页转换**：将 PDF 每一页渲染为 PNG / JPG 图片
- **可配置输出**：自定义输出目录 `-o`、分辨率 `-d`（DPI）、图片格式 `-f`
- **页数查询** `--get-page-count`：不转换，仅向 stdout 打印总页数，供 VBA 精准循环控制
- **自动清理** `--clean`：转换完成后删除本次生成的图片
- **规范错误反馈**：所有日志走 stderr，通过进程退出码 0~6 反馈错误类型
- **零依赖交付**：PyInstaller 打包为单个 `.exe`，目标机无需安装 Python 或任何运行库
- **纯本地离线**：不访问网络、不采集数据

## 快速使用

```bat
:: 基本转换（默认输出到 ./output，150 DPI，PNG）
pdf2img.exe "标书.pdf"

:: 自定义输出目录、DPI、格式
pdf2img.exe "标书.pdf" -o "D:\temp\imgs" -d 300 -f jpg

:: 只获取页数（打印数字到 stdout）
pdf2img.exe "标书.pdf" --get-page-count

:: 转换后自动清理图片
pdf2img.exe "标书.pdf" -o "D:\temp\imgs" --clean

:: 静默模式 + 日志写入文件
pdf2img.exe "标书.pdf" --quiet --log-file "D:\temp\run.log"

:: 查看版本
pdf2img.exe --version
```

## 命令行参数

| 参数 | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `pdf_file` | 位置参数（必选） | 无 | 待转换的 PDF 完整路径 |
| `-o, --output` | 字符串 | `./output` | 图片输出目录（自动创建） |
| `-d, --dpi` | 整数 | `150` | 输出分辨率，推荐 150~300 |
| `-f, --format` | 枚举 | `png` | 图片格式：`png` / `jpg` / `jpeg` |
| `--get-page-count` | 布尔 | `false` | 仅打印总页数到 stdout 后退出，不转换 |
| `--clean` | 布尔 | `false` | 转换完成后删除本次生成的图片 |
| `--quiet` | 布尔 | `false` | 屏蔽 stderr 中的非关键提示 |
| `--log-file` | 字符串 | 无 | 运行日志写入指定文件 |
| `--version` | 布尔 | 无 | 打印版本号后退出 |

**输出命名规则**：`{PDF文件名}_{页码}.{扩展名}`，页码从 1 开始，例如 `标书_1.png` ~ `标书_10.png`。

## 退出码规范

调用方（如 VBA）依据退出码判断执行结果：

| 退出码 | 含义 | 处理建议 |
|---|---|---|
| 0 | 转换/页数查询成功 | 正常执行下一步 |
| 1 | 输入参数错误 | 检查命令行语法 |
| 2 | PDF 文件不存在或无法访问 | 校验文件路径 |
| 3 | PDF 损坏、加密或无法解析 | 手动打开确认文件完整性 |
| 4 | 某页渲染失败 | 检查是否含特殊字体/复杂矢量图 |
| 5 | 无法创建输出目录或磁盘空间不足 | 清理磁盘、检查写入权限 |
| 6 | 写入图片文件失败 | 检查输出目录访问权限 |
| 7 | 未预期的运行时错误 | 复现并收集日志反馈 |

## 与 Word VBA 集成

工具仅负责转换，Word 操作由 VBA 完成。完整示例宏见 [`vba/PDF转图片插入Word示例.bas`](vba/PDF转图片插入Word示例.bas)，导入 Word 后修改顶部常量即可使用。核心流程：

```vba
' 1. 获取页数
Dim shell As Object, exec As Object, n As Long
Set shell = CreateObject("WScript.Shell")
Set exec = shell.Exec("""C:\Tools\pdf2img.exe"" ""标书.pdf"" --get-page-count --quiet")
n = Val(exec.StdOut.ReadLine)

' 2. 执行转换并等待，读取退出码
exitCode = shell.Run("""C:\Tools\pdf2img.exe"" ""标书.pdf"" --output ""D:\temp\imgs"" --dpi 150 --format png --quiet", 0, True)

' 3. 循环插入图片
For i = 1 To n
    Selection.InlineShapes.AddPicture "D:\temp\imgs\标书_" & i & ".png"
Next

' 4. 插入完成后清理临时图片（比 --clean 更可控，推荐）
' Kill "D:\temp\imgs\标书_*.png"
```

## 自动构建（重新打包 exe）

环境：已安装 Python 3.9+（目标机运行 exe 则无需 Python）。

```powershell
.\build.ps1
```

脚本会：安装依赖 → 以可编辑模式安装本包 → PyInstaller 打包为 `dist\pdf2img.exe`。已针对本机可能安装的 PyTorch 等重型无关库做了 `--exclude-module` 排除以保证构建速度。

## 目录结构与源码

```
.
├── src\pdf2img\          # 源码包
│   ├── cli.py            # CLI 交互层：参数解析、日志、路由、退出码
│   ├── processor.py      # 业务逻辑层：PDF 打开、页数、逐页渲染
│   ├── saver.py          # 文件操作层：目录创建、图片写入、清理
│   └── errors.py         # 异常定义与退出码映射
├── main.py               # 打包入口
├── build.ps1             # 一键构建脚本
├── requirements.txt      # 依赖声明
├── pyproject.toml        # 包定义
├── vba\                  # Word VBA 示例宏
└── tests\                # 测试 PDF 生成脚本
```

源码运行方式：

```bash
# 方式一：可编辑安装后通过包调用
python -m pip install -e .
python main.py "标书.pdf"
```

## 部署建议

- 将 `dist\pdf2img.exe` 放置于固定目录（如 `C:\Tools\`）或加入系统 `PATH` 环境变量
- 运行时需具备目标 PDF 的**读取权限**及输出目录的**写入权限**
- 支持 Windows 10 / 11（64 位），纯本地离线运行

## 说明与限制

- 加密 PDF：本工具不支持输入密码解密，加密 PDF 将返回退出码 3
- 图片统一用**白底**渲染输出（避免 JPG 黑底问题）
- 某页渲染失败即中止转换并返回退出码 4（fail-fast）