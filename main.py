"""PyInstaller 打包入口。

依赖已通过 `python -m pip install -e .` 安装到 site-packages，故此处直接顶层
import，PyInstaller 可自动将 pdf2img 包及其依赖（PyMuPDF）收集进单文件 exe。
"""

import sys

from pdf2img.cli import main

if __name__ == "__main__":
    sys.exit(main())