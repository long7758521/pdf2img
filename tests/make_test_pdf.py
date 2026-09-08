"""生成用于测试的多页 PDF。

用法: python tests/make_test_pdf.py [输出路径]
默认输出到项目根目录 tests/test_输入.pdf（3 页）。
"""

import os
import sys

import pymupdf


def make_test_pdf(output_path: str):
    doc = pymupdf.open()
    for i in range(1, 4):
        page = doc.new_page()
        page.insert_text((72, 120), f"PDF 测试文档 - 第 {i} 页", fontsize=24)
        page.draw_rect(pymupdf.Rect(72, 180, 540, 220), color=(0.1, 0.3, 0.8), width=2)
        page.insert_text((72, 260), f"Page number: {i}", fontsize=14)
    doc.save(output_path)
    doc.close()
    print(f"已生成测试 PDF: {output_path}（共 3 页）")


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "test_输入.pdf"
    )
    make_test_pdf(out)