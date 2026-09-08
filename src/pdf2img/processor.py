"""业务逻辑层：PDF 的打开、校验与逐页渲染。"""

import os

import pymupdf  # PyMuPDF 1.24+ 推荐导入名

from .errors import PdfNotFoundError, PdfParseError, RenderError

# 基准分辨率：PDF 内部坐标以 72 DPI 为单位
BASE_DPI = 72


class PdfProcessor:
    """封装 PDF 打开、页数读取与逐页渲染逻辑。"""

    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self._doc = None

    @property
    def page_count(self) -> int:
        """返回 PDF 总页数（1-based 索引的文档页数）。"""
        return self._doc.page_count

    def open(self):
        """打开并校验 PDF 文件。

        文件不存在抛 PdfNotFoundError；无法解析（损坏/加密/格式错误）抛 PdfParseError。
        """
        if not os.path.exists(self.pdf_path):
            raise PdfNotFoundError(f"PDF 文件不存在: {self.pdf_path}")
        if not os.path.isfile(self.pdf_path):
            raise PdfNotFoundError(f"指定路径不是文件: {self.pdf_path}")
        try:
            self._doc = pymupdf.open(self.pdf_path)
        except Exception as exc:
            raise PdfParseError(f"无法打开 PDF（文件可能损坏、加密或格式错误）: {self.pdf_path}") from exc

    def close(self):
        """释放文档资源。"""
        if self._doc is not None:
            self._doc.close()
            self._doc = None

    def render_page(self, index: int, dpi: int) -> bytes:
        """渲染指定页（index 从 0 开始）为 PNG 像素字节流。

        统一 alpha=False，保证 PNG/JPG 输出均为白底，避免 JPG 出现黑底。
        """
        if self._doc is None:
            raise PdfParseError("PDF 尚未打开")
        try:
            page = self._doc.load_page(index)
            zoom = dpi / BASE_DPI
            pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
            return pix.tobytes("png")
        except PdfParseError:
            raise
        except Exception as exc:
            raise RenderError(f"渲染第 {index + 1} 页失败: {exc}") from exc