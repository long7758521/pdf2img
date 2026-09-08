"""业务异常定义，与进程退出码一一对应。

退出码规范（供调用方如 Word VBA 判定）：
    0  成功
    1  输入参数错误
    2  PDF 文件不存在 / 无法访问
    3  PDF 损坏、加密或无法解析
    4  某页渲染失败
    5  无法创建输出目录 / 磁盘空间不足
    6  写入图片文件失败
    7  未预期的运行时错误（防御性兜底）
"""


class Pdf2imgError(Exception):
    """所有业务异常的基类。"""

    exit_code = 1

    def __init__(self, message: str):
        super().__init__(message)
        self.message = message


class ParamError(Pdf2imgError):
    """输入参数错误。"""

    exit_code = 1


class PdfNotFoundError(Pdf2imgError):
    """指定的 PDF 文件不存在或无法访问。"""

    exit_code = 2


class PdfParseError(Pdf2imgError):
    """PDF 文件损坏、加密或无法解析。"""

    exit_code = 3


class RenderError(Pdf2imgError):
    """转换过程中渲染某一页失败。"""

    exit_code = 4


class OutputDirError(Pdf2imgError):
    """无法创建输出目录或磁盘空间不足。"""

    exit_code = 5


class SaveError(Pdf2imgError):
    """写入图片文件失败（权限不足、路径非法）。"""

    exit_code = 6