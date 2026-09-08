"""CLI 交互层：参数解析、日志初始化、功能路由与退出码映射。"""

import argparse
import logging
import os
import sys

import pymupdf  # 用于获取 PDF 页数（get-page-count 时也需解析）

from . import __version__
from .errors import ParamError, Pdf2imgError, PdfNotFoundError
from .processor import PdfProcessor
from .saver import ImageSaver

# 支持的输出格式（jpeg 归一为 jpg）
FORMAT_ALIASES = {"jpeg": "jpg"}
FORMATS = ["png", "jpg"]


class _ArgumentParser(argparse.ArgumentParser):
    """重写 error()：参数错误时抛 ParamError（退出码 1），而非默认的退出码 2。"""

    def error(self, message):
        raise ParamError(message)


def build_parser() -> argparse.ArgumentParser:
    parser = _ArgumentParser(
        prog="pdf2img",
        description="PDF 转图片 CLI 工具：将 PDF 每一页渲染为 PNG/JPG 图片，供脚本（如 Word VBA）调用。",
    )
    parser.add_argument("pdf_file", help="待转换的 PDF 文件完整路径。")
    parser.add_argument("-o", "--output", default="./output", help="图片输出目录，自动创建（默认 ./output）。")
    parser.add_argument("-d", "--dpi", type=int, default=150, help="图片分辨率 DPI，推荐 150~300（默认 150）。")
    parser.add_argument(
        "-f",
        "--format",
        default="png",
        choices=FORMATS + ["jpeg"],
        help="图片格式，可选 png / jpg / jpeg（默认 png）。",
    )
    parser.add_argument(
        "--get-page-count",
        action="store_true",
        help="仅打印 PDF 总页数到 stdout 后退出，不执行转换。",
    )
    parser.add_argument(
        "--clean",
        action="store_true",
        help="转换完成后自动删除本次生成的图片文件。",
    )
    parser.add_argument("--quiet", action="store_true", help="静默模式，屏蔽 stderr 中的非关键提示信息。")
    parser.add_argument("--log-file", help="将运行日志写入指定文件路径。")
    parser.add_argument("--version", action="version", version=f"pdf2img {__version__}")
    return parser


def setup_logging(quiet: bool, log_file: str | None):
    """初始化日志：始终输出到 stderr；可选追加到日志文件。"""
    level = logging.WARNING if quiet else logging.INFO
    root = logging.getLogger()
    root.setLevel(level)
    fmt = logging.Formatter("%(levelname)s - %(message)s")

    stream = logging.StreamHandler(sys.stderr)
    stream.setLevel(level)
    stream.setFormatter(fmt)
    root.addHandler(stream)

    if log_file:
        try:
            fh = logging.FileHandler(log_file, encoding="utf-8")
            fh.setLevel(logging.DEBUG)
            fh.setFormatter(fmt)
            root.addHandler(fh)
        except OSError as exc:
            raise ParamError(f"无法打开日志文件: {log_file} ({exc})") from exc


def _get_page_count(args) -> int:
    """返回 PDF 总页数；文件不存在时抛出针对性异常。"""
    if not os.path.exists(args.pdf_file):
        raise PdfNotFoundError(f"PDF 文件不存在: {args.pdf_file}")
    if not os.path.isfile(args.pdf_file):
        raise PdfNotFoundError(f"指定路径不是文件: {args.pdf_file}")
    try:
        with pymupdf.open(args.pdf_file) as doc:
            return doc.page_count
    except Exception as exc:
        raise Pdf2imgError(f"无法打开 PDF（文件可能损坏、加密或格式错误）: {args.pdf_file}") from exc


def _convert(args) -> tuple[int, ImageSaver]:
    """执行标准转换流程，返回 (生成的图片数量, saver)。"""
    extension = FORMAT_ALIASES.get(args.format, args.format)
    base_name = os.path.splitext(os.path.basename(args.pdf_file))[0]

    processor = PdfProcessor(args.pdf_file)
    saver = ImageSaver(args.output, base_name, extension)
    try:
        processor.open()
        total = processor.page_count
        logging.info("PDF 共 %d 页，开始转换 (DPI=%d, 格式=%s)", total, args.dpi, extension)
        for i in range(total):
            data = processor.render_page(i, args.dpi)
            path = saver.save(i + 1, data)
            logging.info("已生成 %s", path)
    finally:
        processor.close()
    return len(saver.saved_paths), saver


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    setup_logging(args.quiet, args.log_file)
    logger = logging.getLogger("pdf2img")

    try:
        if args.dpi <= 0:
            raise ParamError("DPI 必须为正整数。")
        if args.quiet and args.log_file is None:
            logger.info("进入静默模式。")

        if args.get_page_count:
            count = _get_page_count(args)
            print(count)
            return 0

        count, saver = _convert(args)
        logger.info("转换完成，共生成 %d 张图片。", count)
        if args.clean:
            logger.info("--clean 已启用，开始清理临时图片。")
            saver.cleanup()
        return 0
    except Pdf2imgError as exc:
        logger.error("%s", exc.message)
        return exc.exit_code
    except Exception:  # 全局兜底，保证任意未预期异常也能以非 0 退出
        logger.exception("未预期的错误：")
        return 7


def run():
    """包级入口入口函数，供 __main__ 调用。"""
    sys.exit(main())


if __name__ == "__main__":
    sys.exit(main())