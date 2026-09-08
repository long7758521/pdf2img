"""文件操作层：输出目录创建、图片写入、路径跟踪与清理。"""

import logging
import os

from .errors import OutputDirError, SaveError

logger = logging.getLogger(__name__)


class ImageSaver:
    """负责将渲染图像写入磁盘，并按规则命名、跟踪已生成路径。"""

    def __init__(self, output_dir: str, base_name: str, extension: str):
        """
        :param output_dir: 图片输出目录
        :param base_name: 输出文件名主干（取自 PDF 文件名，不含扩展名）
        :param extension: 图片扩展名，如 "png" / "jpg"
        """
        self.output_dir = output_dir
        self.base_name = base_name
        self.extension = extension
        self.saved_paths: list[str] = []
        self._ensure_output_dir()

    def _ensure_output_dir(self):
        """创建输出目录；失败（权限/磁盘满/路径非法）抛 OutputDirError。"""
        try:
            if not os.path.exists(self.output_dir):
                os.makedirs(self.output_dir)
            elif not os.path.isdir(self.output_dir):
                raise OutputDirError(f"输出路径存在但不是目录: {self.output_dir}")
        except OutputDirError:
            raise
        except Exception as exc:
            raise OutputDirError(f"无法创建输出目录（检查磁盘空间与写入权限）: {self.output_dir}") from exc

    def build_path(self, page_number: int) -> str:
        """按命名规则生成文件名：{主干}_{页码}.{扩展名}，页码从 1 开始。"""
        return os.path.join(self.output_dir, f"{self.base_name}_{page_number}.{self.extension}")

    def save(self, page_number: int, data: bytes) -> str:
        """将图像字节写入磁盘，成功后记录路径。失败抛 SaveError。"""
        path = self.build_path(page_number)
        try:
            with open(path, "wb") as f:
                f.write(data)
        except Exception as exc:
            raise SaveError(f"写入图片失败（检查输出目录权限）: {path}") from exc
        self.saved_paths.append(path)
        return path

    def cleanup(self):
        """删除本次生成的所有图片文件。

        删除失败仅记录 warning，不向上抛出（辅助清理功能，不影响主流程退出码）。
        """
        failed = 0
        for path in self.saved_paths:
            try:
                os.remove(path)
                logger.info("已清理临时图片: %s", path)
            except OSError as exc:
                failed += 1
                logger.warning("删除临时图片失败: %s (%s)", path, exc)
        if failed:
            logger.warning("%d 个临时图片清理失败", failed)