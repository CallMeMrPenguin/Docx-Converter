"""
services/image_service.py - High-performance Image Format & Thumbnail Management Service.
Handles image conversions (WEBP, HEIC, AVIF, BMP, GIF) to Word-compatible formats,
maintains an on-disk cache with automated TTL cleanup, and generates memory-efficient thumbnails.
"""

import os
import time
import tempfile
from typing import Optional, Tuple, List
from PIL import Image

try:
    import pillow_heif
    pillow_heif.register_heif_opener()
except ImportError:
    pass

SUPPORTED_IMAGE_EXTENSIONS = (
    '.png', '.jpg', '.jpeg', '.bmp', '.gif', '.webp', '.tif', '.tiff', '.heic', '.heif', '.avif'
)

WORD_DIRECT_EXTENSIONS = ('.png', '.jpg', '.jpeg')


class ImageService:
    """Manages image conversions, caching, and thumbnail generation."""

    _cache_dir: Optional[str] = None

    @classmethod
    def get_cache_dir(cls) -> str:
        if cls._cache_dir is None or not os.path.exists(cls._cache_dir):
            cls._cache_dir = os.path.join(tempfile.gettempdir(), "DocxConverter_ImgCache")
            os.makedirs(cls._cache_dir, exist_ok=True)
            cls.cleanup_expired_cache(max_age_days=3)
        return cls._cache_dir

    @classmethod
    def cleanup_expired_cache(cls, max_age_days: int = 3):
        """Removes converted cache files older than max_age_days to prevent disk bloat."""
        if not cls._cache_dir or not os.path.exists(cls._cache_dir):
            return
        now = time.time()
        max_age_sec = max_age_days * 86400
        try:
            for entry in os.scandir(cls._cache_dir):
                if entry.is_file():
                    try:
                        if now - entry.stat().st_mtime > max_age_sec:
                            os.remove(entry.path)
                    except Exception:
                        pass
        except Exception:
            pass

    @classmethod
    def ensure_word_compatible_image(cls, filepath: str) -> str:
        """
        Ensures the image is in a format Word COM can reliably insert (.png or .jpg).
        Converts .webp, .heic, .avif, .bmp, etc. into cached .png files.
        """
        if not filepath or not os.path.exists(filepath):
            return filepath

        ext = os.path.splitext(filepath)[1].lower()
        if ext in WORD_DIRECT_EXTENSIONS:
            return filepath

        cache_dir = cls.get_cache_dir()
        file_stat = os.stat(filepath)
        cache_filename = f"conv_{abs(hash(filepath))}_{file_stat.st_size}_{int(file_stat.st_mtime)}.png"
        cached_path = os.path.join(cache_dir, cache_filename)

        if os.path.exists(cached_path):
            return cached_path

        try:
            with Image.open(filepath) as img:
                if img.mode not in ("RGB", "RGBA"):
                    img = img.convert("RGBA" if "A" in img.mode else "RGB")
                img.save(cached_path, format="PNG")
                return cached_path
        except Exception as e:
            print(f"[ImageService] Warning: Failed to convert {filepath}: {e}")
            return filepath

    @classmethod
    def create_thumbnail(cls, filepath: str, max_size: Tuple[int, int] = (160, 160)) -> Optional[Image.Image]:
        """
        Loads and scales down an image for GUI display with minimal memory footprint.
        Returns a PIL Image object.
        """
        if not filepath or not os.path.exists(filepath):
            return None
        try:
            with Image.open(filepath) as img:
                img_copy = img.copy()
                img_copy.thumbnail(max_size, Image.Resampling.LANCZOS)
                return img_copy
        except Exception as e:
            print(f"[ImageService] Failed to generate thumbnail for {filepath}: {e}")
            return None


# Helper wrappers for backward compatibility with renderer_utils
def ensure_word_compatible_image(filepath: str) -> str:
    return ImageService.ensure_word_compatible_image(filepath)
