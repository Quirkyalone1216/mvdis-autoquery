#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
englishAlphanumericOcrApi.py

使用 ddddocr 辨識英數 CAPTCHA。

唯一對外公開函式：
    ocrImage(image: PIL.Image.Image) -> str

修正內容：
1. 不傳入 ddddocr 1.6.1 不支援的 max_image_bytes/max_image_side。
2. 不使用 set_ranges(6)，改用明確的英數字元字串。
3. 針對很寬的 td 截圖，先裁掉背景；必要時再做水平滑動裁切。
4. 一般模型無法得到四碼時，才使用 beta 模型。
"""

from __future__ import annotations

import io
import os
import re
import threading
import unicodedata
from collections import Counter
from dataclasses import dataclass
from typing import Iterable

try:
    from PIL import (
        Image,
        ImageChops,
        ImageEnhance,
        ImageFilter,
        ImageOps,
    )
except ImportError as exc:
    raise SystemExit(
        "缺少 Pillow。請執行：\n"
        "python -m pip install --upgrade ddddocr Pillow"
    ) from exc


# 舊版 ddddocr 對新版 Pillow 的相容處理。
if not hasattr(Image, "ANTIALIAS"):
    Image.ANTIALIAS = Image.Resampling.LANCZOS  # type: ignore[attr-defined]


try:
    import ddddocr
except ImportError as exc:
    raise SystemExit(
        "缺少 ddddocr。請執行：\n"
        "python -m pip install --upgrade ddddocr Pillow"
    ) from exc


__all__ = ["ocrImage"]


# ============================================================
# OCR 設定
# ============================================================

CAPTCHA_LENGTH = max(
    1,
    int(
        os.getenv(
            "OCR_CAPTCHA_LENGTH",
            "4",
        )
    ),
)

MAX_IMAGE_PIXELS = max(
    1,
    int(
        os.getenv(
            "OCR_MAX_IMAGE_PIXELS",
            "40000000",
        )
    ),
)

MAX_IMAGE_SIDE = max(
    64,
    int(
        os.getenv(
            "OCR_MAX_IMAGE_SIDE",
            "4096",
        )
    ),
)


# auto：
#   先用一般模型；
#   沒有得到預期長度結果時，再使用 beta 模型。
#
# default：
#   只使用一般模型。
#
# beta：
#   只使用 beta 模型。
#
# both：
#   每張圖片同時使用一般與 beta 模型。
MODEL_MODE = (
    os.getenv(
        "OCR_DDDDOCR_MODEL",
        "auto",
    )
    .strip()
    .lower()
)

if MODEL_MODE not in {
    "auto",
    "default",
    "beta",
    "both",
}:
    MODEL_MODE = "auto"


# 關鍵修正：
# 不可改成 classifier.set_ranges(6)。
#
# ddddocr 1.6.1 的整數 range 實作，
# 會把整數 6 當成字元表前七個字元，
# 而不是文件所說的英文大小寫加數字。
ALLOWED_CHARACTERS = os.getenv(
    "OCR_ALLOWED_CHARACTERS",
    (
        "0123456789"
        "abcdefghijklmnopqrstuvwxyz"
        "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    ),
)

Image.MAX_IMAGE_PIXELS = MAX_IMAGE_PIXELS


# ============================================================
# OCR 候選結果
# ============================================================

@dataclass(frozen=True)
class Candidate:
    text: str
    source: str
    order: int

    @property
    def exact(self) -> bool:
        return bool(
            re.fullmatch(
                rf"[A-Za-z0-9]"
                rf"{{{CAPTCHA_LENGTH}}}",
                self.text,
            )
        )


# ============================================================
# 圖片正規化
# ============================================================

def normalize_image(
    image: Image.Image,
) -> Image.Image:
    """
    驗證圖片並轉成 RGB。

    不會修改呼叫端傳入的原始圖片。
    """
    if not isinstance(
        image,
        Image.Image,
    ):
        raise TypeError(
            "ocrImage() 只接受 "
            "PIL.Image.Image；"
            f"目前收到："
            f"{type(image).__name__}"
        )

    width, height = image.size

    if width <= 0 or height <= 0:
        raise ValueError(
            "圖片寬度或高度無效。"
        )

    pixel_count = width * height

    if pixel_count > MAX_IMAGE_PIXELS:
        raise ValueError(
            "圖片像素數超過限制："
            f"{pixel_count:,} > "
            f"{MAX_IMAGE_PIXELS:,}"
        )

    copied = image.copy()

    has_transparency = (
        copied.mode in {
            "RGBA",
            "LA",
        }
        or (
            copied.mode == "P"
            and "transparency"
            in copied.info
        )
    )

    if has_transparency:
        rgba = copied.convert(
            "RGBA"
        )

        background = Image.new(
            "RGBA",
            rgba.size,
            (
                255,
                255,
                255,
                255,
            ),
        )

        background.alpha_composite(
            rgba
        )

        rgb = background.convert(
            "RGB"
        )

        rgba.close()
        background.close()
        copied.close()

        return rgb

    rgb = copied.convert(
        "RGB"
    )

    copied.close()

    return rgb


def background_mask(
    image: Image.Image,
) -> Image.Image:
    """
    以圖片四角平均顏色作為背景色，
    建立與背景不同的二值遮罩。
    """
    rgb = image.convert(
        "RGB"
    )

    width, height = rgb.size

    corners = [
        rgb.getpixel(
            (0, 0)
        ),
        rgb.getpixel(
            (
                width - 1,
                0,
            )
        ),
        rgb.getpixel(
            (
                0,
                height - 1,
            )
        ),
        rgb.getpixel(
            (
                width - 1,
                height - 1,
            )
        ),
    ]

    background_color = tuple(
        round(
            sum(
                color[channel]
                for color in corners
            )
            / len(corners)
        )
        for channel in range(3)
    )

    background = Image.new(
        "RGB",
        rgb.size,
        background_color,
    )

    difference = ImageChops.difference(
        rgb,
        background,
    )

    mask = ImageOps.grayscale(
        difference
    )

    mask = mask.point(
        lambda value: (
            255
            if value >= 10
            else 0
        )
    )

    rgb.close()
    background.close()
    difference.close()

    return mask


def trim_background(
    image: Image.Image,
) -> Image.Image:
    """
    去除 td 截圖四周的大面積空白。
    """
    rgb = image.convert(
        "RGB"
    )

    width, height = rgb.size

    if width < 4 or height < 4:
        return rgb

    mask = background_mask(
        rgb
    )

    bounding_box = mask.getbbox()

    mask.close()

    if bounding_box is None:
        return rgb

    left, top, right, bottom = (
        bounding_box
    )

    margin_x = max(
        2,
        round(
            width * 0.01
        ),
    )

    margin_y = max(
        2,
        round(
            height * 0.05
        ),
    )

    left = max(
        0,
        left - margin_x,
    )

    top = max(
        0,
        top - margin_y,
    )

    right = min(
        width,
        right + margin_x,
    )

    bottom = min(
        height,
        bottom + margin_y,
    )

    if (
        right - left < 8
        or bottom - top < 8
    ):
        return rgb

    cropped = rgb.crop(
        (
            left,
            top,
            right,
            bottom,
        )
    )

    rgb.close()

    return cropped


def resize_small(
    image: Image.Image,
) -> Image.Image:
    """
    放大小型 CAPTCHA。

    最短邊會盡量放大到約 96 pixels。
    """
    width, height = image.size

    short_side = min(
        width,
        height,
    )

    long_side = max(
        width,
        height,
    )

    scale = min(
        4.0,
        max(
            1.0,
            96.0
            / max(
                short_side,
                1,
            ),
        ),
    )

    if (
        long_side * scale
        > MAX_IMAGE_SIDE
    ):
        scale = (
            MAX_IMAGE_SIDE
            / max(
                long_side,
                1,
            )
        )

    if 0.99 <= scale <= 1.01:
        return image.copy()

    target_width = max(
        1,
        round(
            width * scale
        ),
    )

    target_height = max(
        1,
        round(
            height * scale
        ),
    )

    return image.resize(
        (
            target_width,
            target_height,
        ),
        Image.Resampling.LANCZOS,
    )


# ============================================================
# Otsu 二值化
# ============================================================

def otsu_threshold(
    grayscale: Image.Image,
) -> int:
    """
    使用 Pillow histogram 計算 Otsu 二值化門檻。
    """
    histogram = grayscale.histogram()

    total_pixels = sum(
        histogram
    )

    if total_pixels <= 0:
        return 127

    total_sum = sum(
        index * count
        for index, count
        in enumerate(histogram)
    )

    background_weight = 0
    background_sum = 0.0

    best_variance = -1.0
    best_threshold = 127

    for threshold, count in enumerate(
        histogram
    ):
        background_weight += count

        if background_weight == 0:
            continue

        foreground_weight = (
            total_pixels
            - background_weight
        )

        if foreground_weight == 0:
            break

        background_sum += (
            threshold * count
        )

        background_mean = (
            background_sum
            / background_weight
        )

        foreground_mean = (
            (
                total_sum
                - background_sum
            )
            / foreground_weight
        )

        variance = (
            background_weight
            * foreground_weight
            * (
                background_mean
                - foreground_mean
            )
            ** 2
        )

        if variance > best_variance:
            best_variance = variance
            best_threshold = threshold

    return best_threshold


# ============================================================
# 圖片候選版本
# ============================================================

def build_primary_variants(
    image: Image.Image,
) -> list[
    tuple[
        str,
        Image.Image,
    ]
]:
    """
    建立主要 OCR 圖片版本：

    1. 裁切後原圖
    2. 灰階高對比圖
    3. 銳化圖
    4. 二值化圖
    5. 反相二值圖
    """
    trimmed = trim_background(
        image
    )

    base = resize_small(
        trimmed
    )

    trimmed.close()

    gray_source = ImageOps.grayscale(
        base
    )

    gray = ImageOps.autocontrast(
        gray_source,
        cutoff=1,
    )

    gray_source.close()

    contrast = ImageEnhance.Contrast(
        gray
    ).enhance(
        1.35
    )

    sharpened = contrast.filter(
        ImageFilter.UnsharpMask(
            radius=1.0,
            percent=160,
            threshold=2,
        )
    )

    threshold = otsu_threshold(
        contrast
    )

    binary = contrast.point(
        lambda value: (
            255
            if value >= threshold
            else 0
        )
    ).convert(
        "RGB"
    )

    inverted = ImageOps.invert(
        binary
    )

    variants = [
        (
            "base",
            base,
        ),
        (
            "gray",
            contrast.convert(
                "RGB"
            ),
        ),
        (
            "sharpened",
            sharpened.convert(
                "RGB"
            ),
        ),
        (
            "binary",
            binary,
        ),
        (
            "inverted",
            inverted,
        ),
    ]

    gray.close()
    contrast.close()
    sharpened.close()

    return variants


def build_horizontal_windows(
    image: Image.Image,
) -> list[
    tuple[
        str,
        Image.Image,
    ]
]:
    """
    td 截圖很寬時，建立水平滑動裁切。

    例如目前的圖片為 399 x 56，
    CAPTCHA 可能只佔圖片其中一小段。
    """
    rgb = image.convert(
        "RGB"
    )

    width, height = rgb.size

    if (
        width <= height * 4
        or width < 120
    ):
        rgb.close()
        return []

    windows: list[
        tuple[
            str,
            Image.Image,
        ]
    ] = []

    window_index = 0

    candidate_widths = sorted(
        {
            min(
                width,
                max(
                    80,
                    round(
                        height * ratio
                    ),
                ),
            )
            for ratio in (
                2.5,
                3.0,
                3.5,
                4.0,
                5.0,
            )
        }
    )

    for crop_width in candidate_widths:
        if crop_width >= width:
            continue

        step = max(
            20,
            crop_width // 3,
        )

        starts = list(
            range(
                0,
                width - crop_width + 1,
                step,
            )
        )

        final_start = (
            width - crop_width
        )

        if final_start not in starts:
            starts.append(
                final_start
            )

        for left in starts:
            right = (
                left + crop_width
            )

            raw_crop = rgb.crop(
                (
                    left,
                    0,
                    right,
                    height,
                )
            )

            trimmed_crop = trim_background(
                raw_crop
            )

            raw_crop.close()

            resized_crop = resize_small(
                trimmed_crop
            )

            trimmed_crop.close()

            window_index += 1

            windows.append(
                (
                    (
                        f"window_"
                        f"{window_index}_"
                        f"{left}"
                    ),
                    resized_crop,
                )
            )

    rgb.close()

    return windows


def to_png_bytes(
    image: Image.Image,
) -> bytes:
    """
    將 PIL.Image.Image 轉成
    ddddocr.classification() 接受的 bytes。
    """
    buffer = io.BytesIO()

    try:
        image.save(
            buffer,
            format="PNG",
            optimize=False,
        )

        image_bytes = (
            buffer.getvalue()
        )

    finally:
        buffer.close()

    if not image_bytes:
        raise RuntimeError(
            "圖片轉為 PNG bytes 後內容為空。"
        )

    return image_bytes


# ============================================================
# OCR 結果處理
# ============================================================

def clean_text(
    value: object,
) -> str:
    """
    正規化結果，只保留 ASCII 英文字母與數字。
    """
    if value is None:
        return ""

    normalized = unicodedata.normalize(
        "NFKC",
        str(value),
    )

    return re.sub(
        r"[^A-Za-z0-9]",
        "",
        normalized,
    )


def select_best(
    candidates: Iterable[
        Candidate
    ],
) -> str:
    """
    選出最佳 OCR 候選。

    優先順序：
    1. 恰好四碼英數。
    2. 多個圖片版本得到相同結果。
    3. 長度最接近預期。
    4. 較早產生的候選。
    """
    candidate_list = [
        candidate
        for candidate in candidates
        if candidate.text
    ]

    if not candidate_list:
        return ""

    frequencies = Counter(
        candidate.text
        for candidate
        in candidate_list
    )

    def score(
        candidate: Candidate,
    ) -> tuple[
        int,
        int,
        int,
        int,
    ]:
        return (
            (
                1
                if candidate.exact
                else 0
            ),
            frequencies[
                candidate.text
            ],
            -abs(
                len(
                    candidate.text
                )
                - CAPTCHA_LENGTH
            ),
            -candidate.order,
        )

    return max(
        candidate_list,
        key=score,
    ).text


def has_exact(
    candidates: Iterable[
        Candidate
    ],
) -> bool:
    return any(
        candidate.exact
        for candidate in candidates
    )


# ============================================================
# ddddocr 引擎
# ============================================================

class DdddOcrEngine:
    """
    共用 ddddocr 模型。

    不會針對每張圖片重新初始化模型。
    """

    def __init__(self) -> None:
        self.lock = threading.Lock()

        self.default_classifier = None
        self.beta_classifier = None

        if MODEL_MODE in {
            "auto",
            "default",
            "both",
        }:
            self.default_classifier = (
                self.create_classifier(
                    beta=False
                )
            )

        if MODEL_MODE in {
            "beta",
            "both",
        }:
            self.beta_classifier = (
                self.create_classifier(
                    beta=True
                )
            )

    @staticmethod
    def create_classifier(
        beta: bool,
    ):
        """
        建立 ddddocr OCR 模型。

        只傳入 ddddocr 1.6.1 支援的建構參數。
        """
        classifier = ddddocr.DdddOcr(
            ocr=True,
            det=False,
            old=False,
            beta=beta,
            use_gpu=False,
            device_id=0,
            show_ad=False,
        )

        # 關鍵修正：
        # 使用明確的英數字元字串，
        # 不使用 classifier.set_ranges(6)。
        classifier.set_ranges(
            ALLOWED_CHARACTERS
        )

        return classifier

    def get_beta_classifier(self):
        """
        auto 模式下延遲建立 beta 模型。
        """
        if self.beta_classifier is None:
            self.beta_classifier = (
                self.create_classifier(
                    beta=True
                )
            )

        return self.beta_classifier

    @staticmethod
    def classify(
        classifier,
        image: Image.Image,
    ) -> str:
        """
        執行一次 ddddocr OCR。
        """
        image_bytes = to_png_bytes(
            image
        )

        try:
            raw_result = (
                classifier.classification(
                    image_bytes,
                    png_fix=True,
                )
            )

        except TypeError:
            # 相容不接受 png_fix 的舊版本。
            raw_result = (
                classifier.classification(
                    image_bytes
                )
            )

        return clean_text(
            raw_result
        )

    def recognize_variants(
        self,
        classifier,
        source_prefix: str,
        variants: list[
            tuple[
                str,
                Image.Image,
            ]
        ],
        start_order: int,
    ) -> list[Candidate]:
        """
        使用同一個模型辨識全部圖片版本。
        """
        output: list[
            Candidate
        ] = []

        for offset, (
            name,
            image,
        ) in enumerate(
            variants
        ):
            result = self.classify(
                classifier,
                image,
            )

            output.append(
                Candidate(
                    text=result,
                    source=(
                        f"{source_prefix}:"
                        f"{name}"
                    ),
                    order=(
                        start_order
                        + offset
                    ),
                )
            )

        return output

    def recognize(
        self,
        image: Image.Image,
    ) -> str:
        """
        辨識一張 CAPTCHA。
        """
        normalized = normalize_image(
            image
        )

        primary_variants: list[
            tuple[
                str,
                Image.Image,
            ]
        ] = []

        window_variants: list[
            tuple[
                str,
                Image.Image,
            ]
        ] = []

        candidates: list[
            Candidate
        ] = []

        try:
            primary_variants = (
                build_primary_variants(
                    normalized
                )
            )

            with self.lock:
                if (
                    self.default_classifier
                    is not None
                ):
                    candidates.extend(
                        self.recognize_variants(
                            classifier=(
                                self.default_classifier
                            ),
                            source_prefix=(
                                "default"
                            ),
                            variants=(
                                primary_variants
                            ),
                            start_order=0,
                        )
                    )

                # 主要圖片版本沒有得到四碼時，
                # 才針對很寬的 td 截圖做水平裁切。
                if not has_exact(
                    candidates
                ):
                    window_variants = (
                        build_horizontal_windows(
                            normalized
                        )
                    )

                    if (
                        window_variants
                        and self.default_classifier
                        is not None
                    ):
                        candidates.extend(
                            self.recognize_variants(
                                classifier=(
                                    self.default_classifier
                                ),
                                source_prefix=(
                                    "default-window"
                                ),
                                variants=(
                                    window_variants
                                ),
                                start_order=(
                                    len(candidates)
                                ),
                            )
                        )

                use_beta = (
                    MODEL_MODE
                    in {
                        "beta",
                        "both",
                    }
                    or (
                        MODEL_MODE
                        == "auto"
                        and not has_exact(
                            candidates
                        )
                    )
                )

                if use_beta:
                    beta_classifier = (
                        self.get_beta_classifier()
                    )

                    candidates.extend(
                        self.recognize_variants(
                            classifier=(
                                beta_classifier
                            ),
                            source_prefix=(
                                "beta"
                            ),
                            variants=(
                                primary_variants
                            ),
                            start_order=(
                                len(candidates)
                            ),
                        )
                    )

                    if (
                        window_variants
                        and not has_exact(
                            candidates
                        )
                    ):
                        candidates.extend(
                            self.recognize_variants(
                                classifier=(
                                    beta_classifier
                                ),
                                source_prefix=(
                                    "beta-window"
                                ),
                                variants=(
                                    window_variants
                                ),
                                start_order=(
                                    len(candidates)
                                ),
                            )
                        )

        finally:
            normalized.close()

            for _, variant in (
                primary_variants
            ):
                try:
                    variant.close()
                except Exception:
                    pass

            for _, variant in (
                window_variants
            ):
                try:
                    variant.close()
                except Exception:
                    pass

        return select_best(
            candidates
        )


# ============================================================
# OCR Singleton
# ============================================================

_OCR_ENGINE: DdddOcrEngine | None = None
_OCR_ENGINE_LOCK = threading.Lock()


def get_ocr_engine() -> DdddOcrEngine:
    """
    取得共用 OCR 模型。
    """
    global _OCR_ENGINE

    if _OCR_ENGINE is not None:
        return _OCR_ENGINE

    with _OCR_ENGINE_LOCK:
        if _OCR_ENGINE is None:
            _OCR_ENGINE = (
                DdddOcrEngine()
            )

    return _OCR_ENGINE


# ============================================================
# 唯一對外公開函式
# ============================================================

def ocrImage(
    image: Image.Image,
) -> str:
    """
    傳入 PIL.Image.Image，回傳 OCR 字串。
    """
    if not isinstance(
        image,
        Image.Image,
    ):
        raise TypeError(
            "ocrImage() 只接受 "
            "PIL.Image.Image；"
            f"目前收到："
            f"{type(image).__name__}"
        )

    try:
        result = (
            get_ocr_engine()
            .recognize(
                image
            )
        )

    except Exception as exc:
        raise RuntimeError(
            "ddddocr 執行失敗："
            f"{type(exc).__name__}: "
            f"{exc}"
        ) from exc

    return str(
        result
    )