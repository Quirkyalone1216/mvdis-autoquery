#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
testOcr.py

用途：
1. 預設讀取專案根目錄下 results\mvdis_captcha 內的圖片。
2. 將每張圖片以 PIL.Image.Image 傳給：

       from englishAlphanumericOcrApi import ocrImage

3. englishAlphanumericOcrApi.py 內部使用 ddddocr 進行 OCR。
4. 在終端機逐張輸出檔名、尺寸、OCR 結果、格式判定與耗時。
5. 支援 PNG、JPG、JPEG、BMP、TIF、TIFF、WEBP。
6. 預設遞迴搜尋子資料夾。
7. 可選 ddddocr 模型模式：auto、default、beta、both。

建議放置位置：
    專案根目錄\src\testOcr.py

預設圖片位置：
    專案根目錄\results\mvdis_captcha

安裝：
    python -m pip install --upgrade ddddocr Pillow

執行：
    python .\src\testOcr.py

只測試前 5 張：
    python .\src\testOcr.py --limit 5

比較一般模型與 beta 模型：
    python .\src\testOcr.py --model-mode both
"""

from __future__ import annotations

import argparse
import importlib
import importlib.metadata
import os
import re
import sys
import time
from pathlib import Path
from typing import Callable, Sequence

try:
    from PIL import Image, UnidentifiedImageError
except ImportError as exc:
    raise SystemExit(
        "缺少 Pillow。請先執行：\n"
        "python -m pip install --upgrade ddddocr Pillow"
    ) from exc


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = (
    SCRIPT_DIR.parent
    if SCRIPT_DIR.name.lower() == "src"
    else SCRIPT_DIR
)
SRC_DIR = PROJECT_ROOT / "src"

if SRC_DIR.is_dir() and str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

DEFAULT_INPUT_DIR = (
    PROJECT_ROOT
    / "results"
    / "mvdis_captcha"
)

IMAGE_EXTENSIONS = {
    ".png",
    ".jpg",
    ".jpeg",
    ".bmp",
    ".tif",
    ".tiff",
    ".webp",
}

OcrFunction = Callable[
    [Image.Image],
    str,
]


def natural_sort_key(
    path: Path,
) -> list[object]:
    """
    讓 image_2.png 排在 image_10.png 前面。
    """
    parts = re.split(
        r"(\d+)",
        str(path).lower(),
    )

    return [
        (
            int(part)
            if part.isdigit()
            else part
        )
        for part in parts
    ]


def find_images(
    input_dir: Path,
    recursive: bool,
) -> list[Path]:
    """
    找出資料夾內支援的圖片。
    """
    iterator = (
        input_dir.rglob("*")
        if recursive
        else input_dir.iterdir()
    )

    images = [
        path
        for path in iterator
        if (
            path.is_file()
            and path.suffix.lower()
            in IMAGE_EXTENSIONS
        )
    ]

    return sorted(
        images,
        key=natural_sort_key,
    )


def load_ocr_function() -> OcrFunction:
    """
    延遲匯入 ocrImage()。

    main() 會先設定：
    - OCR_DDDDOCR_MODEL
    - OCR_CAPTCHA_LENGTH

    再匯入 englishAlphanumericOcrApi，
    使設定可以生效。
    """
    try:
        module = importlib.import_module(
            "englishAlphanumericOcrApi"
        )

    except ImportError as exc:
        raise SystemExit(
            "無法匯入 "
            "englishAlphanumericOcrApi.py。\n"
            "請確認檔案存在：\n"
            f"{SRC_DIR / 'englishAlphanumericOcrApi.py'}\n"
            "並安裝：\n"
            "python -m pip install --upgrade "
            "ddddocr Pillow"
        ) from exc

    ocr_function = getattr(
        module,
        "ocrImage",
        None,
    )

    if not callable(
        ocr_function
    ):
        raise SystemExit(
            "englishAlphanumericOcrApi.py "
            "內找不到：\n"
            "def ocrImage("
            "image: Image.Image"
            ") -> str"
        )

    return ocr_function


def get_package_version(
    package_name: str,
) -> str:
    """
    取得套件版本。
    """
    try:
        return importlib.metadata.version(
            package_name
        )

    except importlib.metadata.PackageNotFoundError:
        return "unknown"


def recognize_image(
    image_path: Path,
    ocr_function: OcrFunction,
) -> tuple[
    str,
    int,
    int,
    str,
    float,
]:
    """
    開啟單張圖片並呼叫 ocrImage()。
    """
    started_at = (
        time.perf_counter()
    )

    with Image.open(
        image_path
    ) as opened_image:
        opened_image.load()

        width, height = (
            opened_image.size
        )

        image_mode = (
            opened_image.mode
        )

        # copy() 後，即使 Image.open() 已關閉，
        # ocrImage() 仍能安全使用圖片。
        input_image = (
            opened_image.copy()
        )

    try:
        ocr_text = str(
            ocr_function(
                input_image
            )
        )

    finally:
        input_image.close()

    elapsed_seconds = (
        time.perf_counter()
        - started_at
    )

    return (
        ocr_text,
        width,
        height,
        image_mode,
        elapsed_seconds,
    )


def is_expected_captcha(
    value: str,
    expected_length: int,
) -> bool:
    """
    判斷結果是否為指定長度的 ASCII 英數字元。
    """
    return bool(
        re.fullmatch(
            rf"[A-Za-z0-9]"
            rf"{{{expected_length}}}",
            value,
        )
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "使用 "
            "englishAlphanumericOcrApi."
            "ocrImage() 與 ddddocr，"
            "辨識 results\\mvdis_captcha "
            "內的圖片。"
        )
    )

    parser.add_argument(
        "--input-dir",
        type=Path,
        default=DEFAULT_INPUT_DIR,
        help=(
            "要辨識的圖片資料夾；"
            "預設為專案根目錄下的 "
            r"results\mvdis_captcha"
        ),
    )

    parser.add_argument(
        "--recursive",
        action=(
            argparse.BooleanOptionalAction
        ),
        default=True,
        help=(
            "是否遞迴搜尋子資料夾；"
            "可用 --no-recursive 關閉"
        ),
    )

    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help=(
            "最多辨識幾張圖片；"
            "不指定代表處理全部圖片"
        ),
    )

    parser.add_argument(
        "--model-mode",
        choices=(
            "auto",
            "default",
            "beta",
            "both",
        ),
        default="auto",
        help=(
            "ddddocr 模型模式："
            "auto=一般模型失敗時再用 beta；"
            "default=只用一般模型；"
            "beta=只用 beta；"
            "both=每張圖都比較兩套模型"
        ),
    )

    parser.add_argument(
        "--expected-length",
        type=int,
        default=4,
        help=(
            "預期 CAPTCHA 字元數；"
            "預設為 4"
        ),
    )

    return parser


def main(
    argv: Sequence[str] | None = None,
) -> int:
    args = (
        build_parser()
        .parse_args(
            argv
        )
    )

    if (
        args.limit is not None
        and args.limit <= 0
    ):
        raise SystemExit(
            "--limit 必須大於 0"
        )

    if args.expected_length <= 0:
        raise SystemExit(
            "--expected-length 必須大於 0"
        )

    # englishAlphanumericOcrApi.py
    # 會在匯入時讀取這兩個設定，
    # 因此必須先設定再匯入。
    os.environ[
        "OCR_DDDDOCR_MODEL"
    ] = args.model_mode

    os.environ[
        "OCR_CAPTCHA_LENGTH"
    ] = str(
        args.expected_length
    )

    input_dir = (
        args.input_dir
        .expanduser()
    )

    if not input_dir.is_absolute():
        input_dir = (
            PROJECT_ROOT
            / input_dir
        )

    input_dir = (
        input_dir.resolve()
    )

    if not input_dir.exists():
        raise SystemExit(
            "找不到 OCR 圖片資料夾：\n"
            f"{input_dir}"
        )

    if not input_dir.is_dir():
        raise SystemExit(
            "OCR 輸入路徑不是資料夾：\n"
            f"{input_dir}"
        )

    image_paths = find_images(
        input_dir=input_dir,
        recursive=args.recursive,
    )

    if args.limit is not None:
        image_paths = (
            image_paths[
                :args.limit
            ]
        )

    if not image_paths:
        supported = ", ".join(
            sorted(
                IMAGE_EXTENSIONS
            )
        )

        raise SystemExit(
            "資料夾內找不到可辨識的圖片。\n"
            f"資料夾：{input_dir}\n"
            f"支援格式：{supported}"
        )

    # 必須在環境變數設定完成後，
    # 才匯入 OCR 模組。
    ocr_function = (
        load_ocr_function()
    )

    print(
        "=" * 72
    )
    print(
        "ddddocr OCR 測試開始"
    )
    print(
        f"圖片資料夾：{input_dir}"
    )
    print(
        f"圖片數量：{len(image_paths)}"
    )
    print(
        "遞迴搜尋："
        f"{'是' if args.recursive else '否'}"
    )
    print(
        "OCR 副程式："
        "englishAlphanumericOcrApi."
        "ocrImage(image)"
    )
    print(
        "ddddocr 版本："
        f"{get_package_version('ddddocr')}"
    )
    print(
        "Pillow 版本："
        f"{get_package_version('Pillow')}"
    )
    print(
        f"模型模式：{args.model_mode}"
    )
    print(
        "預期格式："
        f"{args.expected_length} "
        "個 ASCII 英數字元"
    )
    print(
        "提醒：第一次辨識時需要初始化 "
        "ddddocr ONNX 模型。"
    )
    print(
        "=" * 72
    )

    total_started_at = (
        time.perf_counter()
    )

    valid_count = 0
    invalid_count = 0
    empty_count = 0
    error_count = 0

    for index, image_path in enumerate(
        image_paths,
        start=1,
    ):
        try:
            relative_path = (
                image_path.relative_to(
                    input_dir
                )
            )

        except ValueError:
            relative_path = (
                image_path
            )

        print()
        print(
            f"[{index}/"
            f"{len(image_paths)}] "
            f"{relative_path}"
        )

        try:
            (
                ocr_text,
                width,
                height,
                image_mode,
                elapsed_seconds,
            ) = recognize_image(
                image_path=(
                    image_path
                ),
                ocr_function=(
                    ocr_function
                ),
            )

            print(
                "  圖片資訊："
                f"{width} x {height}，"
                f"模式={image_mode}"
            )

            if not ocr_text:
                print(
                    "  OCR 結果：''"
                    "（沒有辨識到文字）"
                )
                print(
                    "  格式判定：空白"
                )

                empty_count += 1

            elif is_expected_captcha(
                ocr_text,
                args.expected_length,
            ):
                print(
                    f"  OCR 結果："
                    f"{ocr_text!r}"
                )
                print(
                    "  格式判定："
                    "符合預期長度的純英數"
                )

                valid_count += 1

            else:
                print(
                    f"  OCR 結果："
                    f"{ocr_text!r}"
                )
                print(
                    "  格式判定："
                    "不符合預期格式；"
                    f"實際長度="
                    f"{len(ocr_text)}"
                )

                invalid_count += 1

            print(
                "  OCR 耗時："
                f"{elapsed_seconds:.3f} 秒"
            )

        except UnidentifiedImageError as exc:
            error_count += 1

            print(
                "  OCR 錯誤："
                "Pillow 無法辨識圖片；"
                f"{type(exc).__name__}: "
                f"{exc}"
            )

        except KeyboardInterrupt:
            print(
                "\n使用者中止 OCR 測試。"
            )

            return 130

        except Exception as exc:
            error_count += 1

            print(
                "  OCR 錯誤："
                f"{type(exc).__name__}: "
                f"{exc}"
            )

    total_elapsed_seconds = (
        time.perf_counter()
        - total_started_at
    )

    average_seconds = (
        total_elapsed_seconds
        / len(image_paths)
    )

    print()
    print(
        "=" * 72
    )
    print(
        "ddddocr OCR 測試完成"
    )
    print(
        f"格式正確：{valid_count}"
    )
    print(
        f"格式不符：{invalid_count}"
    )
    print(
        f"空白結果：{empty_count}"
    )
    print(
        f"處理錯誤：{error_count}"
    )
    print(
        "總耗時："
        f"{total_elapsed_seconds:.3f} 秒"
    )
    print(
        "平均每張："
        f"{average_seconds:.3f} 秒"
    )
    print(
        "=" * 72
    )

    return (
        0
        if error_count == 0
        else 2
    )


if __name__ == "__main__":
    raise SystemExit(
        main()
    )