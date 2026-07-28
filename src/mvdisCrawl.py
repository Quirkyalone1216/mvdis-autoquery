#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# VERSION: DUAL_CURRENT_AND_PAID_QUERY_2026-07-28
"""
監理服務網「法人交通違規繳費紀錄查詢」批次輔助工具。

功能：
1. 從既有 Excel 找出「統一編號／登記編號」與登記名稱。
2. 使用 Playwright 開啟監理服務網，自動填入統一編號。
3. 依指定 XPath／語意備援定位 CAPTCHA 所在的 td 區塊，保存成 PNG 或 JPG。
4. 將驗證碼圖片開啟成 PIL.Image.Image。
5. 使用指定的匯入方式：

       from englishAlphanumeric import imageInput

   並直接呼叫：

       imageInput(image)

   由 englishAlphanumeric.py 內的人工 UI 顯示圖片、等待使用者輸入，
   再接收 imageInput() 回傳的四個英數字元。
6. 自動把回傳字串填入網頁驗證碼欄位，按下查詢。
7. 每筆公司先查詢「交通違規（含強制險）查詢及繳納」，擷取可線上與不可線上繳納表格。
8. 再查詢原本的「法人交通違規繳費紀錄」，並將兩階段結果合併。
9. 明細工作表沿用來源 Excel 的「155598」格式：A:G、無額外標題與中繼資料。
9. 更新主表的交通違規筆數、查詢狀態、查詢時間與錯誤訊息。
10. 永遠寫入新的 Excel；支援逐筆儲存及中斷續跑。
11. 若某筆 CAPTCHA 在同一輪內連續失敗，會先移至工作佇列尾端，
    讓其他資料先完成，再自動回頭重做，並限制重排次數避免無限循環。

englishAlphanumeric.py 必須與本腳本放在同一資料夾，並提供：

    def imageInput(
        image: Image.Image,
    ) -> str:
        ...

imageInput() 必須在人工 UI 完成輸入後回傳恰好四個英數字元。
本程式不包含 OCR、自動辨識、破解或繞過驗證碼。

僅可查詢您有權處理的法人／商業資料，並請遵守監理服務網使用規範。
"""

from __future__ import annotations

import argparse
from collections import deque
from copy import copy
import hashlib
from html import unescape
import json
import os
import random
import re
import shutil
import tempfile
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable, Sequence

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError as exc:
    raise SystemExit(
        "缺少 openpyxl。請先執行：\n"
        "python -m pip install openpyxl playwright pillow\n"
        "python -m playwright install chromium"
    ) from exc

try:
    from PIL import Image
except ImportError as exc:
    raise SystemExit(
        "缺少 Pillow。請先執行：python -m pip install pillow"
    ) from exc

try:
    from englishAlphanumericOcrApi import ocrImage
except ImportError as exc:
    raise SystemExit(
        "無法從 englishAlphanumericOcrApi.py 匯入 ocrImage。\n"
        "請將 englishAlphanumericOcrApi.py 放在本腳本同一資料夾，\n"
        "並確認其中有 def ocrImage(image: Image.Image) -> str"
    ) from exc

try:
    from playwright.sync_api import (
        Browser,
        Error as PlaywrightError,
        Locator,
        Page,
        TimeoutError as PlaywrightTimeoutError,
        sync_playwright,
    )
except ImportError as exc:
    raise SystemExit(
        "缺少 Playwright。請先執行：\n"
        "python -m pip install playwright\n"
        "python -m playwright install chromium"
    ) from exc


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent

# mvdisCrawl.py 位於 src 資料夾，因此從專案根目錄讀取：
# Data\高雄市.xlsx
DEFAULT_INPUT_PATH = PROJECT_ROOT / "Data" / "高雄市.xlsx"

# 輸出資料夾放在專案根目錄下的 results。
DEFAULT_RESULTS_DIR = PROJECT_ROOT / "results"
DEFAULT_CAPTCHA_DIR = DEFAULT_RESULTS_DIR / "mvdis_captcha"
DEFAULT_DEBUG_DIR = DEFAULT_RESULTS_DIR / "mvdis_debug"

DEFAULT_URL = (
    "https://www.mvdis.gov.tw/m3-emv-vil/vil/penaltyQueryPayRecord/legal"
)

CAPTCHA_CAPTURE_XPATH = (
    "/html/body/table/tbody/tr[2]/td[1]/div[2]/div/div/form/"
    "table/tbody/tr[4]/td"
)

# CAPTCHA 本身的精確圖片 XPath。優先截取 img，避免把整個 td 的
# 標籤、輸入框及「換一張」文字一起送進 OCR。
CAPTCHA_IMAGE_XPATH = (
    "/html/body/table/tbody/tr[2]/td[1]/div[2]/div/div/form/"
    "table/tbody/tr[4]/td/img"
)

# 監理服務網法人違規查詢頁面的實際查詢按鈕是 <a> 元素。
QUERY_BUTTON_XPATH = (
    "/html/body/table/tbody/tr[2]/td[1]/div[2]/div/div/form/div/a"
)

# 查詢結果頁中，完整「罰鍰繳納紀錄」JSON 所在的 hidden input。
# 使用者提供的實際 XPath：
RESULT_JSON_XPATH = (
    "/html/body/table/tbody/tr[2]/td[1]/div[3]/form/input"
)

# 查詢結果頁顯示「查無已繳納罰鍰資料」的精確 XPath。
# 這個節點出現時，代表網站已正常完成查詢，只是結果筆數為 0，
# 不應寫成「錯誤」或「無資料」，而應寫成「成功」與 0 筆。
NO_PAID_DATA_XPATH = (
    "/html/body/table/tbody/tr[2]/td[1]/div[1]/table/"
    "tbody/tr/td[1]"
)

# 第一階段：「交通違規（含強制險）查詢及繳納」。
# 依使用者指定，先從目前的法人繳費紀錄入口頁進入；若該入口頁
# 沒有載入新版查詢頁，則使用監理服務網正式的 penaltyQueryPay 路徑備援。
DEFAULT_CURRENT_QUERY_URL = DEFAULT_URL
CURRENT_QUERY_DIRECT_FALLBACK_URL = (
    "https://www.mvdis.gov.tw/m3-emv-vil/vil/penaltyQueryPay"
)

# 「交通違規（含強制險）查詢及繳納」法人頁籤。
CURRENT_QUERY_CORPORATE_TAB_XPATH = (
    "/html/body/table/tbody/tr[2]/td[1]/div[3]/table/"
    "tbody/tr/td[2]/span/img"
)

# 法人資料與 CAPTCHA 填完後的查詢按鈕。
CURRENT_QUERY_BUTTON_XPATH = (
    "/html/body/table/tbody/tr[2]/td[1]/div[3]/div/"
    "div[2]/form/div/a"
)

# 第一階段查詢結果容器與目前頁籤所顯示的表格。
CURRENT_RESULT_CONTAINER_XPATH = (
    "/html/body/table/tbody/tr[2]/td[1]/form/div[2]/div[2]/div"
)
CURRENT_RESULT_TABLE_XPATH = (
    "/html/body/table/tbody/tr[2]/td[1]/form/div[2]/div[2]/"
    "div/div[2]/table"
)

# 結果頁第二個頁籤；點擊後 CURRENT_RESULT_TABLE_XPATH 會切換成
# 另一組（不可線上繳納／需臨櫃處理）資料。
CURRENT_RESULT_SECOND_TAB_XPATH = (
    "/html/body/table/tbody/tr[2]/td[1]/form/div[2]/"
    "table[2]/tbody/tr/td[2]/span/img"
)

COMBINED_QUERY_MESSAGE_MARKER = "雙階段查詢完成"
CURRENT_ONLINE_CATEGORY = "可線上繳納"
CURRENT_OFFLINE_CATEGORY = "不可線上繳納"

DETAIL_TEMPLATE_SHEET_NAME = "__違規明細範本__"
DEFAULT_DETAIL_TEMPLATE_SHEET = "155598"

TABLE_KIND_UNPAID = "unpaid"
TABLE_KIND_PAID = "paid"

LEGACY_BORDER_COLOR = "CACACA"
LEGACY_PAID_FONT_COLOR = "4D4D4D"
LEGACY_UNPAID_FONT_COLOR = "FF0000"
LEGACY_NEEDS_APPEARANCE_FONT_COLOR = "1F4E78"
LEGACY_ALT_FILL_COLOR = "F5F5F5"

STATUS_SUCCESS = "成功"
STATUS_NO_DATA = "無資料"
STATUS_ERROR = "錯誤"
STATUS_RETRY = "待重試"
STATUS_SKIPPED = "略過"

HEADER_ID_PATTERNS = ("統一編號", "登記編號")
HEADER_NAME_PATTERNS = ("登記名稱", "公司名稱", "商業名稱", "名稱")

RESULT_KEYWORDS = (
    "違規日期",
    "違規事由",
    "罰鍰",
    "舉發單號",
    "車牌",
    "應到案",
    "繳款",
    "處理方式",
)

NO_DATA_PATTERNS = (
    "查無資料",
    "查無違規",
    "無違規資料",
    "目前無違規",
    "沒有違規資料",
    "無符合資料",
    "查無繳費紀錄",
    # 結果頁實際使用的訊息。舊版遺漏這句，導致正常無資料
    # 被錯誤寫成「錯誤」。
    "查無已繳納罰鍰資料",
    "無已繳納罰鍰資料",
    "查無罰鍰繳納資料",
    "查無可線上繳納罰單資料",
    "查無不可線上繳納罰單資料",
    "查無交通違規資料",
    "無交通違規資料",
)

CAPTCHA_ERROR_PATTERNS = (
    "驗證碼錯誤",
    "驗證碼不正確",
    "驗證碼有誤",
    "請輸入驗證碼",
    "驗證碼輸入錯誤",
    "驗證碼輸入有誤",
    "驗證碼驗證失敗",
)

FORM_ERROR_PATTERNS = (
    "請輸入統一編號",
    "統一編號錯誤",
    "統一編號格式",
    "系統忙碌",
    "查詢失敗",
    "操作逾時",
)

# 這些不是永久資料錯誤，應重新查詢，而不是直接寫成「錯誤」。
TRANSIENT_FORM_ERROR_PATTERNS = (
    "系統忙碌",
    "查詢失敗",
    "操作逾時",
)

THIN_GRAY = Side(style="thin", color="D9E2F3")
TITLE_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FILL = PatternFill("solid", fgColor="5B9BD5")
SUBTITLE_FILL = PatternFill("solid", fgColor="D9EAF7")
WHITE_FONT = Font(color="FFFFFF", bold=True)
HEADER_FONT = Font(bold=True, color="FFFFFF")
BOLD_FONT = Font(bold=True)


@dataclass(frozen=True)
class CompanyRow:
    excel_row: int
    raw_id: str
    query_id: str
    name: str


@dataclass
class ExtractedTable:
    kind: str
    title: str
    headers: list[str]
    rows: list[list[Any]]

    @property
    def record_count(self) -> int:
        return len(self.rows)


@dataclass
class QueryOutcome:
    status: str
    record_count: int
    tables: list[ExtractedTable]
    message: str
    result_url: str
    retry_later: bool = False


class CaptchaRecognitionError(RuntimeError):
    """OCR 沒有產生符合格式的四碼字串；可重新取得 CAPTCHA。"""


class TransientPageError(RuntimeError):
    """頁面仍在導向或結果尚未完整載入；應重新查詢。"""


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def normalize_company_id(value: Any) -> tuple[str, str]:
    """Return (raw display id, 8-digit query id)."""
    if value is None:
        raise ValueError("統一編號是空白")

    if isinstance(value, bool):
        raise ValueError("統一編號格式錯誤")

    if isinstance(value, int):
        raw = str(value)
    elif isinstance(value, float):
        if not value.is_integer():
            raise ValueError(f"統一編號不是整數：{value}")
        raw = str(int(value))
    else:
        raw = normalize_text(value)
        if raw.endswith(".0") and raw[:-2].isdigit():
            raw = raw[:-2]

    digits = re.sub(r"\D", "", raw)

    if not digits:
        raise ValueError(f"找不到數字：{raw!r}")

    if len(digits) > 8:
        raise ValueError(f"統一編號超過 8 碼：{raw!r}")

    return raw, digits.zfill(8)


def find_main_sheet_and_columns(
    workbook: Any,
) -> tuple[Any, int, int, int, int]:
    """Return worksheet, header_start_row, data_start_row, id_col, name_col."""
    for worksheet in workbook.worksheets:
        max_scan_row = min(max(worksheet.max_row, 1), 12)
        max_scan_col = min(max(worksheet.max_column, 1), 30)

        id_candidates: list[tuple[int, int]] = []
        name_candidates: list[tuple[int, int]] = []

        for row in range(1, max_scan_row + 1):
            for col in range(1, max_scan_col + 1):
                text = normalize_text(worksheet.cell(row, col).value)

                if not text:
                    continue

                if any(pattern in text for pattern in HEADER_ID_PATTERNS):
                    id_candidates.append((row, col))

                if any(pattern in text for pattern in HEADER_NAME_PATTERNS):
                    name_candidates.append((row, col))

        if not id_candidates or not name_candidates:
            continue

        id_row, id_col = min(id_candidates)

        compatible_names = [
            item
            for item in name_candidates
            if abs(item[0] - id_row) <= 2
        ]

        name_row, name_col = min(
            compatible_names or name_candidates
        )

        header_start = min(id_row, name_row)
        header_end = max(id_row, name_row)

        for probe_row in range(
            header_start,
            min(header_start + 3, worksheet.max_row) + 1,
        ):
            row_text = " ".join(
                normalize_text(worksheet.cell(probe_row, col).value)
                for col in range(1, max_scan_col + 1)
            )

            if any(
                pattern in row_text
                for pattern in HEADER_ID_PATTERNS
            ):
                header_end = max(header_end, probe_row)

        return (
            worksheet,
            header_start,
            header_end + 1,
            id_col,
            name_col,
        )

    raise ValueError(
        "找不到同時包含『統一編號／登記編號』與『登記名稱』的主工作表"
    )


def iter_companies(
    worksheet: Any,
    data_start_row: int,
    id_col: int,
    name_col: int,
    start_row: int | None,
    end_row: int | None,
    limit: int | None,
) -> list[CompanyRow]:
    first = max(data_start_row, start_row or data_start_row)
    last = min(worksheet.max_row, end_row or worksheet.max_row)

    companies: list[CompanyRow] = []

    for row in range(first, last + 1):
        value = worksheet.cell(row, id_col).value
        name = normalize_text(worksheet.cell(row, name_col).value)

        if value in (None, ""):
            continue

        try:
            raw_id, query_id = normalize_company_id(value)
        except ValueError:
            continue

        companies.append(
            CompanyRow(
                excel_row=row,
                raw_id=raw_id,
                query_id=query_id,
                name=name,
            )
        )

        if limit is not None and len(companies) >= limit:
            break

    return companies


def ensure_tracking_columns(
    worksheet: Any,
    header_start_row: int,
) -> dict[str, int]:
    desired = [
        "交通違規筆數",
        "違規查詢狀態",
        "最後查詢時間",
        "違規查詢訊息",
    ]

    existing: dict[str, int] = {}

    for col in range(1, worksheet.max_column + 1):
        text = normalize_text(
            worksheet.cell(header_start_row, col).value
        )

        if text in desired:
            existing[text] = col

    if "交通違規筆數" not in existing:
        numeric_blank_candidates: list[tuple[int, int]] = []

        for col in range(1, worksheet.max_column + 1):
            if normalize_text(
                worksheet.cell(header_start_row, col).value
            ):
                continue

            numeric_count = 0

            for row in range(
                header_start_row + 1,
                min(
                    worksheet.max_row,
                    header_start_row + 250,
                ) + 1,
            ):
                value = worksheet.cell(row, col).value

                if (
                    isinstance(value, (int, float))
                    and not isinstance(value, bool)
                ):
                    numeric_count += 1

            if numeric_count:
                numeric_blank_candidates.append(
                    (col, numeric_count)
                )

        if numeric_blank_candidates:
            existing["交通違規筆數"] = max(
                numeric_blank_candidates
            )[0]

    next_col = worksheet.max_column + 1

    for header in desired:
        if header not in existing:
            existing[header] = next_col
            worksheet.cell(
                header_start_row,
                next_col,
                header,
            )
            next_col += 1

    widths = {
        "交通違規筆數": 14,
        "違規查詢狀態": 14,
        "最後查詢時間": 20,
        "違規查詢訊息": 42,
    }

    for header, col in existing.items():
        cell = worksheet.cell(header_start_row, col)
        cell.value = header
        cell.fill = TITLE_FILL
        cell.font = WHITE_FONT
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )
        cell.border = Border(
            left=THIN_GRAY,
            right=THIN_GRAY,
            top=THIN_GRAY,
            bottom=THIN_GRAY,
        )

        worksheet.column_dimensions[
            get_column_letter(col)
        ].width = widths[header]

    return existing


def reset_tracking_values(
    worksheet: Any,
    data_start_row: int,
    tracking_columns: dict[str, int],
) -> None:
    for row in range(
        data_start_row,
        worksheet.max_row + 1,
    ):
        for col in tracking_columns.values():
            worksheet.cell(row, col).value = None


def normalized_numeric_sheet_title(value: Any) -> str:
    text = normalize_text(value)
    digits = re.sub(r"\D", "", text)

    if not digits:
        return ""

    return digits.lstrip("0") or "0"


def detail_sheet_name(
    raw_id: str,
    query_id: str,
) -> str:
    """
    與來源檔一致，工作表名稱使用 Excel 中顯示的登記編號。

    例如：
        查詢用編號：00155598
        工作表名稱：155598
    """
    normalized_raw = normalized_numeric_sheet_title(
        raw_id
    )

    normalized_query = normalized_numeric_sheet_title(
        query_id
    )

    return safe_sheet_name(
        normalized_raw
        or normalized_query
        or query_id
    )


def clear_stale_numeric_detail_sheets(
    workbook: Any,
    main_title: str,
) -> None:
    for worksheet in list(workbook.worksheets):
        title = worksheet.title.strip()

        if title in {
            main_title,
            DETAIL_TEMPLATE_SHEET_NAME,
        }:
            continue

        if re.fullmatch(r"\d{1,8}", title):
            workbook.remove(worksheet)


def safe_sheet_name(company_id: str) -> str:
    return (
        re.sub(r"[\\/*?:\[\]]", "_", company_id)[:31]
        or "明細"
    )


def remove_existing_detail_variants(
    workbook: Any,
    raw_id: str,
    query_id: str,
) -> None:
    candidates = {
        safe_sheet_name(raw_id),
        safe_sheet_name(query_id),
        detail_sheet_name(
            raw_id,
            query_id,
        ),
    }

    candidates.discard(
        DETAIL_TEMPLATE_SHEET_NAME
    )

    for title in list(candidates):
        if title in workbook.sheetnames:
            workbook.remove(
                workbook[title]
            )


def legacy_medium_side() -> Side:
    return Side(
        style="medium",
        color=LEGACY_BORDER_COLOR,
    )


def legacy_border(
    *,
    left: bool,
    right: bool = True,
    top: bool,
    bottom: bool = True,
) -> Border:
    return Border(
        left=(
            legacy_medium_side()
            if left
            else Side()
        ),
        right=(
            legacy_medium_side()
            if right
            else Side()
        ),
        top=(
            legacy_medium_side()
            if top
            else Side()
        ),
        bottom=(
            legacy_medium_side()
            if bottom
            else Side()
        ),
    )


def apply_legacy_detail_row_style(
    worksheet: Any,
    row: int,
    *,
    column_count: int,
    font_color: str,
    first_row: bool,
    fill_color: str | None,
    hyperlink_column: int | None,
    row_height: float,
) -> None:
    """
    建立與來源檔 155598 工作表相同的基本視覺格式。

    - Arial
    - 一般內容 6 pt
    - 「檢視」10 pt、底線
    - #CACACA medium 框線
    - 垂直置中、文字換行
    - 交替灰底 #F5F5F5
    """
    for column in range(
        1,
        column_count + 1,
    ):
        cell = worksheet.cell(
            row,
            column,
        )

        is_hyperlink = (
            hyperlink_column == column
        )

        cell.font = Font(
            name="Arial",
            size=(
                10
                if is_hyperlink
                else 6
            ),
            color=font_color,
            underline=(
                "single"
                if is_hyperlink
                else None
            ),
        )

        cell.fill = (
            PatternFill(
                "solid",
                fgColor=fill_color,
            )
            if fill_color
            else PatternFill(
                fill_type=None
            )
        )

        cell.border = legacy_border(
            left=(column == 1),
            top=first_row,
        )

        cell.alignment = Alignment(
            vertical="center",
            wrap_text=True,
        )

    worksheet.row_dimensions[
        row
    ].height = row_height


def build_fallback_detail_template(
    worksheet: Any,
) -> None:
    """
    當續跑檔案中不存在來源 155598 工作表時，
    以程式建立等價的隱藏範本。
    """
    worksheet.sheet_format.defaultRowHeight = 12.5
    worksheet.freeze_panes = None
    worksheet.sheet_view.showGridLines = True

    try:
        worksheet.sheet_properties.pageSetUpPr.fitToPage = False
    except Exception:
        pass

    worksheet.page_setup.orientation = "portrait"
    worksheet.page_setup.paperSize = (
        worksheet.PAPERSIZE_A4
    )

    worksheet.page_margins.left = 0.7
    worksheet.page_margins.right = 0.7
    worksheet.page_margins.top = 0.75
    worksheet.page_margins.bottom = 0.75
    worksheet.page_margins.header = 0.3
    worksheet.page_margins.footer = 0.3

    # 155598 範本：
    # 1：第一筆一般未繳
    # 2～4：後續一般未繳
    # 5：需到案
    # 6：第一筆繳納紀錄
    # 7～10：中間繳納紀錄
    # 11：最後一筆繳納紀錄（灰底）
    apply_legacy_detail_row_style(
        worksheet,
        1,
        column_count=6,
        font_color=(
            LEGACY_UNPAID_FONT_COLOR
        ),
        first_row=True,
        fill_color=None,
        hyperlink_column=6,
        row_height=24.5,
    )

    for row in range(2, 5):
        apply_legacy_detail_row_style(
            worksheet,
            row,
            column_count=6,
            font_color=(
                LEGACY_UNPAID_FONT_COLOR
            ),
            first_row=False,
            fill_color=None,
            hyperlink_column=6,
            row_height=24.5,
        )

    apply_legacy_detail_row_style(
        worksheet,
        5,
        column_count=6,
        font_color=(
            LEGACY_NEEDS_APPEARANCE_FONT_COLOR
        ),
        first_row=True,
        fill_color=(
            LEGACY_ALT_FILL_COLOR
        ),
        hyperlink_column=6,
        row_height=16.5,
    )

    apply_legacy_detail_row_style(
        worksheet,
        6,
        column_count=7,
        font_color=(
            LEGACY_PAID_FONT_COLOR
        ),
        first_row=True,
        fill_color=None,
        hyperlink_column=None,
        row_height=16.5,
    )

    for row in range(7, 11):
        apply_legacy_detail_row_style(
            worksheet,
            row,
            column_count=7,
            font_color=(
                LEGACY_PAID_FONT_COLOR
            ),
            first_row=False,
            fill_color=None,
            hyperlink_column=None,
            row_height=16.5,
        )

    apply_legacy_detail_row_style(
        worksheet,
        11,
        column_count=7,
        font_color=(
            LEGACY_PAID_FONT_COLOR
        ),
        first_row=False,
        fill_color=(
            LEGACY_ALT_FILL_COLOR
        ),
        hyperlink_column=None,
        row_height=16.5,
    )



def is_legacy_detail_template_sheet(
    worksheet: Any,
) -> bool:
    if (
        worksheet.max_column > 7
        or worksheet.max_row < 1
    ):
        return False

    first_values = {
        normalize_text(
            worksheet.cell(
                row,
                column,
            ).value
        )
        for row in range(
            1,
            min(
                worksheet.max_row,
                8,
            ) + 1,
        )
        for column in range(
            1,
            min(
                worksheet.max_column,
                7,
            ) + 1,
        )
    }

    if first_values.intersection(
        {
            "統一編號",
            "登記名稱",
            "查詢狀態",
            "違規筆數",
            "查詢時間",
            "來源網址",
        }
    ):
        return False

    return True


def find_detail_template_candidate(
    workbook: Any,
    main_title: str,
    preferred_title: str,
) -> Any | None:
    preferred_normalized = normalize_text(
        preferred_title
    )

    if (
        preferred_normalized
        and preferred_normalized
        in workbook.sheetnames
    ):
        preferred_sheet = workbook[
            preferred_normalized
        ]

        if is_legacy_detail_template_sheet(
            preferred_sheet
        ):
            return preferred_sheet

    preferred_numeric = (
        normalized_numeric_sheet_title(
            preferred_normalized
        )
    )

    for worksheet in workbook.worksheets:
        title = normalize_text(
            worksheet.title
        )

        if title in {
            main_title,
            DETAIL_TEMPLATE_SHEET_NAME,
            "違規查詢紀錄",
        }:
            continue

        numeric_title = (
            normalized_numeric_sheet_title(
                title
            )
        )

        if (
            preferred_numeric
            and numeric_title
            == preferred_numeric
            and is_legacy_detail_template_sheet(
                worksheet
            )
        ):
            return worksheet

    candidates: list[Any] = []

    for worksheet in workbook.worksheets:
        title = normalize_text(
            worksheet.title
        )

        if title in {
            main_title,
            DETAIL_TEMPLATE_SHEET_NAME,
            "違規查詢紀錄",
        }:
            continue

        if not re.fullmatch(
            r"\d{1,8}",
            title,
        ):
            continue

        if is_legacy_detail_template_sheet(
            worksheet
        ):
            candidates.append(
                worksheet
            )

    if not candidates:
        return None

    return min(
        candidates,
        key=lambda item: (
            abs(item.max_row - 11),
            item.max_column,
        ),
    )



def copy_external_detail_template(
    source: Any,
    target: Any,
) -> None:
    """
    將來源活頁簿的 155598 格式複製到既有結果活頁簿。

    不直接複製 StyleArray 編號，改為逐項複製字型、底色、
    框線與對齊，避免跨活頁簿 style id 不一致。
    """
    for row in source.iter_rows():
        for source_cell in row:
            target_cell = target.cell(
                source_cell.row,
                source_cell.column,
            )

            target_cell.value = (
                source_cell.value
            )

            if source_cell.has_style:
                target_cell.font = copy(
                    source_cell.font
                )
                target_cell.fill = copy(
                    source_cell.fill
                )
                target_cell.border = copy(
                    source_cell.border
                )
                target_cell.alignment = copy(
                    source_cell.alignment
                )
                target_cell.number_format = (
                    source_cell.number_format
                )
                target_cell.protection = copy(
                    source_cell.protection
                )

            if source_cell.hyperlink:
                target_cell.hyperlink = copy(
                    source_cell.hyperlink
                )

    for row_index, dimension in (
        source.row_dimensions.items()
    ):
        target_dimension = (
            target.row_dimensions[
                row_index
            ]
        )

        target_dimension.height = (
            dimension.height
        )
        target_dimension.hidden = (
            dimension.hidden
        )
        target_dimension.outlineLevel = (
            dimension.outlineLevel
        )
        target_dimension.thickTop = getattr(
            dimension,
            "thickTop",
            False,
        )
        target_dimension.thickBot = getattr(
            dimension,
            "thickBot",
            False,
        )

    for column_letter, dimension in (
        source.column_dimensions.items()
    ):
        target_dimension = (
            target.column_dimensions[
                column_letter
            ]
        )

        target_dimension.width = (
            dimension.width
        )
        target_dimension.hidden = (
            dimension.hidden
        )
        target_dimension.bestFit = (
            dimension.bestFit
        )
        target_dimension.outlineLevel = (
            dimension.outlineLevel
        )

    target.sheet_format = copy(
        source.sheet_format
    )

    target.sheet_properties = copy(
        source.sheet_properties
    )

    target.page_margins = copy(
        source.page_margins
    )

    target.page_setup = copy(
        source.page_setup
    )

    target.print_options = copy(
        source.print_options
    )

    target.freeze_panes = (
        source.freeze_panes
    )

    target.sheet_view.showGridLines = (
        source.sheet_view.showGridLines
    )

    for merged_range in (
        source.merged_cells.ranges
    ):
        target.merge_cells(
            str(merged_range)
        )


def ensure_detail_template(
    workbook: Any,
    main_title: str,
    preferred_title: str,
    source_path: Path | None = None,
) -> str:
    """
    建立可供所有公司複製的隱藏明細範本。

    新輸出檔會優先複製來源 Excel 的 155598 工作表，
    因此字型、底色、框線、列高、頁面設定可直接沿用。
    """
    if (
        DETAIL_TEMPLATE_SHEET_NAME
        in workbook.sheetnames
    ):
        template = workbook[
            DETAIL_TEMPLATE_SHEET_NAME
        ]
        template.sheet_state = "veryHidden"
        return template.title

    candidate = find_detail_template_candidate(
        workbook,
        main_title,
        preferred_title,
    )

    if candidate is not None:
        template = workbook.copy_worksheet(
            candidate
        )
        template.title = (
            DETAIL_TEMPLATE_SHEET_NAME
        )
    else:
        imported = False

        if (
            source_path is not None
            and source_path.exists()
        ):
            source_workbook = None

            try:
                source_workbook = (
                    load_workbook(
                        source_path
                    )
                )

                source_main = (
                    main_title
                    if main_title
                    in source_workbook.sheetnames
                    else source_workbook
                    .worksheets[0]
                    .title
                )

                source_candidate = (
                    find_detail_template_candidate(
                        source_workbook,
                        source_main,
                        preferred_title,
                    )
                )

                if source_candidate is not None:
                    template = (
                        workbook.create_sheet(
                            DETAIL_TEMPLATE_SHEET_NAME
                        )
                    )

                    copy_external_detail_template(
                        source_candidate,
                        template,
                    )

                    imported = True

            finally:
                if source_workbook is not None:
                    source_workbook.close()

        if not imported:
            template = workbook.create_sheet(
                DETAIL_TEMPLATE_SHEET_NAME
            )
            build_fallback_detail_template(
                template
            )

    template.sheet_state = "veryHidden"

    return template.title


def copy_template_row_style(
    template: Any,
    template_row: int,
    target: Any,
    target_row: int,
) -> None:
    for column in range(1, 8):
        source_cell = template.cell(
            template_row,
            column,
        )

        target_cell = target.cell(
            target_row,
            column,
        )

        if source_cell.has_style:
            target_cell._style = copy(
                source_cell._style
            )

        target_cell.number_format = (
            source_cell.number_format
        )

        target_cell.alignment = copy(
            source_cell.alignment
        )

        target_cell.protection = copy(
            source_cell.protection
        )

    source_dimension = (
        template.row_dimensions[
            template_row
        ]
    )

    target_dimension = (
        target.row_dimensions[
            target_row
        ]
    )

    target_dimension.height = (
        source_dimension.height
    )

    target_dimension.hidden = (
        source_dimension.hidden
    )

    target_dimension.outlineLevel = (
        source_dimension.outlineLevel
    )


def display_width_units(value: Any) -> int:
    text = normalize_text(value)

    units = 0

    for character in text:
        units += (
            1
            if ord(character) < 128
            else 2
        )

    return units


def estimate_legacy_row_height(
    row: Sequence[Any],
    kind: str,
) -> float:
    """
    155598 使用約 8.43 欄寬與 6 pt 字型。

    每一行約可容納 14 個顯示寬度單位；
    來源檔列高以 8 pt 為一級：
        16.5、24.5、32.5、40.5、...
    """
    if kind == TABLE_KIND_UNPAID:
        values = [
            row[1],
            row[2],
            row[4],
        ]
    else:
        values = [
            row[1],
            row[2],
            row[3],
            row[4],
            row[5],
        ]

    line_count = 1

    for value in values:
        units = display_width_units(
            value
        )

        estimated = max(
            1,
            (units + 13) // 14,
        )

        line_count = max(
            line_count,
            estimated,
        )

    return max(
        16.5,
        float(
            line_count * 8
            + 0.5
        ),
    )


def clear_copied_detail_sheet(
    worksheet: Any,
) -> None:
    for row in worksheet.iter_rows():
        for cell in row:
            cell.value = None
            cell.hyperlink = None
            cell.comment = None

    worksheet.freeze_panes = None

    try:
        worksheet.auto_filter.ref = None
    except Exception:
        pass


def collect_outcome_rows(
    outcome: QueryOutcome,
    kind: str,
) -> list[list[Any]]:
    rows: list[list[Any]] = []

    seen_tables: set[str] = set()

    for table in outcome.tables:
        if table.kind != kind:
            continue

        signature = hashlib.sha256(
            json.dumps(
                table.rows,
                ensure_ascii=False,
                sort_keys=False,
                default=str,
            ).encode("utf-8")
        ).hexdigest()

        if signature in seen_tables:
            continue

        seen_tables.add(
            signature
        )

        rows.extend(
            [list(row) for row in table.rows]
        )

    return rows


def write_detail_sheet(
    workbook: Any,
    template_sheet_name: str,
    company: CompanyRow,
    outcome: QueryOutcome,
    queried_at: datetime,
) -> None:
    """
    建立與來源 Excel「155598」相同的明細版面。

    輸出規則：
    - 工作表名稱使用 155598，而不是 00155598。
    - 不寫入「統一編號／登記名稱／查詢時間」等中繼資料。
    - 不寫入自製標題列。
    - 未繳／需到案資料放在最前面 A:F。
    - 繳納紀錄接在後面 A:G。
    - 直接沿用隱藏範本的字型、顏色、框線與頁面設定。
    """
    del queried_at

    remove_existing_detail_variants(
        workbook,
        company.raw_id,
        company.query_id,
    )

    template = workbook[
        template_sheet_name
    ]

    worksheet = workbook.copy_worksheet(
        template
    )

    worksheet.title = detail_sheet_name(
        company.raw_id,
        company.query_id,
    )

    worksheet.sheet_state = "visible"

    clear_copied_detail_sheet(
        worksheet
    )

    unpaid_rows = collect_outcome_rows(
        outcome,
        TABLE_KIND_UNPAID,
    )

    paid_rows = collect_outcome_rows(
        outcome,
        TABLE_KIND_PAID,
    )

    total_rows = (
        len(unpaid_rows)
        + len(paid_rows)
    )

    current_row = 1
    normal_unpaid_index = 0
    needs_appearance_index = 0

    for row in unpaid_rows:
        status = normalize_text(
            row[0]
        )

        is_needs_appearance = (
            "需到案" in status
        )

        if is_needs_appearance:
            template_row = 5
            needs_appearance_index += 1
        else:
            template_row = (
                1
                if normal_unpaid_index == 0
                else 2
            )
            normal_unpaid_index += 1

        copy_template_row_style(
            template,
            template_row,
            worksheet,
            current_row,
        )

        canonical = list(row[:7])
        canonical += [None] * (
            7 - len(canonical)
        )

        canonical[6] = None

        for column, value in enumerate(
            canonical,
            start=1,
        ):
            worksheet.cell(
                current_row,
                column,
                value,
            )

        view_cell = worksheet.cell(
            current_row,
            6,
        )

        if not normalize_text(
            view_cell.value
        ):
            view_cell.value = "檢視"

        if outcome.result_url:
            view_cell.hyperlink = (
                outcome.result_url
            )

        worksheet.row_dimensions[
            current_row
        ].height = estimate_legacy_row_height(
            canonical,
            TABLE_KIND_UNPAID,
        )

        current_row += 1

    for paid_index, row in enumerate(
        paid_rows,
    ):
        if paid_index == 0:
            template_row = 6
        elif paid_index == len(paid_rows) - 1:
            template_row = 11
        else:
            template_row = 7

        copy_template_row_style(
            template,
            template_row,
            worksheet,
            current_row,
        )

        canonical = list(row[:7])
        canonical += [None] * (
            7 - len(canonical)
        )

        for column, value in enumerate(
            canonical,
            start=1,
        ):
            worksheet.cell(
                current_row,
                column,
                value,
            )

        worksheet.row_dimensions[
            current_row
        ].height = estimate_legacy_row_height(
            canonical,
            TABLE_KIND_PAID,
        )

        current_row += 1

    if total_rows <= 0:
        workbook.remove(
            worksheet
        )
        return

    if worksheet.max_row > total_rows:
        worksheet.delete_rows(
            total_rows + 1,
            worksheet.max_row - total_rows,
        )

    worksheet.sheet_view.showGridLines = True
def ensure_log_sheet(workbook: Any) -> Any:
    title = "違規查詢紀錄"

    if title in workbook.sheetnames:
        worksheet = workbook[title]
    else:
        worksheet = workbook.create_sheet(title)

        headers = [
            "查詢時間",
            "Excel列",
            "統一編號",
            "登記名稱",
            "狀態",
            "違規筆數",
            "訊息",
            "結果網址",
        ]

        worksheet.append(headers)

        for cell in worksheet[1]:
            cell.fill = TITLE_FILL
            cell.font = WHITE_FONT
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

        worksheet.freeze_panes = "A2"

        widths = [
            20,
            10,
            14,
            32,
            12,
            12,
            50,
            55,
        ]

        for index, width in enumerate(
            widths,
            start=1,
        ):
            worksheet.column_dimensions[
                get_column_letter(index)
            ].width = width

    return worksheet


def append_log(
    worksheet: Any,
    company: CompanyRow,
    outcome: QueryOutcome,
    queried_at: datetime,
) -> None:
    worksheet.append(
        [
            queried_at.strftime(
                "%Y-%m-%d %H:%M:%S"
            ),
            company.excel_row,
            company.query_id,
            company.name,
            outcome.status,
            outcome.record_count,
            outcome.message,
            outcome.result_url,
        ]
    )

    worksheet.cell(
        worksheet.max_row,
        3,
    ).number_format = "@"

    for cell in worksheet[worksheet.max_row]:
        cell.alignment = Alignment(
            vertical="top",
            wrap_text=True,
        )

        cell.border = Border(
            left=THIN_GRAY,
            right=THIN_GRAY,
            top=THIN_GRAY,
            bottom=THIN_GRAY,
        )


def atomic_save(
    workbook: Any,
    output_path: Path,
) -> None:
    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with tempfile.NamedTemporaryFile(
        prefix=f".{output_path.stem}_",
        suffix=".xlsx",
        dir=output_path.parent,
        delete=False,
    ) as handle:
        temp_path = Path(handle.name)

    try:
        workbook.save(temp_path)

        check = load_workbook(
            temp_path,
            read_only=True,
            data_only=False,
        )
        check.close()

        os.replace(
            temp_path,
            output_path,
        )
    finally:
        if temp_path.exists():
            temp_path.unlink(
                missing_ok=True
            )


def locator_is_usable(locator: Locator) -> bool:
    try:
        return (
            locator.count() > 0
            and locator.first.is_visible()
            and locator.first.is_enabled()
        )
    except PlaywrightError:
        return False


def first_usable(
    locators: Iterable[Locator],
) -> Locator:
    for locator in locators:
        if locator_is_usable(locator):
            return locator.first

    raise RuntimeError(
        "找不到可用的網頁欄位；網站版面可能已改版"
    )


def locate_company_id_input(
    page: Page,
) -> Locator:
    label_patterns = [
        re.compile(r"統一編號", re.I),
        re.compile(r"登記編號", re.I),
        re.compile(r"公司.*編號", re.I),
    ]

    candidates: list[Locator] = [
        page.get_by_label(pattern)
        for pattern in label_patterns
    ]

    candidates += [
        page.locator("input[name*='ban' i]"),
        page.locator("input[id*='ban' i]"),
        page.locator("input[name*='tax' i]"),
        page.locator("input[id*='tax' i]"),
        page.locator("input[name*='company' i]"),
        page.locator("input[id*='company' i]"),
    ]

    try:
        return first_usable(candidates)
    except RuntimeError:
        pass

    handle = page.evaluate_handle(
        """
        () => {
          const visible = el => {
            const s = getComputedStyle(el);
            const r = el.getBoundingClientRect();

            return (
              s.visibility !== 'hidden'
              && s.display !== 'none'
              && r.width > 0
              && r.height > 0
            );
          };

          const inputs = [
            ...document.querySelectorAll('input')
          ].filter(el => {
            const type = (
              el.type || 'text'
            ).toLowerCase();

            return (
              visible(el)
              && !el.disabled
              && [
                'text',
                'tel',
                'number',
                ''
              ].includes(type)
            );
          });

          for (const el of inputs) {
            const nearby = [
              el.getAttribute('aria-label') || '',
              el.placeholder || '',
              el.name || '',
              el.id || '',
              el.closest('label')?.innerText || '',
              el.parentElement?.innerText || '',
              el.parentElement
                ?.previousElementSibling
                ?.innerText || ''
            ].join(' ');

            if (
              /(統一編號|登記編號|公司.{0,4}編號)/
                .test(nearby)
              && !/(驗證碼|captcha|verify)/i
                .test(nearby)
            ) {
              return el;
            }
          }

          return null;
        }
        """
    )

    element = handle.as_element()

    if element is None:
        raise RuntimeError(
            "找不到『統一編號／登記編號』輸入欄位"
        )

    return element


def locate_captcha_input(
    page: Page,
) -> Locator:
    candidates: list[Locator] = [
        page.get_by_label(
            re.compile(r"驗證碼", re.I)
        ),
        page.locator(
            "input[name*='captcha' i]"
        ),
        page.locator(
            "input[id*='captcha' i]"
        ),
        page.locator(
            "input[name*='verify' i]"
        ),
        page.locator(
            "input[id*='verify' i]"
        ),
        page.locator(
            "input[name*='valid' i]"
        ),
        page.locator(
            "input[id*='valid' i]"
        ),
    ]

    try:
        return first_usable(candidates)
    except RuntimeError:
        pass

    handle = page.evaluate_handle(
        """
        () => {
          const visible = el => {
            const s = getComputedStyle(el);
            const r = el.getBoundingClientRect();

            return (
              s.visibility !== 'hidden'
              && s.display !== 'none'
              && r.width > 0
              && r.height > 0
            );
          };

          const inputs = [
            ...document.querySelectorAll('input')
          ].filter(el => {
            const type = (
              el.type || 'text'
            ).toLowerCase();

            return (
              visible(el)
              && !el.disabled
              && [
                'text',
                'tel',
                'number',
                ''
              ].includes(type)
            );
          });

          for (const el of inputs) {
            const nearby = [
              el.getAttribute('aria-label') || '',
              el.placeholder || '',
              el.name || '',
              el.id || '',
              el.closest('label')?.innerText || '',
              el.parentElement?.innerText || '',
              el.parentElement
                ?.previousElementSibling
                ?.innerText || ''
            ].join(' ');

            if (
              /(驗證碼|captcha|verify|validation.?code)/i
                .test(nearby)
            ) {
              return el;
            }
          }

          return null;
        }
        """
    )

    element = handle.as_element()

    if element is None:
        raise RuntimeError(
            "找不到驗證碼輸入欄位"
        )

    return element


def locate_captcha_capture_area(
    page: Page,
) -> Locator:
    """
    找到要截圖保存的 CAPTCHA 圖片。

    第一優先使用 CAPTCHA_IMAGE_XPATH，只截取真正的 <img>；
    這可避免整個 td 內的輸入框、標籤及「換一張」文字干擾 OCR。
    若圖片 XPath 因網站微調失效，才使用語意圖片定位；最後才退回
    CAPTCHA_CAPTURE_XPATH 的 td 區塊。
    """
    exact_image = page.locator(
        f"xpath={CAPTCHA_IMAGE_XPATH}"
    )

    try:
        exact_image.first.wait_for(
            state="visible",
            timeout=10000,
        )

        if exact_image.count() > 0:
            return exact_image.first
    except PlaywrightError:
        pass

    image_candidates: list[Locator] = [
        page.get_by_role(
            "img",
            name=re.compile(
                r"驗證碼|驗證用圖|captcha",
                re.I,
            ),
        ),
        page.locator("img[alt*='驗證碼']"),
        page.locator("img[title*='驗證碼']"),
        page.locator("img[src*='captcha' i]"),
        page.locator("img[id*='captcha' i]"),
        page.locator("img[name*='captcha' i]"),
        page.locator("img[src*='verify' i]"),
        page.locator("img[id*='verify' i]"),
        page.locator("img[src*='valid' i]"),
    ]

    for image_candidate in image_candidates:
        try:
            if (
                image_candidate.count() > 0
                and image_candidate.first.is_visible()
            ):
                return image_candidate.first
        except PlaywrightError:
            continue

    all_images = page.locator("img")

    try:
        image_count = all_images.count()
    except PlaywrightError:
        image_count = 0

    for index in range(image_count):
        image = all_images.nth(index)

        try:
            if not image.is_visible():
                continue

            nearby = image.evaluate(
                """
                el => [
                  el.alt || '',
                  el.title || '',
                  el.name || '',
                  el.id || '',
                  el.className || '',
                  el.getAttribute('src') || '',
                  el.parentElement?.innerText || '',
                  el.parentElement
                    ?.previousElementSibling
                    ?.innerText || '',
                  el.closest('tr')?.innerText || ''
                ].join(' ')
                """
            )
        except PlaywrightError:
            continue

        if re.search(
            r"驗證碼|驗證用圖|captcha|verify|validation.?code",
            normalize_text(nearby),
            re.I,
        ):
            return image

    exact_capture_area = page.locator(
        f"xpath={CAPTCHA_CAPTURE_XPATH}"
    )

    if locator_is_usable(exact_capture_area):
        return exact_capture_area.first

    raise RuntimeError(
        "找不到 CAPTCHA 圖片或擷取區塊；"
        "網站版面可能已改版。"
        f"圖片 XPath：{CAPTCHA_IMAGE_XPATH}；"
        f"區塊 XPath：{CAPTCHA_CAPTURE_XPATH}"
    )

def locate_query_button(
    page: Page,
) -> Locator:
    """
    找到監理服務網的查詢按鈕。

    第一優先使用目前頁面的精確 XPath：

        /html/body/table/tbody/tr[2]/td[1]/div[2]/div/div/form/div/a

    該控制項是 <a> 元素，不一定具有標準 button role、文字或
    input[type=submit] 特徵，因此舊版語意定位可能完全找不到。
    精確 XPath 失效時，才使用語意與 DOM 特徵備援。
    """
    exact_button = page.locator(
        f"xpath={QUERY_BUTTON_XPATH}"
    )

    try:
        exact_button.first.wait_for(
            state="attached",
            timeout=10000,
        )

        if exact_button.count() > 0:
            return exact_button.first

    except PlaywrightError:
        pass

    candidates: list[Locator] = [
        page.get_by_role(
            "button",
            name=re.compile(
                r"^(查詢|送出|確定)$"
            ),
        ),
        page.get_by_role(
            "link",
            name=re.compile(
                r"^(查詢|送出|確定)$"
            ),
        ),
        page.locator(
            "input[type='submit']"
            "[value*='查詢']"
        ),
        page.locator(
            "input[type='button']"
            "[value*='查詢']"
        ),
        page.locator(
            "button:has-text('查詢')"
        ),
        page.locator(
            "a:has-text('查詢')"
        ),
        page.locator(
            "form div a"
        ),
    ]

    try:
        return first_usable(candidates)
    except RuntimeError as original_error:
        pass

    # 最後從包含 CAPTCHA 欄位的 form 內找最可能的 <a>／按鈕。
    try:
        handle = page.evaluate_handle(
            """
            () => {
              const visible = el => {
                const style = getComputedStyle(el);
                const rect = el.getBoundingClientRect();

                return (
                  style.visibility !== 'hidden'
                  && style.display !== 'none'
                  && rect.width > 0
                  && rect.height > 0
                );
              };

              const captchaInputs = [
                ...document.querySelectorAll('input')
              ].filter(el => {
                const nearby = [
                  el.getAttribute('aria-label') || '',
                  el.placeholder || '',
                  el.name || '',
                  el.id || '',
                  el.parentElement?.innerText || ''
                ].join(' ');

                return (
                  visible(el)
                  && /(驗證碼|captcha|verify|validation.?code)/i
                    .test(nearby)
                );
              });

              const form = (
                captchaInputs[0]?.closest('form')
                || document.querySelector('form')
              );

              if (!form) {
                return null;
              }

              const controls = [
                ...form.querySelectorAll(
                  'a, button, input[type="submit"], '
                  + 'input[type="button"], [role="button"], [onclick]'
                )
              ].filter(visible);

              for (const el of controls) {
                const descriptor = [
                  el.innerText || '',
                  el.textContent || '',
                  el.getAttribute('value') || '',
                  el.getAttribute('title') || '',
                  el.getAttribute('aria-label') || '',
                  el.id || '',
                  el.className || '',
                  el.getAttribute('href') || '',
                  el.getAttribute('onclick') || ''
                ].join(' ');

                if (/(查詢|送出|確定|query|search|submit)/i.test(descriptor)) {
                  return el;
                }
              }

              // 此頁的查詢按鈕位在 form 直屬 div 裡的 a。
              const directAnchor = form.querySelector(':scope > div > a');

              if (directAnchor && visible(directAnchor)) {
                return directAnchor;
              }

              return controls.find(el => el.tagName === 'A') || null;
            }
            """
        )

        element = handle.as_element()

        if element is not None:
            return element

    except PlaywrightError:
        pass

    raise RuntimeError(
        "找不到查詢按鈕。"
        f"指定 XPath：{QUERY_BUTTON_XPATH}；"
        "網站版面可能已再次改版"
    )


def click_query_button(
    page: Page,
    timeout_ms: int,
) -> str:
    """
    點擊查詢按鈕。

    依序嘗試：
    1. Playwright 一般點擊。
    2. Playwright 強制點擊。
    3. JavaScript element.click()。
    4. 直接用 document.evaluate() 依 XPath 點擊。
    """
    query_button = locate_query_button(
        page
    )

    try:
        query_button.scroll_into_view_if_needed(
            timeout=min(
                timeout_ms,
                5000,
            )
        )
    except (AttributeError, PlaywrightError):
        pass

    errors: list[str] = []

    try:
        query_button.click(
            timeout=min(
                timeout_ms,
                10000,
            ),
        )
        return "精確 XPath／定位器一般點擊"
    except PlaywrightError as exc:
        errors.append(
            f"一般點擊：{exc}"
        )

    try:
        query_button.click(
            timeout=min(
                timeout_ms,
                10000,
            ),
            force=True,
        )
        return "精確 XPath／定位器強制點擊"
    except PlaywrightError as exc:
        errors.append(
            f"強制點擊：{exc}"
        )

    try:
        query_button.evaluate(
            "element => element.click()"
        )
        return "定位器 JavaScript 點擊"
    except Exception as exc:
        errors.append(
            "定位器 JavaScript 點擊："
            f"{type(exc).__name__}: {exc}"
        )

    try:
        clicked = page.evaluate(
            """
            xpath => {
              const element = document.evaluate(
                xpath,
                document,
                null,
                XPathResult.FIRST_ORDERED_NODE_TYPE,
                null
              ).singleNodeValue;

              if (!element) {
                return false;
              }

              element.click();
              return true;
            }
            """,
            QUERY_BUTTON_XPATH,
        )

        if clicked:
            return "document.evaluate() 精確 XPath 點擊"

    except PlaywrightError as exc:
        errors.append(
            f"document.evaluate()：{exc}"
        )

    raise RuntimeError(
        "查詢按鈕已定位，但所有點擊方式均失敗。"
        f"指定 XPath：{QUERY_BUTTON_XPATH}；"
        f"錯誤：{' | '.join(errors)}"
    )


def body_text(page: Page) -> str:
    try:
        return normalize_text(
            page.locator("body").inner_text(
                timeout=5000
            )
        )
    except PlaywrightError:
        return ""


def contains_any(
    text: str,
    patterns: Sequence[str],
) -> bool:
    return any(
        pattern in text
        for pattern in patterns
    )


def is_captcha_related_exception(
    exception: BaseException,
) -> bool:
    """判斷例外是否發生在 CAPTCHA 擷取、辨識、輸入或驗證流程。"""
    if isinstance(
        exception,
        CaptchaRecognitionError,
    ):
        return True

    message = normalize_text(
        f"{type(exception).__name__}: {exception}"
    ).lower()

    captcha_keywords = (
        "captcha",
        "驗證碼",
        "imageinput",
        "ocrimage",
        "englishalphanumeric",
        "四個英數",
        "四碼",
    )

    return any(
        keyword in message
        for keyword in captcha_keywords
    )


def is_transient_page_exception(
    exception: BaseException,
) -> bool:
    """判斷是否為導頁、執行環境切換或暫時性瀏覽器錯誤。"""
    if isinstance(
        exception,
        TransientPageError,
    ):
        return True

    message = normalize_text(
        f"{type(exception).__name__}: {exception}"
    ).lower()

    transient_patterns = (
        "execution context was destroyed",
        "cannot find context",
        "most likely because of a navigation",
        "navigation interrupted",
        "frame was detached",
        "target page, context or browser has been closed",
        "page.goto: timeout",
        "net::err_",
    )

    return any(
        pattern in message
        for pattern in transient_patterns
    )

def wait_for_result_change(
    page: Page,
    before_text: str,
    timeout_ms: int,
) -> None:
    """等待一般頁面或分頁內容改變。"""
    try:
        page.wait_for_function(
            """
            before => {
              const now = (
                document.body?.innerText || ''
              )
                .replace(/\\s+/g, ' ')
                .trim();

              return now !== before;
            }
            """,
            arg=before_text,
            timeout=timeout_ms,
        )
    except PlaywrightTimeoutError:
        pass
    except PlaywrightError:
        # 導頁時 execution context 會被銷毀，後續以 load state 接手。
        pass

    try:
        page.wait_for_load_state(
            "domcontentloaded",
            timeout=min(timeout_ms, 15000),
        )
    except PlaywrightTimeoutError:
        pass
    except PlaywrightError:
        pass

    try:
        page.wait_for_load_state(
            "networkidle",
            timeout=min(timeout_ms, 10000),
        )
    except PlaywrightTimeoutError:
        pass
    except PlaywrightError:
        pass


def read_no_paid_data_message(
    page: Page,
) -> str:
    """
    讀取結果頁「查無已繳納罰鍰資料」訊息。

    第一優先使用使用者提供的精確 XPath；若網站只做小幅版面調整，
    再使用 id/class/文字語意備援。回傳空字串代表尚未確認為零筆結果。
    """
    candidates = (
        page.locator(
            f"xpath={NO_PAID_DATA_XPATH}"
        ),
        page.locator(
            "#headerMessage"
        ),
        page.get_by_text(
            re.compile(
                r"查無已繳納罰鍰資料",
                re.I,
            ),
            exact=False,
        ),
        page.locator(
            "td:has-text('查無已繳納罰鍰資料')"
        ),
    )

    for candidate in candidates:
        try:
            count = candidate.count()
        except PlaywrightError:
            continue

        for index in range(count):
            item = candidate.nth(index)

            try:
                if not item.is_visible():
                    continue

                message = normalize_text(
                    item.inner_text(
                        timeout=3000
                    )
                )
            except PlaywrightError:
                continue

            if contains_any(
                message,
                (
                    "查無已繳納罰鍰資料",
                    "無已繳納罰鍰資料",
                    "查無罰鍰繳納資料",
                ),
            ):
                return message

    return ""


def page_has_query_form(page: Page) -> bool:
    """判斷目前是否仍停留在統一編號／CAPTCHA 查詢頁。"""
    selectors = (
        "form#form2",
        f"xpath={QUERY_BUTTON_XPATH}",
        f"xpath={CURRENT_QUERY_BUTTON_XPATH}",
        "input[name*='captcha' i]",
        "input[id*='captcha' i]",
        "input[name*='verify' i]",
        "input[id*='verify' i]",
    )

    for selector in selectors:
        try:
            if page.locator(selector).count() > 0:
                return True
        except PlaywrightError:
            continue

    return False


def page_has_result_marker(page: Page) -> bool:
    """判斷結果頁的正式容器、表格或訊息是否已出現。"""
    selectors = (
        "form#penaltyQueryPayRecordForm",
        "table#info",
        "#headerMessage",
        "div.cont90 div.caption_std",
        f"xpath={RESULT_JSON_XPATH}",
        f"xpath={NO_PAID_DATA_XPATH}",
        f"xpath={CURRENT_RESULT_CONTAINER_XPATH}",
        f"xpath={CURRENT_RESULT_TABLE_XPATH}",
        f"xpath={CURRENT_RESULT_SECOND_TAB_XPATH}",
    )

    for selector in selectors:
        try:
            if page.locator(selector).count() > 0:
                return True
        except PlaywrightError:
            continue

    return False


def wait_for_query_result_page(
    page: Page,
    before_text: str,
    timeout_ms: int,
) -> str:
    """
    等待查詢結果真正穩定，而不是只等到第一次 DOM 文字變動。

    舊版在導頁剛開始時就繼續讀取，會出現：
    - 尚未顯示「驗證碼輸入錯誤」便開始解析；
    - Page.evaluate 的 execution context 被導頁銷毀；
    - 只取得 735 bytes 的暫時空白頁。
    """
    deadline = time.monotonic() + max(
        timeout_ms,
        1000,
    ) / 1000.0

    last_signature: tuple[str, str] | None = None
    stable_count = 0
    latest_text = ""

    while time.monotonic() < deadline:
        try:
            page.wait_for_load_state(
                "domcontentloaded",
                timeout=1000,
            )
        except (
            PlaywrightTimeoutError,
            PlaywrightError,
        ):
            pass

        latest_text = body_text(page)

        marker_ready = (
            page_has_result_marker(page)
            or bool(
                read_no_paid_data_message(page)
            )
            or contains_any(
                latest_text,
                CAPTCHA_ERROR_PATTERNS,
            )
            or contains_any(
                latest_text,
                FORM_ERROR_PATTERNS,
            )
            or contains_any(
                latest_text,
                NO_DATA_PATTERNS,
            )
        )

        changed = bool(
            latest_text
            and latest_text != before_text
        )

        if marker_ready and changed:
            signature = (
                page.url,
                latest_text,
            )

            if signature == last_signature:
                stable_count += 1
            else:
                last_signature = signature
                stable_count = 1

            if stable_count >= 2:
                return latest_text
        else:
            stable_count = 0
            last_signature = None

        time.sleep(0.25)

    return body_text(page)

def compact_header_text(value: Any) -> str:
    return re.sub(
        r"[\s:：、，,。．·()（）\[\]【】]",
        "",
        normalize_text(value),
    )


def header_index(
    headers: Sequence[str],
    aliases: Sequence[str],
) -> int | None:
    normalized_aliases = {
        compact_header_text(alias)
        for alias in aliases
    }

    for index, header in enumerate(
        headers
    ):
        normalized = compact_header_text(
            header
        )

        # 導覽列或整頁摘要通常非常長，
        # 不可把其中出現的關鍵字當成表頭。
        if (
            not normalized
            or len(normalized) > 16
        ):
            continue

        if normalized in normalized_aliases:
            return index

    return None


def identify_result_header(
    rows: Sequence[Sequence[str]],
) -> tuple[
    str,
    int,
    dict[str, int],
] | None:
    """
    僅接受真正的違規結果表頭。

    paid：
        繳費日期、單號、車號、事由、繳納方式、罰鍰

    unpaid：
        違規日、事由、罰鍰、應到案日

    必須分散在不同儲存格中，避免將網站導覽列、
    整頁文字摘要誤判成明細表格。
    """
    for row_index, row in enumerate(
        rows
    ):
        headers = [
            normalize_text(value)
            for value in row
        ]

        paid_map = {
            "date": header_index(
                headers,
                (
                    "繳費日期",
                    "繳納日期",
                ),
            ),
            "ticket": header_index(
                headers,
                (
                    "單號",
                    "舉發單號",
                ),
            ),
            "plate": header_index(
                headers,
                (
                    "車號",
                    "車牌",
                    "車牌號碼",
                ),
            ),
            "reason": header_index(
                headers,
                (
                    "事由",
                    "違規事由",
                ),
            ),
            "method": header_index(
                headers,
                (
                    "繳納方式",
                    "繳費方式",
                    "處理方式",
                ),
            ),
            "fine": header_index(
                headers,
                (
                    "罰鍰",
                    "罰款",
                    "金額",
                ),
            ),
        }

        if all(
            value is not None
            for value in paid_map.values()
        ):
            distinct = {
                int(value)
                for value in paid_map.values()
                if value is not None
            }

            if len(distinct) == len(
                paid_map
            ):
                return (
                    TABLE_KIND_PAID,
                    row_index,
                    {
                        key: int(value)
                        for key, value
                        in paid_map.items()
                        if value is not None
                    },
                )

        unpaid_map = {
            "date": header_index(
                headers,
                (
                    "違規日",
                    "違規日期",
                ),
            ),
            "reason": header_index(
                headers,
                (
                    "事由",
                    "違規事由",
                ),
            ),
            "fine": header_index(
                headers,
                (
                    "罰鍰",
                    "罰款",
                    "金額",
                ),
            ),
            "due": header_index(
                headers,
                (
                    "應到案日",
                    "應到案日期",
                    "應到案",
                ),
            ),
        }

        if all(
            value is not None
            for value in unpaid_map.values()
        ):
            distinct = {
                int(value)
                for value in unpaid_map.values()
                if value is not None
            }

            if len(distinct) == len(
                unpaid_map
            ):
                return (
                    TABLE_KIND_UNPAID,
                    row_index,
                    {
                        key: int(value)
                        for key, value
                        in unpaid_map.items()
                        if value is not None
                    },
                )

    return None


def row_value(
    row: Sequence[str],
    index: int | None,
) -> str:
    if (
        index is None
        or index < 0
        or index >= len(row)
    ):
        return ""

    return normalize_text(
        row[index]
    )


def normalize_amount(value: Any) -> int | str:
    text = normalize_text(value)

    if not text:
        return ""

    compact = re.sub(
        r"[,，\s元]",
        "",
        text,
    )

    if re.fullmatch(
        r"-?\d+",
        compact,
    ):
        try:
            return int(compact)
        except ValueError:
            pass

    return text


def is_roc_date_text(value: Any) -> bool:
    text = normalize_text(value)

    return bool(
        re.fullmatch(
            r"\d{2,3}\s*年\s*"
            r"\d{1,2}\s*月\s*"
            r"\d{1,2}\s*日",
            text,
        )
    )


def shifted_row_value(
    row: Sequence[str],
    index: int,
    offset: int,
) -> str:
    return row_value(
        row,
        index + offset,
    )


def select_paid_column_offset(
    row: Sequence[str],
    columns: dict[str, int],
) -> int:
    best_offset = 0
    best_score = -1

    for offset in (
        -1,
        0,
        1,
        2,
    ):
        payment_date = shifted_row_value(
            row,
            columns["date"],
            offset,
        )

        ticket_number = shifted_row_value(
            row,
            columns["ticket"],
            offset,
        )

        plate_number = shifted_row_value(
            row,
            columns["plate"],
            offset,
        )

        reason = shifted_row_value(
            row,
            columns["reason"],
            offset,
        )

        amount = normalize_amount(
            shifted_row_value(
                row,
                columns["fine"],
                offset,
            )
        )

        score = 0

        if (
            not payment_date
            or is_roc_date_text(
                payment_date
            )
        ):
            score += 1

        if re.fullmatch(
            r"[A-Za-z0-9-]{5,24}",
            ticket_number,
        ):
            score += 3

        if (
            plate_number
            and len(plate_number) <= 20
        ):
            score += 2

        if reason:
            score += 2

        if amount not in ("", None):
            score += 2

        if score > best_score:
            best_score = score
            best_offset = offset

    return best_offset


def select_unpaid_column_offset(
    row: Sequence[str],
    columns: dict[str, int],
) -> int:
    best_offset = 0
    best_score = -1

    for offset in (
        -1,
        0,
        1,
        2,
    ):
        violation_date = shifted_row_value(
            row,
            columns["date"],
            offset,
        )

        reason = shifted_row_value(
            row,
            columns["reason"],
            offset,
        )

        amount = normalize_amount(
            shifted_row_value(
                row,
                columns["fine"],
                offset,
            )
        )

        due_date = shifted_row_value(
            row,
            columns["due"],
            offset,
        )

        score = 0

        if is_roc_date_text(
            violation_date
        ):
            score += 4

        if reason:
            score += 2

        if amount not in ("", None):
            score += 2

        if (
            not due_date
            or is_roc_date_text(
                due_date
            )
        ):
            score += 2

        if score > best_score:
            best_score = score
            best_offset = offset

    return best_offset


def canonicalize_paid_row(
    row: Sequence[str],
    columns: dict[str, int],
) -> list[Any] | None:
    offset = select_paid_column_offset(
        row,
        columns,
    )

    payment_date = shifted_row_value(
        row,
        columns["date"],
        offset,
    )

    ticket_number = shifted_row_value(
        row,
        columns["ticket"],
        offset,
    )

    plate_number = shifted_row_value(
        row,
        columns["plate"],
        offset,
    )

    reason = shifted_row_value(
        row,
        columns["reason"],
        offset,
    )

    payment_method = shifted_row_value(
        row,
        columns["method"],
        offset,
    )

    fine = normalize_amount(
        shifted_row_value(
            row,
            columns["fine"],
            offset,
        )
    )

    if (
        compact_header_text(payment_date)
        in {
            "繳費日期",
            "繳納日期",
        }
    ):
        return None

    # 來源檔中有少數繳費日期空白的紀錄，
    # 因此以單號、車號、事由及罰鍰判斷資料列。
    if not (
        ticket_number
        and plate_number
        and reason
        and fine not in ("", None)
    ):
        return None

    if (
        payment_date
        and not is_roc_date_text(
            payment_date
        )
    ):
        return None

    return [
        None,
        payment_date or None,
        ticket_number,
        plate_number,
        reason,
        payment_method or None,
        fine,
    ]


def canonicalize_unpaid_row(
    row: Sequence[str],
    columns: dict[str, int],
) -> list[Any] | None:
    offset = select_unpaid_column_offset(
        row,
        columns,
    )

    violation_date = shifted_row_value(
        row,
        columns["date"],
        offset,
    )

    reason = shifted_row_value(
        row,
        columns["reason"],
        offset,
    )

    fine = normalize_amount(
        shifted_row_value(
            row,
            columns["fine"],
            offset,
        )
    )

    due_date = shifted_row_value(
        row,
        columns["due"],
        offset,
    )

    if (
        compact_header_text(
            violation_date
        )
        in {
            "違規日",
            "違規日期",
        }
    ):
        return None

    if not (
        is_roc_date_text(
            violation_date
        )
        and reason
        and fine not in ("", None)
    ):
        return None

    if (
        due_date
        and not is_roc_date_text(
            due_date
        )
    ):
        return None

    combined = " ".join(
        normalize_text(value)
        for value in row
    )

    status = (
        "需到案"
        if "需到案" in combined
        else None
    )

    view_text = next(
        (
            normalize_text(value)
            for value in row
            if "檢視" in normalize_text(
                value
            )
        ),
        "檢視",
    )

    return [
        status,
        violation_date,
        reason,
        fine,
        due_date or None,
        view_text,
        None,
    ]

ENGLISH_MONTH_NUMBERS = {
    "jan": 1,
    "feb": 2,
    "mar": 3,
    "apr": 4,
    "may": 5,
    "jun": 6,
    "jul": 7,
    "aug": 8,
    "sep": 9,
    "oct": 10,
    "nov": 11,
    "dec": 12,
}


def parse_result_datetime(
    value: Any,
) -> datetime | None:
    """
    解析結果 JSON 常見日期格式。

    支援例如：
    - Jul 2, 2026 8:50:26 AM
    - Thu Jul 02 08:50:26 CST 2026
    - 2026/07/02 08:50:26
    - 2026-07-02 08:50:26
    """
    if isinstance(value, datetime):
        return value

    if (
        isinstance(value, (int, float))
        and not isinstance(value, bool)
    ):
        timestamp = float(value)

        if abs(timestamp) >= 10_000_000_000:
            timestamp /= 1000.0

        try:
            return datetime.fromtimestamp(
                timestamp
            )
        except (
            OverflowError,
            OSError,
            ValueError,
        ):
            return None

    text = normalize_text(value)

    if not text:
        return None

    normalized = re.sub(
        r"\s+",
        " ",
        text,
    ).strip()

    for date_format in (
        "%b %d, %Y %I:%M:%S %p",
        "%b %d, %Y %I:%M %p",
        "%b %d, %Y",
        "%Y/%m/%d %H:%M:%S",
        "%Y/%m/%d %H:%M",
        "%Y/%m/%d",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
        "%a %b %d %H:%M:%S %Z %Y",
    ):
        try:
            return datetime.strptime(
                normalized,
                date_format,
            )
        except ValueError:
            continue

    match = re.search(
        r"(?P<month>[A-Za-z]{3,9})\s+"
        r"(?P<day>\d{1,2})(?:,)?\s+"
        r"(?P<year>\d{4})",
        normalized,
    )

    if match:
        month = ENGLISH_MONTH_NUMBERS.get(
            match.group("month")[:3].lower()
        )

        if month is not None:
            try:
                return datetime(
                    int(match.group("year")),
                    month,
                    int(match.group("day")),
                )
            except ValueError:
                return None

    match = re.search(
        r"(?:[A-Za-z]{3}\s+)?"
        r"(?P<month>[A-Za-z]{3,9})\s+"
        r"(?P<day>\d{1,2})\s+"
        r"(?:\d{1,2}:\d{2}(?::\d{2})?\s+)?"
        r"(?:[A-Za-z]{2,5}\s+)?"
        r"(?P<year>\d{4})",
        normalized,
    )

    if match:
        month = ENGLISH_MONTH_NUMBERS.get(
            match.group("month")[:3].lower()
        )

        if month is not None:
            try:
                return datetime(
                    int(match.group("year")),
                    month,
                    int(match.group("day")),
                )
            except ValueError:
                return None

    return None


def format_roc_date(
    value: Any,
) -> str:
    """
    將西元日期轉為來源 Excel 使用的民國日期文字。

    例如：
        Jul 2, 2026 8:50:26 AM
        -> 115年7月2日
    """
    parsed = parse_result_datetime(
        value
    )

    if parsed is None:
        return normalize_text(
            value
        )

    return (
        f"{parsed.year - 1911}年"
        f"{parsed.month}月"
        f"{parsed.day}日"
    )


def first_present_value(
    record: dict[str, Any],
    keys: Sequence[str],
) -> Any:
    for key in keys:
        value = record.get(key)

        if value is None:
            continue

        if isinstance(value, str):
            if normalize_text(value):
                return value
            continue

        return value

    return None


def build_paid_table_from_json_payload(
    payload: Any,
) -> ExtractedTable | None:
    """
    將 hidden input 的 JSON 轉成來源 155598 的 A:G 結構。

    JSON 欄位對應：
    - A：空白
    - B：繳費日期（updateTime）
    - C：單號（vilTicket）
    - D：車號（plateNo）
    - E：事由（vilFact）
    - F：繳納方式（payWay）
    - G：罰鍰（payment，依序備援 penalty、penaltyAmount）
    """
    records: Any = payload

    if isinstance(records, dict):
        for key in (
            "data",
            "records",
            "items",
            "rows",
            "result",
        ):
            candidate = records.get(key)

            if isinstance(candidate, list):
                records = candidate
                break

    if not isinstance(records, list):
        return None

    canonical_rows: list[list[Any]] = []

    for raw_record in records:
        if not isinstance(
            raw_record,
            dict,
        ):
            continue

        payment_date_value = first_present_value(
            raw_record,
            (
                "updateTimeStr",
                "updateTime",
                "paymentDateStr",
                "paymentDate",
                "payDateStr",
                "payDate",
            ),
        )

        ticket_number = normalize_text(
            first_present_value(
                raw_record,
                (
                    "vilTicket",
                    "ticket",
                    "ticketNo",
                ),
            )
        )

        plate_number = normalize_text(
            first_present_value(
                raw_record,
                (
                    "plateNo",
                    "plate",
                    "plateNumber",
                ),
            )
        )

        reason = normalize_text(
            first_present_value(
                raw_record,
                (
                    "vilFact",
                    "reason",
                    "fact",
                ),
            )
        )

        payment_method = normalize_text(
            first_present_value(
                raw_record,
                (
                    "payWay",
                    "paymentMethod",
                    "payMethod",
                ),
            )
        )

        fine = normalize_amount(
            first_present_value(
                raw_record,
                (
                    "payment",
                    "penalty",
                    "penaltyAmount",
                    "amount",
                ),
            )
        )

        if not (
            ticket_number
            and plate_number
            and reason
            and fine not in ("", None)
        ):
            continue

        payment_date = format_roc_date(
            payment_date_value
        )

        canonical_rows.append(
            [
                None,
                payment_date or None,
                ticket_number,
                plate_number,
                reason,
                payment_method or None,
                fine,
            ]
        )

    if not canonical_rows:
        return None

    return ExtractedTable(
        kind=TABLE_KIND_PAID,
        title="罰鍰繳納紀錄",
        headers=[
            "",
            "繳費日期",
            "單號",
            "車號",
            "事由",
            "繳納方式",
            "罰鍰",
        ],
        rows=canonical_rows,
    )


def extract_paid_table_from_hidden_json(
    page: Page,
) -> tuple[bool, ExtractedTable | None]:
    """
    從結果頁 hidden input 讀取完整 JSON。

    回傳：
        (是否找到 JSON input, 解析後的繳納紀錄表)

    找到 input 但 value 為 [] 時，回傳 (True, None)。
    """
    candidates = [
        page.locator(
            f"xpath={RESULT_JSON_XPATH}"
        ),
        page.locator(
            "form#penaltyQueryPayRecordForm input#json"
        ),
        page.locator(
            "form#penaltyQueryPayRecordForm input[name='json']"
        ),
        page.locator(
            "input#json[name='json']"
        ),
    ]

    found_input = False
    last_error: Exception | None = None

    for candidate in candidates:
        try:
            count = candidate.count()
        except PlaywrightError as exc:
            last_error = exc
            continue

        if count <= 0:
            continue

        for index in range(count):
            item = candidate.nth(index)

            try:
                item.wait_for(
                    state="attached",
                    timeout=5000,
                )

                element_id = normalize_text(
                    item.get_attribute("id")
                )

                element_name = normalize_text(
                    item.get_attribute("name")
                )

                if (
                    element_id
                    and element_id != "json"
                    and element_name != "json"
                ):
                    continue

                raw_value = item.get_attribute(
                    "value"
                )

                found_input = True

                if raw_value is None:
                    continue

                stripped = raw_value.strip()

                if not stripped:
                    return True, None

                try:
                    payload = json.loads(
                        stripped
                    )
                except json.JSONDecodeError:
                    # 若取得的是 HTML 原始屬性文字，
                    # 將 &quot; 等 entity 還原後再解析。
                    payload = json.loads(
                        unescape(
                            stripped
                        )
                    )

                return (
                    True,
                    build_paid_table_from_json_payload(
                        payload
                    ),
                )

            except (
                json.JSONDecodeError,
                PlaywrightError,
                TypeError,
                ValueError,
            ) as exc:
                last_error = exc
                continue

    if (
        found_input
        and last_error is not None
    ):
        raise RuntimeError(
            "已找到結果 JSON input，"
            "但無法解析其 value："
            f"{type(last_error).__name__}: "
            f"{last_error}"
        ) from last_error

    return False, None

def extract_visible_tables(
    page: Page,
) -> list[ExtractedTable]:
    """
    從頁面中擷取真正的結果表格。

    原版本只要表格文字含有「交通違規、罰鍰、車牌」等字樣，
    就可能把網站導覽選單當成結果，造成輸出出現 A:P 的巨大雜訊。

    本版本要求表格具有完整且分欄的正式表頭，
    並將資料直接轉成來源 155598 的 A:G 欄位結構。
    """
    payload = page.evaluate(
        """
        () => {
          const norm = value => (
            value || ''
          )
            .replace(/\\s+/g, ' ')
            .trim();

          const visible = element => {
            const style = getComputedStyle(
              element
            );

            const rect = (
              element
              .getBoundingClientRect()
            );

            return (
              style.visibility !== 'hidden'
              && style.display !== 'none'
              && rect.width > 0
              && rect.height > 0
            );
          };

          const directRows = table => {
            const selectors = [
              ':scope > thead > tr',
              ':scope > tbody > tr',
              ':scope > tfoot > tr',
              ':scope > tr'
            ].join(', ');

            return [
              ...table.querySelectorAll(
                selectors
              )
            ].filter(visible);
          };

          const directCells = row => [
            ...row.children
          ]
            .filter(cell => (
              (
                cell.tagName === 'TH'
                || cell.tagName === 'TD'
              )
              && visible(cell)
            ))
            .map(cell => ({
              text: norm(
                cell.innerText
              ),
              header: (
                cell.tagName === 'TH'
              )
            }));

          const previousTitle = table => {
            if (
              table.caption
              && norm(
                table.caption.innerText
              )
            ) {
              return norm(
                table.caption.innerText
              );
            }

            let current = table;

            for (
              let depth = 0;
              depth < 6 && current;
              depth++
            ) {
              let previous = (
                current.previousElementSibling
              );

              while (previous) {
                if (
                  visible(previous)
                  && norm(
                    previous.innerText
                  )
                  && (
                    /^H[1-6]$/.test(
                      previous.tagName
                    )
                    || /(title|header|caption)/i
                      .test(
                        String(
                          previous.className
                          || ''
                        )
                      )
                  )
                ) {
                  return norm(
                    previous.innerText
                  );
                }

                previous = (
                  previous
                  .previousElementSibling
                );
              }

              current = (
                current.parentElement
              );
            }

            return '';
          };

          return [
            ...document.querySelectorAll(
              'table'
            )
          ]
            .filter(visible)
            .map(table => ({
              title: previousTitle(
                table
              ),
              rows: directRows(
                table
              )
                .map(directCells)
                .filter(row => (
                  row.length > 0
                ))
            }))
            .filter(table => (
              table.rows.length > 0
            ));
        }
        """
    )

    extracted: list[
        ExtractedTable
    ] = []

    seen: set[str] = set()

    for raw_table in payload:
        raw_rows = (
            raw_table.get("rows")
            or []
        )

        plain_rows = [
            [
                normalize_text(
                    cell.get("text")
                )
                for cell in row
            ]
            for row in raw_rows
        ]

        identified = identify_result_header(
            plain_rows
        )

        if identified is None:
            continue

        kind, header_row, columns = (
            identified
        )

        canonical_rows: list[
            list[Any]
        ] = []

        for row in plain_rows[
            header_row + 1:
        ]:
            if kind == TABLE_KIND_PAID:
                canonical = (
                    canonicalize_paid_row(
                        row,
                        columns,
                    )
                )
            else:
                canonical = (
                    canonicalize_unpaid_row(
                        row,
                        columns,
                    )
                )

            if canonical is not None:
                canonical_rows.append(
                    canonical
                )

        if not canonical_rows:
            continue

        headers = plain_rows[
            header_row
        ]

        signature = hashlib.sha256(
            json.dumps(
                [
                    kind,
                    headers,
                    canonical_rows,
                ],
                ensure_ascii=False,
                sort_keys=False,
                default=str,
            ).encode("utf-8")
        ).hexdigest()

        if signature in seen:
            continue

        seen.add(
            signature
        )

        extracted.append(
            ExtractedTable(
                kind=kind,
                title=normalize_text(
                    raw_table.get(
                        "title"
                    )
                ),
                headers=headers,
                rows=canonical_rows,
            )
        )

    return extracted

def extract_visible_tables_with_retry(
    page: Page,
    timeout_ms: int,
) -> list[ExtractedTable]:
    """在導頁 execution context 切換期間重試 table extraction。"""
    deadline = time.monotonic() + min(
        max(timeout_ms, 1000),
        15000,
    ) / 1000.0

    last_error: Exception | None = None

    while time.monotonic() < deadline:
        try:
            return extract_visible_tables(page)
        except PlaywrightError as exc:
            last_error = exc

            if not is_transient_page_exception(exc):
                raise

            try:
                page.wait_for_load_state(
                    "domcontentloaded",
                    timeout=1000,
                )
            except (
                PlaywrightTimeoutError,
                PlaywrightError,
            ):
                pass

            time.sleep(0.25)

    if last_error is not None:
        raise TransientPageError(
            "結果頁在導向期間無法穩定讀取表格："
            f"{type(last_error).__name__}: {last_error}"
        ) from last_error

    return []

def is_record_row(
    row: Sequence[str],
) -> bool:
    text = " ".join(
        normalize_text(value)
        for value in row
    )

    if not text:
        return False

    if contains_any(
        text,
        NO_DATA_PATTERNS,
    ):
        return False

    if re.search(
        r"(^|\s)"
        r"(合計|總計|共計|總筆數|資料筆數)"
        r"(\s|$)",
        text,
    ):
        return False

    if sum(
        keyword in text
        for keyword in RESULT_KEYWORDS
    ) >= 2:
        return False

    if re.search(
        r"\d{2,3}\s*年\s*"
        r"\d{1,2}\s*月\s*"
        r"\d{1,2}\s*日",
        text,
    ):
        return True

    if (
        re.search(
            r"\b[A-Z0-9]{6,12}\b",
            text,
        )
        and len(
            [value for value in row if value]
        ) >= 3
    ):
        return True

    if (
        re.search(
            r"(?:^|\s)"
            r"\d{2,6}"
            r"(?:元)?"
            r"(?:\s|$)",
            text,
        )
        and len(
            [value for value in row if value]
        ) >= 3
    ):
        return True

    return False


def find_next_page_control(
    page: Page,
) -> Locator | None:
    candidates = [
        page.get_by_role(
            "link",
            name=re.compile(
                r"^(下一頁|下頁|Next|›|»)$",
                re.I,
            ),
        ),
        page.get_by_role(
            "button",
            name=re.compile(
                r"^(下一頁|下頁|Next|›|»)$",
                re.I,
            ),
        ),
        page.locator(
            "input[type='button']"
            "[value='下一頁']"
        ),
        page.locator(
            "input[type='submit']"
            "[value='下一頁']"
        ),
    ]

    for candidate in candidates:
        try:
            for index in range(
                candidate.count()
            ):
                item = candidate.nth(index)

                if (
                    not item.is_visible()
                    or not item.is_enabled()
                ):
                    continue

                disabled = (
                    item.get_attribute(
                        "aria-disabled"
                    )
                    == "true"
                )

                classes = (
                    item.get_attribute("class")
                    or ""
                ).lower()

                if (
                    disabled
                    or "disabled" in classes
                ):
                    continue

                return item
        except PlaywrightError:
            continue

    return None


def collect_all_result_pages(
    page: Page,
    timeout_ms: int,
    max_pages: int = 100,
) -> list[ExtractedTable]:
    """
    收集結果資料。

    hidden JSON 已包含完整繳納紀錄，因此找到 JSON 後不再點分頁；
    這可避免在已取得完整資料後又導頁，造成 execution context destroyed。
    """
    all_tables: list[ExtractedTable] = []
    seen_page_signatures: set[str] = set()
    seen_table_signatures: set[str] = set()

    for _ in range(max_pages):
        try:
            (
                json_input_found,
                paid_json_table,
            ) = extract_paid_table_from_hidden_json(
                page
            )
        except Exception as exc:
            if is_transient_page_exception(exc):
                raise TransientPageError(
                    "讀取結果 JSON 時頁面仍在導向："
                    f"{type(exc).__name__}: {exc}"
                ) from exc
            raise

        page_tables = extract_visible_tables_with_retry(
            page,
            timeout_ms,
        )

        if json_input_found:
            page_tables = [
                table
                for table in page_tables
                if table.kind != TABLE_KIND_PAID
            ]

            if paid_json_table is not None:
                page_tables.append(
                    paid_json_table
                )

        page_signature = hashlib.sha256(
            json.dumps(
                [
                    {
                        "kind": table.kind,
                        "title": table.title,
                        "headers": table.headers,
                        "rows": table.rows,
                    }
                    for table in page_tables
                ],
                ensure_ascii=False,
                sort_keys=True,
                default=str,
            ).encode("utf-8")
        ).hexdigest()

        if page_signature in seen_page_signatures:
            break

        seen_page_signatures.add(
            page_signature
        )

        for table in page_tables:
            signature = hashlib.sha256(
                json.dumps(
                    [
                        table.kind,
                        table.title,
                        table.headers,
                        table.rows,
                    ],
                    ensure_ascii=False,
                    sort_keys=True,
                    default=str,
                ).encode("utf-8")
            ).hexdigest()

            if signature in seen_table_signatures:
                continue

            seen_table_signatures.add(
                signature
            )
            all_tables.append(table)

        # hidden JSON 是完整紀錄，不需要再操作 pagination。
        if json_input_found:
            break

        next_control = find_next_page_control(page)

        if next_control is None:
            break

        before = body_text(page)

        try:
            next_control.click(
                timeout=5000
            )
            wait_for_result_change(
                page,
                before,
                timeout_ms,
            )
        except PlaywrightError as exc:
            if is_transient_page_exception(exc):
                raise TransientPageError(
                    "結果分頁導向尚未完成："
                    f"{type(exc).__name__}: {exc}"
                ) from exc
            break

    return all_tables

def save_debug_artifacts(
    page: Page,
    debug_dir: Path,
    company_id: str,
    label: str,
) -> None:
    debug_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    stamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    safe_label = re.sub(
        r"[^A-Za-z0-9_-]",
        "_",
        label,
    )

    base = debug_dir / (
        f"{company_id}_"
        f"{stamp}_"
        f"{safe_label}"
    )

    try:
        base.with_suffix(".html").write_text(
            page.content(),
            encoding="utf-8",
        )
    except PlaywrightError:
        pass

    try:
        page.screenshot(
            path=str(
                base.with_suffix(".png")
            ),
            full_page=True,
        )
    except PlaywrightError:
        pass


def normalize_captcha_image_format(
    value: str,
) -> str:
    normalized = normalize_text(value).lower()

    if normalized == "jpeg":
        normalized = "jpg"

    if normalized not in {"png", "jpg"}:
        raise ValueError(
            "驗證碼圖片格式只支援 "
            "png、jpg 或 jpeg"
        )

    return normalized


def save_captcha_image(
    page: Page,
    captcha_dir: Path,
    company: CompanyRow,
    image_format: str,
) -> Path:
    """
    優先使用指定 XPath 擷取目前瀏覽器中的 CAPTCHA 圖片。

    回傳的 PNG/JPG 路徑會交給
    call_english_alphanumeric_image_input()，
    再由該函式呼叫 imageInput(image)。
    """
    normalized_format = (
        normalize_captcha_image_format(
            image_format
        )
    )

    captcha_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    captcha_capture_area = locate_captcha_capture_area(page)

    captcha_capture_area.wait_for(
        state="visible",
        timeout=10000,
    )

    captcha_capture_area.scroll_into_view_if_needed(
        timeout=5000
    )

    stamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S_%f"
    )

    archive_path = captcha_dir / (
        f"{company.query_id}_"
        f"{stamp}_"
        f"captcha.{normalized_format}"
    )

    screenshot_arguments: dict[str, Any] = {
        "path": str(archive_path),
        "timeout": 10000,
        "type": (
            "png"
            if normalized_format == "png"
            else "jpeg"
        ),
    }

    if normalized_format == "jpg":
        screenshot_arguments["quality"] = 95

    captcha_capture_area.screenshot(
        **screenshot_arguments
    )

    if (
        not archive_path.exists()
        or archive_path.stat().st_size == 0
    ):
        raise RuntimeError(
            "驗證碼圖片截圖失敗或檔案為空"
        )

    latest_path = captcha_dir / (
        f"current_captcha."
        f"{normalized_format}"
    )

    try:
        shutil.copy2(
            archive_path,
            latest_path,
        )
    except OSError:
        pass

    return archive_path



def validate_captcha_code(value: Any) -> str:
    """
    驗證 OCR 回傳值。格式不符屬於可重試的 CAPTCHA 辨識失敗，
    不應直接跳出單筆查詢並寫成「錯誤」。
    """
    if not isinstance(value, str):
        raise CaptchaRecognitionError(
            "ocrImage() 必須回傳 str；"
            f"目前回傳型別：{type(value).__name__}"
        )

    code = normalize_text(value)

    if not re.fullmatch(r"[A-Za-z0-9]{4}", code):
        raise CaptchaRecognitionError(
            "ocrImage() 回傳值格式錯誤；"
            "必須恰好是四個英數字元，"
            f"目前收到：{code!r}"
        )

    return code

def call_english_alphanumeric_image_input(
    captcha_image_path: Path,
) -> str:
    """
    將 CAPTCHA 圖片載入為 PIL.Image.Image，直接呼叫：

        imageInput(image)

    englishAlphanumeric.py 內的 imageInput() 應負責：
    1. 顯示外部人工輸入 UI。
    2. 顯示傳入的驗證碼圖片。
    3. 等待使用者輸入並送出。
    4. 關閉 UI。
    5. 回傳四個英數字元字串。

    本函式接收 imageInput() 的回傳值，驗證格式後再回傳給
    acquire_captcha()。
    """
    resolved_image_path = (
        captcha_image_path
        .expanduser()
        .resolve()
    )

    if not resolved_image_path.exists():
        raise RuntimeError(
            "要傳給 imageInput() 的圖片不存在："
            f"{resolved_image_path}"
        )

    if not callable(ocrImage):
        raise RuntimeError(
            "從 englishAlphanumeric 匯入的 imageInput 不是可呼叫函式"
        )

    ui_image: Image.Image | None = None

    print(
        "  呼叫 imageInput(image)",
        flush=True,
    )
    print(
        f"  傳入 CAPTCHA 圖片：{resolved_image_path}",
        flush=True,
    )

    try:
        with Image.open(resolved_image_path) as opened_image:
            opened_image.load()

            # copy() 後，即使 Image.open() 的檔案已關閉，
            # imageInput() 仍可正常使用圖片。
            ui_image = opened_image.copy()

        returned_code = ocrImage(
            ui_image
        )

    except Exception as exc:
        raise RuntimeError(
            "imageInput() 執行失敗："
            f"{type(exc).__name__}: {exc}"
        ) from exc

    finally:
        if ui_image is not None:
            try:
                ui_image.close()
            except Exception:
                pass

    captcha_code = validate_captcha_code(
        returned_code
    )

    print(
        "  imageInput() 已回傳四個英數字元，"
        "準備自動填入網頁驗證碼欄位。",
        flush=True,
    )

    return captcha_code


def acquire_captcha(
    page: Page,
    company: CompanyRow,
    captcha_dir: Path,
    captcha_image_format: str,
) -> str:
    """
    完整 CAPTCHA 人工輸入流程：

    1. 找到網頁驗證碼輸入欄位。
    2. 依 XPath 擷取目前 CAPTCHA 所在的 td 區塊。
    3. 保存成 PNG/JPG。
    4. 將圖片開啟成 PIL.Image.Image。
    5. 呼叫 imageInput(image)。
    6. 接收 imageInput() 回傳的四碼字串。
    7. 將字串回傳給 perform_query() 自動填入網頁。
    """
    captcha_input = locate_captcha_input(page)
    captcha_input.click()

    captcha_path = save_captcha_image(
        page=page,
        captcha_dir=captcha_dir,
        company=company,
        image_format=captcha_image_format,
    )

    normalized_format = (
        normalize_captcha_image_format(
            captcha_image_format
        )
    )

    print(
        f"  CAPTCHA 圖片已提取：{captcha_path}",
        flush=True,
    )

    print(
        "  最新驗證碼圖片固定路徑："
        f"{captcha_dir / ('current_captcha.' + normalized_format)}",
        flush=True,
    )

    return call_english_alphanumeric_image_input(
        captcha_image_path=captcha_path,
    )

def is_login_page(page: Page) -> bool:
    text = body_text(page)

    try:
        has_password = (
            page.locator(
                "input[type='password']"
            ).count()
            > 0
        )
    except PlaywrightError:
        has_password = False

    return (
        has_password
        or contains_any(
            text,
            (
                "會員登入",
                "登入監理服務網",
                "監理服務網會員",
            ),
        )
    )


def ensure_query_form_after_optional_login(
    page: Page,
    query_url: str,
    timeout_ms: int,
    headless: bool,
) -> Locator:
    try:
        return locate_company_id_input(page)
    except RuntimeError as original_error:
        if not is_login_page(page):
            raise

        if headless:
            raise RuntimeError(
                "網站要求會員登入，"
                "但目前使用 --headless；"
                "請移除 --headless 後重新執行"
            ) from original_error

        print(
            "網站目前要求會員登入。"
            "請在瀏覽器完成登入，"
            "完成後回到此視窗按 Enter。",
            flush=True,
        )

        try:
            input()
        except EOFError as exc:
            raise RuntimeError(
                "無法等待人工登入"
            ) from exc

        page.goto(
            query_url,
            wait_until="domcontentloaded",
            timeout=timeout_ms,
        )

        try:
            page.wait_for_load_state(
                "networkidle",
                timeout=min(
                    timeout_ms,
                    15000,
                ),
            )
        except PlaywrightTimeoutError:
            pass

        return locate_company_id_input(page)



def click_locator_with_fallback(
    page: Page,
    locator: Locator,
    exact_xpath: str,
    timeout_ms: int,
    label: str,
) -> str:
    """使用一般、強制與 JavaScript 三層方式點擊指定控制項。"""
    errors: list[str] = []

    try:
        locator.scroll_into_view_if_needed(
            timeout=min(timeout_ms, 5000)
        )
    except PlaywrightError:
        pass

    try:
        locator.click(
            timeout=min(timeout_ms, 10000)
        )
        return f"{label}一般點擊"
    except PlaywrightError as exc:
        errors.append(f"一般點擊：{exc}")

    try:
        locator.click(
            timeout=min(timeout_ms, 10000),
            force=True,
        )
        return f"{label}強制點擊"
    except PlaywrightError as exc:
        errors.append(f"強制點擊：{exc}")

    try:
        locator.evaluate(
            "element => element.click()"
        )
        return f"{label}JavaScript 點擊"
    except Exception as exc:
        errors.append(
            "定位器 JavaScript 點擊："
            f"{type(exc).__name__}: {exc}"
        )

    try:
        clicked = page.evaluate(
            """
            xpath => {
              const element = document.evaluate(
                xpath,
                document,
                null,
                XPathResult.FIRST_ORDERED_NODE_TYPE,
                null
              ).singleNodeValue;

              if (!element) {
                return false;
              }

              element.click();
              return true;
            }
            """,
            exact_xpath,
        )

        if clicked:
            return f"{label}document.evaluate() 點擊"
    except PlaywrightError as exc:
        errors.append(f"document.evaluate()：{exc}")

    raise RuntimeError(
        f"{label}已定位，但所有點擊方式均失敗；"
        f"XPath：{exact_xpath}；"
        f"錯誤：{' | '.join(errors)}"
    )


def locate_current_query_corporate_tab(
    page: Page,
) -> Locator | None:
    """定位第一階段查詢頁的「法人」頁籤。"""
    exact = page.locator(
        f"xpath={CURRENT_QUERY_CORPORATE_TAB_XPATH}"
    )

    try:
        if exact.count() > 0:
            return exact.first
    except PlaywrightError:
        pass

    candidates = (
        page.get_by_role(
            "img",
            name=re.compile(r"法人", re.I),
        ),
        page.locator(
            "img[alt*='法人']"
        ),
        page.locator(
            "img[title*='法人']"
        ),
        page.locator(
            "td:has-text('法人') span img"
        ),
    )

    for candidate in candidates:
        try:
            if candidate.count() > 0:
                return candidate.first
        except PlaywrightError:
            continue

    return None


def locate_current_query_button(
    page: Page,
) -> Locator:
    """定位第一階段法人查詢表單的查詢按鈕。"""
    exact = page.locator(
        f"xpath={CURRENT_QUERY_BUTTON_XPATH}"
    )

    try:
        exact.first.wait_for(
            state="attached",
            timeout=5000,
        )
        if exact.count() > 0:
            return exact.first
    except PlaywrightError:
        pass

    # 只在第一階段指定的 form 範圍內備援定位，避免誤抓到
    # 原本「法人交通違規繳費紀錄查詢」頁面的查詢按鈕。
    scoped = page.locator(
        "xpath=/html/body/table/tbody/tr[2]/td[1]/div[3]/"
        "div/div[2]/form//a[contains(normalize-space(.), '查詢')]"
    )

    try:
        if scoped.count() > 0:
            return scoped.first
    except PlaywrightError:
        pass

    raise RuntimeError(
        "找不到交通違規（含強制險）法人查詢按鈕；"
        f"XPath：{CURRENT_QUERY_BUTTON_XPATH}"
    )


def open_current_query_corporate_form(
    page: Page,
    entry_url: str,
    timeout_ms: int,
) -> None:
    """
    開啟「交通違規（含強制險）查詢及繳納」法人表單。

    第一個網址依使用者指定使用原本 legal 入口；若頁面沒有法人頁籤，
    再改用監理服務網正式 penaltyQueryPay 路徑。
    """
    urls: list[str] = []

    for candidate_url in (
        entry_url,
        CURRENT_QUERY_DIRECT_FALLBACK_URL,
    ):
        if candidate_url and candidate_url not in urls:
            urls.append(candidate_url)

    errors: list[str] = []

    for candidate_url in urls:
        try:
            page.goto(
                candidate_url,
                wait_until="domcontentloaded",
                timeout=timeout_ms,
            )

            try:
                page.wait_for_load_state(
                    "networkidle",
                    timeout=min(timeout_ms, 12000),
                )
            except PlaywrightTimeoutError:
                pass

            # 若已經是法人表單，不必重複點頁籤。
            try:
                locate_current_query_button(page)
                locate_company_id_input(page)
                return
            except RuntimeError:
                pass

            corporate_tab = (
                locate_current_query_corporate_tab(
                    page
                )
            )

            if corporate_tab is None:
                raise RuntimeError(
                    "找不到交通違規（含強制險）查詢頁的法人頁籤"
                )

            method = click_locator_with_fallback(
                page=page,
                locator=corporate_tab,
                exact_xpath=(
                    CURRENT_QUERY_CORPORATE_TAB_XPATH
                ),
                timeout_ms=timeout_ms,
                label="法人頁籤",
            )

            print(
                f"  第一階段法人頁籤點擊方式：{method}",
                flush=True,
            )

            locate_current_query_button(page)
            locate_company_id_input(page)
            return

        except Exception as exc:
            errors.append(
                f"{candidate_url}："
                f"{type(exc).__name__}: {exc}"
            )

    raise RuntimeError(
        "無法開啟交通違規（含強制險）法人查詢表單；"
        + " | ".join(errors)
    )


def click_current_query_button(
    page: Page,
    timeout_ms: int,
) -> str:
    button = locate_current_query_button(page)
    return click_locator_with_fallback(
        page=page,
        locator=button,
        exact_xpath=CURRENT_QUERY_BUTTON_XPATH,
        timeout_ms=timeout_ms,
        label="交通違規（含強制險）查詢按鈕",
    )


def table_rows_from_locator(
    table_locator: Locator,
) -> list[list[dict[str, Any]]]:
    """只讀取目標 table 的直屬列與儲存格，避免抓到外層導覽表格。"""
    return table_locator.evaluate(
        """
        table => {
          const norm = value => (value || '')
            .replace(/\\s+/g, ' ')
            .trim();

          const rows = [
            ...table.querySelectorAll(
              ':scope > thead > tr, '
              + ':scope > tbody > tr, '
              + ':scope > tfoot > tr, '
              + ':scope > tr'
            )
          ];

          return rows.map(row => [
            ...row.children
          ]
            .filter(cell => (
              cell.tagName === 'TH'
              || cell.tagName === 'TD'
            ))
            .map(cell => ({
              text: norm(cell.innerText),
              header: cell.tagName === 'TH'
            })))
            .filter(row => row.length > 0);
        }
        """
    )


def add_current_category_to_row(
    row: list[Any],
    category: str,
) -> list[Any]:
    canonical = list(row[:7])
    canonical += [None] * (7 - len(canonical))

    existing_status = normalize_text(
        canonical[0]
    )

    canonical[0] = (
        f"{category}／{existing_status}"
        if existing_status
        else category
    )

    return canonical


def extract_current_table_at_xpath(
    page: Page,
    category: str,
) -> ExtractedTable | None:
    """擷取目前結果頁籤的精確表格並轉成 155598 的 A:G 結構。"""
    locator = page.locator(
        f"xpath={CURRENT_RESULT_TABLE_XPATH}"
    )

    try:
        if locator.count() <= 0:
            return None

        locator.first.wait_for(
            state="visible",
            timeout=5000,
        )

        payload = table_rows_from_locator(
            locator.first
        )
    except PlaywrightError:
        return None

    plain_rows = [
        [
            normalize_text(
                cell.get("text")
            )
            for cell in row
        ]
        for row in payload
    ]

    combined_text = " ".join(
        " ".join(row)
        for row in plain_rows
    )

    if contains_any(
        combined_text,
        NO_DATA_PATTERNS,
    ):
        return None

    identified = identify_result_header(
        plain_rows
    )

    canonical_rows: list[list[Any]] = []
    headers: list[str] = []

    if identified is not None:
        kind, header_row, columns = identified
        headers = plain_rows[header_row]

        for row in plain_rows[header_row + 1:]:
            canonical = (
                canonicalize_unpaid_row(
                    row,
                    columns,
                )
                if kind == TABLE_KIND_UNPAID
                else None
            )

            if canonical is None:
                continue

            canonical_rows.append(
                add_current_category_to_row(
                    canonical,
                    category,
                )
            )
    else:
        # 網站若只更換表頭名稱，仍保留每一列的原始文字，避免整張表遺失。
        data_start = 0
        for index, raw_row in enumerate(payload):
            if any(
                bool(cell.get("header"))
                for cell in raw_row
            ):
                headers = plain_rows[index]
                data_start = index + 1
                break

        for row in plain_rows[data_start:]:
            if not any(row):
                continue

            row_text = " ".join(row)
            if contains_any(
                row_text,
                NO_DATA_PATTERNS,
            ):
                continue

            raw_values: list[Any] = list(row[:7])
            raw_values += [None] * (
                7 - len(raw_values)
            )

            raw_values[0] = (
                f"{category}／{normalize_text(raw_values[0])}"
                if normalize_text(raw_values[0])
                else category
            )

            canonical_rows.append(
                raw_values
            )

    if not canonical_rows:
        return None

    return ExtractedTable(
        kind=TABLE_KIND_UNPAID,
        title=category,
        headers=headers,
        rows=canonical_rows,
    )


def current_result_table_text(
    page: Page,
) -> str:
    locator = page.locator(
        f"xpath={CURRENT_RESULT_TABLE_XPATH}"
    )

    try:
        if locator.count() <= 0:
            return ""
        return normalize_text(
            locator.first.inner_text(
                timeout=3000
            )
        )
    except PlaywrightError:
        return ""


def wait_for_current_table_change(
    page: Page,
    before_text: str,
    timeout_ms: int,
) -> bool:
    deadline = time.monotonic() + min(
        max(timeout_ms, 1000),
        12000,
    ) / 1000.0

    while time.monotonic() < deadline:
        now = current_result_table_text(page)
        body = body_text(page)

        if (
            now != before_text
            or "查無不可線上繳納" in body
            or "無不可線上繳納" in body
        ):
            return True

        time.sleep(0.25)

    return False


def collect_current_violation_tables(
    page: Page,
    timeout_ms: int,
) -> list[ExtractedTable]:
    """依序收集可線上繳納及第二頁籤的表格。"""
    tables: list[ExtractedTable] = []

    online_table = extract_current_table_at_xpath(
        page,
        CURRENT_ONLINE_CATEGORY,
    )

    if online_table is not None:
        tables.append(online_table)

    before_table_text = current_result_table_text(
        page
    )

    second_tab = page.locator(
        f"xpath={CURRENT_RESULT_SECOND_TAB_XPATH}"
    )

    try:
        second_tab_exists = second_tab.count() > 0
    except PlaywrightError:
        second_tab_exists = False

    if second_tab_exists:
        method = click_locator_with_fallback(
            page=page,
            locator=second_tab.first,
            exact_xpath=(
                CURRENT_RESULT_SECOND_TAB_XPATH
            ),
            timeout_ms=timeout_ms,
            label="不可線上繳納頁籤",
        )

        print(
            f"  第一階段第二頁籤點擊方式：{method}",
            flush=True,
        )

        changed = wait_for_current_table_change(
            page,
            before_table_text,
            timeout_ms,
        )

        offline_table = extract_current_table_at_xpath(
            page,
            CURRENT_OFFLINE_CATEGORY,
        )

        # 若內容沒有切換，避免把第一頁籤資料重複寫入兩次。
        if (
            offline_table is not None
            and (
                changed
                or current_result_table_text(page)
                != before_table_text
            )
        ):
            tables.append(offline_table)

    return tables


def perform_current_violation_query(
    page: Page,
    company: CompanyRow,
    entry_url: str,
    timeout_ms: int,
    max_captcha_attempts: int,
    debug_dir: Path,
    captcha_dir: Path,
    captcha_image_format: str,
) -> QueryOutcome:
    """第一階段：查詢尚未結案的交通違規及強制險違規。"""
    last_retry_message = ""

    for attempt in range(
        1,
        max_captcha_attempts + 1,
    ):
        try:
            open_current_query_corporate_form(
                page=page,
                entry_url=entry_url,
                timeout_ms=timeout_ms,
            )

            company_input = locate_company_id_input(
                page
            )
            company_input.fill(
                company.query_id
            )

            before = body_text(page)

            try:
                captcha_code = acquire_captcha(
                    page=page,
                    company=company,
                    captcha_dir=captcha_dir,
                    captcha_image_format=(
                        captcha_image_format
                    ),
                )
            except Exception as exc:
                if not is_captcha_related_exception(exc):
                    raise

                last_retry_message = (
                    "第一階段 CAPTCHA 辨識未取得四碼："
                    f"{type(exc).__name__}: {exc}"
                )
                print(
                    "  第一階段 CAPTCHA 辨識失敗，將重新查詢"
                    f"（{attempt}/{max_captcha_attempts}）。",
                    flush=True,
                )
                continue

            locate_captcha_input(page).fill(
                captcha_code
            )

            submit_method = click_current_query_button(
                page,
                timeout_ms,
            )

            print(
                f"  第一階段查詢按鈕點擊方式：{submit_method}",
                flush=True,
            )

            text = wait_for_query_result_page(
                page,
                before,
                timeout_ms,
            )

            if contains_any(
                text,
                CAPTCHA_ERROR_PATTERNS,
            ):
                last_retry_message = (
                    "第一階段網站回報驗證碼輸入錯誤"
                )
                continue

            if contains_any(
                text,
                FORM_ERROR_PATTERNS,
            ):
                message = next(
                    pattern
                    for pattern in FORM_ERROR_PATTERNS
                    if pattern in text
                )

                if message in TRANSIENT_FORM_ERROR_PATTERNS:
                    last_retry_message = (
                        f"第一階段網站暫時性回應：{message}"
                    )
                    continue

                return QueryOutcome(
                    STATUS_ERROR,
                    0,
                    [],
                    f"交通違規（含強制險）：{message}",
                    page.url,
                )

            tables = collect_current_violation_tables(
                page,
                timeout_ms,
            )

            count = sum(
                table.record_count
                for table in tables
            )

            online_count = sum(
                table.record_count
                for table in tables
                if table.title
                == CURRENT_ONLINE_CATEGORY
            )

            offline_count = sum(
                table.record_count
                for table in tables
                if table.title
                == CURRENT_OFFLINE_CATEGORY
            )

            if count > 0:
                return QueryOutcome(
                    STATUS_SUCCESS,
                    count,
                    tables,
                    (
                        "交通違規（含強制險）查詢完成："
                        f"可線上繳納 {online_count} 筆；"
                        f"不可線上繳納 {offline_count} 筆"
                    ),
                    page.url,
                )

            text = body_text(page)

            if contains_any(
                text,
                NO_DATA_PATTERNS,
            ):
                return QueryOutcome(
                    STATUS_SUCCESS,
                    0,
                    [],
                    "交通違規（含強制險）查詢完成：0 筆",
                    page.url,
                )

            current_result_exists = False
            for xpath in (
                CURRENT_RESULT_CONTAINER_XPATH,
                CURRENT_RESULT_TABLE_XPATH,
                CURRENT_RESULT_SECOND_TAB_XPATH,
            ):
                try:
                    if page.locator(
                        f"xpath={xpath}"
                    ).count() > 0:
                        current_result_exists = True
                        break
                except PlaywrightError:
                    continue

            if current_result_exists:
                save_debug_artifacts(
                    page,
                    debug_dir,
                    company.query_id,
                    "current_query_zero_result",
                )
                return QueryOutcome(
                    STATUS_SUCCESS,
                    0,
                    [],
                    (
                        "交通違規（含強制險）查詢完成；"
                        "結果頁無可辨識資料，依 0 筆處理"
                    ),
                    page.url,
                )

            last_retry_message = (
                "第一階段結果頁尚未完整載入"
            )

        except Exception as exc:
            if (
                is_captcha_related_exception(exc)
                or is_transient_page_exception(exc)
            ):
                last_retry_message = (
                    "第一階段可重試錯誤："
                    f"{type(exc).__name__}: {exc}"
                )
                continue
            raise

    return QueryOutcome(
        STATUS_RETRY,
        0,
        [],
        (
            last_retry_message
            or (
                "交通違規（含強制險）查詢連續無法完成 "
                f"{max_captcha_attempts} 次"
            )
        ),
        page.url,
        retry_later=True,
    )


def combine_query_outcomes(
    current_outcome: QueryOutcome,
    paid_outcome: QueryOutcome,
) -> QueryOutcome:
    """將第一階段未結案違規與第二階段繳費紀錄合併。"""
    completed_statuses = {
        STATUS_SUCCESS,
        STATUS_NO_DATA,
    }

    if (
        current_outcome.status == STATUS_ERROR
        or paid_outcome.status == STATUS_ERROR
    ):
        return QueryOutcome(
            STATUS_ERROR,
            0,
            [],
            (
                "雙階段查詢失敗；"
                f"第一階段：{current_outcome.message}；"
                f"第二階段：{paid_outcome.message}"
            ),
            current_outcome.result_url
            or paid_outcome.result_url,
        )

    if (
        current_outcome.status not in completed_statuses
        or paid_outcome.status not in completed_statuses
    ):
        return QueryOutcome(
            STATUS_RETRY,
            0,
            [],
            (
                "雙階段查詢尚未全部完成；"
                f"第一階段：{current_outcome.message}；"
                f"第二階段：{paid_outcome.message}"
            ),
            current_outcome.result_url
            or paid_outcome.result_url,
            retry_later=True,
        )

    tables = (
        list(current_outcome.tables)
        + list(paid_outcome.tables)
    )

    total_count = (
        current_outcome.record_count
        + paid_outcome.record_count
    )

    return QueryOutcome(
        STATUS_SUCCESS,
        total_count,
        tables,
        (
            f"{COMBINED_QUERY_MESSAGE_MARKER}；"
            f"第一階段 {current_outcome.record_count} 筆；"
            f"第二階段 {paid_outcome.record_count} 筆；"
            f"合計 {total_count} 筆"
        ),
        current_outcome.result_url
        or paid_outcome.result_url,
    )


def perform_paid_record_query(
    page: Page,
    company: CompanyRow,
    query_url: str,
    timeout_ms: int,
    max_captcha_attempts: int,
    debug_dir: Path,
    captcha_dir: Path,
    captcha_image_format: str,
    headless: bool,
) -> QueryOutcome:
    """
    查詢單一公司。

    OCR 格式不符、網站回報 CAPTCHA 錯誤、導頁 context 切換及暫時空白頁
    都在單筆的嘗試迴圈內重新查詢；只有真正不可恢復的欄位／資料錯誤
    才回傳 STATUS_ERROR。
    """
    last_retry_message = ""

    for attempt in range(
        1,
        max_captcha_attempts + 1,
    ):
        try:
            page.goto(
                query_url,
                wait_until="domcontentloaded",
                timeout=timeout_ms,
            )

            try:
                page.wait_for_load_state(
                    "networkidle",
                    timeout=min(
                        timeout_ms,
                        15000,
                    ),
                )
            except PlaywrightTimeoutError:
                pass

            company_input = (
                ensure_query_form_after_optional_login(
                    page,
                    query_url,
                    timeout_ms,
                    headless,
                )
            )

            company_input.fill(
                company.query_id
            )

            before = body_text(page)

            try:
                captcha_code = acquire_captcha(
                    page=page,
                    company=company,
                    captcha_dir=captcha_dir,
                    captcha_image_format=(
                        captcha_image_format
                    ),
                )
            except Exception as exc:
                if not is_captcha_related_exception(exc):
                    raise

                last_retry_message = (
                    f"CAPTCHA 辨識未取得四碼："
                    f"{type(exc).__name__}: {exc}"
                )

                print(
                    "  CAPTCHA OCR 結果不符合四碼格式，"
                    "將換一張並重新查詢"
                    f"（{attempt}/{max_captcha_attempts}）。",
                    flush=True,
                )
                continue

            captcha_input = locate_captcha_input(
                page
            )
            captcha_input.fill(
                captcha_code
            )

            submit_method = click_query_button(
                page=page,
                timeout_ms=timeout_ms,
            )

            print(
                f"  查詢按鈕點擊方式：{submit_method}",
                flush=True,
            )

            text = wait_for_query_result_page(
                page,
                before,
                timeout_ms,
            )

            if contains_any(
                text,
                CAPTCHA_ERROR_PATTERNS,
            ):
                last_retry_message = (
                    "網站回報驗證碼輸入錯誤"
                )

                print(
                    "  網站回報驗證碼錯誤，"
                    "將重新取得圖片並再次辨識"
                    f"（{attempt}/{max_captcha_attempts}）。",
                    flush=True,
                )
                continue

            if contains_any(
                text,
                FORM_ERROR_PATTERNS,
            ):
                message = next(
                    pattern
                    for pattern in FORM_ERROR_PATTERNS
                    if pattern in text
                )

                if message in TRANSIENT_FORM_ERROR_PATTERNS:
                    last_retry_message = (
                        f"網站暫時性回應：{message}"
                    )
                    print(
                        f"  {message}，將重新查詢"
                        f"（{attempt}/{max_captcha_attempts}）。",
                        flush=True,
                    )
                    continue

                save_debug_artifacts(
                    page,
                    debug_dir,
                    company.query_id,
                    "form_error",
                )

                return QueryOutcome(
                    STATUS_ERROR,
                    0,
                    [],
                    message,
                    page.url,
                )

            try:
                tables = collect_all_result_pages(
                    page,
                    timeout_ms=timeout_ms,
                )
            except Exception as exc:
                if not is_transient_page_exception(exc):
                    raise

                last_retry_message = (
                    "結果頁導向尚未完成："
                    f"{type(exc).__name__}: {exc}"
                )

                print(
                    "  結果頁仍在導向，將整筆重新查詢"
                    f"（{attempt}/{max_captcha_attempts}）。",
                    flush=True,
                )
                continue

            # 表格解析後重新讀取頁面文字。舊版使用解析前的舊 text，
            # 因此後來才顯示的 CAPTCHA 錯誤或無資料訊息會被漏掉。
            text = body_text(page)

            if contains_any(
                text,
                CAPTCHA_ERROR_PATTERNS,
            ):
                last_retry_message = (
                    "網站回報驗證碼輸入錯誤"
                )
                print(
                    "  結果解析時確認為驗證碼錯誤，"
                    "將重新查詢"
                    f"（{attempt}/{max_captcha_attempts}）。",
                    flush=True,
                )
                continue

            count = sum(
                table.record_count
                for table in tables
            )

            if count > 0:
                unpaid_count = sum(
                    table.record_count
                    for table in tables
                    if table.kind
                    == TABLE_KIND_UNPAID
                )

                paid_count = sum(
                    table.record_count
                    for table in tables
                    if table.kind
                    == TABLE_KIND_PAID
                )

                return QueryOutcome(
                    STATUS_SUCCESS,
                    count,
                    tables,
                    (
                        f"未繳／需到案 {unpaid_count} 筆；"
                        f"繳納紀錄 {paid_count} 筆；"
                        f"合計 {count} 筆"
                    ),
                    page.url,
                )

            # 精確 XPath 出現時，網站已正常完成查詢，只是沒有已繳資料。
            # 依需求：交通違規筆數寫 0，違規查詢狀態寫「成功」。
            no_paid_message = (
                read_no_paid_data_message(
                    page
                )
            )

            if no_paid_message:
                return QueryOutcome(
                    STATUS_SUCCESS,
                    0,
                    [],
                    no_paid_message,
                    page.url,
                )

            if contains_any(
                text,
                NO_DATA_PATTERNS,
            ):
                return QueryOutcome(
                    STATUS_SUCCESS,
                    0,
                    [],
                    "查詢成功，交通違規筆數為 0",
                    page.url,
                )

            # 若仍停在 CAPTCHA 表單或頁面尚未完整載入，不可誤判為 0 筆；
            # 應重新取得 CAPTCHA 並重做這筆查詢。
            if page_has_query_form(page):
                last_retry_message = (
                    "查詢送出後仍停留在 CAPTCHA 查詢頁，"
                    "可能是驗證碼訊息尚未完成顯示"
                )
                print(
                    "  查詢後仍停留在驗證碼頁面，"
                    "將重新查詢"
                    f"（{attempt}/{max_captcha_attempts}）。",
                    flush=True,
                )
                continue

            if not page_has_result_marker(page):
                last_retry_message = (
                    "結果頁尚未完整載入或只取得暫時空白頁"
                )
                print(
                    "  結果頁尚未完整載入，將重新查詢"
                    f"（{attempt}/{max_captcha_attempts}）。",
                    flush=True,
                )
                continue

            # 依需求：如果結果頁已完整載入，但找不到可辨識的違規明細
            # 表格，仍視為查詢成功且筆數為 0。保留除錯檔供日後確認網站
            # 結構是否改版，但不再把主表狀態標成「錯誤」。
            save_debug_artifacts(
                page,
                debug_dir,
                company.query_id,
                "unrecognized_zero_result",
            )

            return QueryOutcome(
                STATUS_SUCCESS,
                0,
                [],
                (
                    "查詢成功；找不到可辨識的違規明細表格，"
                    "依 0 筆處理；已保存除錯 HTML/截圖"
                ),
                page.url,
            )

        except Exception as exc:
            if (
                is_captcha_related_exception(exc)
                or is_transient_page_exception(exc)
            ):
                last_retry_message = (
                    f"{type(exc).__name__}: {exc}"
                )
                print(
                    "  遇到可重試的暫時錯誤，將重新查詢"
                    f"（{attempt}/{max_captcha_attempts}）："
                    f"{last_retry_message}",
                    flush=True,
                )
                continue
            raise

    return QueryOutcome(
        STATUS_RETRY,
        0,
        [],
        (
            last_retry_message
            or (
                "CAPTCHA／結果頁連續無法完成 "
                f"{max_captcha_attempts} 次"
            )
        ),
        page.url,
        retry_later=True,
    )

def should_skip_row(
    workbook: Any,
    worksheet: Any,
    company: CompanyRow,
    tracking_columns: dict[str, int],
    resume: bool,
) -> bool:
    if not resume:
        return False

    status = normalize_text(
        worksheet.cell(
            company.excel_row,
            tracking_columns[
                "違規查詢狀態"
            ],
        ).value
    )

    # 舊版的「無資料」也屬於已完成的零筆查詢；新版不再產生此狀態，
    # 但仍保留續跑相容性。
    if status == STATUS_NO_DATA:
        return True

    if status != STATUS_SUCCESS:
        return False

    query_message = normalize_text(
        worksheet.cell(
            company.excel_row,
            tracking_columns[
                "違規查詢訊息"
            ],
        ).value
    )

    # 舊版只做繳費紀錄查詢；沒有雙階段標記時必須重新處理，
    # 才能補上可線上與不可線上繳納的目前違規資料。
    if (
        COMBINED_QUERY_MESSAGE_MARKER
        not in query_message
    ):
        return False

    record_count_value = worksheet.cell(
        company.excel_row,
        tracking_columns[
            "交通違規筆數"
        ],
    ).value

    try:
        record_count = int(
            record_count_value
        )
    except (
        TypeError,
        ValueError,
    ):
        record_count = None

    # 成功且 0 筆時本來就不會建立明細工作表，續跑應直接略過。
    if record_count == 0:
        return True

    expected_title = detail_sheet_name(
        company.raw_id,
        company.query_id,
    )

    if expected_title not in workbook.sheetnames:
        # 舊版可能建立成 00155598；
        # 新版要求使用 155598，因此必須重做。
        return False

    detail_sheet = workbook[
        expected_title
    ]

    if detail_sheet.max_column > 7:
        # 舊版 A:P 大型頁面導覽雜訊。
        return False

    first_values = {
        normalize_text(
            detail_sheet.cell(
                row,
                column,
            ).value
        )
        for row in range(
            1,
            min(
                detail_sheet.max_row,
                8,
            ) + 1,
        )
        for column in range(
            1,
            min(
                detail_sheet.max_column,
                7,
            ) + 1,
        )
    }

    if first_values.intersection(
        {
            "統一編號",
            "登記名稱",
            "查詢狀態",
            "違規筆數",
            "查詢時間",
            "來源網址",
        }
    ):
        return False

    return True
def update_main_row(
    worksheet: Any,
    company: CompanyRow,
    outcome: QueryOutcome,
    queried_at: datetime,
    tracking_columns: dict[str, int],
) -> None:
    values = {
        "交通違規筆數": (
            outcome.record_count
            if outcome.status in {
                STATUS_SUCCESS,
                STATUS_NO_DATA,
            }
            else None
        ),
        "違規查詢狀態": outcome.status,
        "最後查詢時間": queried_at,
        "違規查詢訊息": outcome.message,
    }

    style_source = worksheet.cell(
        company.excel_row,
        tracking_columns["交通違規筆數"],
    )

    for header, value in values.items():
        cell = worksheet.cell(
            company.excel_row,
            tracking_columns[header],
            value,
        )

        if header != "交通違規筆數":
            cell.fill = copy(
                style_source.fill
            )
            cell.font = copy(
                style_source.font
            )

        cell.alignment = Alignment(
            vertical="top",
            wrap_text=True,
        )

        cell.border = Border(
            left=THIN_GRAY,
            right=THIN_GRAY,
            top=THIN_GRAY,
            bottom=THIN_GRAY,
        )

    worksheet.cell(
        company.excel_row,
        tracking_columns[
            "最後查詢時間"
        ],
    ).number_format = (
        "yyyy-mm-dd hh:mm:ss"
    )


def launch_browser(
    playwright: Any,
    headless: bool,
) -> Browser:
    errors: list[Exception] = []

    try:
        return playwright.chromium.launch(
            headless=headless
        )
    except PlaywrightError as exc:
        errors.append(exc)

    try:
        return playwright.chromium.launch(
            channel="chrome",
            headless=headless,
        )
    except PlaywrightError as exc:
        errors.append(exc)

    for executable in (
        shutil.which("chromium"),
        shutil.which("chromium-browser"),
        shutil.which("google-chrome"),
        shutil.which("google-chrome-stable"),
    ):
        if not executable:
            continue

        try:
            return playwright.chromium.launch(
                executable_path=executable,
                headless=headless,
                args=["--no-sandbox"],
            )
        except PlaywrightError as exc:
            errors.append(exc)

    raise RuntimeError(
        "無法啟動瀏覽器。請執行："
        "python -m playwright install chromium"
    ) from (
        errors[-1]
        if errors
        else None
    )



def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "依 Excel 統一編號先查交通違規（含強制險），"
            "再查法人交通違規繳費紀錄；"
            "結果優先從 hidden JSON 讀取，"
            "明細工作表沿用來源 155598 格式。"
        )
    )

    parser.add_argument(
        "input",
        nargs="?",
        type=Path,
        default=DEFAULT_INPUT_PATH,
        help=(
            "來源 Excel；預設為專案根目錄下的 "
            r"Data\高雄市.xlsx"
        ),
    )

    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help=(
            "輸出 Excel；預設輸出到 results "
            "資料夾，檔名為 "
            "<來源檔名>_違規明細查詢結果.xlsx"
        ),
    )

    parser.add_argument(
        "--detail-template-sheet",
        default=DEFAULT_DETAIL_TEMPLATE_SHEET,
        help=(
            "來源 Excel 中作為明細格式範本的工作表；"
            "預設 155598"
        ),
    )

    parser.add_argument(
        "--current-query-url",
        default=DEFAULT_CURRENT_QUERY_URL,
        help=(
            "第一階段交通違規（含強制險）入口網址；"
            "預設先使用原本 legal 入口，找不到時自動改用 penaltyQueryPay"
        ),
    )

    parser.add_argument(
        "--url",
        default=DEFAULT_URL,
        help="查詢網址",
    )

    parser.add_argument(
        "--start-row",
        type=int,
        default=None,
        help="從指定 Excel 列開始",
    )

    parser.add_argument(
        "--end-row",
        type=int,
        default=None,
        help="查到指定 Excel 列為止",
    )

    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="最多處理幾筆，適合先小量測試",
    )

    parser.add_argument(
        "--headless",
        action="store_true",
        help=(
            "Playwright 使用無頭模式；"
            "englishAlphanumeric.py 的人工 UI "
            "仍會在本機顯示"
        ),
    )

    parser.add_argument(
        "--resume",
        action=argparse.BooleanOptionalAction,
        default=True,
        help=(
            "輸出檔存在時續跑，"
            "略過狀態為成功／無資料的列"
        ),
    )

    parser.add_argument(
        "--overwrite-output",
        action="store_true",
        help=(
            "若輸出檔已存在，"
            "刪除並從來源檔重新開始"
        ),
    )

    parser.add_argument(
        "--timeout",
        type=int,
        default=30000,
        help="一般網頁逾時毫秒",
    )

    parser.add_argument(
        "--max-captcha-attempts",
        type=int,
        default=5,
        help=(
            "同一輪處理單筆資料時，最多重新取得 CAPTCHA "
            "並呼叫 imageInput() 的次數"
        ),
    )

    parser.add_argument(
        "--max-captcha-requeues",
        type=int,
        default=1,
        help=(
            "同一筆資料耗盡 CAPTCHA 嘗試次數後，"
            "最多移到工作佇列尾端重做幾次；預設 1 次"
        ),
    )

    parser.add_argument(
        "--delay-min",
        type=float,
        default=1.5,
        help="每筆查詢後最短等待秒數",
    )

    parser.add_argument(
        "--delay-max",
        type=float,
        default=3.0,
        help="每筆查詢後最長等待秒數",
    )

    parser.add_argument(
        "--captcha-dir",
        type=Path,
        default=DEFAULT_CAPTCHA_DIR,
        help=(
            "保存每次驗證碼圖片的資料夾；"
            "另會更新 current_captcha.png "
            "或 current_captcha.jpg"
        ),
    )

    parser.add_argument(
        "--captcha-image-format",
        choices=(
            "png",
            "jpg",
            "jpeg",
        ),
        default="png",
        help=(
            "傳入 imageInput() "
            "前保存的圖片格式；預設 png"
        ),
    )

    parser.add_argument(
        "--debug-dir",
        type=Path,
        default=DEFAULT_DEBUG_DIR,
        help=(
            "無法辨識頁面時保存 "
            "HTML 與截圖的資料夾"
        ),
    )

    return parser

def main(
    argv: Sequence[str] | None = None,
) -> int:
    args = build_parser().parse_args(argv)

    input_path = (
        args.input
        .expanduser()
        .resolve()
    )

    results_dir = (
        DEFAULT_RESULTS_DIR
        .expanduser()
        .resolve()
    )

    if not input_path.exists():
        raise SystemExit(
            f"找不到來源檔：{input_path}"
        )

    if input_path.suffix.lower() != ".xlsx":
        raise SystemExit(
            "目前只支援 .xlsx"
        )

    if (
        args.delay_min < 0
        or args.delay_max < args.delay_min
    ):
        raise SystemExit(
            "delay-min／delay-max 設定錯誤"
        )

    if args.max_captcha_attempts <= 0:
        raise SystemExit(
            "max-captcha-attempts 必須大於 0"
        )

    if args.max_captcha_requeues < 0:
        raise SystemExit(
            "max-captcha-requeues 不可小於 0"
        )

    try:
        normalize_captcha_image_format(
            args.captcha_image_format
        )
    except ValueError as exc:
        raise SystemExit(
            str(exc)
        ) from exc

    if not callable(ocrImage):
        raise SystemExit(
            "from englishAlphanumeric import imageInput "
            "匯入的物件不是可呼叫函式"
        )

    output_path = (
        args.output
        .expanduser()
        .resolve()
        if args.output
        else (
            results_dir
            / (
                f"{input_path.stem}_"
                "違規明細查詢結果.xlsx"
            )
        )
    )

    if output_path == input_path:
        raise SystemExit(
            "輸出檔不能與來源檔相同；"
            "本程式不會覆寫原始 Excel"
        )

    new_output = (
        not output_path.exists()
        or args.overwrite_output
    )

    if (
        args.overwrite_output
        and output_path.exists()
    ):
        output_path.unlink()

    if new_output:
        output_path.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        shutil.copy2(
            input_path,
            output_path,
        )

        print(
            f"已建立新輸出檔：{output_path}"
        )
    else:
        print(
            "偵測到既有輸出檔，將續跑："
            f"{output_path}"
        )

    workbook = load_workbook(output_path)

    (
        main_sheet,
        header_start_row,
        data_start_row,
        id_col,
        name_col,
    ) = find_main_sheet_and_columns(
        workbook
    )

    detail_template_sheet_name = (
        ensure_detail_template(
            workbook,
            main_sheet.title,
            args.detail_template_sheet,
            input_path,
        )
    )

    if new_output:
        clear_stale_numeric_detail_sheets(
            workbook,
            main_sheet.title,
        )

    tracking_columns = ensure_tracking_columns(
        main_sheet,
        header_start_row,
    )

    if new_output:
        reset_tracking_values(
            main_sheet,
            data_start_row,
            tracking_columns,
        )

    log_sheet = ensure_log_sheet(workbook)

    companies = iter_companies(
        main_sheet,
        data_start_row,
        id_col,
        name_col,
        args.start_row,
        args.end_row,
        args.limit,
    )

    if not companies:
        raise SystemExit(
            "指定範圍內沒有可用的"
            "統一編號／登記編號"
        )

    captcha_dir = (
        args.captcha_dir
        .expanduser()
        .resolve()
    )

    debug_dir = (
        args.debug_dir
        .expanduser()
        .resolve()
    )

    print(
        f"主工作表：{main_sheet.title}；"
        f"ID 欄：{get_column_letter(id_col)}；"
        f"名稱欄：{get_column_letter(name_col)}；"
        f"待掃描：{len(companies)} 筆。"
    )

    print(
        f"來源 Excel：{input_path}"
    )

    print(
        f"輸出 Excel：{output_path}"
    )

    print(
        "CAPTCHA 圖片 XPath："
        f"{CAPTCHA_IMAGE_XPATH}"
    )

    print(
        "CAPTCHA 擷取區塊備援 XPath："
        f"{CAPTCHA_CAPTURE_XPATH}"
    )

    print(
        "第一階段法人頁籤 XPath："
        f"{CURRENT_QUERY_CORPORATE_TAB_XPATH}"
    )

    print(
        "第一階段查詢按鈕 XPath："
        f"{CURRENT_QUERY_BUTTON_XPATH}"
    )

    print(
        "第一階段結果表格 XPath："
        f"{CURRENT_RESULT_TABLE_XPATH}"
    )

    print(
        "第一階段第二頁籤 XPath："
        f"{CURRENT_RESULT_SECOND_TAB_XPATH}"
    )

    print(
        "查詢按鈕 XPath："
        f"{QUERY_BUTTON_XPATH}"
    )

    print(
        "結果 JSON XPath："
        f"{RESULT_JSON_XPATH}"
    )

    print(
        "查無已繳納罰鍰資料 XPath："
        f"{NO_PAID_DATA_XPATH}"
    )

    print(
        "違規明細格式範本："
        f"{args.detail_template_sheet} "
        f"→ {detail_template_sheet_name}"
    )

    print(
        "人工驗證碼副程式："
        "imageInput(image)"
    )

    print(
        f"驗證碼圖片資料夾：{captcha_dir}"
    )

    processed = 0
    success = 0
    zero_result_success = 0
    no_data = 0
    errors = 0
    pending = 0
    skipped = 0
    deferred = 0
    queue_runs = 0

    # 每個項目為：(公司資料, 已移到佇列尾端的次數)。
    # CAPTCHA 失敗時 append 回 deque 尾端，其他公司會先被處理。
    work_queue = deque(
        (company, 0)
        for company in companies
    )

    with sync_playwright() as playwright:
        browser = launch_browser(
            playwright,
            args.headless,
        )

        context = browser.new_context(
            locale="zh-TW",
            viewport={
                "width": 1440,
                "height": 1000,
            },
            ignore_https_errors=False,
        )

        page = context.new_page()
        page.set_default_timeout(
            args.timeout
        )

        try:
            while work_queue:
                company, requeue_count = (
                    work_queue.popleft()
                )

                if should_skip_row(
                    workbook,
                    main_sheet,
                    company,
                    tracking_columns,
                    args.resume,
                ):
                    skipped += 1
                    print(
                        f"[已完成 {processed}/"
                        f"{len(companies)}] "
                        f"{company.query_id} "
                        f"{company.name}："
                        "先前已成功／無資料，略過。"
                    )
                    continue

                queue_runs += 1
                retry_label = (
                    ""
                    if requeue_count == 0
                    else (
                        "；隊尾重試 "
                        f"{requeue_count}/"
                        f"{args.max_captcha_requeues}"
                    )
                )

                print(
                    f"[執行 {queue_runs}；"
                    f"已完成 {processed}/"
                    f"{len(companies)}；"
                    f"佇列尚有 {len(work_queue)}] "
                    f"查詢 {company.query_id} "
                    f"{company.name}"
                    f"{retry_label}",
                    flush=True,
                )

                queried_at = datetime.now()

                try:
                    print(
                        "  [階段 1/2] 交通違規（含強制險）查詢及繳納",
                        flush=True,
                    )

                    current_outcome = (
                        perform_current_violation_query(
                            page=page,
                            company=company,
                            entry_url=(
                                args.current_query_url
                            ),
                            timeout_ms=args.timeout,
                            max_captcha_attempts=(
                                args.max_captcha_attempts
                            ),
                            debug_dir=debug_dir,
                            captcha_dir=captcha_dir,
                            captcha_image_format=(
                                args.captcha_image_format
                            ),
                        )
                    )

                    print(
                        "  [階段 2/2] 法人交通違規繳費紀錄查詢",
                        flush=True,
                    )

                    paid_outcome = (
                        perform_paid_record_query(
                            page=page,
                            company=company,
                            query_url=args.url,
                            timeout_ms=args.timeout,
                            max_captcha_attempts=(
                                args.max_captcha_attempts
                            ),
                            debug_dir=debug_dir,
                            captcha_dir=captcha_dir,
                            captcha_image_format=(
                                args.captcha_image_format
                            ),
                            headless=args.headless,
                        )
                    )

                    outcome = combine_query_outcomes(
                        current_outcome,
                        paid_outcome,
                    )

                except KeyboardInterrupt:
                    print(
                        "\n使用者中止。"
                        "已保存目前進度。"
                    )

                    atomic_save(
                        workbook,
                        output_path,
                    )

                    return 130

                except Exception as exc:
                    save_debug_artifacts(
                        page,
                        debug_dir,
                        company.query_id,
                        "exception",
                    )

                    retryable_exception = (
                        is_captcha_related_exception(exc)
                        or is_transient_page_exception(exc)
                    )

                    outcome = QueryOutcome(
                        (
                            STATUS_RETRY
                            if retryable_exception
                            else STATUS_ERROR
                        ),
                        0,
                        [],
                        (
                            f"{type(exc).__name__}: "
                            f"{exc}"
                        ),
                        page.url,
                        retry_later=retryable_exception,
                    )

                if outcome.retry_later:
                    if (
                        requeue_count
                        < args.max_captcha_requeues
                    ):
                        next_requeue_count = (
                            requeue_count + 1
                        )

                        deferred_outcome = QueryOutcome(
                            STATUS_RETRY,
                            0,
                            [],
                            (
                                f"{outcome.message}；"
                                "已移到工作佇列尾端，"
                                "其他資料完成後將自動重做"
                                f"（隊尾重試 "
                                f"{next_requeue_count}/"
                                f"{args.max_captcha_requeues}）"
                            ),
                            outcome.result_url,
                            retry_later=True,
                        )

                        update_main_row(
                            main_sheet,
                            company,
                            deferred_outcome,
                            queried_at,
                            tracking_columns,
                        )

                        append_log(
                            log_sheet,
                            company,
                            deferred_outcome,
                            queried_at,
                        )

                        work_queue.append(
                            (
                                company,
                                next_requeue_count,
                            )
                        )

                        atomic_save(
                            workbook,
                            output_path,
                        )

                        deferred += 1

                        print(
                            "  → 待重試；"
                            f"{company.query_id} 已排到佇列尾端；"
                            f"目前佇列共有 {len(work_queue)} 筆。",
                            flush=True,
                        )

                        if work_queue:
                            time.sleep(
                                random.uniform(
                                    args.delay_min,
                                    args.delay_max,
                                )
                            )

                        continue

                    outcome = QueryOutcome(
                        STATUS_RETRY,
                        0,
                        [],
                        (
                            f"{outcome.message}；"
                            "已達本次執行允許的隊尾重試上限 "
                            f"{args.max_captcha_requeues} 次；"
                            "保留為待重試，下一次使用 --resume 執行時會自動重做"
                        ),
                        outcome.result_url,
                        retry_later=False,
                    )

                update_main_row(
                    main_sheet,
                    company,
                    outcome,
                    queried_at,
                    tracking_columns,
                )

                if outcome.status == STATUS_SUCCESS:
                    if outcome.record_count > 0:
                        write_detail_sheet(
                            workbook,
                            detail_template_sheet_name,
                            company,
                            outcome,
                            queried_at,
                        )
                    else:
                        # 零筆成功不應保留舊的明細工作表。
                        remove_existing_detail_variants(
                            workbook,
                            company.raw_id,
                            company.query_id,
                        )
                        zero_result_success += 1

                    success += 1

                elif outcome.status == STATUS_NO_DATA:
                    remove_existing_detail_variants(
                        workbook,
                        company.raw_id,
                        company.query_id,
                    )
                    no_data += 1

                elif outcome.status == STATUS_RETRY:
                    pending += 1

                else:
                    errors += 1

                append_log(
                    log_sheet,
                    company,
                    outcome,
                    queried_at,
                )

                atomic_save(
                    workbook,
                    output_path,
                )

                processed += 1

                print(
                    f"  → {outcome.status}；"
                    f"違規筆數="
                    f"{outcome.record_count}；"
                    f"{outcome.message}"
                )

                if work_queue:
                    time.sleep(
                        random.uniform(
                            args.delay_min,
                            args.delay_max,
                        )
                    )

        finally:
            context.close()
            browser.close()

    print(
        f"完成。最終處理 {processed} 筆："
        f"成功 {success}"
        f"（其中零筆 {zero_result_success}）、"
        f"舊版無資料 {no_data}、"
        f"待重試 {pending}、"
        f"錯誤 {errors}；"
        f"略過 {skipped}、"
        f"曾移至隊尾 {deferred} 次。"
    )

    print(
        f"輸出檔：{output_path}"
    )

    return (
        0
        if errors == 0
        else 2
    )


if __name__ == "__main__":
    raise SystemExit(main())
