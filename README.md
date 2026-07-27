# MVDIS FineTrace

> 以 Excel 批次查詢臺灣監理服務網法人交通違規資料，整合 Playwright、英數 OCR、結果解析、Excel 明細輸出與中斷續跑。

## 專案簡介

`MVDIS FineTrace` 是一套以 Python 開發的監理服務網法人交通違規查詢工具。

程式會從既有 Excel 讀取「統一編號／登記編號」與登記名稱，透過 Playwright 操作監理服務網，擷取驗證碼圖片並交由 `ddddocr` 辨識，再將查詢結果整理回新的 Excel 檔案。

此專案同時提供獨立 OCR API 與批次 OCR 測試程式，方便驗證圖片前處理、模型模式與英數辨識結果。

## 主要功能

- 從 Excel 自動尋找統一編號與登記名稱欄位。
- 使用 Playwright 操作監理服務網法人違規查詢頁面。
- 優先擷取 CAPTCHA 圖片本身，並提供 XPath 與語意定位備援。
- 使用 `ddddocr` 辨識指定長度的 ASCII 英數字元。
- 提供原圖、灰階、高對比、銳化、二值化與反相圖片候選。
- 對過寬圖片進行背景裁切與水平滑動裁切。
- 支援 default、beta、auto 與 both OCR 模型模式。
- 解析未繳、需到案及罰鍰繳納紀錄。
- 優先從結果頁 hidden JSON 取得完整繳納紀錄。
- 排除網站導覽表格與重複結果。
- 沿用來源 Excel 的 `155598` 工作表格式輸出 A:G 明細。
- 更新查詢筆數、查詢狀態、查詢時間與查詢訊息。
- 每筆處理後以原子寫入方式保存進度。
- 支援中斷續跑、CAPTCHA 重試與工作佇列尾端重排。
- 將正常零筆結果記錄為「成功」。
- 發生頁面解析問題時保存 HTML 與完整頁面截圖。

## 專案結構

```text
MVDIS-FineTrace/
├─ Data/
│  └─ 高雄市.xlsx
├─ src/
│  ├─ mvdisCrawl.py
│  ├─ englishAlphanumericOcrApi.py
│  └─ testOcr.py
├─ results/
│  ├─ mvdis_captcha/
│  ├─ mvdis_debug/
│  └─ 高雄市_違規明細查詢結果.xlsx
└─ README.md
```

## 程式說明

### `src/mvdisCrawl.py`

主要批次查詢程式，負責：

1. 讀取來源 Excel。
2. 找出公司統一編號與名稱。
3. 開啟監理服務網。
4. 擷取並辨識 CAPTCHA。
5. 填入表單並送出查詢。
6. 解析 visible table 與 hidden JSON。
7. 寫入主表、查詢紀錄與個別違規明細工作表。
8. 保存進度並處理續跑與重試。

### `src/englishAlphanumericOcrApi.py`

提供唯一公開函式：

```python
from PIL import Image
from englishAlphanumericOcrApi import ocrImage

result: str = ocrImage(image)
```

輸入必須是 `PIL.Image.Image`，回傳值為 OCR 字串。

### `src/testOcr.py`

批次讀取 `results/mvdis_captcha` 內的圖片，逐張輸出：

- 圖片檔名
- 圖片尺寸與模式
- OCR 結果
- 格式判定
- 單張辨識耗時
- 總耗時與平均耗時

## 執行環境

- Python 3.10 以上
- Windows、Linux 或 macOS
- Chromium 或 Google Chrome

主要 Python 套件：

- `playwright`
- `openpyxl`
- `Pillow`
- `ddddocr`

## 安裝

### 1. 建立虛擬環境

PowerShell：

```powershell
python -m venv .venv
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
.\.venv\Scripts\Activate.ps1
```

### 2. 安裝 Python 套件

```powershell
python -m pip install --upgrade pip
python -m pip install --upgrade ddddocr Pillow openpyxl playwright
```

### 3. 安裝 Playwright Chromium

```powershell
python -m playwright install chromium
```

## 輸入 Excel 規格

程式會在活頁簿前幾列中，自動尋找下列欄位：

- `統一編號` 或 `登記編號`
- `登記名稱`、`公司名稱`、`商業名稱` 或 `名稱`

預設輸入檔案：

```text
Data\高雄市.xlsx
```

來源 Excel 若包含名稱為 `155598` 的明細工作表，程式會優先將其作為輸出格式範本；若不存在，則建立相容的備援格式。

## 快速開始

在專案根目錄執行：

```powershell
python .\src\mvdisCrawl.py
```

程式預設讀取：

```text
Data\高雄市.xlsx
```

並輸出至：

```text
results\高雄市_違規明細查詢結果.xlsx
```

## 常用執行方式

### 指定來源 Excel

```powershell
python .\src\mvdisCrawl.py ".\Data\高雄市.xlsx"
```

### 指定輸出檔案

```powershell
python .\src\mvdisCrawl.py `
  ".\Data\高雄市.xlsx" `
  --output ".\results\高雄市_違規明細查詢結果.xlsx"
```

### 先測試前 5 筆

```powershell
python .\src\mvdisCrawl.py --limit 5
```

### 指定 Excel 列範圍

```powershell
python .\src\mvdisCrawl.py --start-row 2 --end-row 100
```

### 使用無頭模式

```powershell
python .\src\mvdisCrawl.py --headless
```

網站若要求人工登入，請不要使用 `--headless`。

### 刪除既有輸出並重新執行

```powershell
python .\src\mvdisCrawl.py --overwrite-output
```

### 關閉續跑

```powershell
python .\src\mvdisCrawl.py --no-resume
```

### 調整 CAPTCHA 嘗試次數

```powershell
python .\src\mvdisCrawl.py `
  --max-captcha-attempts 5 `
  --max-captcha-requeues 1
```

### 調整每筆查詢間隔

```powershell
python .\src\mvdisCrawl.py --delay-min 1.5 --delay-max 3.0
```

## OCR 測試

### 測試全部 CAPTCHA 圖片

```powershell
python .\src\testOcr.py
```

### 只測試前 5 張

```powershell
python .\src\testOcr.py --limit 5
```

### 比較一般模型與 beta 模型

```powershell
python .\src\testOcr.py --model-mode both
```

### 指定圖片資料夾

```powershell
python .\src\testOcr.py `
  --input-dir ".\results\mvdis_captcha"
```

### 指定預期字元數

```powershell
python .\src\testOcr.py --expected-length 4
```

## OCR 模型模式

| 模式 | 說明 |
|---|---|
| `auto` | 先使用一般模型，未得到預期長度時再使用 beta 模型 |
| `default` | 只使用一般模型 |
| `beta` | 只使用 beta 模型 |
| `both` | 同時執行一般模型與 beta 模型並比較候選結果 |

## OCR 環境變數

| 環境變數 | 預設值 | 用途 |
|---|---:|---|
| `OCR_DDDDOCR_MODEL` | `auto` | 設定 OCR 模型模式 |
| `OCR_CAPTCHA_LENGTH` | `4` | 設定預期字元數 |
| `OCR_ALLOWED_CHARACTERS` | 英文大小寫與數字 | 限制可辨識字元 |
| `OCR_MAX_IMAGE_PIXELS` | `40000000` | 限制圖片總像素數 |
| `OCR_MAX_IMAGE_SIDE` | `4096` | 限制放大後最長邊 |

PowerShell 範例：

```powershell
$env:OCR_DDDDOCR_MODEL = "auto"
$env:OCR_CAPTCHA_LENGTH = "4"
python .\src\testOcr.py
```

## 輸出內容

### 主工作表新增或更新欄位

- `交通違規筆數`
- `違規查詢狀態`
- `最後查詢時間`
- `違規查詢訊息`

### 查詢紀錄工作表

程式會建立 `違規查詢紀錄`，包含：

- 查詢時間
- Excel 列號
- 統一編號
- 登記名稱
- 查詢狀態
- 違規筆數
- 查詢訊息
- 結果網址

### 公司明細工作表

有違規紀錄時，程式會依公司登記編號建立個別工作表：

- 未繳／需到案資料使用 A:F
- 繳納紀錄使用 A:G
- 不加入額外標題與中繼資料
- 沿用 `155598` 範本的字型、框線、底色、列高與頁面設定

### CAPTCHA 圖片

```text
results\mvdis_captcha\
```

每次擷取的圖片會以時間戳記保存，並同步更新：

```text
current_captcha.png
```

或：

```text
current_captcha.jpg
```

### 除錯資料

```text
results\mvdis_debug\
```

當頁面無法辨識、結果結構異常或發生例外時，程式會保存：

- HTML
- 完整頁面 PNG 截圖

## 續跑與錯誤處理

預設啟用 `--resume`。

以下資料會被視為已完成並略過：

- 狀態為成功且違規筆數大於 0，並存在有效明細工作表。
- 狀態為成功且違規筆數為 0。
- 舊版狀態為無資料的紀錄。

可重試情況包括：

- OCR 沒有取得指定長度英數字元。
- 網站回報 CAPTCHA 錯誤。
- 頁面導向期間 execution context 被銷毀。
- 結果頁尚未完整載入。
- 系統忙碌、查詢失敗或操作逾時。

單筆資料耗盡 CAPTCHA 嘗試次數後，可移至工作佇列尾端，讓其他公司先完成，再回頭重新查詢。

## 已知注意事項

1. 監理服務網若更改 HTML、XPath、表單欄位或結果 JSON 結構，定位與解析規則可能需要同步更新。
2. OCR 結果必須符合指定長度的 ASCII 英數格式，否則會重新取得 CAPTCHA。
3. 第一次執行 OCR 時，需要初始化 `ddddocr` 的 ONNX 模型，因此通常較慢。
4. 目前 `mvdisCrawl.py` 的部分說明文字與終端訊息仍使用舊名稱 `imageInput()` 或「人工 UI」，但實際程式匯入並呼叫的是 `englishAlphanumericOcrApi.ocrImage()`。
5. 建議不要將來源 Excel、查詢輸出、CAPTCHA 圖片與除錯頁面提交到公開 GitHub Repo。

建議 `.gitignore` 至少包含：

```gitignore
.venv/
__pycache__/
*.py[cod]
Data/*.xlsx
results/
.env
```

## 使用限制與責任聲明

本專案僅供合法、授權及內部資料處理用途。

使用者應：

- 僅查詢自己或已獲授權處理的法人／商業資料。
- 遵守監理服務網的服務條款、使用規範與存取限制。
- 避免高頻率或大量請求影響網站服務。
- 妥善保護統一編號、公司資料、違規資料及輸出 Excel。
- 自行確認查詢結果的正確性與法律效力。

本專案不是監理服務網或任何政府機關的官方工具。

## License

此專案目前尚未指定開放原始碼授權條款。公開 Repo 前，請依實際用途加入適合的 `LICENSE`。
