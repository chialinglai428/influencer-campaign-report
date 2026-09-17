# influencer-campaign-report

**[English](#english) | [繁體中文](#繁體中文)**

## English

An AI-assisted workflow that turns Instagram and Facebook creator analytics screenshots into a consolidated Excel campaign report. It combines a Claude skill for screenshot interpretation and user confirmation with a Python script for workbook generation.

### The marketing problem

Campaign reporting often means collecting screenshots from multiple creators, transcribing metrics, and assembling a spreadsheet by hand. Different analytics layouts and inconsistent metric labels make that process harder to repeat reliably.

This project defines a reusable reporting workflow: extract the figures, confirm them with the marketer, and generate a structured workbook with formulas.

### Workflow

1. Provide post or Reel insights screenshots, profile screenshots, manually entered metrics, or a combination.
2. Claude interprets the inputs and organizes the metrics for each creator.
3. Review and confirm each creator's figures before they are added.
4. The Python report builder generates an `.xlsx` file with creator rows, totals, and engagement-rate formulas.

Screenshot interpretation is handled by Claude; the Python script accepts structured data and builds the workbook. This is a workflow with human review, not a standalone OCR application.

### Features

- Combines creator profile and post insights data into a single row.
- Records followers, reach, likes, comments, shares, saves, and optional link clicks.
- Uses editable Excel formulas for engagement rates and totals.
- Calculates the campaign engagement rate using reach weighting rather than a simple average of creator rates.
- Leaves missing link-click values blank.
- Produces a formatted workbook with Traditional Chinese column labels.

### Calculation rules

```text
Creator engagement rate = (likes + comments + shares) / reach
Campaign engagement rate = sum(likes + comments + shares) / sum(reach)
```

Here, shares include reposts and sends. Saves are recorded separately and are not included in this project's engagement-rate formula.

Reach is preferred as the denominator. If only view counts are available, the skill instructs Claude to disclose that they are being used as a proxy. Views and unique reach are not equivalent, so review denominator consistency before comparing creators.

Summed reach and follower counts are not deduplicated audiences. Inputs should have positive reach values; the current builder does not guard against division by zero.

### Getting started

Use a Claude environment that supports skills, image inputs, and Python execution. Place this repository in your skills directory, for example:

```text
~/.claude/skills/influencer-campaign-report/
```

Install the report builder's Python dependency:

```bash
python -m pip install openpyxl
```

Ask Claude:

> Help me consolidate these creators' Instagram campaign results into an Excel report.

Then provide the screenshots or numbers and confirm the extracted figures.

The skill also references a formula-recalculation utility supplied by its original host environment. That utility is not included in this repository. In another environment, use an available spreadsheet engine to recalculate and inspect the workbook; `openpyxl` writes formulas but does not evaluate them.

### Use the report builder directly

Run this example from the repository root:

```python
from scripts.build_report import build

# 合成範例資料，非真實活動成效。
rows = [
    {
        "name": "Creator A",
        "followers": 15000,
        "reach": 10000,
        "likes": 1000,
        "comments": 50,
        "shares": 150,
        "saves": 80,
        "link_clicks": 100,
    },
    {
        "name": "Creator B",
        "followers": 8000,
        "reach": 5000,
        "likes": 200,
        "comments": 20,
        "shares": 30,
        "saves": 40,
        "link_clicks": None,
    },
]

build(rows, "campaign_report.xlsx")
```

For this synthetic example, creator engagement rates are 12.0% and 5.0%; the reach-weighted campaign rate is approximately 9.7%. Open the workbook in a spreadsheet application to calculate and view the formula results.

### Project files

- [SKILL.md](SKILL.md): input interpretation, confirmation, metric rules, and reporting workflow.
- [scripts/build_report.py](scripts/build_report.py): Python workbook generator using `openpyxl`.

---

## 繁體中文

網紅專案結案數據整理 skill。把多位網紅的 IG/FB 貼文洞察報告或 Reel 洞察報告截圖，轉換成一份含加總與加權平均互動率的成效總表（xlsx）。

## 功能

- 讀取社群後台截圖（貼文洞察報告 / Reel 洞察報告 / profile 頁面）辨識數據，也支援手動輸入或截圖混用
- 每位網紅資料寫入前先覆誦確認，避免辨識錯誤
- 自動處理不同後台版型的觸及數標籤差異（觸及帳號 / 瀏覽人數 / 觀看次數）
- 產出 xlsx 總表：互動率以 Excel 公式寫入，總計列用觸及帳號數做加權平均互動率（非簡單平均）

## 安裝

這是一個 Claude skill。將整個資料夾放入你的 skills 目錄（例如 `~/.claude/skills/influencer-campaign-report/`），或打包成 `.skill` 檔匯入 Claude。

需要 Python 套件 `openpyxl` 來產生報表。

## 使用方式

對 Claude 說：

- 「幫我整理網紅成效」
- 「這是XX的後台截圖，粉絲數是XX」
- 「網紅專案結案報表」

然後依序提供每位網紅的截圖或數字，最後產出成效總表 xlsx。

## 報表範例

| 網紅名稱 | 粉絲數 | 觸及帳號數 | 愛心數 | 留言數 | 轉發數(轉發+傳送) | 收藏數 | 互動率 | 連結點擊數 |
|---|---|---|---|---|---|---|---|---|
| 網紅A | 15,000 | 10,104 | 1,174 | 13 | 245 | 185 | 14.2% | 134 |
| 網紅B | 18,000 | 7,262 | 260 | 10 | 22 | 13 | 4.0% | |
| 網紅C | 52,000 | 31,500 | 2,880 | 96 | 410 | 520 | 10.7% | 371 |
| **總計** | **85,000** | **48,866** | **4,314** | **119** | **677** | **718** | **10.5%** | **505** |

- 互動率公式：`(愛心數 + 留言數 + 轉發數) ÷ 觸及帳號數`，以 Excel 公式寫入儲存格
- 總計列的互動率為**加權平均**（以觸及帳號數加權），不是各列互動率的簡單平均
- 連結點擊數為選填，未提供則留空、不計入加總
- xlsx 樣式：標題列深紫底白字粗體，總計列淺紫底粗體，字型 Arial

## 檔案結構

```
influencer-campaign-report/
├── SKILL.md                    # skill 定義與流程說明
└── scripts/
    └── build_report.py         # xlsx 報表產生器
```

