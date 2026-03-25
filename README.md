# Easy Knowledge Graph

以 Excel 檔案為資料來源的互動式知識圖譜，靈感來自 [Obsidian Graph View](https://obsidian.md/)。透過瀏覽器即可視覺化探索本地 `.xlsx` 檔案之間的關聯網路。

![架構示意](https://img.shields.io/badge/Python-3.11+-blue) ![FastAPI](https://img.shields.io/badge/FastAPI-0.115+-green) ![D3.js](https://img.shields.io/badge/D3.js-v7-orange) ![License](https://img.shields.io/badge/license-MIT-lightgrey)

---

## 功能特色

- **互動式力導向圖** — D3.js 力模擬，支援拖曳、縮放、平移
- **Excel 原生支援** — 直接讀取 `.xlsx`，無需安裝 Office
- **Meta Sheet 規範** — 每個檔案可加一張 `Meta` 工作表定義標題、標籤、連結、說明
- **顯式 + 隱式連結** — Meta 明確指定的連結（顯式），以及儲存格內容比對到其他文件名稱的自動偵測連結（隱式）
- **即時同步** — watchdog 監聽資料夾，變更時透過 WebSocket 自動更新圖譜
- **內容預覽面板** — 點擊節點展開側面板，逐頁籤瀏覽工作表資料（預設顯示前 100 列）
- **標籤篩選 & 搜尋** — 依標籤或文字過濾節點，反白相關連結
- **局部視圖** — 雙擊節點僅顯示該節點及其直接鄰居
- **匯出獨立 HTML** — 一鍵產生內嵌所有資料的單一 HTML 檔，無需 Python 即可分享瀏覽
- **深色介面** — 預設深色主題，易於長時間閱讀

---

## 系統需求

- Python 3.11+
- [uv](https://docs.astral.sh/uv/) 套件管理工具

---

## 快速開始

### 1. 安裝依賴

```bash
git clone https://github.com/heinzyao/easy-knowledge-graph.git
cd easy-knowledge-graph
uv sync
```

### 2. 建立範例資料（可選）

```bash
uv run python create_sample.py
```

會在 `sample_data/` 目錄下產生 6 個範例 Excel 檔，包含相互連結的專案文件（專案計畫書、技術規格、客戶需求等）。

### 3. 啟動伺服器

```bash
# 使用預設 sample_data/ 資料夾
uv run python main.py

# 指定自訂資料夾
uv run python main.py /path/to/your/excel/folder
```

啟動後瀏覽器會自動開啟 `http://localhost:8000`。

---

## Excel 檔案格式

### Meta 工作表（選填）

在 `.xlsx` 檔案中新增一張名為 `Meta` 的工作表，第一欄為鍵、第二欄為值：

| 鍵          | 值（範例）                    | 說明                       |
|-------------|-------------------------------|----------------------------|
| title       | 技術規格文件                  | 節點顯示名稱（預設為檔名）  |
| tags        | 技術, 後端, API               | 逗號分隔的標籤              |
| links       | 專案計畫書, 資料庫設計        | 逗號分隔的連結目標（檔名或標題） |
| description | 系統技術架構與 API 規格說明   | 節點說明文字                |

### 隱式連結

若工作表中的儲存格內容完全符合另一個檔案的標題或檔名，系統會自動建立隱式連結（預設隱藏，可透過工具列開關顯示）。

---

## 操作說明

| 操作           | 功能                               |
|----------------|------------------------------------|
| 點擊節點       | 開啟右側內容預覽面板               |
| 雙擊節點       | 進入局部視圖（僅顯示鄰居）         |
| 雙擊空白處     | 退出局部視圖                       |
| 滾輪 / 觸控板  | 縮放圖譜                           |
| 拖曳背景       | 平移圖譜                           |
| 拖曳節點       | 固定節點位置                       |
| 搜尋框         | 即時過濾節點（標題比對）           |
| 標籤下拉選單   | 依標籤篩選節點                     |
| 顯示隱式連結   | 切換顯示/隱藏自動偵測的隱式連結    |
| 重新整理按鈕   | 手動觸發全圖重建                   |
| 📥 匯出按鈕    | 匯出為可離線分享的獨立 HTML 檔     |

### 匯出獨立 HTML

點擊工具列的「📥 匯出」按鈕，瀏覽器會下載一個 `knowledge-graph.html` 檔案。此檔案：

- 內嵌所有節點資料、圖結構與工作表內容
- 內嵌 D3.js（匯出時自動下載，或保留 CDN 連結作為備援）
- 支援搜尋、標籤篩選、隱式連結切換、節點面板等所有互動功能
- 可直接用瀏覽器開啟，**無需安裝 Python 或啟動任何伺服器**

---

## 專案結構

```
easy-knowledge-graph/
├── main.py              # FastAPI 應用程式進入點
├── pyproject.toml       # 專案設定與依賴宣告
├── create_sample.py     # 產生範例 Excel 資料
│
├── core/
│   ├── scanner.py       # 掃描 .xlsx 並解析 Meta 工作表
│   ├── graph_builder.py # NetworkX 圖建構與增量更新
│   └── watcher.py       # watchdog 資料夾監聽（asyncio 橋接）
│
├── api/
│   └── routes.py        # FastAPI 路由（REST + WebSocket）
│
├── frontend/
│   ├── index.html       # 主頁面（工具列、圖容器、面板框架）
│   ├── style.css        # 深色主題樣式
│   ├── graph.js         # D3.js 力導向圖、互動邏輯
│   └── panel.js         # 節點預覽面板、工作表切換
│
└── sample_data/         # 範例 Excel 檔（執行 create_sample.py 產生）
```

---

## API 端點

| 方法      | 路徑              | 說明                                       |
|-----------|-------------------|--------------------------------------------|
| `GET`     | `/`               | 前端主頁面                                 |
| `GET`     | `/api/graph`      | 圖資料（D3 格式），支援 `?include_implicit=true` |
| `GET`     | `/api/tags`       | 所有標籤列表                               |
| `GET`     | `/api/node`       | 節點詳情與工作表預覽，支援 `?id=&sheet=&limit=` |
| `POST`    | `/api/refresh`    | 觸發全圖重建                               |
| `GET`     | `/api/export`     | 匯出內嵌所有資料的獨立 HTML                |
| `WS`      | `/ws/updates`     | WebSocket 即時更新推播                     |

---

## 技術棧

| 層級     | 技術                                                              |
|----------|-------------------------------------------------------------------|
| 後端     | [FastAPI](https://fastapi.tiangolo.com/) + [uvicorn](https://www.uvicorn.org/) |
| 圖計算   | [NetworkX](https://networkx.org/)                                 |
| Excel 解析 | [openpyxl](https://openpyxl.readthedocs.io/)                   |
| 檔案監聽 | [watchdog](https://python-watchdog.readthedocs.io/)               |
| 前端視覺化 | [D3.js v7](https://d3js.org/)（CDN，無需建置工具）               |
| 套件管理 | [uv](https://docs.astral.sh/uv/)                                  |

---

## License

MIT
