# 翊鼎國際廣告 · 代銷投組控制台

靜態版本，部署在 GitHub Pages，所有資料存放在 `db.json`。

## 一、首次設定

```bash
# 安裝相依套件（只需要 openpyxl，一次性）
pip install openpyxl
```

## 二、更新資料（每次有新 Excel 時）

```bash
# 1. 執行更新腳本，指定你的 Excel 路徑
python update_data.py "C:/Users/你的名字/Downloads/5131公司2025每周盈虧.xlsx"

# 2. 推送到 GitHub（Pages 自動更新）
git add db.json
git commit -m "更新資料 114.05"
git push
```

等 1–2 分鐘，儀表板即自動更新。

## 三、Repo 檔案說明

| 檔案 | 說明 |
|------|------|
| `index.html` | 儀表板前端（GitHub Pages 入口） |
| `db.json` | 所有資料（由 update_data.py 產生） |
| `update_data.py` | 本機執行的資料更新腳本 |

## 四、儀表板功能

- **月份切換**：114.01 ～ 最新月份
- **六個 KPI**：在售建案、規劃戶數、剩餘可售、本月損益、來人、成交（含與上月比較）
- **建案損益排行**：雙向長條圖，綠賺紅賠
- **去化進度**：每案已售比例進度條
- **月損益趨勢**：投組合計折線圖（損益 + 來人 + 成交）
- **來人→成交漏斗**：轉換率
- **建案明細表**：可排序；點列開啟單案抽屜（週損益 + 歷月趨勢 + 財務數據）
