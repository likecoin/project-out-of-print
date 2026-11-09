# 絕版重生計劃簡報

按網站 `index.html` 的內容製作的簡報，共 11 版，含講者備註。

| 檔案 | 說明 |
|---|---|
| `poop-deck.pptx` | 簡報檔（16:9，1280×720） |
| `build.js` | 產生 `poop-deck.pptx` 的程式。改簡報請改這裡再重新產生，不要直接編輯 .pptx |

## 版面

1. 封面 — POOP · 絕版重生計劃 · 標語
2. 計劃理念
3. 為甚麼要做這件事 — 三項背景
4. 由 3ook.com 圖書館發起
5. 計劃涵蓋甚麼 — 地域語言、與上架的關係、敏感書
6. 版權仍然屬於版權持有人
7. 版權持有人需要提供甚麼
8. 免費參加，那如何持續營運
9. 已經重生的作品
10. 三種參與方式
11. 結尾

配色取自網站：3ook.com 品牌薄荷綠 `#50E3C2`、墨黑 `#16130F`、米白 `#F5ECE0`。

## 重新產生

```sh
cd deck
npm install pptxgenjs     # jszip 會一併裝上，程式會從 pptxgenjs 旁邊取用
node build.js             # 覆寫 poop-deck.pptx
```

`build.js` 是自足的，除 `pptxgenjs` 外沒有其他相依。

## 預覽與轉檔

```sh
soffice --headless --convert-to pdf poop-deck.pptx   # 需要 libreoffice-impress
```

## 轉成 Google Slides

把 `poop-deck.pptx` 拖進 Google Drive，Drive 會自動轉成 Google Slides；或在 Google Slides 選 檔案 → 匯入投影片。
