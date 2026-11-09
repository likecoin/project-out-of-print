# 絕版重生計劃

> 好書不該被遺忘，讓絕版重生，讓思想流通。

「絕版重生計劃」的單頁網站，由 [3ook.com 圖書館](https://3ook.com/) 發起，現正招募社會大眾參與。

## 內容

- `index.html` — 首頁，全部內容內嵌（HTML／CSS，無需建置步驟，無 JavaScript）
- `progress.html` — 計劃進展頁。書目即時讀取 Google 試算表「已發佈到網路」的公開分頁 CSV，
  因此這一頁需要 JavaScript；來源網址是檔案內 `CSV_URL` 一個常數。
  公開分頁只含書名、作者、出版社、系列、出版年月、狀態、上架連結；
  工作表上的作者聯絡線索、義工姓名與跟進記錄不會離開原表。
- `og-image.jpg` — 社交平台預覽圖

## 本機預覽

直接用瀏覽器打開 `index.html`，或：

```sh
python3 -m http.server 8000
# 然後開啟 http://localhost:8000
```

## 部署

靜態網站，把 `index.html` 放到任何靜態主機即可。若用 GitHub Pages，在 repository settings → Pages 選擇由 branch 的根目錄發佈。

## 授權

程式碼以 [GPL-3.0](LICENSE) 授權；網站文案版權屬「絕版重生計劃」。
