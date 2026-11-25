/**
 * 【絕版重生】書目 — 有人改動某一行，就在該行的「Last Update」欄蓋上時間。
 *
 * 安裝：在試算表選 擴充功能 > Apps Script，把整個檔案的內容貼進去，存檔。
 * 不用自己建觸發器，也不用授權：onEdit 是簡單觸發器，存檔後下一次編輯就生效。
 *
 * 因此它只在「人手編輯」時執行。以下情況不會蓋時間：
 *   - 公式重算（例如狀態欄因為授權欄改了而變）
 *   - 透過 API 或其他指令碼寫入
 *   - 匯入、還原版本
 *
 * 欄位按標題名尋找，所以日後調動欄的位置不用改這裡。
 *
 * 時間用試算表本身的時區。要改在 檔案 > 設定 > 時區。
 */

var SHEET_NAME   = '書目';
var STAMP_HEADER = 'Last Update';

function onEdit(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  var headers  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var stampCol = headers.indexOf(STAMP_HEADER) + 1;   // indexOf 回 -1 時變成 0
  if (!stampCol) return;                              // 找不到該欄就什麼都不做

  // 改的就是時間戳本身（例如有人手動清走），不要立刻再蓋一次
  if (e.range.getColumn() === stampCol && e.range.getNumColumns() === 1) return;

  var firstRow = Math.max(e.range.getRow(), 2);       // 標題列不蓋
  var lastRow  = e.range.getLastRow();
  if (lastRow < firstRow) return;

  // 一次貼上多行時，每一行都蓋
  var now    = new Date();
  var stamps = [];
  for (var r = firstRow; r <= lastRow; r++) stamps.push([now]);

  sheet.getRange(firstRow, stampCol, stamps.length, 1)
       .setValues(stamps)
       .setNumberFormat('yyyy-mm-dd hh:mm');
}
