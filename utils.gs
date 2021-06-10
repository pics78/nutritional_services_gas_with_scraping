/**
 * [util]セルの書式と値を設定する関数
 */
function setCellProp(sheet, range, backgroundColor, fontColor, value) {
  sheet.getRange(range).setBackground(backgroundColor).setFontColor(fontColor).setValue(value);
}

/**
 * [util]初期化関数
 * フォントカラーと値を初期化しつつ背景をグレーにする
 */
function resetResult(sheet, range) { setCellProp(sheet, range, COMMON_INFO.InactiveBackColor, null, null); }

/**
 * [util]値入力関数
 * 背景とフォントの色を初期化しつつ値を入力する
 */
function setVal(sheet, range, value) { setCellProp(sheet, range, COMMON_INFO.ActiveBackColor, null, value); }

/**
 * [util]エラー値入力関数
 * 背景を初期化しつつ赤色のフォントでエラー文言を入力する
 */
function setError(sheet, range, value) { setCellProp(sheet, range, COMMON_INFO.ActiveBackColor, 'red', value); }

/**
 * [util]指定セルに設定する関数を作成する
 * cell: 関数を設定するセル
 * value: 100gあたり栄養価
 */
function makeFuncForCell(cell, value) {
  const col = cell.replace(/\d+$/, "");
  const row = cell.replace(/^[A-Z]+/, "");
  const weightCell = COMMON_INFO.calculateNutritionalValues.ranges.weightCol + row;
  const digits = Object.values(COMMON_INFO.calculateNutritionalValues.nutInfo)
    .find(elem => elem.sheetCol == col).digits;
  return Utilities.formatString("=round(%f*(%s*0.01), %d)", value, weightCell, digits);
}

/**
 * [util]指定した列の合計値セルに合計値を算出する関数を作成する
 * col: 関数を設定する列
 * total: 総食品数
 */
function makeSumFunc(col, total) {
  const firstRow = COMMON_INFO.calculateNutritionalValues.ranges.partialResultFirstRow;
  return Utilities.formatString("=sum(%s%d:%s%d)", col, firstRow, col, firstRow + total - 1);
}

/**
 * [util]成分値が通常の数値か計算不要の記号等かどうか判定する
 */
function checkNormalValue(value) {
  return value == 'Tr' || value == '-' || /^\([\d\.Tr]+\)$/.test(value) ? false : true;
}

/**
 * [util]正常のHTMLが取得できたかをチェックする
 * 食品成分データベースはエラーが出るとtitleタグに[ERROR]という文言が入るようなのでこれを見る
 */
function isSearchError(resultHtml) {
  const title = Parser.data(resultHtml).from('<title>').to('</title>').build();
  return /^\[ERROR\].*$/.test(title) ? true : false;
}