function onEdit(e) {
  const editedSheet = SpreadsheetApp.getActiveSpreadsheet();
  const editedSheetName = editedSheet.getSheetName();
  const editedRange = e.range;
  const editedCell = editedRange.getA1Notation();
  
  if (editedSheetName == COMMON_INFO.searchFoods.sheetName) {
    // 食品検索シートが編集された場合
    const RANGE = COMMON_INFO.searchFoods.ranges;
    if (editedCell == RANGE.searchWordCell) {
      // 検索ワードセルが編集されていた場合
      editedRange.setBackground(
        editedRange.getValue() != '' ?
          COMMON_INFO.ActiveBackColor : COMMON_INFO.InactiveBackColor
      );
    }
  } else if (editedSheetName == COMMON_INFO.calculateNutritionalValues.sheetName) {
    // 栄養価計算シートが編集された場合
    const editedCol = editedCell.split('')[0];
    const RANGE = COMMON_INFO.calculateNutritionalValues.ranges;
    if (editedCol == RANGE.foodNumberCol) {
      const editedRow = editedRange.getRow();
      if (RANGE.partialResultFirstRow <= editedRow && editedRow <= COMMON_INFO.MAX_ROW) {
        // 食品番号列が編集された場合
        editedRange.setBackground(
          editedRange.getValue() != '' && editedRange.offset(-1, 0).getBackground() != COMMON_INFO.InactiveBackColor ?
            COMMON_INFO.ActiveBackColor : COMMON_INFO.InactiveBackColor
        );
      }
    } else if (editedCol == RANGE.weightCol) {
      const editedRow = editedRange.getRow();
      if (RANGE.partialResultFirstRow <= editedRow && editedRow <= COMMON_INFO.MAX_ROW) {
        // 重量列が編集された場合
        if (editedRange.getValue() == '') {
          editedRange.setBackground(COMMON_INFO.InactiveBackColor);
        } else {
          if (editedRange.offset(-1, 0).getBackground() != COMMON_INFO.InactiveBackColor) {
            editedRange.setBackground(COMMON_INFO.ActiveBackColor);
          }
        }
      }
    }
  }
}

/**
 * 検索ワードから食品成分表に載っている食品とその食品番号を検索して表示する
 */
function searchFoods() {
  const RANGE = COMMON_INFO.searchFoods.ranges;
  const sheet = SpreadsheetApp.getActiveSpreadsheet();

  // 前回の検索結果をクリアする
  const prevHits = sheet.getRange(RANGE.hitCountCell).getValue();
  if (prevHits != '') {
    if (/^\d+$/.test(prevHits) && prevHits > 0) {
      const lastRow = RANGE.resultFirstRow + prevHits - 1;
      resetResult(sheet, `${RANGE.foodNumberCol}${RANGE.resultFirstRow}:${RANGE.foodNameCol}${lastRow}`);
    }
    resetResult(sheet, RANGE.hitCountCell);
  }
  // 検索ワードを取得
  const searchWord = sheet.getRange(RANGE.searchWordCell).getValue();
  if (searchWord == '') {
    Browser.msgBox('検索ワードを入力してください。', Browser.Buttons.OK);
    return;
  }
  const sourceInfo = COMMON_INFO.searchFoods.sourceInfo;
  const options = {
    'method': sourceInfo.method,
    'payload': JSON.parse(Utilities.formatString(sourceInfo.payloadTemp, searchWord)),
  };
  const html = UrlFetchApp.fetch(sourceInfo.url, options).getContentText('UTF-8');
  let hits = 0;
  
  if (!isSearchError(html)) {
    const foodsTable = Parser.data(html).from('<table id="food_table" ').to('</table>').build();
    const tableContent = Parser.data(foodsTable).from('<tbody>').to('</tbody>').build();
    const foodRows = Parser.data(tableContent).from('<tr>').to('</tr>').iterate();
    hits = foodRows.length;
    for (i=0; i<hits; i++) {
      const foodNumber = foodRows[i].replace(/^<td><label\sfor="c_7_\d{1,2}_(\d{5})">.*$/, '$1');
      const foodName = foodRows[i].replace(/^<td><label\s.*><input\s.*\schecked>(.*)<\/label><\/td>$/, '$1');
      const currentRow = RANGE.resultFirstRow + i;
      setVal(sheet, RANGE.foodNumberCol + currentRow, foodNumber);
      setVal(sheet, RANGE.foodNameCol + currentRow, foodName);
    }
  }
  setVal(sheet, RANGE.hitCountCell, hits);
}

/**
 * 食品検索シートを初期化する
 */
function resetSearchFoods() {
  const RANGE = COMMON_INFO.searchFoods.ranges;
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  resetResult(sheet, RANGE.searchWordCell);
  resetResult(sheet, RANGE.hitCountCell);
  resetResult(sheet, `${RANGE.foodNumberCol}${RANGE.resultFirstRow}:${RANGE.foodNameCol}${COMMON_INFO.MAX_ROW}`);
}

/**
 * 入力された食品番号と重量から各食品毎と合計の栄養価を算出する
 */
function calculateNutritionalValues() {
  const RANGE = COMMON_INFO.calculateNutritionalValues.ranges;
  const NUT_INFO = COMMON_INFO.calculateNutritionalValues.nutInfo;
  const sheet = SpreadsheetApp.getActiveSpreadsheet();

  // 前回の検索結果をクリアする
  const preTotal = sheet.getRange(RANGE.totalFoodCountCell).getValue();
  if (preTotal != '') {
    if (/^\d+$/.test(preTotal) && preTotal > 0) {
      const lastRow = RANGE.partialResultFirstRow + preTotal - 1;
      resetResult(sheet, `${RANGE.foodNameCol}${RANGE.partialResultFirstRow}:${RANGE.foodNameCol}${lastRow}`);
      resetResult(sheet, `${NUT_INFO.E.sheetCol}${RANGE.partialResultFirstRow}:${NUT_INFO.Sal.sheetCol}${lastRow}`);
      resetResult(sheet, `${RANGE.weightCol}${RANGE.totalResultRow}:${NUT_INFO.Sal.sheetCol}${RANGE.totalResultRow}`);
      sheet.getRange(`${RANGE.cacheCol}${RANGE.partialResultFirstRow}:${RANGE.cacheCol}${lastRow}`).clearContent();
    }
    resetResult(sheet, RANGE.totalFoodCountCell);
  }
  // 栄養価の入力セル情報と桁数情報のテンプレを作成
  const requiredElementMappingTemp = {};
  Object.values(NUT_INFO).map(function(elem) {
    requiredElementMappingTemp[elem.sourceTableRow] = {
      'cell': elem.sheetCol, // まだ列名のみ ループの中でそれぞれの行数を加える
      'digits': elem.digits,
    };
  });
  const requiredIndex = Object.keys(requiredElementMappingTemp);

  let i = RANGE.partialResultFirstRow;
  let total = 0;
  while (true) {
    // 食品番号を取得
    const fn = sheet.getRange(RANGE.foodNumberCol + i).getValue();
    if (fn == '') break;
    total++;
    if (/^\d{5}$/.test(fn)) {
      const html = UrlFetchApp.fetch(
        Utilities.formatString(COMMON_INFO.calculateNutritionalValues.sourceInfo.urlTemp, fn)
      ).getContentText('UTF-8');

      if (isSearchError(html)) {
        setError(sheet, RANGE.foodNameCol + i++, '存在しない食品番号です。');
        continue;
      }
      // 重量を取得
      const weight = sheet.getRange(RANGE.weightCol + i).getValue();
      if (!/^\d+\.?\d*$/.test(weight)) {
        setError(sheet, RANGE.foodNameCol + i++, '重量の値が不正です。');
        continue;
      }
      sheet.getRange(RANGE.preWeightCol + i).setValue(weight);
      // 食品名を取得
      const foodFullName = Parser.data(html).from('<span class="foodfullname">').to('</span>').build();

      const nutTable = Parser.data(html).from('<table id="nut"').to('</table>').build();
      const tbody = Parser.data(nutTable).from('<tbody>').to('</tbody>').build();
      const elements = Parser.data(tbody).from('<tr>').to('</tr>').iterate();

      // 現在の食品の栄養価入力セル情報と桁数情報の作成
      const requiredElementMapping = JSON.parse(JSON.stringify(requiredElementMappingTemp));
      requiredIndex.map(key => requiredElementMapping[key].cell += i);

      setVal(sheet, RANGE.foodNameCol + i, foodFullName);
      for (let num=0; num<elements.length; num++) {
        if (requiredIndex.indexOf(num.toString()) != -1) {
          const columns = Parser.data(elements[num]).from('<td ').to('</td>').iterate();
          const val = columns[columns.length-2].replace(/^class="(num|marker)">([\d\(\)Tr\.\-]*)$/, '$2');
          const targetCell = requiredElementMapping[num].cell;
          setVal(sheet, targetCell, checkNormalValue(val) ? makeFuncForCell(targetCell, val) : val);
        }
      }
      i++;
    } else {
      setError(sheet, RANGE.foodNameCol + i++, '食品番号の形式が正しくありません。');
    }
  }
  setVal(sheet, RANGE.weightCol + RANGE.totalResultRow, makeSumFunc(RANGE.weightCol, total));
  Object.values(NUT_INFO).map(elem =>
    setVal(sheet, elem.sheetCol + RANGE.totalResultRow, makeSumFunc(elem.sheetCol, total)));
  setVal(sheet, RANGE.totalFoodCountCell, total);
}

/**
 * 栄養価計算結果を初期化する
 */
function resetCalculateNutritionalValues() {
  const RANGE = COMMON_INFO.calculateNutritionalValues.ranges;
  const NUT_INFO = COMMON_INFO.calculateNutritionalValues.nutInfo;
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  resetResult(sheet, RANGE.totalFoodCountCell);
  resetResult(sheet, `${RANGE.weightCol}${RANGE.totalResultRow}:${NUT_INFO.Sal.sheetCol}${RANGE.totalResultRow}`);
  resetResult(sheet, `${RANGE.foodNumberCol}${RANGE.partialResultFirstRow}:${NUT_INFO.Sal.sheetCol}${COMMON_INFO.MAX_ROW}`);
}