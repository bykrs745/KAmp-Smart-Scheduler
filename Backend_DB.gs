/**
 * Backend_DB.gs
 * シートのヘッダーを読み込み、列名と列番号を動的にマッピングする機能
 * （旧 02_ColumnManager.gs をDeck向けにリネーム）
 */

/**
 * 対象シートの1行目を読み込み、{"列名": 列番号(1始まり)} のMapを返します。
 */
function getColumnMap(sheet) {
  const maxCol = sheet.getMaxColumns();
  if (maxCol === 0) return {};
  const headers = sheet.getRange(1, 1, 1, maxCol).getValues()[0];
  const colMap = {};
  headers.forEach((name, index) => {
    const trimmedName = name.toString().trim();
    if (trimmedName !== "") {
      colMap[trimmedName] = index + 1;
    }
  });
  return colMap;
}

/**
 * 列マップを使って、指定行のデータを一括でJSON(Object)として取得する便利関数
 * 🌟 今後、Deck（一覧画面）を作る際に、スプレッドシートのデータを
 * JSON配列としてフロントエンドに返すために重宝します。
 */
function getRowDataAsObject(sheet, rowNumber, colMap) {
  const lastCol = sheet.getLastColumn();
  const rowValues = sheet.getRange(rowNumber, 1, 1, lastCol).getValues()[0];
  const rowData = {};
  for (const [colName, colIndex] of Object.entries(colMap)) {
    rowData[colName] = rowValues[colIndex - 1];
  }
  return rowData;
}