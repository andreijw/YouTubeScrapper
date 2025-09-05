const TITLE_COLOR = '#d9ead3'

/**
 * Styles a range of rows in the active sheet.
 * @param {number} startRow - The starting row number (1-based).
 * @param {number} numRows - The number of rows to style.
 * @param {number} numCols - The number of columns to style.
 * @param {string} bgColor - Background color (e.g. '#d9ead3').
 * @param {boolean} bold - Whether to apply bold text.
 * @param {boolean} borders - Whether to apply borders.
 */
function styleRows(startRow, numRows, numCols, bgColor, bold, borders) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange(startRow, 1, numRows, numCols);
  if (bgColor) range.setBackground(bgColor);
  if (bold) range.setFontWeight('bold');
  if (borders) range.setBorder(true, true, true, true, true, true);
}