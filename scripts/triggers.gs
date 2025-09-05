/**
 * Add sheets triggers
 */
function onOpen(e) {
  // Add a custom menu to the spreadsheet.
  SpreadsheetApp.getUi() // Or DocumentApp, SlidesApp, or FormApp.
      .createMenu('YouTube Api Tools')
      .addItem('Search channels by subject', 'promptSubject')
      .addItem('Query channel ids for metrics', 'collectPublicMetrics')
      .addToUi();
}