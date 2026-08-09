/**
 * ==========================================================
 * HYPEMARK CRM v1.0
 * Utility Functions
 * ==========================================================
 */

const Utils = {

  /**
   * Returns active spreadsheet
   */
  getSpreadsheet() {
    return SpreadsheetApp.getActiveSpreadsheet();
  },

  /**
   * Returns sheet by name
   */
  getSheet(sheetName) {

    const sheet = this.getSpreadsheet().getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found.`);
    }

    return sheet;

  },

  /**
   * Check if sheet exists
   */
  sheetExists(sheetName) {

    return this.getSpreadsheet().getSheetByName(sheetName) !== null;

  },

  /**
   * Create sheet if not exists
   */
  createSheet(sheetName) {

    if (this.sheetExists(sheetName)) {
      return this.getSheet(sheetName);
    }

    return this.getSpreadsheet().insertSheet(sheetName);

  },

  /**
   * Get last row safely
   */
  lastRow(sheet) {

    return Math.max(sheet.getLastRow(), 1);

  },

  /**
   * Get last column safely
   */
  lastColumn(sheet) {

    return Math.max(sheet.getLastColumn(), 1);

  },

  /**
   * Current timestamp
   */
  now() {

    return new Date();

  },

  /**
   * Format date
   */
  formatDate(date) {

    return Utilities.formatDate(
      date,
      CONFIG.TIMEZONE,
      CONFIG.DATE_FORMAT
    );

  },

  /**
   * Generate Unique ID
   */
  generateID(prefix) {

    const time = Date.now();

    const random = Math.floor(Math.random() * 1000);

    return `${prefix}-${time}-${random}`;

  }

};