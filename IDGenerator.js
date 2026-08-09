/**
 * ==========================================================
 * HYPEMARK CRM v1.0
 * ID Generator
 * ==========================================================
 */

const IDGenerator = {

  /**
   * Generates Client ID
   * Format: CDDMMYYNNN
   * Example: C070826009
   */
  generateClientID() {

    const sheet = Utils.getSheet(APP.SHEETS.CLIENTS);

    // Header row is row 1
    const clientCount = Math.max(sheet.getLastRow() - 1, 0);

    const nextNumber = clientCount + 1;

    const datePart = Utilities.formatDate(
      new Date(),
      CONFIG.TIMEZONE,
      "ddMMyy"
    );

    const serialPart = String(nextNumber).padStart(3, "0");

    return `C${datePart}${serialPart}`;

  }

};



/**
 * Generates Project ID
 * Format:
 * P07082601
 */
function generateProjectID() {

  const sheet = Utils.getSheet(APP.SHEETS.PROJECTS);

  const row = Math.max(sheet.getLastRow(), 1);

  const number = String(row).padStart(2, "0");

  const date = Utilities.formatDate(
    new Date(),
    CONFIG.TIMEZONE,
    "ddMMyy"
  );

  return "P" + date + number;

}







/**
 * Generates Engagement ID
 * Format:
 * E07082601
 */
function generateEngagementID() {

  const sheet = Utils.getSheet(APP.SHEETS.ENGAGEMENTS);

  const row = Math.max(sheet.getLastRow(), 1);

  const number = String(row).padStart(2, "0");

  const date = Utilities.formatDate(
    new Date(),
    CONFIG.TIMEZONE,
    "ddMMyy"
  );

  return "E" + date + number;

}