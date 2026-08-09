/**
 * ==========================================================
 * HYPEMARK CRM v1.0
 * Data Service
 * ==========================================================
 */

const DataService = {

  /**
   * Returns the requested sheet.
   */
  getSheet(sheetName) {

    return Utils.getSheet(sheetName);

  },

  /**
   * Inserts one row into a sheet.
   */
  insert(sheetName, rowData) {

    const sheet = this.getSheet(sheetName);

    sheet.appendRow(rowData);

    CRMLogger.info(`Inserted record into ${sheetName}`);

  },

  /**
   * Returns all data from a sheet.
   */
  getAll(sheetName) {

    const sheet = this.getSheet(sheetName);

    return sheet.getDataRange().getValues();

  },

  /**
   * Finds a row using the first column (ID).
   * Returns the sheet row number or -1 if not found.
   */
  findRowByID(sheetName, id) {

    const sheet = this.getSheet(sheetName);

    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {

      if (String(values[i][0]) === String(id)) {
        return i + 1;
      }

    }

    return -1;

  }

};

/**
 * Creates default project after client onboarding
 */
function createProject(clientID, clientName, engagementID) {

  const sheet = Utils.getSheet(APP.SHEETS.PROJECTS);

  const projectID = generateProjectID();

  sheet.appendRow([

    projectID,

    clientID,

    engagementID,

    clientName + " Project",

    "Monthly Retainer",

    new Date(),

    "",

    "Active",

    "",

    "Medium",

    new Date(),

    new Date()

  ]);

  CRMLogger.info("Project Created : " + projectID);

  return projectID;

}




/**
 * Creates default engagement after client onboarding
 */
function createEngagement(clientID, clientName) {

  const sheet = Utils.getSheet(APP.SHEETS.ENGAGEMENTS);

  const engagementID = generateEngagementID();

  sheet.appendRow([

    engagementID,

    clientID,

    "Monthly Retainer",

    clientName + " Monthly Retainer",

    "",

    "Monthly",

    "Advance",

    new Date(),

    "",

    "No",

    "",

    "Active",

    new Date()

  ]);

  CRMLogger.info("Engagement Created : " + engagementID);

  return engagementID;

}







/**
 * Creates a client and automatically creates
 * Engagement and Project
 */
function createClient(clientData) {

  const clientSheet = Utils.getSheet(APP.SHEETS.CLIENTS);

  const clientID = generateClientID();

  clientSheet.appendRow([

  clientID,

  clientData.clientName || "",

  clientData.company || "",

  clientData.contactPerson || "",

  clientData.phone || "",

  clientData.email || "",

  clientData.website || "",

  clientData.address || "",

  clientData.gst || "",

  clientData.referredBy || "",

  clientData.industry || "",

  clientData.clientStage || "Onboarded",

  clientData.accountManager || "",

  clientData.priority || "Medium",

  new Date(),

  "Active"

]);
}