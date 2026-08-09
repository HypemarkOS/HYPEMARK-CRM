/**
 * ==========================================================
 * HYPEMARK CRM v1.0
 * CRM Initializer
 * ==========================================================
 */

function initializeCRM() {

  try {

    CRMLogger.info("Initializing CRM...");

    const SHEETS = [

      {
  name: APP.SHEETS.CLIENTS,
  headers: [

  "Client ID",

  "Client Name",

  "Company",

  "Contact Person",

  "Phone",

  "Email",

  "Website",

  "Address",

  "GST",

  "Referred By",

  "Industry",

  "Client Stage",

  "Account Manager",

  "Priority",

  "Created On",

  "Status"

]
},



{
  name: APP.SHEETS.ENGAGEMENTS,
  headers: [
    "Engagement ID",
    "Client ID",
    "Engagement Type",
    "Package Name",
    "Budget",
    "Billing Cycle",
    "Payment Terms",
    "Start Date",
    "End Date",
    "Agreement Signed",
    "Agreement Date",
    "Status",
    "Created On"
  ]
},

{
  name: APP.SHEETS.DELIVERABLES,
  headers: [
    "Deliverable ID",
    "Client ID",
    "Engagement ID",
    "Deliverable Type",
    "Quantity",
    "Frequency",
    "Remarks",
    "Created On"
  ]
},


{
  name: APP.SHEETS.SERVICES,
  headers: [
    "Service ID",
    "Service Name",
    "Category",
    "Unit",
    "Default Price",
    "Status"
  ]
},

{
  name: APP.SHEETS.PROJECTS,
  headers: [
    "Project ID",
    "Client ID",
    "Engagement ID",
    "Project Name",
    "Project Type",
    "Start Date",
    "End Date",
    "Status",
    "Project Manager",
    "Priority",
    "Created On",
    "Modified On"
  ]
},

{
  name: APP.SHEETS.ACTIVITY_LOG,
  headers: [
    "Activity ID",
    "Date & Time",
    "Module",
    "Record ID",
    "Activity",
    "Description",
    "Performed By"
  ]
},



      {
        name: APP.SHEETS.LEADS,
        headers: [
          "Lead ID",
          "Lead Name",
          "Phone",
          "Source",
          "Assigned To",
          "Status",
          "Created On"
        ]
      },

      {
        name: APP.SHEETS.PAYMENTS,
        headers: [
          "Payment ID",
          "Client ID",
          "Amount",
          "Received In",
          "Received By",
          "Date",
          "Remarks"
        ]
      },

      {
        name: APP.SHEETS.TASKS,
        headers: [
          "Task ID",
          "Task",
          "Assigned To",
          "Due Date",
          "Priority",
          "Status"
        ]
      },

      {
        name: APP.SHEETS.USERS,
        headers: [
          "User ID",
          "Name",
          "Email",
          "Role",
          "Status"
        ]
      },

      {
        name: APP.SHEETS.SETTINGS,
        headers: [
          "Setting",
          "Value"
        ]
      }

    ];

    SHEETS.forEach(createCRMTable);
    seedServices();

    SpreadsheetApp.flush();

    //SpreadsheetApp.getUi().alert("✅ HYPEMARK CRM initialized successfully.");

    CRMLogger.info("CRM initialization completed.");

  } catch (error) {

    CRMLogger.error(error);

    SpreadsheetApp.getUi().alert(
      "CRM Initialization Failed\n\n" + error.message
    );

    throw error;

  }

}
















/**
 * Populates the Services master sheet.
 */
function seedServices() {

  const sheet = Utils.getSheet(APP.SHEETS.SERVICES);

  if (sheet.getLastRow() > 1) {
    return;
  }

  const services = [

    ["SER001","Reel","Social Media","Nos","0","Active"],
    ["SER002","Poster","Social Media","Nos","0","Active"],
    ["SER003","Carousel","Social Media","Nos","0","Active"],
    ["SER004","Story","Social Media","Nos","0","Active"],
    ["SER005","YouTube Video","Video","Nos","0","Active"],
    ["SER006","Meta Ads","Advertising","Campaign","0","Active"],
    ["SER007","Google Ads","Advertising","Campaign","0","Active"],
    ["SER008","Website","Development","Project","0","Active"],
    ["SER009","SEO","Marketing","Month","0","Active"],
    ["SER010","Product Shoot","Production","Day","0","Active"],
    ["SER011","UGC Video","Content","Nos","0","Active"],
    ["SER012","Influencer Campaign","Marketing","Campaign","0","Active"]

  ];

  sheet.getRange(2,1,services.length,6).setValues(services);

}

















/**
 * Creates a sheet and prepares its headers
 */
function createCRMTable(config) {

  let sheet;

  if (Utils.sheetExists(config.name)) {

    sheet = Utils.getSheet(config.name);

  } else {

    sheet = Utils.createSheet(config.name);

    CRMLogger.info(`Created sheet : ${config.name}`);

  }

  createHeaders(sheet, config.headers);

}



/**
 * Creates and formats the header row
 * Only creates headers if the sheet is empty.
 */
function createHeaders(sheet, headers) {

  // If the sheet already has data, don't overwrite it
  if (sheet.getLastRow() > 0) {

    CRMLogger.info(`Headers already exist for ${sheet.getName()}`);
    return;

  }

  // Write headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);

  headerRange
    .setBackground(APP.COLORS.PRIMARY)
    .setFontColor(APP.COLORS.WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  CRMLogger.info(`Headers created for ${sheet.getName()}`);

}