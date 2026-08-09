/**
 * ==========================================================
 * HYPEMARK CRM v1.0
 * Activity Service
 * ==========================================================
 */

const ActivityService = {

  log(module, recordID, activity, description, user = "System") {

    const sheet = Utils.getSheet(APP.SHEETS.ACTIVITY_LOG);

    const activityID = "ACT" + Utilities.formatDate(
      new Date(),
      CONFIG.TIMEZONE,
      "ddMMyyHHmmss"
    );

    sheet.appendRow([
      activityID,
      new Date(),
      module,
      recordID,
      activity,
      description,
      user
    ]);

    CRMLogger.info(`${module} Activity Logged`);

  }

};