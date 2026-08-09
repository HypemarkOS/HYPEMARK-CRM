function recordActivity_(entity, entityId, action) {
  appendRow_(APP.SHEETS.ACTIVITIES, ['Activity ID','Entity','Entity ID','Action','User','Date Time'], {
    'Activity ID': generateId_(CONFIG.ID_PREFIX.ACTIVITY),
    'Entity': entity,
    'Entity ID': entityId,
    'Action': action,
    'User': Session.getActiveUser().getEmail() || 'System',
    'Date Time': now_()
  });
}
