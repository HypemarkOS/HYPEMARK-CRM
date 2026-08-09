function generateId_(prefix) {
  return prefix + '-' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMddHHmmss') + '-' + Math.floor(100 + Math.random() * 900);
}
