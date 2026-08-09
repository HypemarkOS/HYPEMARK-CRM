function logInfo_(message, data) { console.log('[INFO] ' + message + (data ? ' ' + JSON.stringify(data) : '')); }
function logError_(message, error) { console.error('[ERROR] ' + message + (error && error.stack ? '\n' + error.stack : '')); }
