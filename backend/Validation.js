function validateClient_(data) {
  var errors = [];
  if (!data || !asText_(data.clientName)) errors.push('Client Name is required.');
  if (!data || !asText_(data.businessName)) errors.push('Business Name is required.');
  if (data && data.mobile && normalizeMobile_(data.mobile).length < 10) errors.push('Mobile number is invalid.');
  if (data && data.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalizeEmail_(data.email))) errors.push('Email is invalid.');
  return { valid: errors.length === 0, errors: errors };
}
