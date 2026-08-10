const CONFIG = {
  COMPANY: { NAME: 'Hypemark Digital Solutions Pvt Ltd', BRAND: 'HYPEMARK CRM' },
  DATE_FORMAT: 'dd-MMM-yyyy',
  TIMEZONE: 'Asia/Kolkata',
  CURRENCY: 'INR',
  DEFAULTS: {
    CLIENT_STATUS: 'Active',
    PROJECT_STATUS: 'Active',
    DELIVERABLE_STATUS: 'Pending',
    PRIORITY: 'Medium'
  },
  ID_PREFIX: {
    CLIENT: 'CLI', ENGAGEMENT: 'ENG', PROJECT: 'PRJ', DELIVERABLE: 'DLV', CONTENT: 'CNT',
    PAYMENT: 'PAY', CLIENT_RECEIPT: 'REC', EMPLOYEE_PAYMENT: 'EMP', BUSINESS_EXPENSE: 'EXP', ACTIVITY: 'ACT'
  },
  PAYMENT_ACCOUNTS: ['Ajay', 'Vineeth', 'HYPEMARK', 'Whitecoat Media'],
  PAYMENT_RECIPIENTS: ['Ajay', 'Vineeth'],
  EMPLOYEE_PAYMENT_TYPES: ['Salary', 'Commission', 'Freelance', 'Reimbursement', 'Advance', 'Other'],
  EXPENSE_CATEGORIES: ['Petrol', 'Camera Rental', 'Transport', 'Food', 'Accommodation', 'Equipment', 'Office', 'Other'],
  EXPENSE_CONTEXTS: ['Client Shoot', 'Client Meeting', 'Client Travel', 'Office', 'General Business', 'Other'],
  PAYMENT_MODES: ['Cash', 'UPI', 'Bank Transfer', 'Card', 'Other']
};
