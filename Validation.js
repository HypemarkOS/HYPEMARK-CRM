/**
 * ==========================================================
 * HYPEMARK CRM v1.0
 * Validation Library
 * ==========================================================
 */

const Validation = {

  /**
   * Check if a value is empty
   */
  isEmpty(value) {
    return value === null ||
           value === undefined ||
           value === "";
  },

  /**
   * Check if a string is a valid email
   */
  isEmail(email) {

    if (this.isEmpty(email)) return false;

    const regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

    return regex.test(email);

  },

  /**
   * Check if a mobile number is valid
   * Accepts 10-digit Indian mobile numbers
   */
  isMobile(mobile) {

    if (this.isEmpty(mobile)) return false;

    const regex = /^[6-9]\d{9}$/;

    return regex.test(String(mobile));

  },

  /**
   * Check if value is numeric
   */
  isNumber(value) {

    return !isNaN(value);

  }

};