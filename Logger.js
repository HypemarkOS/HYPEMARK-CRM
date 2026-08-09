/**
 * ==========================================================
 * HYPEMARK CRM v1.0
 * Logger Library
 * ==========================================================
 */

const CRMLogger = {

  info(message) {

    console.log(`INFO : ${message}`);

  },

  warning(message) {

    console.warn(`WARNING : ${message}`);

  },

  error(message) {

    console.error(`ERROR : ${message}`);

  }

};