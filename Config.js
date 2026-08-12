const CONFIG = Object.freeze({
  SHEETS: {
    TEMPLATE: 'EUR',
    HISTORY: 'EUR_XERO',
    TRANSACTIONS: 'EUR_transactions',
    CONTACTS: 'SA Contacts'
  },

  TEMPLATE: {
    HEADER_ROW: 1,
    FIRST_DATA_ROW: 2,
    COLUMN_COUNT: 6,

    COLUMNS: {
      DATE: 1,
      AMOUNT: 2,
      PAYEE: 3,
      DESCRIPTION: 4,
      REFERENCE: 5,
      CHECK_NUMBER: 6
    }
  },

  TRANSACTIONS: {
    HEADER_ROW: 10,
    FIRST_DATA_ROW: 11,
    COLUMN_COUNT: 14,

    COLUMNS: {
      VALUE_DATE: 4,
      DEBIT: 6,
      CREDIT: 7,
      DESCRIPTION_1: 11,
      DESCRIPTION_3: 13
    }
  },

  CONTACTS: {
    FIRST_DATA_ROW: 2,
    COLUMN_COUNT: 2,

    COLUMNS: {
      XERO_NAME: 1,
      BANK_NAME: 2
    }
  }
});