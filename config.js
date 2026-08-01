// ============================================================
// HOUSEHOLD LEDGER — CONFIGURATION
// Fill in the four values below after completing the Google Cloud
// and Google Sheets setup steps in SETUP.md. Nothing else in this
// app needs to be touched.
// ============================================================
const CONFIG = {
  // The long ID in your Google Sheet's URL:
  // https://docs.google.com/spreadsheets/d/THIS_PART_HERE/edit
  SHEET_ID: "1iNqPPd3GAs-wZSTtu75QvlTPyQAY6bTtDmXq0veKCxE",
  // Google Cloud Console → APIs & Services → Credentials → API key
  // (restricted to Google Sheets API)
  API_KEY: "AIzaSyDp0-aRZ96aLnQ4GtEEEQidgub6UWDRrQg",
  // Google Cloud Console → APIs & Services → Credentials → OAuth 2.0 Client ID
  // (type: Web application)
  OAUTH_CLIENT_ID: "1063197035929-bdk4su5j148q96idv512vqglaath96u3.apps.googleusercontent.com",
  // Tab names — change only if you rename the tabs in your Sheet
  TABS: {
    CASH_FLOW: "Cash Flow",
    ACTUALS: "Actuals",
    USERS: "Users",
  },
};
