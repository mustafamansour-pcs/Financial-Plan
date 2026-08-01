// ============================================================
// Static reference data pulled directly from the household sheet's
// dropdown lists (Category / Sub Category) so the Add Entry form
// works instantly without waiting on a network round trip.
// Edit these arrays if the categories in your sheet ever change.
// ============================================================

const CATEGORIES = [
  "Home Expenses",
  "Car Installments",
  "Gam3ya",
  "Car Expenses",
  "Contingency (Bank Account BM)",
  "Outings",
  "Pocket Money",
  "Car Maintinance (BM)",
];

// Sub categories shown regardless of parent category, matching the
// sheet's existing single flat dropdown (Actuals!N4:N26).
const SUBCATEGORIES = [
  "Hygene",
  "Meat (Beef & Chicken)",
  "Seafood",
  "Rice & Macaroni & Fries",
  "Diaries",
  "Drinks",
  "Eggs",
  "Bread",
  "Oils",
  "Sause",
  "Spread",
  "Snacks",
  "Washing",
  "Electricity",
  "Other Utilities",
  "Internet",
  "Mobile",
  "Sub total (1) Fuel",
  "Sub total (2) Tires",
  "Sub total (3) Maintinance",
  "Pocket Money",
  "Medications",
  "Fruites",
  "Miscl.",
  "Vegetables",
  "Supplements",
];
