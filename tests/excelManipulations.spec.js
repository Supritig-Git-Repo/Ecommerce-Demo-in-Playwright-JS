const { test } = require("@playwright/test");
const { excelUtils } = require("../utils/excelUtils");

test("excel implementation", async () => {
  const excelUtilities = new excelUtils();
  await excelUtilities.traceCordinate(
    "Inventory",
    "C:/Users/GhoshSupriti/Downloads/Samsung-s23-Inventory.xlsx",
    "128GB"
  );

});