function generateBidSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const takeoffSheet = ss.getSheetByName("takeoff");
  const materialsSheet = ss.getSheetByName("materials");
  const laborSheet = ss.getSheetByName("labor");

  let bidSheet = ss.getSheetByName("bid_sheet");
  if (bidSheet) {
    ss.deleteSheet(bidSheet);
  }
  bidSheet = ss.insertSheet("bid_sheet");

  // Fetch and parse reference data
  const materialData = materialsSheet.getDataRange().getValues();
  const laborData = laborSheet.getDataRange().getValues();

  const materialMap = {};
  for (let i = 1; i < materialData.length; i++) {
    const row = materialData[i];
    materialMap[row[0]] = {
      category: row[1],
      description: row[2],
      unit: row[3],
      cost: parseFloat(row[4]),
      vendor: row[5]
    };
  }

  const laborMap = {};
  for (let i = 1; i < laborData.length; i++) {
    const row = laborData[i];
    laborMap[row[0]] = parseFloat(row[1]); // Key: Category
  }

  // Write header
  const header = ["Item", "Category", "Unit", "Qty", "Material $", "Labor $", "Line Total"];
  bidSheet.getRange(1, 1, 1, header.length).setValues([header]);

  const takeoffData = takeoffSheet.getDataRange().getValues();
  let output = [];
  for (let i = 1; i < takeoffData.length; i++) {
    const [itemDesc, , unit, qty] = takeoffData[i];
    if (!itemDesc || !qty) continue;

    // Extract item code (assumes code is before first space or parenthesis)
    const match = itemDesc.match(/^([A-Z0-9-]+)/);
    const code = match ? match[1] : null;

    let category = "Misc", materialCost = 0, laborCost = 0, vendor = "";

    if (code && materialMap[code]) {
      const material = materialMap[code];
      category = material.category || "Misc";
      materialCost = material.cost || 0;
      vendor = material.vendor || "";
    }

    if (laborMap[category]) {
      laborCost = laborMap[category];
    }

    const lineTotal = qty * (materialCost + laborCost);
    output.push([
      itemDesc, category, unit, qty, materialCost, laborCost, lineTotal
    ]);
  }

  // Write bid line items
  bidSheet.getRange(2, 1, output.length, header.length).setValues(output);

  // Add totals
  const lastRow = output.length + 2;
  bidSheet.getRange(lastRow, 6).setValue("Subtotal:");
  bidSheet.getRange(lastRow, 7).setFormula(`=SUM(G2:G${lastRow - 1})`);

  bidSheet.getRange(lastRow + 1, 6).setValue("Tax:");
  bidSheet.getRange(lastRow + 1, 7).setFormula(`=G${lastRow} * 0.05`); // Example 5% tax

  bidSheet.getRange(lastRow + 2, 6).setValue("Overhead:");
  bidSheet.getRange(lastRow + 2, 7).setValue(500); // Flat overhead

  bidSheet.getRange(lastRow + 3, 6).setValue("Margin:");
  bidSheet.getRange(lastRow + 3, 7).setFormula(`=(G${lastRow} + G${lastRow + 1} + G${lastRow + 2}) * 0.1`); // 10% margin

  bidSheet.getRange(lastRow + 4, 6).setValue("Grand Total:");
  bidSheet.getRange(lastRow + 4, 7).setFormula(`=SUM(G${lastRow}:G${lastRow + 3})`);

  bidSheet.autoResizeColumns(1, 7);
}
