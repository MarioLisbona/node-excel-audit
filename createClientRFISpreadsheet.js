import { getGraphClient } from "./lib/msAuth.js";

// Function to copy a worksheet to a new spreadsheet
export const copyWorksheetToNewSpreadsheet = async (
  sourceSpreadsheetId,
  worksheetName,
  newSpreadsheetName
) => {
  const client = await getGraphClient();

  // Step 1: Create a new Excel file in the root of the drive
  const newSpreadsheet = await client
    .api(`/drives/${process.env.ONEDRIVE_ID}/root/children`)
    .post({
      name: `${newSpreadsheetName}.xlsx`, // Name for the new Excel file
      file: {}, // Specify that this is a file
      "@microsoft.graph.conflictBehavior": "rename", // Handle conflicts by renaming
    });

  // Step 2: Get the ID of the new spreadsheet
  const newSpreadsheetId = newSpreadsheet.id;

  // Step 3: Get the data from the existing worksheet
  const existingData = await client
    .api(
      `/drives/${process.env.ONEDRIVE_ID}/items/${sourceSpreadsheetId}/workbook/worksheets/${worksheetName}/usedRange`
    )
    .get();

  // Check if existingData is valid
  if (!existingData || !existingData.values) {
    throw new Error("No data found in the existing worksheet.");
  }

  // Step 4: Create a new worksheet in the new spreadsheet
  const newWorksheet = await client
    .api(
      `/drives/${process.env.ONEDRIVE_ID}/items/${newSpreadsheetId}/workbook/worksheets`
    )
    .post({
      name: worksheetName, // Name for the new worksheet
    });

  console.log({ newWorksheet, id: newWorksheet.id });

  // Step 5: Write the data to the new worksheet
  const rangeAddress = `${worksheetName}!A1`; // Specify the starting cell for the data
  await client
    .api(
      `/drives/${process.env.ONEDRIVE_ID}/items/${newSpreadsheetId}/workbook/worksheets/${newWorksheet.id}/range(address='${rangeAddress}')`
    )
    .patch({
      values: existingData.values, // Write the existing data to the new worksheet
    });

  return newSpreadsheetId;
};

// Call the function
copyWorksheetToNewSpreadsheet(
  process.env.SOURCE_SPREADSHEET_ID,
  process.env.WORKSHEET_NAME,
  process.env.NEW_SPREADSHEET_NAME
);
