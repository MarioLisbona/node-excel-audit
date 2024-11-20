import { getGraphClient } from "./lib/msAuth.js";

// Function to calculate the cell range based on the data array
const getCellRange = (data) => {
  const startCell = "A1"; // Starting cell
  const numRows = data.length; // Number of rows in the data
  const numCols = data[0] ? data[0].length : 0; // Number of columns in the first row

  // Calculate the end cell based on the number of rows and columns
  const endCell =
    String.fromCharCode("A".charCodeAt(0) + numCols - 1) + numRows;

  return `${startCell}:${endCell}`; // Return the range in A1:B2 format
};

// Function to copy a worksheet to a new spreadsheet
export const copyWorksheetToNewSpreadsheet = async (
  sourceSpreadsheetId,
  sourceWorksheetName,
  newWorksheetName,
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

  // Step 6: Create a new worksheet with the desired name
  const newWorksheet = await client
    .api(
      `/drives/${process.env.ONEDRIVE_ID}/items/${newSpreadsheetId}/workbook/worksheets`
    )
    .post({
      name: newWorksheetName, // Name for the new worksheet
    });

  // Step 5: Delete "Sheet1" in the new spreadsheet
  await client
    .api(
      `/drives/${process.env.ONEDRIVE_ID}/items/${newSpreadsheetId}/workbook/worksheets('Sheet1')`
    )
    .delete();

  // Step 3: Get the data from the existing worksheet
  const existingData = await client
    .api(
      `/drives/${process.env.ONEDRIVE_ID}/items/${sourceSpreadsheetId}/workbook/worksheets/${sourceWorksheetName}/usedRange`
    )
    .get();

  // Check if existingData is valid
  if (!existingData || !existingData.values) {
    throw new Error("No data found in the existing worksheet.");
  }

  const cellValuesData = existingData.values;

  // Step 7: Write the data to the newly created worksheet
  const newRangeAddress = getCellRange(cellValuesData);

  // Log the new range address and the filtered data being sent
  console.log(`Writing to range: ${newRangeAddress}`);

  await client
    .api(
      `/drives/${process.env.ONEDRIVE_ID}/items/${newSpreadsheetId}/workbook/worksheets/${newWorksheetName}/range(address='${newRangeAddress}')`
    )
    .patch({
      values: cellValuesData, // Write the filtered data to the new worksheet
    });
};

// Call the function
copyWorksheetToNewSpreadsheet(
  process.env.SOURCE_SPREADSHEET_ID,
  process.env.SOURCE_WORKSHEET_NAME,
  process.env.NEW_WORKSHEET_NAME,
  process.env.NEW_SPREADSHEET_NAME
);
