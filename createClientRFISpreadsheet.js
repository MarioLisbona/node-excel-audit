import { getGraphClient } from "./lib/msAuth.js";
import { getCellRange } from "./lib/utils.js";

// Function to copy a worksheet to a new spreadsheet
export const copyWorksheetToNewSpreadsheet = async (
  sourceWorkbookId,
  sourceWorksheetName,
  newWorkbookId,
  newWorksheetName
) => {
  // Create a Graph client with caching disabled
  const client = await getGraphClient({ cache: false });

  // Create a new Excel spreadsheet
  const newSpreadsheet = await client
    .api(`/drives/${process.env.ONEDRIVE_ID}/root/children`)
    .post({
      name: `${newWorkbookId}.xlsx`, // Name for the new Excel file
      file: {}, // Specify that this is a file
      "@microsoft.graph.conflictBehavior": "rename", // Handle conflicts by renaming
    });

  // Extract the ID of the new spreadsheet
  const newSpreadsheetId = newSpreadsheet.id;

  // Create a new worksheet in the new spreadsheet
  const newWorksheet = await client
    .api(
      `/drives/${process.env.ONEDRIVE_ID}/items/${newSpreadsheetId}/workbook/worksheets`
    )
    .post({
      name: newWorksheetName, // Name for the new worksheet
    });

  // Delete the default "Sheet1" in the new spreadsheet
  await client
    .api(
      `/drives/${process.env.ONEDRIVE_ID}/items/${newSpreadsheetId}/workbook/worksheets('Sheet1')`
    )
    .delete();

  // Extract the data from the source worksheet
  const existingData = await client
    .api(
      `/drives/${process.env.ONEDRIVE_ID}/items/${sourceWorkbookId}/workbook/worksheets/${sourceWorksheetName}/usedRange`
    )
    .get();

  // Check if existingData is valid
  if (!existingData || !existingData.values) {
    throw new Error("No data found in the existing worksheet.");
  }

  const cellValuesData = existingData.values;

  // Calculate the cell range for the data
  const newRangeAddress = getCellRange(cellValuesData);

  // Log the new range address
  console.log(`Writing to range: ${newRangeAddress}`);

  // Write the data to the new worksheet
  await client
    .api(
      `/drives/${process.env.ONEDRIVE_ID}/items/${newSpreadsheetId}/workbook/worksheets/${newWorksheetName}/range(address='${newRangeAddress}')`
    )
    .patch({
      values: cellValuesData, // Write the filtered data to the new worksheet
    });
};
