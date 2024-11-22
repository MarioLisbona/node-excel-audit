import { getCellRange } from "./lib/utils.js";

// Function to copy a worksheet to a new spreadsheet
export const copyWorksheetToNewWorkbook = async (
  client,
  sourceWorkbookId,
  sourceWorksheetName,
  clientName,
  newWorksheetName
) => {
  // Create a new Excel spreadsheet
  const newWorkbook = await client
    .api(`/drives/${process.env.ONEDRIVE_ID}/root/children`)
    .post({
      name: `RFI Spreadsheet - ${clientName}.xlsx`, // Name for the new Excel file
      file: {}, // Specify that this is a file
      "@microsoft.graph.conflictBehavior": "rename", // Handle conflicts by renaming
    });

  // Extract the ID and name of the new workbook
  const newWorkbookId = newWorkbook.id;
  const newWorkbookName = newWorkbook.name;

  // Create a new worksheet in the new spreadsheet
  await client
    .api(
      `/drives/${process.env.ONEDRIVE_ID}/items/${newWorkbookId}/workbook/worksheets`
    )
    .post({
      name: newWorksheetName, // Name for the new worksheet
    });

  // Delete the default "Sheet1" in the new spreadsheet
  await client
    .api(
      `/drives/${process.env.ONEDRIVE_ID}/items/${newWorkbookId}/workbook/worksheets('Sheet1')`
    )
    .delete();

  // Extract the data from the source worksheet
  const existingData = await client
    .api(
      `/drives/${process.env.ONEDRIVE_ID}/items/${sourceWorkbookId}/workbook/worksheets/${sourceWorksheetName}/range(address='A7:B46')`
    )
    .get();

  // Check if existingData is valid
  if (!existingData || !existingData.values) {
    throw new Error("No data found in the existing worksheet.");
  }

  // Extract the cell values from the existing data
  const cellValuesData = existingData.values;

  // Calculate the cell range for the data
  const newRangeAddress = getCellRange(cellValuesData);

  // Log the new range address
  console.log(`Writing to range: ${newRangeAddress}`);

  // Write the data to the new worksheet
  await client
    .api(
      `/drives/${process.env.ONEDRIVE_ID}/items/${newWorkbookId}/workbook/worksheets/${newWorksheetName}/range(address='${newRangeAddress}')`
    )
    .patch({
      values: cellValuesData, // Write the filtered data to the new worksheet
    });

  return { newWorkbookId, newWorkbookName };
};
