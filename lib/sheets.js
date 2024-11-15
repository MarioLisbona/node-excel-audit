import { filterRowsForRFICells } from "./utils.js";

// Retrieve data from an Excel file
export const getExcelData = async (client, userId, workbookId, sheetName) => {
  try {
    // Construct the URL for the Excel file's used range
    const range = `https://graph.microsoft.com/v1.0/users/${userId}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/usedRange`;

    // Fetch the data from all non-empty rows in the sheet
    const response = await client.api(range).get();

    // Extract the values from the response
    const data = response.values;

    // Filter the data to include only rows containing "RFI"
    const filteredRows = filterRowsForRFICells(data);

    // Log the last row of the filtered data for verification
    console.log(filteredRows[filteredRows.length - 1]);

    // Return the filtered rows
    return filteredRows;
  } catch (error) {
    // Log the error if data retrieval fails
    console.error("Error retrieving data:", error.message);
    console.error("Full error details:", error);
  }
};

// Update data in an Excel file
export const updateExcelData = async (
  client,
  userId,
  workbookId,
  sheetName,
  range,
  values
) => {
  try {
    const updateRange = `https://graph.microsoft.com/v1.0/users/${userId}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/range(address='${range}')`;
    console.log("Updating data at:", updateRange);
    await client.api(updateRange).patch({
      values: values,
    });
    console.log("Data Updated Successfully");
  } catch (error) {
    console.error("Error updating data:", error.message);
    console.error("Full error details:", error);
  }
};
