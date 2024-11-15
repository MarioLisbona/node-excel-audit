import { filterRowsForRFICells } from "./utils.js";
import { updateRfiCellData } from "./utils.js";
import fs from "fs";

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
    const rfiCellData = filterRowsForRFICells(data);

    const updatedRfiCellData = await updateRfiCellData(rfiCellData);

    // Write updatedRfiCellData to a json file in the root of the project
    const dataToWrite = updatedRfiCellData.map(
      ({ projectsAffected, ...rest }) => rest
    ); // Exclude projectsAffected
    fs.writeFileSync("updatedRfiCellData.json", JSON.stringify(dataToWrite));

    // return updatedRows;
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
