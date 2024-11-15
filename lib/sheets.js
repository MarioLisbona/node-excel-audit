import { filterRowsForRFICells } from "./utils.js";
import {
  updateRfiCellData,
  prepareRfiCellDataForRfiSpreadsheet,
} from "./utils.js";
import fs from "fs";

// Retrieve data from an Excel file
export const processTestingSheet = async (
  client,
  userId,
  workbookId,
  sheetName
) => {
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

    // // Write updatedRfiCellData to a json file in the root of the project
    // const dataToWrite = updatedRfiCellData.map(
    //   ({ projectsAffected, ...rest }) => rest
    // ); // Exclude projectsAffected
    // fs.writeFileSync("updatedRfiCellData.json", JSON.stringify(dataToWrite));

    return updatedRfiCellData;
  } catch (error) {
    // Log the error if data retrieval fails
    console.error("Error retrieving data:", error.message);
    console.error("Full error details:", error);
  }
};

// Function to update an Excel spreadsheet with new data
export const updateRfiSpreadsheet = async (
  client,
  userId,
  workbookId,
  sheetName,
  rfiCellData
) => {
  // Specify the range you want to update (e.g., A1:B2)
  const rangeAddress = "A1:B2"; // Adjust this as needed
  // const range = `https://graph.microsoft.com/v1.0/users/${userId}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/usedRange`;

  // Filter groupedData into two arrays: one where projectsAffected.length >= 4, and one where projectsAffected.length < 4
  const generalIssuesRfi = rfiCellData.filter(
    (group) => group.projectsAffected.length >= 4
  );
  const specificIssuesRfi = rfiCellData.filter(
    (group) => group.projectsAffected.length < 4
  );

  // Prepare data for both sets of groups
  const generalIssuesRfiData =
    prepareRfiCellDataForRfiSpreadsheet(generalIssuesRfi);
  const specificIssuesRfiData =
    prepareRfiCellDataForRfiSpreadsheet(specificIssuesRfi);

  console.log("generalIssuesRfiData", generalIssuesRfiData);
  console.log("specificIssuesRfiData", specificIssuesRfiData);

  // Define starting row for both cases
  const startRowGeneralIssuesRfi = 7; // Start from row 7 for groups with enough projects
  const startRowSpecificIssuesRfi = 18; // Start from row 18 for groups with fewer projects

  // Update data for general issues RFI (>= 4)
  const rangeForGeneralIssuesRfi = `A${startRowGeneralIssuesRfi}:B${
    startRowGeneralIssuesRfi + generalIssuesRfiData.length - 1
  }`;

  // Update data for general issues RFI (>= 4)
  const rangeForSpecificIssuesRfi = `A${startRowSpecificIssuesRfi}:B${
    startRowSpecificIssuesRfi + specificIssuesRfiData.length - 1
  }`;

  const urlGeneralIssuesRfi = `https://graph.microsoft.com/v1.0/users/${userId}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/range(address='${rangeForGeneralIssuesRfi}')`;
  const urlSpecificIssuesRfi = `https://graph.microsoft.com/v1.0/users/${userId}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/range(address='${rangeForSpecificIssuesRfi}')`;

  // Prepare the request body with the data to update
  const requestBody = {
    values: generalIssuesRfiData,
  };

  try {
    // Send a PATCH request to update the data in the specified range for general issues RFI
    await client.api(urlGeneralIssuesRfi).patch(requestBody);

    console.log(
      "Excel spreadsheet updated successfully for general issues RFI."
    );
  } catch (error) {
    // Log the error if the update fails for general issues RFI
    console.error(
      "Error updating Excel data for general issues RFI:",
      error.message
    );
    console.error("Full error details for general issues RFI:", error);
    // Log the request details for better debugging for general issues RFI
    console.error("Request URL for general issues RFI:", urlGeneralIssuesRfi);
    console.error(
      "Request Body for general issues RFI:",
      JSON.stringify(requestBody)
    );
  }

  // Prepare the request body with the data to update for specific issues RFI
  const requestBodySpecificIssuesRfi = {
    values: specificIssuesRfiData,
  };

  try {
    // Send a PATCH request to update the data in the specified range for specific issues RFI
    await client.api(urlSpecificIssuesRfi).patch(requestBodySpecificIssuesRfi);

    console.log(
      "Excel spreadsheet updated successfully for specific issues RFI."
    );
  } catch (error) {
    // Log the error if the update fails for specific issues RFI
    console.error(
      "Error updating Excel data for specific issues RFI:",
      error.message
    );
    console.error("Full error details for specific issues RFI:", error);
    // Log the request details for better debugging for specific issues RFI
    console.error("Request URL for specific issues RFI:", urlSpecificIssuesRfi);
    console.error(
      "Request Body for specific issues RFI:",
      JSON.stringify(requestBodySpecificIssuesRfi)
    );
  }
};
