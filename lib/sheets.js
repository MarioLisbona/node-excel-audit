import { filterRowsForRFICells } from "./utils.js";
import {
  updateRfiCellData,
  prepareRfiCellDataForRfiSpreadsheet,
  getRfiRanges,
  updateExcelData,
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
    // fs.writeFileSync(
    //   "updatedRfiCellTestData.json",
    //   JSON.stringify(updatedRfiCellData)
    // );

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

  // Define starting row for both cases
  const startRowGeneralIssuesRfi = 7; // Start from row 7 for groups with enough projects
  const startRowSpecificIssuesRfi = 18; // Start from row 18 for groups with fewer projects

  // Get the ranges for general and specific issues RFI for the update request
  const { rangeForGeneralIssuesRfi, rangeForSpecificIssuesRfi } = getRfiRanges(
    startRowGeneralIssuesRfi,
    startRowSpecificIssuesRfi,
    generalIssuesRfiData.length,
    specificIssuesRfiData.length
  );

  const urlGeneralIssuesRfi = `https://graph.microsoft.com/v1.0/users/${userId}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/range(address='${rangeForGeneralIssuesRfi}')`;
  const urlSpecificIssuesRfi = `https://graph.microsoft.com/v1.0/users/${userId}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/range(address='${rangeForSpecificIssuesRfi}')`;

  // Prepare the request body with the data to update
  const requestBodyGeneralIssuesRfi = {
    values: generalIssuesRfiData,
  };

  // Call the new function for general issues RFI
  await updateExcelData(
    client,
    urlGeneralIssuesRfi,
    requestBodyGeneralIssuesRfi,
    "general issues RFI"
  );

  // Prepare the request body with the data to update for specific issues RFI
  const requestBodySpecificIssuesRfi = {
    values: specificIssuesRfiData,
  };

  // Call the new function for specific issues RFI
  await updateExcelData(
    client,
    urlSpecificIssuesRfi,
    requestBodySpecificIssuesRfi,
    "specific issues RFI"
  );
};
