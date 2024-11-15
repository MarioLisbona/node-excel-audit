import { getGraphClient } from "./lib/msAuth.js";
import { getExcelData, updateExcelData } from "./lib/sheets.js";
import dotenv from "dotenv";

dotenv.config();

const client = await getGraphClient();

// Replace with your OneDrive file ID, sheet name, and user ID
const workbookId = process.env.WORKBOOK_ID;
const sheetName = process.env.SHEET_NAME;
const userId = process.env.USER_ID;

// Retrieve existing data only once
const data = await getExcelData(client, userId, workbookId, sheetName);

// Example of updating data
// const rangeToUpdate = "C6:D6"; // Specify the range to update
// const newValues = [["Updated Value 1", "Updated Value 2"]]; // New data to update
// await updateExcelData(
//   client,
//   userId,
//   workbookId,
//   sheetName,
//   rangeToUpdate,
//   newValues
// );
