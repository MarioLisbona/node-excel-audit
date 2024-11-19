import { getGraphClient } from "./lib/msAuth.js";
import { processTestingSheet, updateRfiSpreadsheet } from "./lib/sheets.js";
import dotenv from "dotenv";

// load the environment variables
dotenv.config();

// Create a Graph client with caching disabled
const client = await getGraphClient({ cache: false });

// Workbook ID and User ID
const workbookId = process.env.WORKBOOK_ID;
const userId = process.env.USER_ID;

// Process the testing sheet and return the updated RFI cell data
// Incliudes the
const updatedRfiCellData = await processTestingSheet(
  client,
  userId,
  workbookId,
  "Testing"
);

// // Update the RFI spreadsheet with the updated RFI cell data
// updateRfiSpreadsheet(
//   client,
//   userId,
//   workbookId,
//   "RFI Spreadsheet",
//   updatedRfiCellData
// );
