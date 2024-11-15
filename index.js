import { getGraphClient } from "./lib/msAuth.js";
import { processTestingSheet, updateRfiSpreadsheet } from "./lib/sheets.js";
import dotenv from "dotenv";

dotenv.config();

const client = await getGraphClient();

// Replace with your OneDrive file ID, sheet name, and user ID
const workbookId = process.env.WORKBOOK_ID;
const userId = process.env.USER_ID;

// Retrieve existing data only once
const updatedRfiCellData = await processTestingSheet(
  client,
  userId,
  workbookId,
  "Testing"
);

updateRfiSpreadsheet(
  client,
  userId,
  workbookId,
  "RFI Spreadsheet",
  updatedRfiCellData
);
