import { getGraphClient } from "./lib/msAuth.js";
import { getExcelData } from "./lib/sheets.js";
import dotenv from "dotenv";

dotenv.config();

const client = await getGraphClient();

// Replace with your OneDrive file ID, sheet name, and user ID
const workbookId = process.env.WORKBOOK_ID;
const sheetName = process.env.SHEET_NAME;
const userId = process.env.USER_ID;

// Retrieve existing data only once
const updatedRfiCellData = await getExcelData(
  client,
  userId,
  workbookId,
  sheetName
);
