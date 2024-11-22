import { getGraphClient } from "./lib/msAuth.js";
import {
  processTestingSheet,
  updateRfiSpreadsheet,
  copyWorksheetToNewWorkbook,
} from "./lib/worksheetProcessing.js";
import { emailRfiToClient } from "./lib/email.js";
import dotenv from "dotenv";

// load the environment variables
dotenv.config();

// Create a Graph client with caching disabled
const client = await getGraphClient({ cache: false });

// Workbook ID and User ID
const workbookId = process.env.WORKBOOK_ID;
const userId = process.env.USER_ID;
const testingSheetName = "Testing";

// Process the testing sheet and return the updated RFI cell data
// Incliudes the
const updatedRfiCellData = await processTestingSheet(
  client,
  userId,
  workbookId,
  testingSheetName
);

// Update the RFI spreadsheet with the updated RFI cell data
await updateRfiSpreadsheet(
  client,
  userId,
  workbookId,
  "RFI Spreadsheet",
  updatedRfiCellData
);

// Call the function
const { newWorkbookId, newWorkbookName } = await copyWorksheetToNewWorkbook(
  client,
  process.env.SOURCE_WORKBOOK_ID,
  process.env.SOURCE_WORKSHEET_NAME,
  "Mario Lisbona Dev",
  process.env.NEW_WORKSHEET_NAME
);

// await emailRfiToClient(newWorkbookId, newWorkbookName);
