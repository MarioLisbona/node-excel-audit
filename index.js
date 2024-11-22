import { getGraphClient } from "./lib/msAuth.js";
import {
  processTestingSheet,
  updateRfiSpreadsheet,
  copyWorksheetToNewWorkbook,
} from "./lib/worksheetProcessing.js";
import { emailRfiToClient } from "./lib/email.js";
import dotenv from "dotenv";
import { getFileIdByName } from "./lib/oneDrive.js";

// load the environment variables
dotenv.config();

// Create a Graph client with caching disabled
const client = await getGraphClient({ cache: false });

// use filename to get workbook ID
const workbookId = await getFileIdByName(
  process.env.ONEDRIVE_ID,
  "els-testing-client-XXXX.xlsx"
  // "els-testing.xlsx"
);
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
  workbookId,
  process.env.SOURCE_WORKSHEET_NAME,
  "Mario Lisbona Dev",
  process.env.NEW_WORKSHEET_NAME
);

await emailRfiToClient(newWorkbookId, newWorkbookName);
