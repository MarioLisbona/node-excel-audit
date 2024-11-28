import { getGraphClient } from "./lib/msAuth.js";
import {
  processTesting,
  updateRfiSpreadsheet,
  copyWorksheetToNewWorkbook,
} from "./lib/worksheetProcessing.js";
import { emailRfiToClient } from "./lib/email.js";
import dotenv from "dotenv";
import { getFileIdByName } from "./lib/oneDrive.js";
import { copyWorksheetToClientWorkbook } from "./lib/worksheetProcessing.js";
// load the environment variables
dotenv.config();

// Create a Graph client with caching disabled
const client = await getGraphClient({ cache: false });

// use filename to get workbook ID
const workbookId = await getFileIdByName(
  process.env.ONEDRIVE_ID,
  // "els-testing-client-XXXX.xlsx"
  "els-testing.xlsx"
);
const userId = process.env.USER_ID;
const testingSheetName = "Testing";

// Process the testing sheet and return the updated RFI cell data
// Incliudes the
const updatedRfiCellData = await processTesting(
  client,
  userId,
  workbookId,
  testingSheetName
);

// only update the RFI spreadsheet, copy ane email if there is RFI data to process
if (updatedRfiCellData.length > 0) {
  // Update the RFI spreadsheet with the updated RFI cell data
  await updateRfiSpreadsheet(
    client,
    userId,
    workbookId,
    "RFI Spreadsheet",
    updatedRfiCellData
  );

  // Copy the data in the updated RFI spreadsheet to a new workbook
  // using the template and email it to the client
  const { newWorkbookId, newWorkbookName } =
    await copyWorksheetToClientWorkbook(
      client,
      userId,
      workbookId,
      process.env.SOURCE_WORKSHEET_NAME,
      "Mario Lisbona Dev"
    );

  // await emailRfiToClient(newWorkbookId, newWorkbookName);
} else {
  console.log("No RFI data to process");
}
