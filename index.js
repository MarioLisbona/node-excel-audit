import { getGraphClient } from "./lib/msAuth.js";
import { getExcelData, updateExcelData } from "./lib/sheets.js";
const client = await getGraphClient();

// Replace with your OneDrive file ID, sheet name, and user ID
const workbookId = "01FNQELGBWKBKRZQZJYZFLPYTIA4IS3HB7"; // Use this ID in your API calls
const sheetName = "test";
const userId = "c9641b08-94b7-4988-99a9-c2ddad0268d4"; // Specify the user ID here

// Retrieve existing data only once
const data = await getExcelData(client, userId, workbookId, sheetName);

// Example of updating data
const rangeToUpdate = "C6:D6"; // Specify the range to update
const newValues = [["Updated Value 1", "Updated Value 2"]]; // New data to update
await updateExcelData(
  client,
  userId,
  workbookId,
  sheetName,
  rangeToUpdate,
  newValues
);
