import { ConfidentialClientApplication } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import "dotenv/config";

// MSAL client configuration
const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET,
  },
};

const getAccessToken = async () => {
  const cca = new ConfidentialClientApplication(msalConfig);

  const authResponse = await cca.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });

  return authResponse.accessToken;
};

// Initialize Microsoft Graph client
const getGraphClient = async () => {
  const accessToken = await getAccessToken();
  return Client.init({
    authProvider: (done) => done(null, accessToken),
  });
};

// Retrieve data from an Excel file
const getExcelData = async (client, userId, workbookId, sheetName) => {
  try {
    const range = `https://graph.microsoft.com/v1.0/users/${userId}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/usedRange`;
    console.log("Fetching data from:", range);
    const response = await client.api(range).get();
    console.log("Data Retrieved:", response.values);
    return response.values;
  } catch (error) {
    console.error("Error retrieving data:", error.message);
    console.error("Full error details:", error);
  }
};

// Update data in an Excel file
const updateExcelData = async (
  client,
  userId,
  workbookId,
  sheetName,
  range,
  values
) => {
  try {
    const updateRange = `https://graph.microsoft.com/v1.0/users/${userId}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/range(address='${range}')`;
    console.log("Updating data at:", updateRange);
    await client.api(updateRange).patch({
      values: values,
    });
    console.log("Data Updated Successfully");
  } catch (error) {
    console.error("Error updating data:", error.message);
    console.error("Full error details:", error);
  }
};

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
