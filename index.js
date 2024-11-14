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
    const range = `https://graph.microsoft.com/v1.0/users/${userId}/drive/items/${workbookId}/workbook/worksheets/${sheetName}/range(address='A1:D10')`;
    const response = await client.api(range).get();
    console.log("Data Retrieved:", response.values);
    return response.values;
  } catch (error) {
    console.error("Error retrieving data:", error);
  }
};

// Update Excel data
const updateExcelData = async (client, workbookId, sheetName, values) => {
  try {
    const range = `https://graph.microsoft.com/v1.0/me/drive/items/${workbookId}/workbook/worksheets/${sheetName}/range(address='A1')`;
    await client.api(range).patch({
      values: values,
    });
    console.log("Data updated successfully");
  } catch (error) {
    console.error("Error updating data:", error);
  }
};

// Main function
const main = async () => {
  const client = await getGraphClient();

  // Replace with your OneDrive file ID, sheet name, and user ID
  const workbookId = "80FC2B37B766D18D!282"; // Use this ID in your API calls
  const sheetName = "simple-test";
  const userId = "9287cd73-637d-4d90-8c3e-585e4a4ae6c2"; // Specify the user ID here

  // Retrieve existing data
  const data = await getExcelData(client, userId, workbookId, sheetName);

  // Update data (e.g., adding a new row)
  const newData = [...data, ["New Data", 123, "Another Value", 456]];
  await updateExcelData(client, workbookId, sheetName, newData);
};

main().catch(console.error);
