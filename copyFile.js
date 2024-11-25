import { getGraphClient, getAccessToken } from "./lib/msAuth.js";
import { getFileIdByName } from "./lib/oneDrive.js";

// Function to copy a file in OneDrive
export const copyRfiTemplate = async (userId, workbookId, newFileName) => {
  const client = await getGraphClient();

  // Step 1: Copy the file
  await client.api(`/users/${userId}/drive/items/${workbookId}/copy`).post({
    parentReference: {
      id: "root", // or specify the folder ID if needed
    },
    name: newFileName,
  });

  console.log(`File copied to new file: ${newFileName}`);
};

const main = async () => {
  const userId = process.env.USER_ID;
  // use filename to get workbook ID
  const workbookId = await getFileIdByName(
    process.env.ONEDRIVE_ID,
    "Templates.xlsx"
  );
  const clientName = "Test Client";

  console.log({ workbookId });

  // Usage
  copyRfiTemplate(
    userId,
    workbookId,
    `RFI Responses - ${clientName}.xlsx` // New file name
  );
};

main();
