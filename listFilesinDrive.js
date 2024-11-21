import axios from "axios";
import { getAccessToken } from "./lib/msAuth.js";

async function listFilesInOneDrive(driveId) {
  const accessToken = await getAccessToken();
  const response = await axios.get(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    }
  );
  return response.data.value; // Returns the list of files
}

async function getFileNamesAndIds(driveId) {
  const files = await listFilesInOneDrive(driveId);
  return files.map((file) => ({
    id: file.id,
    name: file.name,
  }));
}

const fileNamesAndIds = await getFileNamesAndIds(process.env.ONEDRIVE_ID);
console.log(fileNamesAndIds);
