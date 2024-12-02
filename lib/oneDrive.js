import axios from "axios";
import fs from "fs/promises";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { getAccessToken } from "./msAuth.js";

// Function to download a file from OneDrive
export const downloadFile = async (graphClient, userId, fileId) => {
  try {
    const fileUrl = `https://graph.microsoft.com/v1.0/users/${userId}/drive/items/${fileId}/content`;
    const response = await graphClient
      .api(fileUrl)
      .responseType("arraybuffer")
      .get();

    return Buffer.from(response);
  } catch (error) {
    console.error("Error downloading file:", error);
    throw error;
  }
};

// Function to upload a file to OneDrive
export const uploadFile = async (
  graphClient,
  userId,
  folderId,
  fileName,
  fileBuffer
) => {
  try {
    const uploadUrl = `https://graph.microsoft.com/v1.0/users/${userId}/drive/items/${folderId}:/${fileName}:/content`;
    await graphClient.api(uploadUrl).put(fileBuffer);
    console.log(`File uploaded successfully: ${fileName}`);
  } catch (error) {
    console.error("Error uploading file:", error);
    throw error;
  }
};

// Function to replace placeholders in a Word document
export const replacePlaceholders = async (fileBuffer, data) => {
  try {
    const zip = new PizZip(fileBuffer);

    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: { start: "{{", end: "}}" },
    });

    try {
      // Pass data directly to render() instead of using setData()
      doc.render(data);
    } catch (error) {
      if (error.properties && error.properties.errors instanceof Array) {
        const errorMessages = error.properties.errors
          .map((e) => `Error in template: ${e.properties.explanation}`)
          .join("\n");
        console.error("Template Error:", errorMessages);
      }
      throw error;
    }

    return doc.getZip().generate({ type: "nodebuffer" });
  } catch (error) {
    console.error("Error processing template:", error);
    throw error;
  }
};

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

export async function getFileIdByName(driveId, fileName) {
  const files = await getFileNamesAndIds(driveId);
  const file = files.find((file) => file.name === fileName);
  return file ? file.id : null; // Returns the file ID or null if not found
}

// const files = await listFilesInOneDrive(process.env.ONEDRIVE_ID);
// console.log(files);

// const response = await getFileNamesAndIds(process.env.ONEDRIVE_ID);
// console.log(response);

// const fileId = await getFileIdByName(
//   process.env.ONEDRIVE_ID,
//   "els-testing-client-XXXX.xlsx"
// );
// console.log(fileId);
