import { getGraphClient } from "./lib/msAuth.js";
import fs from "fs/promises";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";

// Function to download a file from OneDrive
const downloadFile = async (graphClient, userId, fileId) => {
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
const uploadFile = async (
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
const replacePlaceholders = async (fileBuffer, data) => {
  try {
    const zip = new PizZip(fileBuffer);

    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: { start: "{{", end: "}}" },
    });

    doc.setData(data);

    try {
      doc.render();
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

// Main function
const main = async () => {
  try {
    const graphClient = await getGraphClient();

    // Input parameters (replace with actual values or arguments)
    const userId = process.env.USER_ID;
    const templateFileId = "01FNQELGEXFNT2IFI6AFHIZIEME6KZWEGQ"; // Replace with the file ID of your Word doc template
    const folderId = "01FNQELGGVWQQ6RELQ2BE3ETXPHL463HB4"; // Replace with the destination folder ID
    const newFileName = process.argv[2]; // New file name provided as a command-line argument

    if (!newFileName) {
      console.error("Please provide a new file name as an argument.");
      process.exit(1);
    }

    // Data for placeholder replacement
    const data = {
      clientName: "John Doe",
      parra1: "Thank you for reaching out to us. We appreciate your business.",
      parra2:
        "If you have any questions, feel free to contact us at support@example.com.",
      companyName: "Your Company Name",
      address: "123 Street Address, City, ST ZIP Code",
      phone: "(555) 555-5555",
      email: "contact@example.com",
      website: "www.example.com",
    };

    // Step 1: Download the template from OneDrive
    const templateBuffer = await downloadFile(
      graphClient,
      userId,
      templateFileId
    );

    // Step 2: Replace placeholders in the template
    const updatedDocumentBuffer = await replacePlaceholders(
      templateBuffer,
      data
    );

    // Step 3: Upload the updated document to OneDrive
    await uploadFile(
      graphClient,
      userId,
      folderId,
      newFileName,
      updatedDocumentBuffer
    );

    console.log("Document processed and uploaded successfully!");
  } catch (error) {
    console.error("Error:", error);
  }
};

main();
