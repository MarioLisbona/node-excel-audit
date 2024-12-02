import { getGraphClient } from "./lib/msAuth.js";
import {
  downloadFile,
  uploadFile,
  replacePlaceholders,
} from "./lib/oneDrive.js";

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
