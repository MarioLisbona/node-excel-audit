import { getGraphClient } from "./lib/msAuth.js";
import { extractResponsesForReport } from "./lib/worksheetProcessing.js";
import {
  downloadFile,
  uploadFile,
  updatePlaceholders,
} from "./lib/oneDrive.js";

// Main function
const main = async () => {
  try {
    const graphClient = await getGraphClient();

    // Input parameters (replace with actual values or arguments)
    const userId = process.env.USER_ID;
    // const templateFileId = "01FNQELGER3IH53IAQ4JCYCVCPQARA7SAZ"; // Replace with the file ID of your Word doc template
    const templateFileId = "01FNQELGEXFNT2IFI6AFHIZIEME6KZWEGQ"; // Replace with the file ID of your Word doc template
    const folderId = "01FNQELGGVWQQ6RELQ2BE3ETXPHL463HB4"; // Replace with the destination folder ID
    const newFileName = process.argv[2]; // New file name provided as a command-line argument

    if (!newFileName) {
      console.error("Please provide a new file name as an argument.");
      process.exit(1);
    }

    // Step 1: Download the template from OneDrive
    const templateBuffer = await downloadFile(
      graphClient,
      userId,
      templateFileId
    );

    // Step 2: Extract the column data from the RFI spreadsheet
    const columnData = await extractResponsesForReport(
      graphClient,
      userId,
      "01FNQELGFWWY3333PH6ZFJG5GCYY3Y22YV",
      "RFI Spreadsheet"
    );

    // Format the findings for the report
    const formattedFindings = columnData
      .filter((item) => item && item.trim()) // Remove empty or whitespace-only entries
      .map((item, index) => `${index + 1}. ${item.trim()}`); // Add numbering and single newline

    console.log({ columnData, formattedFindings });

    // Data for placeholder replacement
    const data = {
      date: "December 2, 2024",
      projectName: "Water heater replacement",
      clientName: "Mario Lisbona",
      auditorName: "Energy Link Services",
      companyName: "MLD",
      address: "123 Street Address, City, ST ZIP Code",
      phone: "(555) 555-5555",
      email: "contact@example.com",
      website: "www.example.com",
      items: [
        { name: "John", age: 30 },
        { name: "Jane", age: 25 },
      ],
      columnData: formattedFindings,
    };

    // Step 2: Replace placeholders in the template
    const updatedDocumentBuffer = await updatePlaceholders(
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
