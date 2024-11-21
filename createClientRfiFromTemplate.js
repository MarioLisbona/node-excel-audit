import { getGraphClient } from "./lib/msAuth.js";
import { getFileIdByName } from "./returnFileIdfromFileName.js";

export const createClientRfiFromTemplate = async (fileId) => {
  const client = await getGraphClient({ cache: false });

  // Workbook - "els-testing.xlsx"
  // Worksheet - "RFI Spreadsheet"
  const sourceWorkbookId = "01FNQELGCAO75FVWXUEVF23O3LZBKISNGT";
  const sourceWorksheetName = "RFI Spreadsheet";

  // Workbook - "RFI Client Responses.xlsx"
  // Worksheet - "RFI Spreadsheet"
  const destinationWorkbookId = "01FNQELGCXMVOOF4GZN5F27JF4762JUXH3";
  const destinationWorksheetName = "RFI Spreadsheet - Client Name";
  // Extract the data from the source worksheet "RFI Spreadsheet" in workbook "els-testing.xlsx"
  const existingData = await client
    .api(
      `/drives/${process.env.ONEDRIVE_ID}/items/${sourceWorkbookId}/workbook/worksheets/${sourceWorksheetName}/range(address='A7:B46')`
    )
    .get();

  // Create a new worksheet in the existing workbook
  await client
    .api(
      `/drives/${process.env.ONEDRIVE_ID}/items/${destinationWorkbookId}/workbook/worksheets/add()`
    )
    .post({
      name: destinationWorksheetName, // Pass the name of the new worksheet in the request body
    });
};
