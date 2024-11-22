import { getGraphClient } from "./lib/msAuth.js";

// Function to write test data to a worksheet
export const writeTestDataToWorksheet = async () => {
  const graphClient = await getGraphClient();

  const testData = [
    ["Header1", "Header2"],
    ["Row1Col1", "Row1Col2"],
  ];

  await graphClient
    .api(
      `/drives/${process.env.ONEDRIVE_ID}/items/${process.env.NEW_TEST_SPREADSHEET_ID}/workbook/worksheets/${process.env.NEW_TEST_WORKSHEET_NAME}/range(address='A1:B2')`
    )
    .patch({
      values: testData,
    });
};

writeTestDataToWorksheet();
