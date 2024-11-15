// Retrieve data from an Excel file
export const getExcelData = async (client, userId, workbookId, sheetName) => {
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
export const updateExcelData = async (
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
