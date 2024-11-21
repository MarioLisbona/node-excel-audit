import axios from "axios";
import { getAccessToken } from "./lib/msAuth.js";
const getFileIdByName = async (accessToken, fileName) => {
  const driveId = process.env.ONEDRIVE_ID;
  const response = await axios.get(
    `https://graph.microsoft.com/v1.0/drives/b!Pw8sQ7LElUmhelxnYnooEDNStA9suUxPkV_boLlDFMZWakTGDVmtRZGLua34fA4e/root/search(q='els-testing.xlsx')`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    }
  );

  const files = response.data.value;
  if (files.length > 0) {
    return files[0].id; // Return the ID of the first matching file
  } else {
    throw new Error("File not found");
  }
};

const accessToken = await getAccessToken();
console.log(accessToken);
const fileId = await getFileIdByName(accessToken, "els-testing.xlsx");
console.log(fileId);
