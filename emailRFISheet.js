import axios from "axios";
import nodemailer from "nodemailer";
import fs from "fs";
import path from "path";
import dotenv from "dotenv";
import { getAccessToken } from "./lib/msAuth.js";
import { fileURLToPath } from "url";
import { dirname } from "path";
import { getFileIdByName } from "./returnFileIdfromFileName.js";
dotenv.config();

// Assuming you have a working getAccessToken function
// const getAccessToken = async () => { ... };

// Get the directory name from the current module
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const downloadFile = async (accessToken, fileId) => {
  const driveId = process.env.ONEDRIVE_ID;
  const response = await axios.get(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}/content`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
      responseType: "stream",
    }
  );

  const filePath = path.join(__dirname, process.env.WORKSHEET_NAME_TO_EMAIL);
  const writer = fs.createWriteStream(filePath);
  response.data.pipe(writer);

  return new Promise((resolve, reject) => {
    writer.on("finish", () => resolve(filePath));
    writer.on("error", reject);
  });
};

const sendEmailWithAttachment = async (filePath) => {
  let transporter = nodemailer.createTransport({
    service: "gmail",
    auth: {
      user: process.env.EMAIL_USER,
      pass: process.env.EMAIL_PASS,
    },
  });

  let mailOptions = {
    from: process.env.EMAIL_USER,
    to: "mario.lisbona@gmail.com",
    subject: "RFI Response needed",
    text: "Hi. please find the attached RFI response spreadsheet. Add the response to the ACP Response column and send it back to the us.",
    attachments: [
      {
        filename: process.env.WORKSHEET_NAME_TO_EMAIL,
        path: filePath,
      },
    ],
  };

  try {
    let info = await transporter.sendMail(mailOptions);
    console.log("Email sent: " + info.response);
  } catch (error) {
    console.error("Error sending email: ", error);
  }
};

export const emailRfiToClient = async () => {
  try {
    const accessToken = await getAccessToken(); // Use your existing function
    const fileId = await getFileIdByName(
      process.env.ONEDRIVE_ID,
      process.env.WORKSHEET_NAME_TO_EMAIL
    ); // Replace with your actual file ID
    console.log(fileId);
    const filePath = await downloadFile(accessToken, fileId);
    await sendEmailWithAttachment(filePath);
  } catch (error) {
    console.error("Error: ", error);
  }
};
