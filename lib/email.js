import axios from "axios";
import nodemailer from "nodemailer";
import fs from "fs";
import path from "path";
import dotenv from "dotenv";
import { getAccessToken } from "./msAuth.js";
import { fileURLToPath } from "url";
import { dirname } from "path";
import { Buffer } from "buffer";
dotenv.config();

const downloadFile = async (accessToken, fileId) => {
  const driveId = process.env.ONEDRIVE_ID;

  // Get the directory name from the current module
  const __filename = fileURLToPath(import.meta.url);
  const __dirname = dirname(__filename);
  const response = await axios.get(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}/content`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
      responseType: "arraybuffer",
    }
  );

  return response.data;
};

const sendEmailWithAttachment = async (fileBuffer, fileName) => {
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
        filename: fileName,
        content: fileBuffer,
        contentType:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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

export const emailRfiToClient = async (fileId, fileName) => {
  try {
    const accessToken = await getAccessToken(); // Use your existing function
    const fileBuffer = await downloadFile(accessToken, fileId);
    await sendEmailWithAttachment(fileBuffer, fileName);
  } catch (error) {
    console.error("Error: ", error);
  }
};
