import { responseKnowledgeBase } from "./lib/constants.js";
import { openAiQuery } from "./lib/openAI.js";
import { azureGptQuery } from "./lib/azureGpt.cjs";

const clientResponse = {
  issuesIdentified:
    "The auditor noted that a CCEW has not been attached. Please provide the CCEW.",
  response: "We cannot locate the form at the moment. We will upload it ASAP.",
};
const clientResponse1 = {
  issuesIdentified:
    "The auditor noted that a CCEW has not been attached. Please provide the CCEW.",
  response:
    "We check with our admin team and they have confirmed that the CCEW has been uploaded to the evidence pack.",
};

const prompt = `
You are a helpful assistant that can help me analyze responses from clients that are being audited.

The knowledge base contains responses from previous requests for information (RFI) from auditors. The format of the knowledge base is JSON with the following fields:
- issuesIdentified: The issue that the auditor has identified.
- acpResponse: The response from the client to the RFI.
- auditorNotes: The notes from the auditor about the response.
- closedOut: Whether the issue has been closed out.

The workflow is as follows:
1. An auditor will identify an issue in a job that is being audited.
  - Attribute used: issuesIdentified
2. The auditor will then send an RFI to the client asking for more information regarding the issue.
3. The client will then respond to the RFI with a response.
  - Attribute used: acpResponse
4. Your job is to analyze the response and see if it is satisfactory and fill in the auditorNotes and closedOut fields:
  - Attribute used: auditorNotes, closedOut

The auditorNotes field should contain one of the following statuses:
- "Finding and recommendation"
- "Observation"
- "No issues"
- "RFI"

Provide additional information in the auditorNotes field if the response does not satisfactorily address the issue.

The closedOut field should be either "Y" or "N".

Your task is to analyze the knowledge base and identify correlations and patterns between the issuesIdentified and acpResponse attributes and their relationships to the auditorNotes and closedOut attributes.

Use these correlations to infer and generate the values for auditorNotes and closedOut when only issuesIdentified and acpResponse are provided. Ensure consistency with the patterns observed in the knowledge base.

You will be provided with an incomplete json object where only issuesIdentified and acpResponse are given, generate detailed auditorNotes using the specified status codes (F, M, or RFI) and relevant context, and determine the appropriate value for closedOut (Y or N).

Here is the knowledge base:
${JSON.stringify(responseKnowledgeBase)}

Here is the response from the client:
${JSON.stringify(clientResponse)}

**Return the completed json object only. DO NOT use markdown formatting.**

`;

const openAiResponse = await openAiQuery(prompt);
const azureGptResponse = await azureGptQuery(prompt);

console.log("OpenAI Response: ", JSON.parse(openAiResponse));
console.log("Azure GPT Response: ", JSON.parse(azureGptResponse));
