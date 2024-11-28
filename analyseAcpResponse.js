import { responseKnowledgeBase } from "./knowledgeBase.js";
import { openAiQuery } from "./lib/openAI.js";

const clientResponse = {
  issuesIdentified:
    "The auditor noted that a CCEW has not been attached. Please provide the CCEW.",
  projectsAffected: ["J2288990"],
  response:
    "the form has been uploaded to the evidence pack. But it seems that you cannot see it?",
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

Asses the "issuesIdentified" and "response" fields in the knowledge base and create relationships between the combination of these two fields and the "auditorNotes" field. Use this relationship to fill in the auditorNotes field.

Add the auditorNotes, closedOut to the clientResponse JSON object.

**Only return the updated clientResponse JSON object, nothing else.**

Here is the knowledge base:
${JSON.stringify(responseKnowledgeBase)}

Here is the response from the client:
${JSON.stringify(clientResponse)}


`;

const response = await openAiQuery(prompt);

console.log(JSON.parse(response));
