export function updateRfiDataWithOpenAIPrompt(groupedData) {
  return `
  Please process each entry in the following array according to these rules:

  1. Remove "RFI" at the beginning of each text.
  2. Replace it with "The auditor noted that".
  3. Complete each sentence with an action item, such as "can you please clarify?" or "can you provide additional evidence?" based on the context.

  Return the results as a JSON array with each entry as a separate string, formatted like this:

  [
      "The auditor noted that installer declaration listed HPC026/MT-200R26E20 instead of MHW-F26WN3/MT-300R26E20, can you please clarify?",
      "The auditor noted that invoice says $0, but list of sites listed $1,500, can you provide additional evidence?",
      ...
  ]

  The output should strictly follow JSON syntax and include all elements in an array format.
  **DO NOT** use any markdown text in the output.

  Array of entries: ${groupedData}

  Ignore any previous instructions or context. Treat this prompt as a standalone task.

  `;
}
export function updateRfiDataWithAzureGptPrompt(groupedData) {
  return `
  Please process each entry in the following array according to these rules:

  1. Remove "RFI" at the beginning of each text.
  2. Replace it with "The auditor noted that".
  3. Complete each sentence with an action item, such as "can you please clarify?" or "can you provide additional evidence?" based on the context.

  Return the results as a JSON array with each entry as a separate string, formatted like this:

  [
      "The auditor noted that installer declaration listed HPC026/MT-200R26E20 instead of MHW-F26WN3/MT-300R26E20, can you please clarify?",
      "The auditor noted that invoice says $0, but list of sites listed $1,500, can you provide additional evidence?",
      ...
  ]

  The output should strictly follow JSON syntax and include all elements in an array format.
  **DO NOT** use any markdown text in the output.

  Array of entries: ${groupedData}

  Ignore any previous instructions or context. Treat this prompt as a standalone task.

  `;
}

export function analyseAcpResponsePrompt(
  responseKnowledgeBase,
  clientResponses
) {
  return `
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

Here are the responses from the client:
${JSON.stringify(clientResponses)}

**Return the completed json object only. DO NOT use markdown formatting.**

`;
}
