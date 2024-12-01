import { responseKnowledgeBase, clientResponses } from "./lib/constants.js";
import { openAiQuery } from "./lib/openAI.js";
import { azureGptQuery } from "./lib/azureGpt.cjs";
import { analyseAcpResponsePrompt } from "./lib/prompts.js";

const prompt = analyseAcpResponsePrompt(responseKnowledgeBase, clientResponses);

const openAiResponse = await openAiQuery(prompt);
const azureGptResponse = await azureGptQuery(prompt);

console.log("OpenAI Response: ", JSON.parse(openAiResponse));
console.log("Azure GPT Response: ", JSON.parse(azureGptResponse));
