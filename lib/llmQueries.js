import { azureGptQuery } from "./azureGpt.cjs";
import { openAiQuery } from "./openAI.js";
import {
  updateRfiDataWithAzureGptPrompt,
  updateRfiDataWithOpenAIPrompt,
} from "./prompts.js";

export async function updateRfiDataWithOpenAI(data) {
  // Create the prompt for the OpenAI API
  const prompt = await updateRfiDataWithOpenAIPrompt(data);

  // Send the prompt to the OpenAI API and return the response
  const amendedRdiData = await openAiQuery(prompt);

  return amendedRdiData;
}

export async function updateRfiDataWithAzureGptQuery(data) {
  // Create the prompt for the OpenAI API
  const prompt = await updateRfiDataWithAzureGptPrompt(data);

  // Send the prompt to the OpenAI API and return the response
  const amendedRdiData = await azureGptQuery(prompt);

  return amendedRdiData;
}
