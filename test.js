import { azureGptQuery } from "./lib/azureGpt.cjs";

const response = await azureGptQuery("What is the capital of France?");
console.log(response);
