import OpenAI from "openai";
import dotenv from "dotenv";
import { embeddingsKnowledgeBase } from "./lib/constants.js";
import { cosineSimilarity } from "./lib/utils.js";

dotenv.config();

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

async function inferMissingData(query, kb) {
  try {
    // Generate embedding for the new input
    const response = await openai.embeddings.create({
      model: "text-embedding-ada-002",
      input: query.issuesIdentified + " " + query.acpResponse,
    });

    const queryEmbedding = response.data[0].embedding;

    // Validate that we have a valid embedding
    if (!queryEmbedding || !Array.isArray(queryEmbedding)) {
      throw new Error("Invalid query embedding received from OpenAI");
    }

    // Find the closest match in the KB
    const matches = kb
      .filter((entry) => entry && entry.embedding)
      .map((entry) => ({
        ...entry.output,
        similarity: cosineSimilarity(queryEmbedding, entry.embedding),
      }));

    if (matches.length === 0) {
      throw new Error("No valid matches found in knowledge base");
    }

    // Sort matches by similarity
    matches.sort((a, b) => b.similarity - a.similarity);
    const bestMatch = matches[0];

    // Return the inferred output
    return {
      input: {
        issuesIdentified: query.issuesIdentified,
        acpResponse: query.acpResponse,
      },
      output: {
        auditorNotes: bestMatch.auditorNotes,
        closedOut: bestMatch.closedOut,
      },
    };
  } catch (error) {
    console.error("Error in inferMissingData:", error);
    throw error;
  }
}

// Test with a new query
const newQuery = {
  issuesIdentified: "The list of sites lists $500 for a service that is free.",
  acpResponse: "This was due to a manual error, the service is indeed free.",
};

const result = await inferMissingData(newQuery, embeddingsKnowledgeBase);
console.log("Result:", result);