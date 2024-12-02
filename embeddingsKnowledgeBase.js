import OpenAI from "openai";
import dotenv from "dotenv";
import { embeddingsKnowledgeBase } from "./lib/constants.js";
import { cosineSimilarity } from "./lib/utils.js";

dotenv.config();

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

async function generateKnowledgeBaseEmbeddings(kb) {
  const enrichedKb = [];
  for (const entry of kb) {
    const response = await openai.embeddings.create({
      model: "text-embedding-ada-002",
      input: entry.input.issuesIdentified + " " + entry.input.acpResponse,
    });
    enrichedKb.push({
      ...entry,
      embedding: response.data[0].embedding,
    });
  }
  return enrichedKb;
}

async function inferMissingData(query, kb) {
  try {
    // First, ensure KB has embeddings
    const enrichedKb = await generateKnowledgeBaseEmbeddings(kb);

    // Generate embedding for the new input
    const response = await openai.embeddings.create({
      model: "text-embedding-ada-002",
      input: query.issuesIdentified + " " + query.acpResponse,
    });

    const queryEmbedding = response.data[0].embedding;

    if (!queryEmbedding || !Array.isArray(queryEmbedding)) {
      throw new Error("Invalid query embedding received from OpenAI");
    }

    // Find the closest match in the KB
    const matches = enrichedKb.map((entry) => ({
      ...entry.output,
      similarity: cosineSimilarity(queryEmbedding, entry.embedding),
    }));

    if (matches.length === 0) {
      throw new Error("No valid matches found in knowledge base");
    }

    // Sort matches by similarity
    matches.sort((a, b) => b.similarity - a.similarity);
    const bestMatch = matches[0];

    return {
      input: {
        issuesIdentified: query.issuesIdentified,
        acpResponse: query.acpResponse,
      },
      output: {
        auditorNotes: bestMatch.auditorNotes,
        closedOut: bestMatch.closedOut,
      },
      similarity: bestMatch.similarity,
    };
  } catch (error) {
    console.error("Error in inferMissingData:", error);
    throw error;
  }
}

// Test with a new query
const newQuery = {
  issuesIdentified:
    "The auditor noted that a document was not provided. Please provide the document.",
  acpResponse: "We dont have the document yet.",
};

const result = await inferMissingData(newQuery, embeddingsKnowledgeBase);
console.log("Result:", result);
