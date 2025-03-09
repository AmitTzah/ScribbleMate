import { GoogleGenerativeAI } from "@google/generative-ai";

async function generateContinuations(
  apiKey,
  prompt,
  n = 5,
  temperature = 0.8,
  topP = 1,
  topK = 40,
  maxOutputTokens = 60,
  stopSequences = ["\n", "."],
  modelName = "gemini-pro",
  systemMessage = ""
) {
  const genAI = new GoogleGenerativeAI(apiKey);
  const model = genAI.getGenerativeModel({
    model: modelName,
    safetySettings: [
      { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_NONE" },
      { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE" },
      { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_NONE" },
      { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE" },
      { category: "HARM_CATEGORY_CIVIC_INTEGRITY", threshold: "BLOCK_NONE" },
    ],
    generationConfig: {
      temperature,
      topP,
      topK,
      maxOutputTokens,
      candidateCount: n,
      stopSequences,
    },
  });

    const result = await model.generateContent(prompt);
    const continuations = [];

    if (result.response.candidates) {
      result.response.candidates.forEach((candidate) => {
        continuations.push(candidate.content.parts[0].text);
      });
    } else {
        continuations.push(result.response.text());
    }

  return continuations;
}

export { generateContinuations };
