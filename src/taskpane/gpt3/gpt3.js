const { Configuration, OpenAIApi } = require("openai");
const { api_key } = require("./../../../api_key.js");

const configuration = new Configuration({
  apiKey: api_key,
});

delete configuration.baseOptions.headers["User-Agent"];

const openai = new OpenAIApi(configuration);

async function generateContinuation(prompt) {
  const requestBody = {
    model: "gpt-3.5-turbo",
    max_tokens: 75,
    temperature: 0.8,
    top_p: 1,
    n: 1,
    stop: ["\n"],
    presence_penalty: 0.5,
    frequency_penalty: 0.5,
    messages: [
      {
        role: "system",
        content:
          "You are an AI writing assistant for a novelist. Your goal is to use creative and descriptive language to complete the provided text without introducing new plot points or dialogue. Your writing should be sophisticated and adult-oriented, dealing with mature topics and avoiding cliches. Your responses should enhance the writer's original vision, providing vivid, concrete details of the setting and characters in the style of J.K. Rowling. Ensure that your responses remain within the same paragraph as the provided text, and avoid introducing new characters or events that are not directly related to the provided idea.",
      },
      {
        role: "user",
        content: prompt,
      },
    ],
  };

  const response = await openai.createChatCompletion(requestBody);
  return response;
}

module.exports = {
  generateContinuation,
};
