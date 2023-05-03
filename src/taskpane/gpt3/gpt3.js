const { Configuration, OpenAIApi } = require("openai");
const { api_key } = require("./../../../api_key.js");

const configuration = new Configuration({
  apiKey: api_key,
});

delete configuration.baseOptions.headers["User-Agent"];

const openai = new OpenAIApi(configuration);

function getContinuationsContent(response, n) {
  //the array of continuations to be returned

  const continuations = [];
  for (let i = 0; i < n; i++) {
    continuations.push(response.data.choices[i].message.content);
  }
  return continuations;
}

async function generateContinuations(
  prompt,
  n = 1,
  temperature = 0.8,
  top_p = 1,
  presence_penalty = 0.5,
  frequency_penalty = 0.5,
  stop = ["\n"],
  model = "gpt-3.5-turbo",
  max_tokens = 75
) {
  //this function generates n continuations of the prompt using the GPT-3 API

  const requestBody = {
    model: model,
    max_tokens: max_tokens,
    temperature: temperature,
    top_p: top_p,
    n: n,
    stop: stop,
    presence_penalty: presence_penalty,
    frequency_penalty: frequency_penalty,
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

  console.log("The response was: ");
  console.log(response);

  return getContinuationsContent(response, n);
}

module.exports = {
  generateContinuations,
};
