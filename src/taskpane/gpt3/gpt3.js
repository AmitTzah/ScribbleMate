const { Configuration, OpenAIApi } = require("openai");

function getContinuationsContent(response, n) {
  //the array of continuations to be returned

  const continuations = [];
  for (let i = 0; i < n; i++) {
    continuations.push(response.data.choices[i].message.content);
  }
  return continuations;
}

async function generateContinuations(
  api_key,
  prompt,
  n = 5,
  temperature = 0.8,
  top_p = 1,
  presence_penalty = 0.6,
  frequency_penalty = 0.6,
  stop = ["\n", "."],
  model = "gpt-3.5-turbo-0301",
  max_tokens = 60
) {
  //this function generates n continuations of the prompt using the GPT-3 API
  //returns an array of continuations

  const configuration = new Configuration({
    apiKey: api_key,
  });
  delete configuration.baseOptions.headers["User-Agent"];

  const openai = new OpenAIApi(configuration);

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

  console.log("The response was from GPT was: ");
  console.log(response);

  return getContinuationsContent(response, n);
}

async function checkApiKey(api_Key) {
  //this function checks if the API key is valid
  //returns true if the API key is valid, false otherwise

  const configuration = new Configuration({
    apiKey: api_Key,
  });
  delete configuration.baseOptions.headers["User-Agent"];

  const openai = new OpenAIApi(configuration);
  try {
    const response = await openai.createCompletion({
      model: "babbage-002",
      prompt: "t",
      temperature: 0,
      max_tokens: 1,
    });
    console.log(response);
    return true;
  } catch (error) {
    console.log(error);
    return false;
  }
}

module.exports = {
  generateContinuations,
  checkApiKey,
};
