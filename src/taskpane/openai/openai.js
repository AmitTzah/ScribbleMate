const OpenAI = require("openai");

function getContinuationsContent(response, n) {
  const continuations = [];
  for (let i = 0; i < n; i++) {
    continuations.push(response.choices[i].message.content);
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
  model = "gpt-4o",
  max_tokens = 60,
  system_message = ""
) {
  const clientConfig = {
    apiKey: api_key,
    dangerouslyAllowBrowser: true,
  };

  const client = new OpenAI(clientConfig);

  const response = await client.chat.completions.create({
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
        content: system_message,
      },
      {
        role: "user",
        content: prompt,
      },
    ],
  });

  console.log("The response was from GPT was: ");
  console.log(response);

  return getContinuationsContent(response, n);
}

async function checkApiKey(api_Key) {
  const client = new OpenAI({
    apiKey: api_Key,
    dangerouslyAllowBrowser: true,
  });

  try {
    const response = await client.completions.create({
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

export { generateContinuations, checkApiKey };
