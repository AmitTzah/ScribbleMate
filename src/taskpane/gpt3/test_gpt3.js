const { generateContinuations, checkApiKey } = require("./gpt3.js");

async function test_generateContinuations(api_key) {
  const prompt =
    "I took a few. Fresh fruit wasnt a rarity for me these days, but the grapes were lovely nonetheless, just on the verge ";
  const continuations = await generateContinuations(
    api_key,
    prompt,
    (n = 5),
    (temperature = 1),
    (top_p = 1),
    (presence_penalty = 0.5),
    (frequency_penalty = 0.5),
    (stop = ["\n"]),
    (model = "gpt-3.5-turbo"),
    (max_tokens = 30)
  );
  console.log(continuations);
}

async function test_checkApiKey(api_key) {
  const valid = await checkApiKey(api_key);
  console.log(valid);
}

api_key = "test_api_key_here";
//test_checkApiKey(api_key);
//test_generateContinuations(api_key);
