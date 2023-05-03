const { generateContinuations } = require("./gpt3.js");

async function test() {
  const prompt =
    "I took a few. Fresh fruit wasnt a rarity for me these days, but the grapes were lovely nonetheless, just on the verge ";
  const continuations = await generateContinuations(
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

test();
