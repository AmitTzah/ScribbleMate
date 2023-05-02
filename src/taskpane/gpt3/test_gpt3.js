const { generateContinuation } = require("./gpt3.js");

async function test() {
  const prompt =
    "I took a few. Fresh fruit wasnt a rarity for me these days, but the grapes were lovely nonetheless, just on the verge ";
  const continuation = await generateContinuation(prompt);
  console.log(continuation.data.choices);
  console.log(continuation.data.usage);
}

test();
