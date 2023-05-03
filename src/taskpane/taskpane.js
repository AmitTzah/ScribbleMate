/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

const { generateContinuations } = require("./gpt3/gpt3.js");

async function suggestText() {
  return Word.run(async (context) => {
    // Get the selected text
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    // Send the selected text to the GPT-3 API to generate a description
    const prompt = selection.text;
    const Continuations = await generateContinuations(prompt);

    console.log("The generated continuation was: ");
    console.log(Continuations[0]);

    // Insert the generated continuation into the Word document
    selection.insertText(Continuations[0], Word.InsertLocation.end);
    await context.sync();
  });
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = suggestText;
  }
});
