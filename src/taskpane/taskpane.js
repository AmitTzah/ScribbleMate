/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

const { generateContinuation } = require("./gpt3/gpt3.js");

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // Get the selected text
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    // Send the selected text to the GPT-3 API to generate a description
    const prompt = selection.text;
    const Continuation = await generateContinuation(prompt);

    console.log("The generated continuation was: ");
    console.log(Continuation);

    // Insert the generated continuation into the Word document
    selection.insertText(Continuation.data.choices[0].message.content, Word.InsertLocation.end);
    await context.sync();
  });
}
