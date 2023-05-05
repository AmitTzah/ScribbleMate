/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

const { generateContinuations, checkApiKey } = require("./gpt3/gpt3.js");

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

async function loadGPT3(api_key) {
  // Get the API key from the input box
  api_key.value = document.getElementById("api-key").value;

  console.log("The API key is: ");
  console.log(api_key.value);

  //remove the error icon if it exists
  const icon = document.querySelector(".icon-alert-triangle");
  if (icon) {
    icon.remove();
  }

  //remove the check icon if it exists
  const icon2 = document.querySelector(".icon-check");
  if (icon2) {
    icon2.remove();
  }

  //to the control-api-input element, add the class is-loading
  document.getElementById("control-api-input").classList.add("is-loading");

  // Check if the API key is valid
  const valid = await checkApiKey(api_key.value);

  //if not valid (false)
  if (!valid) {
    //add an error icon to the api-key input box
    const icon = document.createElement("span");
    icon.className = "icon is-small is-right";
    const icon2 = document.createElement("span");
    icon2.className = "icon-alert-triangle";
    icon.appendChild(icon2);

    //remove the is-loading class from the control-api-input element
    document.getElementById("control-api-input").classList.remove("is-loading");

    document.getElementById("api-key").insertAdjacentElement("afterend", icon);

    //if valid (true)
  } else {
    //add a check icon to the api-key input box
    const icon = document.createElement("span");
    icon.className = "icon is-small is-right";
    const icon2 = document.createElement("span");
    icon2.className = "icon-check";
    icon.appendChild(icon2);

    //remove the is-loading class from the control-api-input element
    document.getElementById("control-api-input").classList.remove("is-loading");

    document.getElementById("api-key").insertAdjacentElement("afterend", icon);
  }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    //define a global variable for the api key
    let api_key = { value: "" };

    //set an event listener for the api-button
    document.getElementById("api-key-button").addEventListener("click", function () {
      loadGPT3(api_key);
    });
  }
});
