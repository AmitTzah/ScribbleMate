/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

const { generateContinuations, checkApiKey } = require("./gpt3/gpt3.js");

async function suggestText(api_key, numOptions = 5) {
  // Get the selected text from the input textarea
  //Send the selected text to the GPT-3 API to generate a description
  //put each generated description into the output textareas: "option 1", "option 2", "option 3", "option 4", "option 5"

  // Get the selected text from input textarea
  const selectedText = document.getElementById("inputTextArea").value;

  //Send the selected text to the GPT-3 API to generate a description
  const continuations = await generateContinuations(api_key.value, selectedText, numOptions);

  //put each generated description into the output textareas: "option 1", "option 2", "option 3", "option 4", "option 5"
  for (let i = 0; i < numOptions; i++) {

    //if the textarea doesn't exist, create it
    if (!document.getElementById(`option ${i + 1}`)) {
      const textarea = document.createElement("textarea");
      textarea.id = `option ${i + 1}`;
      textarea.className = "textarea";
      textarea.readOnly = true;
      textarea.placeholder = "The generations will appear here.";

      const subtitle = document.createElement("p");
      subtitle.className = "subtitle mt-2";
      subtitle.innerText = `Option ${i + 1}:`;

      document.getElementById("generations").appendChild(subtitle);

      document.getElementById("generations").appendChild(textarea);
    }

    document.getElementById(`option ${i + 1}`).value = continuations[i];
  }
}

async function validateAndSaveApiKey(api_key) {
  // Get the API key from the input box
  api_key.value = document.getElementById("api-key").value;

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

  //remove the error message if it exists
  const error = document.getElementById("api-input-error-message");
  if (error) {
    error.remove();
  }

  //to the control-api-input element, add the class is-loading
  document.getElementById("control-api-input").classList.add("is-loading");

  // Check if the API key is valid
  const valid = await checkApiKey(api_key.value);

  if (!valid) {
    //remove the is-loading class from the control-api-input element
    document.getElementById("control-api-input").classList.remove("is-loading");
    //add an error icon to the api-key input box
    const icon = document.createElement("span");
    icon.className = "icon is-small is-right";
    const icon2 = document.createElement("span");
    icon2.className = "icon-alert-triangle";
    icon.appendChild(icon2);

    document.getElementById("api-key").insertAdjacentElement("afterend", icon);

    //add an error message to the api-key input box
    const error = document.createElement("p");
    error.id = "api-input-error-message";
    error.className = "help is-danger";
    error.innerText = "This API key is invalid";

    document.getElementById("api-input-field").insertAdjacentElement("afterend", error);
  } else {
    //remove the is-loading class from the control-api-input element
    document.getElementById("control-api-input").classList.remove("is-loading");

    //add a check icon to the api-key input box
    const icon = document.createElement("span");
    icon.className = "icon is-small is-right";
    const icon2 = document.createElement("span");
    icon2.className = "icon-check";
    icon.appendChild(icon2);

    document.getElementById("api-key").insertAdjacentElement("afterend", icon);

    //set the display of the "login-screen" to none
    //Remove the display:none from the "main-screen"

    document.getElementById("login-screen").style.display = "none";
    document.getElementById("main-screen").style.display = "block";
  }
}

// Function to update the textarea with the selected text
function updateSelectedText() {
  Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const selectedText = result.value;

      // if the selected text is empty then do nothing
      if (selectedText === "") {
        return;
      }
      const textarea = document.getElementById("inputTextArea");
      textarea.value = selectedText;
    }
  });
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    //define a global variable for the api key
    let api_key = { value: "" };

    //set an event listener for the api-button
    document.getElementById("api-key-button").addEventListener("click", function () {
      validateAndSaveApiKey(api_key);
    });

    // Add event handler for selection change
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, updateSelectedText);

    //add an event listener for the suggest-button
    document.getElementById("suggest-text-button").addEventListener("click", function () {
      suggestText(api_key);
    });
  }
});
