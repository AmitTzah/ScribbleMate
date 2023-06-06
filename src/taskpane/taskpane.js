/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

const { generateContinuations, checkApiKey } = require("./gpt3/gpt3.js");

const { optionsSelect } = require("./options.js");

async function suggestText(api_key, numOptions, textInserted, event) {
  // Add a loading icon to the button
  event.target.classList.add("is-loading");

  //add loading indicators to the textareas of all options
  setLoadingAllOptions(numOptions);

  resetOptions(numOptions, textInserted);

  const selectedText = document.getElementById("inputTextArea").value;

  const cleanedText = removeTrailingNewline(selectedText);

  const continuations = await generateContinuations(api_key.value, cleanedText, numOptions.value);

  updateOutputTextareas(continuations, numOptions);

  event.target.classList.remove("is-loading");
  removeLoadingClasses(numOptions);
}

function resetOptions(numOptions, textInserted) {
  // This function resets the options to their default state
  //T textareas are cleared, the insert buttons are enabled, and the remove buttons are hidden

  textInserted.value = false;

  for (let i = 1; i <= numOptions.value; i++) {
    hideRemoveButton(i);
    enableInsertButton(i);
    clearTextarea(i);
  }
}

function hideRemoveButton(optionIndex) {
  const removeButton = document.getElementById(`remove-option-${optionIndex}`);
  removeButton.style.display = "none";
}

function enableInsertButton(optionIndex) {
  const insertButton = document.getElementById(`insert-option-${optionIndex}`);
  insertButton.disabled = false;
}

function clearTextarea(optionIndex) {
  document.getElementById(`option ${optionIndex}`).value = "";
}

function setLoadingClasses(optionIndex) {
  //Adds loading indicators to the textarea of the given option
  const parentDiv = document.getElementById(`option ${optionIndex}`).parentElement;
  parentDiv.classList.add("is-loading");
  parentDiv.classList.add("is-large");
}

//Add loading indicators to the textareas of all options
function setLoadingAllOptions(numOptions) {
  for (let i = 1; i <= numOptions.value; i++) {
    setLoadingClasses(i);
  }
}

function removeTrailingNewline(text) {
  const lastChar = text[text.length - 1];
  if (lastChar === "\n") {
    return text.trimEnd();
  }
  return text;
}

function updateOutputTextareas(continuations, numOptions) {
  for (let i = 0; i < numOptions.value; i++) {
    document.getElementById(`option ${i + 1}`).value = continuations[i];
  }
}

function removeLoadingClasses(numOptions) {
  for (let i = 1; i <= numOptions.value; i++) {
    const parentDiv = document.getElementById(`option ${i}`).parentElement;
    parentDiv.classList.remove("is-loading");
    parentDiv.classList.remove("is-large");
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
async function updateSelectedText(currentRange) {
  return Word.run(async (context) => {
    // Get the selected text
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    const selectedText = selection.text;
    // if the selected text is empty then do nothing
    if (selectedText === "") {
      return;
    }

    if (currentRange.range !== null) {
      currentRange.range.context.trackedObjects.remove(currentRange.range);
    }

    const textarea = document.getElementById("inputTextArea");
    textarea.value = selectedText;

    context.trackedObjects.add(selection);
    currentRange.range = selection;
  });
}

//Function to initialize all the event listeners
function initializeEventListeners(api_key, currentRange, numOptions, textInserted) {
  //set an event listener for the api-button
  document.getElementById("api-key-button").addEventListener("click", function () {
    validateAndSaveApiKey(api_key);
  });

  // Add event handler for text selection change
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, function () {
    updateSelectedText(currentRange);
  });

  //add an event listener for the suggest-button
  document.getElementById("suggest-text-button").addEventListener("click", function (event) {
    suggestText(api_key, numOptions, textInserted, event);
  });

  //add an event listener for the options-select select element to update the number of options and thier event listeners
  document.getElementById("options-select").addEventListener("change", function () {
    optionsSelect(numOptions, currentRange, textInserted);
  });

  //fire the change event on the options-select element to initialize the number of options and thier event listeners
  document.getElementById("options-select").dispatchEvent(new Event("change"));
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    //Global variable for the api key
    let api_key = { value: "" };

    // Global variable to store the range object of the the input text.
    let currentRange = { range: null };

    //global variable to store the number of options
    let numOptions = { value: parseInt(document.getElementById("options-select").value) };

    //Global variable to store whether the text is inserted or not
    let textInserted = { value: false };

    //initialize the event listeners
    initializeEventListeners(api_key, currentRange, numOptions, textInserted);
  }
});
