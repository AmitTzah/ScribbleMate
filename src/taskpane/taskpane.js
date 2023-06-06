/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

const { generateContinuations, checkApiKey } = require("./gpt3/gpt3.js");

const {
  optionsSelect,
  setLoadingAllOptions,
  resetOptions,
  updateOutputTextareas,
  removeLoadingAllClasses,
} = require("./options.js");

function removeTrailingNewline(text) {
  const lastChar = text[text.length - 1];
  if (lastChar === "\n") {
    return text.trimEnd();
  }
  return text;
}

async function suggestText(api_key, numOptions, textInserted, event) {
  // Add a loading icon to the suggest text button
  event.target.classList.add("is-loading");

  //add loading indicators to the textareas of all options
  setLoadingAllOptions(numOptions);

  resetOptions(numOptions, textInserted);

  const selectedText = document.getElementById("inputTextArea").value;

  const cleanedText = removeTrailingNewline(selectedText);

  const continuations = await generateContinuations(api_key.value, cleanedText, numOptions.value);

  updateOutputTextareas(continuations, numOptions);

  event.target.classList.remove("is-loading");
  removeLoadingAllClasses(numOptions);
}

async function validateAndSaveApiKey(api_key) {
  api_key.value = document.getElementById("api-key").value;
  removeErrorIcon();
  removeCheckIcon();
  removeErrorMessage();
  showLoadingState();

  const valid = await checkApiKey(api_key.value);

  if (!valid) {
    removeLoadingState();
    addErrorIcon();
    addErrorMessage();
  } else {
    removeLoadingState();
    addCheckIcon();
    showMainScreen();
  }
}

function removeErrorIcon() {
  const icon = document.querySelector(".icon-alert-triangle");
  if (icon) {
    icon.remove();
  }
}

function removeCheckIcon() {
  const icon = document.querySelector(".icon-check");
  if (icon) {
    icon.remove();
  }
}

function removeErrorMessage() {
  const error = document.getElementById("api-input-error-message");
  if (error) {
    error.remove();
  }
}

function showLoadingState() {
  document.getElementById("control-api-input").classList.add("is-loading");
}

function removeLoadingState() {
  document.getElementById("control-api-input").classList.remove("is-loading");
}

function addErrorIcon() {
  const icon = document.createElement("span");
  icon.className = "icon is-small is-right";
  const icon2 = document.createElement("span");
  icon2.className = "icon-alert-triangle";
  icon.appendChild(icon2);

  document.getElementById("api-key").insertAdjacentElement("afterend", icon);
}

function addErrorMessage() {
  const error = document.createElement("p");
  error.id = "api-input-error-message";
  error.className = "help is-danger";
  error.innerText = "This API key is invalid";

  document.getElementById("api-input-field").insertAdjacentElement("afterend", error);
}

function addCheckIcon() {
  const icon = document.createElement("span");
  icon.className = "icon is-small is-right";
  const icon2 = document.createElement("span");
  icon2.className = "icon-check";
  icon.appendChild(icon2);

  document.getElementById("api-key").insertAdjacentElement("afterend", icon);
}

function showMainScreen() {
  document.getElementById("login-screen").style.display = "none";
  document.getElementById("main-screen").style.display = "block";
}

// Function to update the inputTextArea with the selected text
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
