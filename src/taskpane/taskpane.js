/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

const { generateContinuations } = require("./gpt3/gpt3.js");

const {
  optionsSelect,
  setLoadingAllOptions,
  resetOptions,
  updateOutputTextareas,
  removeLoadingAllClasses,
  CycleOptionsEventListeners,
} = require("./options.js");

const { validateAndSaveApiKey } = require("./login-screen.js");

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
function initializeEventListeners(api_key, currentRange, numOptions, textInserted, currentIndex) {
  //initialize the event listeners for the options arrow buttons
  CycleOptionsEventListeners(numOptions, currentIndex);

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

    console.log("numOptions", numOptions);

    //Global variable to store whether the text is inserted or not
    let textInserted = { value: false };

    //Global variable to store the current index of the inserted option
    let currentIndex = { value: 0 };

    //initialize the event listeners
    initializeEventListeners(api_key, currentRange, numOptions, textInserted, currentIndex);
  }
});
