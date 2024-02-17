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
  removeOptionEventListener,
  highlightOptionEventListener,
  clearOptionsButtonEventListener,
} = require("./options.js");

const { validateAndSaveApiKey, addErrorMessage, removeErrorMessage } = require("./login-screen.js");

function removeTrailingNewline(text) {
  const lastChar = text[text.length - 1];
  if (lastChar === "\n") {
    return text.trimEnd();
  }
  return text;
}

async function suggestText(api_key, numOptions, event, currentIndex, selectedModel) {
  // Add a loading icon to the suggest text button
  event.target.classList.add("is-loading");

  //add loading indicators to the textareas of all options
  setLoadingAllOptions(numOptions);

  resetOptions(numOptions, currentIndex);

  const selectedText = document.getElementById("inputTextArea").value;

  const cleanedText = removeTrailingNewline(selectedText);
  removeErrorMessage("api-input-error-message");

  try {
    const continuations = await generateContinuations(
      api_key.value,
      cleanedText,
      numOptions.value,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      selectedModel.value,
      undefined
    );
    // Handle the successful response and continuations here
    updateOutputTextareas(continuations, numOptions);

    event.target.classList.remove("is-loading");
    removeLoadingAllClasses(numOptions);

    //fire the change event for the nextButton element
    const nextButton = document.getElementById("nextButton");
    nextButton.dispatchEvent(new Event("click"));
  } catch (error) {
    console.error("An error occurred:", error);
    // Handle the error appropriately
    event.target.classList.remove("is-loading");
    removeLoadingAllClasses(numOptions);

    //add error message to the suggest text button
    addErrorMessage(error, "suggest-text-nav");
  }
}

// Function to update the inputTextArea with the selected text
async function updateSelectedText(currentRange) {
  return Word.run(async (context) => {
    // Get the selected text
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    const selectedText = selection.text;
    // if the selected text is empty or a newline character or spaces, do nothing
    //use trim to remove the whitespace characters
    if (selectedText.trim() === "") {
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
function initializeEventListeners(api_key, currentRange, numOptions, currentIndex, selectedModel) {
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
    suggestText(api_key, numOptions, event, currentIndex, selectedModel);
  });

  //add an event listener for the options-select select element to update the number of options and thier event listeners
  document.getElementById("options-select").addEventListener("change", function () {
    optionsSelect(numOptions);
  });

  ////add an event listener for the model-select element to update the selectedModel
  document.getElementById("model-select").addEventListener("change", function () {
    selectedModel.value = document.getElementById("model-select").value;
  });

  //fire the change event on the options-select element to initialize the number of options and thier event listeners
  document.getElementById("options-select").dispatchEvent(new Event("change"));

  //initialize the event listeners for the options arrow buttons
  CycleOptionsEventListeners(numOptions, currentIndex, currentRange);

  removeButton = document.getElementById(`remove-option-button`);
  //initialize the event listener for the remove button
  removeButton.addEventListener("click", function (event) {
    removeOptionEventListener(currentIndex, currentRange);
  });

  //initialize the event listener for the highlight-option-checkbox
  document.getElementById("highlight-option-checkbox").addEventListener("change", function () {
    highlightOptionEventListener(currentRange, currentIndex);
  });

  //initialize the event listener for the clear-option-button
  document.getElementById("clear-option-button").addEventListener("click", function () {
    clearOptionsButtonEventListener(numOptions, currentIndex, currentRange);
  });
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    //Global variable for the api key
    let api_key = { value: "" };

    // Global variable to store the range object of the the input text.
    let currentRange = { range: null };

    //global variable to store the number of options
    let numOptions = { value: parseInt(document.getElementById("options-select").value) };

    //global variable to store the selected model name
    let selectedModel = { value: document.getElementById("model-select").value };

    //Global variable to store the current index of the inserted option (starts at 0)
    //If no text is inserted this will be set to -1
    let currentIndex = { value: -1 };

    //initialize the event listeners
    initializeEventListeners(api_key, currentRange, numOptions, currentIndex, selectedModel);
  }
});
