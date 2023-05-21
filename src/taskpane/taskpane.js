/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

const { generateContinuations, checkApiKey } = require("./gpt3/gpt3.js");

async function suggestText(api_key, numOptions) {
  // Get the selected text from the input textarea
  //Send the selected text to the GPT-3 API to generate a description
  //put each generated description into the output textareas: "option 1", "option 2", "option 3", "option 4", "option 5"

  // Get the selected text from input textarea
  let selectedText = document.getElementById("inputTextArea").value;

  //check if selected text contains a /n character at the end
  const lastChar = selectedText[selectedText.length - 1];
  if (lastChar === "\n") {
    //remove it
    selectedText = selectedText.slice(0, -1);
  }

  //Send the selected text to the GPT-3 API to generate a description
  const continuations = await generateContinuations(api_key.value, selectedText, numOptions.value);

  //put each generated description into the output textareas: "option 1", "option 2", "option 3", "option 4", "option 5"
  for (let i = 0; i < numOptions.value; i++) {
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

//Function for hover over options
function hoverOverOption(currentRange, event) {
  //get the option that was hovered over
  const option = event.target;

  //check if option.value is empty
  if (option.value === "") {
    return;
  }
  if (event.type === "mouseenter") {
    return Word.run(currentRange.range, async (context) => {
      //get the range of the selected text

      range = currentRange.range;
      range.load();
      await context.sync();

      //use the range property of the textarea to insert the option.value into the document
      range.insertText(option.value, Word.InsertLocation.end);
      range.load();
      await context.sync();

      //deselct the text
      range.select("end");
    });
  } else if (event.type === "mouseleave") {
    // Trigger your desired action or event when hovering ends

    return Word.run(currentRange.range, async (context) => {
      //get the range of the selected text

      textToRemove = option.value;

      range = currentRange.range;
      range.load();
      await context.sync();

      //use search method to find the option.value in the range
      var searchResults = range.search(textToRemove);

      //load the search results
      searchResults.load("items");

      await context.sync();

      //get the first search result
      var searchResult = searchResults.items[0];
      //remove the search result
      searchResult.delete();
    });
  }
}

//add an event listener for the options-select select element to update the number of options and thier event listeners
function optionsSelect(numOptions, currentRange) {
  //get the value of the selected option as an integer
  numOptions.value = parseInt(document.getElementById("options-select").value);

  //remove all the textareas above numOptions currently in the generations div
  const generations = document.getElementById("generations");
  while (generations.childElementCount > 2 * numOptions.value - 1) {
    generations.removeChild(generations.lastChild);
  }

  for (let i = 0; i < numOptions.value; i++) {
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

      //add a hover event listener to the textarea
      textarea.addEventListener("mouseenter", function (event) {
        hoverOverOption(currentRange, event);
      });

      textarea.addEventListener("mouseleave", function (event) {
        hoverOverOption(currentRange, event);
      });
    }
  }
}

//Function to initialize all the event listeners
function initializeEventListeners(api_key, currentRange, numOptions) {
  //set an event listener for the api-button
  document.getElementById("api-key-button").addEventListener("click", function () {
    validateAndSaveApiKey(api_key);
  });

  // Add event handler for text selection change
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, function () {
    updateSelectedText(currentRange);
  });

  //add an event listener for the suggest-button
  document.getElementById("suggest-text-button").addEventListener("click", function () {
    suggestText(api_key, numOptions);
  });

  //add an event listener for the options-select select element to update the number of options and thier event listeners
  document.getElementById("options-select").addEventListener("change", function () {
    optionsSelect(numOptions, currentRange);
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

    //initialize the event listeners
    initializeEventListeners(api_key, currentRange, numOptions);
  }
});
