/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

const { generateContinuations, checkApiKey } = require("./gpt3/gpt3.js");

async function suggestText(api_key, numOptions, textInserted) {
  textInserted.value = false;

  //hide remove button from all options, if it exists, the id of that elemnt is "remove-option-i
  for (let i = 1; i <= numOptions.value; i++) {
    const removeButton = document.getElementById(`remove-option-${i}`);
    removeButton.style.display = "none";
    const insert_button = document.getElementById(`insert-option-${i}`);
    insert_button.disabled = false;
  }

  // Get the selected text from the input textarea
  //Send the selected text to the GPT-3 API to generate a description
  //put each generated description into the output textareas: "option 1", "option 2", "option 3", "option 4", "option 5"

  // Get the selected text from input textarea
  let selectedText = document.getElementById("inputTextArea").value;

  //check if selected text contains a /n character at the end
  const lastChar = selectedText[selectedText.length - 1];
  if (lastChar === "\n") {
    //remove it
    selectedText = selectedText.trimEnd();
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

async function basicSearchRemoval(context, inputRange, fullSearchterm) {
  //this function removes the fullSearchterm from the inputRange
  //since range.search doesn't work with search terms longer than 255 characters, we need to split the search term into multiple parts
  //and remove each part individually
  //This function assumes that inputRange indeed contains the fullSearchterm

  let restOfSearchterm = fullSearchterm;

  while (restOfSearchterm.length > 0) {
    //first check if restOfSearchterm is less than 255 characters
    if (restOfSearchterm.length < 255) {
      //search for the rest of the search term
      const searchResults = inputRange.search(restOfSearchterm);
      //load the search results
      searchResults.load("items");
      await context.sync();
      //get the last search result
      var searchResult = searchResults.items[searchResults.items.length - 1];

      //remove the search result
      searchResult.delete();
      await context.sync();
      //set restOfSearchterm to an empty string
      restOfSearchterm = "";
    } else {
      //get the first 255 characters of the search term
      const firstPart = restOfSearchterm.slice(0, 255);

      //get the rest of the search term
      restOfSearchterm = restOfSearchterm.slice(255);
      //search for the first part of the search term
      const searchResults = inputRange.search(firstPart);
      //load the search results
      searchResults.load("items");
      await context.sync();
      //get the first search result
      var searchResult = searchResults.items[0];
      //remove the search result
      searchResult.delete();
    }
  }
}

//Function for hover over options
function hoverOverOption(currentRange, event, textInserted) {
  //get the option that was hovered over
  const option = event.target;

  if (textInserted.value === true) {
    return;
  }

  if (event.type === "mouseenter") {
    return insertOption(currentRange, option);
  } else if (event.type === "mouseleave") {
    return removeOption(currentRange, option);
  }
}

function removeOption(currentRange, option) {
  //check if option.value is empty
  if (option.value === "") {
    return;
  }

  return Word.run(currentRange.range, async (context) => {
    //get the range of the selected text

    textToRemove = option.value;
    range = currentRange.range;
    range.load();
    await context.sync();
    await basicSearchRemoval(context, range, textToRemove);
  });
}

function insertOption(currentRange, option) {
  //check if option.value is empty
  if (option.value === "") {
    return;
  }

  return Word.run(currentRange.range, async (context) => {
    //get the range of the selected text

    range = currentRange.range;
    range.load();
    await context.sync();

    const trimmedText = range.text.trimEnd();
    range.insertText(trimmedText, "Replace");

    //use the range property of the textarea to insert the option.value into the document
    range.insertText(" " + option.value, Word.InsertLocation.end);
    range.load();
    await context.sync();

    //deselct the text
    //this makes the view jump to the inserted text
    //range.select("end");
  });
}

//add an event listener for the options-select select element to update the number of options and thier event listeners
function optionsSelect(numOptions, currentRange, textInserted) {
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

      //add a button underneath each textarea
      //here's the format:<button id="insert-option-i" class="button is-info is-small">Insert</button>
      const insert_button = document.createElement("button");
      insert_button.id = `insert-option-${i + 1}`;
      insert_button.className = "button is-info is-small";
      insert_button.innerText = "Insert";
      remove_button = document.createElement("button");
      remove_button.id = `remove-option-${i + 1}`;
      remove_button.className = "button is-info is-small";
      remove_button.innerText = "Remove";

      //add an event listener to each button, using the insertOption function
      insert_button.addEventListener("click", function () {
        const option = document.getElementById(`option ${i + 1}`);
        insertOption(currentRange, option);

        if (option.value !== "") {
          textInserted.value = true;
          //show remove button
          remove_button = document.getElementById(`remove-option-${i + 1}`);
          remove_button.style.display = "inline-block";

          //grey out all the other insert buttons
          for (let j = 0; j < numOptions.value; j++) {
            if (j !== i) {
              const other_insert_button = document.getElementById(`insert-option-${j + 1}`);
              other_insert_button.disabled = true;
            }
          }
        }
      });

      remove_button.addEventListener("click", function () {
        const option = document.getElementById(`option ${i + 1}`);
        removeOption(currentRange, option);
        textInserted.value = false;

        //hide remove button
        remove_button = document.getElementById(`remove-option-${i + 1}`);
        remove_button.style.display = "none";

        //ungrey out all the other insert buttons
        for (let j = 0; j < numOptions.value; j++) {
          if (j !== i) {
            const other_insert_button = document.getElementById(`insert-option-${j + 1}`);
            other_insert_button.disabled = false;
          }
        }
      });

      const nav = document.createElement("nav");
      nav.className = "level is-mobile mt-4";
      const level_left = document.createElement("div");
      level_left.className = "level-left";
      const level_item_remove = document.createElement("div");
      level_item_remove.className = "level-item has-text-centered";
      const level_item_insert = document.createElement("div");
      level_item_insert.className = "level-item has-text-centered";
      level_item_insert.appendChild(insert_button);
      level_item_remove.appendChild(remove_button);
      level_left.appendChild(level_item_insert);
      level_left.appendChild(level_item_remove);
      nav.appendChild(level_left);

      //hide remove button
      remove_button.style.display = "none";

      document.getElementById("generations").appendChild(subtitle);

      document.getElementById("generations").appendChild(textarea);

      document.getElementById("generations").appendChild(nav);

      //add a hover event listener to the textarea
      textarea.addEventListener("mouseenter", function (event) {
        hoverOverOption(currentRange, event, textInserted);
      });

      textarea.addEventListener("mouseleave", function (event) {
        hoverOverOption(currentRange, event, textInserted);
      });
    }
  }
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
  document.getElementById("suggest-text-button").addEventListener("click", function () {
    suggestText(api_key, numOptions, textInserted);
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
