//Module for handling the options provided by gpt3, including removing and inserting options into the document

async function basicSearchRemoval(context, inputRange, fullSearchterm) {
  //this function removes the fullSearchterm from the inputRange
  //since range.search doesn't work with search terms longer than 255 characters, we need to split the search term into multiple parts
  //and remove each part individually
  //This function assumes that inputRange indeed contains the fullSearchterm

  let restOfSearchterm = fullSearchterm;

  while (restOfSearchterm.length > 0) {
    if (restOfSearchterm.length < 255) {
      await removeSearchResult(context, inputRange, restOfSearchterm);
      restOfSearchterm = "";
    } else {
      const firstPart = restOfSearchterm.slice(0, 255);
      restOfSearchterm = restOfSearchterm.slice(255);
      await removeSearchResult(context, inputRange, firstPart);
    }
  }
}

async function removeSearchResult(context, inputRange, searchTerm) {
    
  //remove the last search result from the inputRange
  //assuming the search term is not longer than 255 characters
  const searchResults = inputRange.search(searchTerm);
  searchResults.load("items");
  await context.sync();

  if (searchResults.items.length > 0) {
    const lastSearchResult = searchResults.items[searchResults.items.length - 1];
    lastSearchResult.delete();
    await context.sync();
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

function createTextarea(i) {
  const textarea = document.createElement("textarea");
  textarea.id = `option ${i + 1}`;
  textarea.className = "textarea";
  textarea.readOnly = true;
  textarea.placeholder = "The generations will appear here.";
  return textarea;
}

function createControl(textarea) {
  const control = document.createElement("div");
  control.className = "control";
  control.appendChild(textarea);
  return control;
}

function createSubtitle(i) {
  const subtitle = document.createElement("p");
  subtitle.className = "subtitle mt-2";
  subtitle.innerText = `Option ${i + 1}:`;
  return subtitle;
}

function createInsertButton(i) {
  const insert_button = document.createElement("button");
  insert_button.id = `insert-option-${i + 1}`;
  insert_button.className = "button is-info is-small";
  insert_button.innerText = "Insert";
  return insert_button;
}

function createRemoveButton(i) {
  const remove_button = document.createElement("button");
  remove_button.id = `remove-option-${i + 1}`;
  remove_button.className = "button is-info is-small";
  remove_button.innerText = "Remove";
  return remove_button;
}

function createNav(insert_button, remove_button) {
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
  return nav;
}

function handleInsertButtonClick(i, numOptions, currentRange, textInserted) {
  const option = document.getElementById(`option ${i + 1}`);
  insertOption(currentRange, option);

  if (option.value !== "") {
    textInserted.value = true;
    //show remove button
    remove_button = document.getElementById(`remove-option-${i + 1}`);
    remove_button.style.display = "inline-block";

    //grey out all other insert buttons
    for (let j = 0; j < numOptions.value; j++) {
      const insert_button = document.getElementById(`insert-option-${j + 1}`);
      insert_button.disabled = true;
    }
  }
}

function handleRemoveButtonClick(i, numOptions, currentRange, textInserted) {
  const option = document.getElementById(`option ${i + 1}`);
  removeOption(currentRange, option);
  textInserted.value = false;

  //hide remove button
  remove_button = document.getElementById(`remove-option-${i + 1}`);
  remove_button.style.display = "none";

  //ungrey out all the other insert buttons
  for (let j = 0; j < numOptions.value; j++) {
    const insert_button = document.getElementById(`insert-option-${j + 1}`);
    insert_button.disabled = false;
  }
}

function optionsSelect(numOptions, currentRange, textInserted) {
  //this function is called when the number of options is changed (via the select element) and when the page is loaded
  //It updates the generations div to have the correct number of textareas and buttons, and sets up the event listeners for the buttons and textareas
  updateNumOptions(numOptions);
  removeExcessOptions(numOptions);
  createMissingOptions(numOptions, currentRange, textInserted);
}

function updateNumOptions(numOptions) {
  const optionsSelectElement = document.getElementById("options-select");
  numOptions.value = parseInt(optionsSelectElement.value);
}

function removeExcessOptions(numOptions) {
  const generations = document.getElementById("generations");
  while (generations.childElementCount > 2 * numOptions.value - 1) {
    generations.removeChild(generations.lastChild);
  }
}

function createMissingOptions(numOptions, currentRange, textInserted) {
  for (let i = 0; i < numOptions.value; i++) {
    const optionId = `option ${i + 1}`;
    if (!document.getElementById(optionId)) {
      const textarea = createTextarea(i);
      const control = createControl(textarea);
      const subtitle = createSubtitle(i);
      const insert_button = createInsertButton(i);
      const remove_button = createRemoveButton(i);

      setupButtonListeners(i, numOptions, currentRange, textInserted, insert_button, remove_button);
      setupOptionHoverHandlers(currentRange, textInserted, textarea);

      const nav = createNav(insert_button, remove_button);
      remove_button.style.display = "none";

      const generations = document.getElementById("generations");
      appendElements(generations, [subtitle, control, nav]);
    }
  }
}

function setupButtonListeners(optionIndex, numOptions, currentRange, textInserted, insert_button, remove_button) {
  insert_button.addEventListener("click", function () {
    handleInsertButtonClick(optionIndex, numOptions, currentRange, textInserted);
  });

  remove_button.addEventListener("click", function () {
    handleRemoveButtonClick(optionIndex, numOptions, currentRange, textInserted);
  });
}

function setupOptionHoverHandlers(currentRange, textInserted, textarea) {
  textarea.addEventListener("mouseenter", function (event) {
    hoverOverOption(currentRange, event, textInserted);
  });

  textarea.addEventListener("mouseleave", function (event) {
    hoverOverOption(currentRange, event, textInserted);
  });
}

function appendElements(parent, elements) {
  for (const element of elements) {
    parent.appendChild(element);
  }
}

module.exports = {
  optionsSelect,
};
