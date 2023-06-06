//Module for handling the options provided by gpt3, including removing and inserting options into the document

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

  //Update the number of options
  numOptions.value = parseInt(document.getElementById("options-select").value);

  //remove all the textareas above numOptions currently in the generations div
  const generations = document.getElementById("generations");
  while (generations.childElementCount > 2 * numOptions.value - 1) {
    generations.removeChild(generations.lastChild);
  }

  for (let i = 0; i < numOptions.value; i++) {
    //if the option doesn't exist, create it (including all the ui elements required)
    if (!document.getElementById(`option ${i + 1}`)) {
      //the following are all the ui elements that need to be created
      const textarea = createTextarea(i);
      const control = createControl(textarea);
      const subtitle = createSubtitle(i);
      const insert_button = createInsertButton(i);
      const remove_button = createRemoveButton(i);

      insert_button.addEventListener("click", function () {
        handleInsertButtonClick(i, numOptions, currentRange, textInserted);
      });

      remove_button.addEventListener("click", function () {
        handleRemoveButtonClick(i, numOptions, currentRange, textInserted);
      });

      const nav = createNav(insert_button, remove_button);

      //hide remove button
      remove_button.style.display = "none";

      document.getElementById("generations").appendChild(subtitle);
      document.getElementById("generations").appendChild(control);
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

module.exports = {
  optionsSelect,
};
