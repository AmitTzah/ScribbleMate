//Module for handling the options provided by gpt3, including removing and inserting options into the document

function showOption(index, currentRange, textInsertedIndex) {
  //index is the index of the option to show
  //currentRange is the range object of the current selection
  //textInsertedIndex is the index of the option that has been inserted into the document, or -1 if no option has been inserted
  if (textInsertedIndex !== -1) {
    //text has been inserted into the document, we need to remove it
    previousOption = document.getElementById("option " + (textInsertedIndex + 1));
    removeOption(currentRange, previousOption);
  }

  //Update the text of the carousel element
  const carouselOption = document.getElementById("carousel-option");
  carouselOption.textContent = "Option " + (index + 1);

  const currentOption = document.getElementById("option " + (index + 1));

  insertOption(currentRange, currentOption);
}

function addFocusToCurrentOption(currentIndex) {
  const textarea = document.getElementById("option " + (currentIndex + 1));
  textarea.classList.add("is-focused");
}

function removeFocusFromAllOptions(numOptions) {
  for (let i = 0; i < numOptions; i++) {
    const textarea = document.getElementById("option " + (i + 1));
    textarea.classList.remove("is-focused");
  }
}

function CycleOptionsEventListeners(numOptions, currentIndex, currentRange) {
  //Initialize event listeners to the prev and next buttons

  const prevButton = document.getElementById("prevButton");
  const nextButton = document.getElementById("nextButton");

  prevButton.addEventListener("click", () => {
    oldIndex = currentIndex.value;
    currentIndex.value = (currentIndex.value - 1 + numOptions.value) % numOptions.value;
    showOption(currentIndex.value, currentRange, oldIndex);
    removeFocusFromAllOptions(numOptions.value);
    addFocusToCurrentOption(currentIndex.value);
  });

  nextButton.addEventListener("click", () => {
    oldIndex = currentIndex.value;
    currentIndex.value = (currentIndex.value + 1) % numOptions.value;
    showOption(currentIndex.value, currentRange, oldIndex);
    removeFocusFromAllOptions(numOptions.value);
    addFocusToCurrentOption(currentIndex.value);
  });
}

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
    range.insertText(" " + option.value + " ", Word.InsertLocation.end);
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

function optionsSelect(numOptions) {
  //this function is called when the number of options is changed (via the select element) and when the page is loaded
  //It updates the generations div to have the correct number of textareas and buttons, and sets up the event listeners for the buttons and textareas
  updateNumOptions(numOptions);
  removeAllOptions();
  createMissingOptions(numOptions);
}

function updateNumOptions(numOptions) {
  const optionsSelectElement = document.getElementById("options-select");
  numOptions.value = parseInt(optionsSelectElement.value);
}

function removeAllOptions() {
  const generations = document.getElementById("generations");
  //remove all generations.childElementCount
  while (generations.childElementCount > 2) {
    generations.removeChild(generations.lastChild);
  }
}

function createMissingOptions(numOptions) {
  for (let i = 0; i < numOptions.value; i++) {
    const optionId = `option ${i + 1}`;
    if (!document.getElementById(optionId)) {
      const textarea = createTextarea(i);
      const control = createControl(textarea);
      const subtitle = createSubtitle(i);

      const generations = document.getElementById("generations");
      appendElements(generations, [subtitle, control]);
    }
  }
}

function appendElements(parent, elements) {
  for (const element of elements) {
    parent.appendChild(element);
  }
}

//Add loading indicators to the textareas of all options
function setLoadingAllOptions(numOptions) {
  for (let i = 1; i <= numOptions.value; i++) {
    setLoadingClasses(i);
  }
}

function resetOptions(numOptions, currentIndex) {
  // This function resets the options to their default state
  //The textareas are cleared, the insert buttons are enabled, and the remove buttons are hidden

  currentIndex.value = -1;
  for (let i = 1; i <= numOptions.value; i++) {
    clearTextarea(i);
  }
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

function updateOutputTextareas(continuations, numOptions) {
  //update the textareas of all options
  for (let i = 0; i < numOptions.value; i++) {
    document.getElementById(`option ${i + 1}`).value = continuations[i];
  }
}

function removeLoadingAllClasses(numOptions) {
  //remove the loading indicators from the textareas of all options
  for (let i = 1; i <= numOptions.value; i++) {
    removeLoadingClasses(i);
  }
}

function removeLoadingClasses(optionIndex) {
  //remove the loading indicators from the textarea of the given option
  const parentDiv = document.getElementById(`option ${optionIndex}`).parentElement;
  parentDiv.classList.remove("is-loading");
  parentDiv.classList.remove("is-large");
}

function removeOptionEventListener(currentIndex, currentRange) {
  //if currentIndex.value is equal to -1, then there is no option to remove
  if (currentIndex.value === -1) {
    return;
  }

  option = document.getElementById(`option ${currentIndex.value + 1}`);

  removeOption(currentRange, option);
}

module.exports = {
  optionsSelect,
  setLoadingAllOptions,
  resetOptions,
  updateOutputTextareas,
  removeLoadingAllClasses,
  CycleOptionsEventListeners,
  removeOptionEventListener,
};
