const { checkApiKey } = require("./gpt3/gpt3.js");

async function validateAndSaveApiKey(api_key) {
  api_key.value = document.getElementById("api-key").value;
  removeErrorIcon();
  removeCheckIcon();
  removeErrorMessage("api-input-error-message");
  showLoadingState();

  const valid = await checkApiKey(api_key.value);

  if (!valid) {
    removeLoadingState();
    addErrorIcon();
    addErrorMessage("Invalid API key or incorrect model", "api-input-field");
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

function removeErrorMessage(elementID) {
  const error = document.getElementById(elementID);
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

function addErrorMessage(message, elementID) {
  //this function takes in a message and an elementID and adds the message as an error message to the element with the given ID
  const error = document.createElement("p");
  error.id = "api-input-error-message";
  error.className = "help is-danger";
  error.innerText = message;

  document.getElementById(elementID).insertAdjacentElement("afterend", error);
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

module.exports = {
  validateAndSaveApiKey,
  addErrorMessage,
  removeErrorMessage,
};
