function validateAndSaveApiKey(api_keys) {
  const openaiKey = document.getElementById("api-key").value.trim();
  const deepseekKey = document.getElementById("deepseek-key").value.trim();

  removeErrorIcon("api-key");
  removeErrorIcon("deepseek-key");
  removeCheckIcon();
  removeErrorMessage("openai-input-field-error-message");
  removeErrorMessage("deepseek-input-field-error-message");

  let hasError = false;

  if (!openaiKey) {
    addErrorIcon("api-key");
    addErrorMessage("OpenAI API key is required", "openai-input-field");
    hasError = true;
  }

  if (!deepseekKey) {
    addErrorIcon("deepseek-key");
    addErrorMessage("Deepseek API key is required", "deepseek-input-field");
    hasError = true;
  }

  if (hasError) {
    return false;
  }

  api_keys.openai = openaiKey;
  api_keys.deepseek = deepseekKey;

  addCheckIcon();
  showMainScreen();
  return true;
}

function removeErrorIcon(inputId) {
  const inputField = document.getElementById(inputId);
  const icon = inputField.parentElement.querySelector(".icon-alert-triangle");
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

function addErrorIcon(inputId) {
  const inputField = document.getElementById(inputId);
  const existingIcon = inputField.parentElement.querySelector(".icon-alert-triangle");
  if (existingIcon) return;

  const icon = document.createElement("span");
  icon.className = "icon is-small is-right";
  const icon2 = document.createElement("span");
  icon2.className = "icon-alert-triangle";
  icon.appendChild(icon2);

  inputField.insertAdjacentElement("afterend", icon);
}

function addErrorMessage(message, elementID) {
  //this function takes in a message and an elementID and adds the message as an error message to the element with the given ID
  const error = document.createElement("p");
  error.id = `${elementID}-error-message`;
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
