const { checkApiKey } = require("./gpt3/gpt3.js");

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

module.exports = {
  validateAndSaveApiKey,
};
