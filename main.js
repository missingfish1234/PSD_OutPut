const { entrypoints } = require("uxp");
const { app } = require("photoshop");

function showAlert() {
  alert("PSD Export Pipeline loaded");
}

entrypoints.setup({
  commands: {
    showAlert,
  },
  panels: {
    vanilla: {
      show() {},
    },
  },
});

function renderLayerNames() {
  const target = document.getElementById("layers");
  if (!target) {
    return;
  }

  try {
    const doc = app.activeDocument;
    if (!doc) {
      target.textContent = "No active document";
      return;
    }

    const names = doc.layers.map((layer) => layer.name).sort((a, b) => (a < b ? -1 : a > b ? 1 : 0));
    target.innerHTML = `<ul>${names.map((name) => `<li>${name}</li>`).join("")}</ul>`;
  } catch (error) {
    target.textContent = `Runtime error: ${error.message}`;
  }
}

document.addEventListener("DOMContentLoaded", () => {
  const button = document.getElementById("btnPopulate");
  if (button) {
    button.addEventListener("click", renderLayerNames);
  }
});
