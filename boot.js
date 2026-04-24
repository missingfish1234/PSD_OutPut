(() => {
  const info = document.getElementById("info");
  if (!info) {
    return;
  }

  try {
    const ps = require("photoshop");
    const appName = ps?.app?.name || "Unknown";
    info.textContent = "Runtime ok. Host: " + appName;
  } catch (error) {
    info.textContent = "Runtime failed: " + (error && error.message ? error.message : String(error));
  }
})();
