const { app, core, constants } = require("photoshop");
const { storage } = require("uxp");

const fs = storage.localFileSystem;
const STORAGE_KEY = "psd-export-pipeline-settings";
const FOLDER_TOKEN_KEY = "psd-export-pipeline-folder-token";
const RELEASE_INFO = {
  version: "1.1.87",
  build: "v77",
  stamp: "2026-04-25-05",
};
const PNG_SAVE_COMPRESSION = 2;
const ENABLE_PNG_LOSSLESS_SLIMMING = false;
const PNG_LOSSLESS_SLIMMING_MIN_BYTES = 1024 * 1024;
const PNG_LOSSLESS_SLIMMING_MAX_BYTES = 8 * 1024 * 1024;
const ENABLE_DEEP_EXPORT_TRANSPARENCY_CHECK = false;
const EXPORT_MODAL_BATCH_SIZE = 4;
const EXPORT_MODAL_BATCH_COOLDOWN_MS = 40;
const ENABLE_COCOS_FAST_TRIM_HINT = true;
const PNG_COMPLETION_STABLE_POLLS = 2;
const QUICK_EXPORT_TIMEOUT_MS = 10000;
const ENABLE_SELECTION_QUICK_EXPORT = false;
const ENABLE_DUPLICATE_SAVEAS_FALLBACK = true;
const ENABLE_DOCUMENT_QUICK_EXPORT_FALLBACK = false;
const DEFAULT_SETTINGS = {
  mode: "normal",
  layerOrder: "auto",
  prefix: "",
  platformFolderKeyword: "",
  spineFolderKeyword: "spine",
  writeSpineFormat: true,
  writeSpineAtlas: true,
  spineJsonFileName: "skeleton.json",
  spineImagesPath: "images/",
  prefabTarget: "none",
  writePrefabPackage: true,
  selectedOnly: false,
  exportHidden: false,
  includeEffects: true,
  useFullPathNames: true,
  writeMetadata: true,
  recursiveNormal: true,
};

const SELECT_OPTIONS = {
  modeSelect: [
    { value: "normal", label: "一般輸出（平台圖）" },
    { value: "spine", label: "Spine 拆圖" },
  ],
  layerOrderSelect: [
    { value: "auto", label: "自動（平台=PSD，Spine=反轉）" },
    { value: "psd", label: "PSD 順序（上到下）" },
    { value: "reverse", label: "反轉順序（下到上）" },
  ],
  prefabTargetSelect: [
    { value: "none", label: "不匯出 Prefab" },
    { value: "cocos-3.8.8", label: "Cocos Creator 3.8.8" },
    { value: "unity-6.0", label: "Unity 6.0" },
    { value: "unity-6.3", label: "Unity 6.3" },
  ],
};

const ENGINE_PRESET_TARGETS = new Set(["unity-6.0", "unity-6.3", "cocos-3.8.8"]);

const state = {
  outputFolder: null,
  candidates: [],
  docInfo: null,
  busy: false,
  quickExportCapability: ENABLE_SELECTION_QUICK_EXPORT ? "unknown" : "unavailable",
  settings: { ...DEFAULT_SETTINGS },
};

const ui = {};
let settingsRefreshTimer = null;

document.addEventListener("DOMContentLoaded", () => {
  try {
    init();
  } catch (error) {
    console.error("Panel init failed", error);
    const fallback = document.createElement("pre");
    fallback.textContent = `Panel init failed:\n${error && error.message ? error.message : String(error)}`;
    fallback.style.whiteSpace = "pre-wrap";
    fallback.style.padding = "10px";
    fallback.style.color = "#ffb3b3";
    document.body.innerHTML = "";
    document.body.appendChild(fallback);
  }
});

function init() {
  ensureSettingsUi();
  bindUi();
  state.settings = loadSettings();
  applySettingsToUi();
  renderFolderState();
  renderSnapshot();
  setStatus("面板已載入。請先選擇輸出資料夾，再手動按「重新掃描 PSD」。", "ok");
}

function ensureSettingsUi() {
  let grid = document.querySelector(".pe-setting-grid");
  if (!grid) {
    const settingsSection = document.getElementById("settingsCard") || document.querySelector("section.pe-card");
    if (!settingsSection) {
      throw new Error("找不到設定區塊，無法初始化設定欄位。");
    }

    grid = document.createElement("div");
    grid.className = "pe-setting-grid";
    settingsSection.appendChild(grid);
  }

  grid.innerHTML = getSettingsGridMarkup();
  forceShowSettingsGrid(grid);
}

function getSettingsGridMarkup() {
  return `
    <div class="pe-basic-grid">
      <label class="pe-field">
        <span>模式</span>
        <button id="modeSelect" class="pe-select-btn" type="button" data-kind="select">-</button>
      </label>
      <label class="pe-field">
        <span>Prefab 目標平台</span>
        ${buildNativeSelectMarkup("prefabTargetSelect", "pe-native-select")}
      </label>
      <label class="pe-field pe-field--wide">
        <span>檔名前綴</span>
        <input id="prefixInput" class="pe-native-input" type="text" placeholder="例如 ui">
      </label>
    </div>
    <div class="pe-toggle-grid pe-toggle-grid-basic">
      <button id="includeEffectsCheckbox" class="pe-toggle-btn" type="button" data-kind="check" data-label="邊界包含圖層效果"></button>
      <button id="fullPathCheckbox" class="pe-toggle-btn" type="button" data-kind="check" data-label="檔名包含群組路徑"></button>
      <button id="metadataCheckbox" class="pe-toggle-btn" type="button" data-kind="check" data-label="輸出 metadata/layout.json"></button>
      <button id="prefabPackageCheckbox" class="pe-toggle-btn" type="button" data-kind="check" data-label="輸出平台 Prefab 建置包"></button>
    </div>
    <div class="pe-advanced-summary">進階選項（固定顯示）</div>
    <div class="pe-advanced-grid">
      <label class="pe-field">
        <span>圖層順序</span>
        <button id="layerOrderSelect" class="pe-select-btn" type="button" data-kind="select">-</button>
      </label>
      <label class="pe-field">
        <span>平台資料夾關鍵字（逗號分隔）</span>
        <input id="platformKeywordInput" class="pe-native-input" type="text" placeholder="例如 ui,platform">
      </label>
      <label class="pe-field">
        <span>Spine 資料夾關鍵字（逗號分隔）</span>
        <input id="spineKeywordInput" class="pe-native-input" type="text" placeholder="例如 spine,bone">
      </label>
      <label class="pe-field">
        <span>Spine JSON 檔名</span>
        <input id="spineJsonFileNameInput" class="pe-native-input" type="text" placeholder="skeleton.json">
      </label>
      <label class="pe-field">
        <span>Spine 圖片路徑（相對 spine 檔案）</span>
        <input id="spineImagesPathInput" class="pe-native-input" type="text" placeholder="images/">
      </label>
    </div>
    <div class="pe-toggle-grid pe-toggle-grid-advanced">
      <button id="selectedOnlyCheckbox" class="pe-toggle-btn" type="button" data-kind="check" data-label="只輸出目前選取圖層"></button>
      <button id="exportHiddenCheckbox" class="pe-toggle-btn" type="button" data-kind="check" data-label="包含隱藏圖層"></button>
      <button id="recursiveNormalCheckbox" class="pe-toggle-btn" type="button" data-kind="check" data-label="一般模式遞迴拆解群組內圖層"></button>
      <button id="spineFormatCheckbox" class="pe-toggle-btn" type="button" data-kind="check" data-label="輸出 Spine 格式（skeleton.json + atlas）"></button>
      <button id="spineAtlasCheckbox" class="pe-toggle-btn" type="button" data-kind="check" data-label="輸出 Spine atlas 檔案"></button>
    </div>
  `;
}

function buildNativeSelectMarkup(id, className) {
  const options = SELECT_OPTIONS[id] || [];
  return `<select id="${id}" class="${className}">${options.map((item) => `<option value="${escapeHtml(item.value)}">${escapeHtml(item.label)}</option>`).join("")}</select>`;
}

function forceShowSettingsGrid(grid) {
  if (!grid || !grid.style) {
    return;
  }

  applyImportantStyles(grid, {
    display: "block",
    "margin-top": "5px",
    visibility: "visible",
    opacity: "1",
    "min-height": "0",
    height: "auto",
    "max-height": "none",
    overflow: "visible",
    position: "static",
    transform: "none",
    "clip-path": "none",
  });

  Array.from(grid.children).forEach((child) => {
    if (!child || !child.style) {
      return;
    }

    const isToggleGrid = child.classList && child.classList.contains("pe-toggle-grid");
    const isGridBlock = child.classList && (child.classList.contains("pe-basic-grid") || child.classList.contains("pe-advanced-grid"));
    applyImportantStyles(child, (isToggleGrid || isGridBlock) ? {
      display: "grid",
      "grid-template-columns": "repeat(2, minmax(0, 1fr))",
      gap: "5px",
      visibility: "visible",
      opacity: "1",
      "min-height": "24px",
      height: "auto",
      "max-height": "none",
      overflow: "visible",
      position: "static",
      margin: "0 0 5px 0",
    } : {
      display: "block",
      visibility: "visible",
      opacity: "1",
      "min-height": "24px",
      height: "auto",
      "max-height": "none",
      overflow: "visible",
      position: "static",
      float: "none",
      clear: "both",
      margin: "0 0 5px 0",
      transform: "none",
    });

    if (child.classList && child.classList.contains("pe-field")) {
      applyImportantStyles(child, {
        "min-height": "48px",
        padding: "0",
      });
      const title = child.querySelector("span");
      if (title) {
        applyImportantStyles(title, {
          display: "block",
          "margin-bottom": "4px",
          "font-size": "12px",
          color: "#cdcdcd",
        });
      }
    }
  });

  grid.querySelectorAll("button, input, textarea, select").forEach((control) => {
    if (!control || !control.style) {
      return;
    }
    const isToggle = control.classList && control.classList.contains("pe-toggle-btn");
    applyImportantStyles(control, {
      display: "block",
      visibility: "visible",
      opacity: "1",
      "min-height": "28px",
      height: isToggle ? "auto" : "28px",
      "max-height": "none",
      padding: "4px 8px",
      border: "1px solid #5d5d5d",
      "border-radius": "6px",
      "background-color": "#2f2f2f",
      color: "#f3f3f3",
      "font-size": "13px",
      "line-height": "1.3",
      "text-align": "left",
      overflow: "visible",
      position: "static",
      float: "none",
      transform: "none",
      "clip-path": "none",
    });
  });

  [0, 50, 150, 300, 600].forEach((delay) => {
    setTimeout(() => {
      Array.from(grid.children).forEach((child) => {
        if (!child || !child.style) {
          return;
        }
        if (child.classList && (child.classList.contains("pe-toggle-grid") || child.classList.contains("pe-basic-grid") || child.classList.contains("pe-advanced-grid"))) {
          applyImportantStyles(child, {
            display: "grid",
            "grid-template-columns": "repeat(2, minmax(0, 1fr))",
            gap: "5px",
            visibility: "visible",
            opacity: "1",
            "min-height": "24px",
            height: "auto",
            "max-height": "none",
            position: "static",
            margin: "0 0 5px 0",
          });
        } else {
          applyImportantStyles(child, {
            display: "block",
            visibility: "visible",
            opacity: "1",
            "min-height": "24px",
            height: "auto",
            "max-height": "none",
            position: "static",
            margin: "0 0 5px 0",
          });
        }
      });

      grid.querySelectorAll("button, input, textarea, select").forEach((control) => {
        if (!control || !control.style) {
          return;
        }
        const isToggle = control.classList && control.classList.contains("pe-toggle-btn");
        applyImportantStyles(control, {
          display: "block",
          visibility: "visible",
          opacity: "1",
          "min-height": "28px",
          height: isToggle ? "auto" : "28px",
          "max-height": "none",
          "font-size": "13px",
          position: "static",
          transform: "none",
        });
      });
    }, delay);
  });
}

function applyImportantStyles(element, styleMap) {
  if (!element || !element.style || !styleMap) {
    return;
  }

  Object.keys(styleMap).forEach((prop) => {
    element.style.setProperty(prop, String(styleMap[prop]), "important");
  });
}

function bindUi() {
  ui.chooseFolderBtn = document.getElementById("chooseFolderBtn");
  ui.refreshBtn = document.getElementById("refreshBtn");
  ui.exportBtn = document.getElementById("exportBtn");
  ui.folderState = document.getElementById("folderState");
  ui.folderPath = document.getElementById("folderPath");
  ui.modeSelect = document.getElementById("modeSelect");
  ui.layerOrderSelect = document.getElementById("layerOrderSelect");
  ui.prefixInput = document.getElementById("prefixInput");
  ui.platformKeywordInput = document.getElementById("platformKeywordInput");
  ui.spineKeywordInput = document.getElementById("spineKeywordInput");
  ui.spineFormatCheckbox = document.getElementById("spineFormatCheckbox");
  ui.spineAtlasCheckbox = document.getElementById("spineAtlasCheckbox");
  ui.spineJsonFileNameInput = document.getElementById("spineJsonFileNameInput");
  ui.spineImagesPathInput = document.getElementById("spineImagesPathInput");
  ui.prefabTargetSelect = document.getElementById("prefabTargetSelect");
  ui.prefabPackageCheckbox = document.getElementById("prefabPackageCheckbox");
  ui.selectedOnlyCheckbox = document.getElementById("selectedOnlyCheckbox");
  ui.exportHiddenCheckbox = document.getElementById("exportHiddenCheckbox");
  ui.includeEffectsCheckbox = document.getElementById("includeEffectsCheckbox");
  ui.fullPathCheckbox = document.getElementById("fullPathCheckbox");
  ui.metadataCheckbox = document.getElementById("metadataCheckbox");
  ui.recursiveNormalCheckbox = document.getElementById("recursiveNormalCheckbox");
  ui.docName = document.getElementById("docName");
  ui.docSize = document.getElementById("docSize");
  ui.assetCount = document.getElementById("assetCount");
  ui.namePreview = document.getElementById("namePreview");
  ui.previewList = document.getElementById("previewList");
  ui.statusBox = document.getElementById("statusBox");

  assertUiBindings([
    "chooseFolderBtn",
    "refreshBtn",
    "exportBtn",
    "folderState",
    "folderPath",
    "modeSelect",
    "layerOrderSelect",
    "prefixInput",
    "platformKeywordInput",
    "spineKeywordInput",
    "spineFormatCheckbox",
    "spineAtlasCheckbox",
    "spineJsonFileNameInput",
    "spineImagesPathInput",
    "prefabTargetSelect",
    "prefabPackageCheckbox",
    "selectedOnlyCheckbox",
    "exportHiddenCheckbox",
    "includeEffectsCheckbox",
    "fullPathCheckbox",
    "metadataCheckbox",
    "recursiveNormalCheckbox",
    "docName",
    "docSize",
    "assetCount",
    "namePreview",
    "previewList",
    "statusBox",
  ]);

  ui.chooseFolderBtn.addEventListener("click", chooseOutputFolder);
  ui.refreshBtn.addEventListener("click", refreshCandidates);
  ui.exportBtn.addEventListener("click", runExport);

  const settingControls = [
    ui.modeSelect,
    ui.layerOrderSelect,
    ui.prefixInput,
    ui.platformKeywordInput,
    ui.spineKeywordInput,
    ui.spineFormatCheckbox,
    ui.spineAtlasCheckbox,
    ui.spineJsonFileNameInput,
    ui.spineImagesPathInput,
    ui.prefabTargetSelect,
    ui.prefabPackageCheckbox,
    ui.selectedOnlyCheckbox,
    ui.exportHiddenCheckbox,
    ui.includeEffectsCheckbox,
    ui.fullPathCheckbox,
    ui.metadataCheckbox,
    ui.recursiveNormalCheckbox,
  ];

  settingControls.forEach(bindSettingControl);
}

function bindSettingControl(element) {
  if (!element) {
    return;
  }

  const kind = element.dataset && element.dataset.kind;
  if (kind === "select") {
    element.addEventListener("click", () => {
      const previousValue = getControlValue(element);
      cycleSelectControl(element);
      if (element.id === "prefabTargetSelect") {
        const nextValue = getControlValue(element);
        if (nextValue !== previousValue) {
          applyEnginePresetForTarget(nextValue);
        }
      }
      queueSettingsRefresh();
    });
    return;
  }

  if (kind === "check") {
    element.addEventListener("click", () => {
      const next = !getControlChecked(element);
      setControlChecked(element, next);
      queueSettingsRefresh();
    });
    return;
  }

  const tagName = (element.tagName || "").toUpperCase();
  if (tagName === "SELECT") {
    element.addEventListener("change", () => {
      if (element.id === "prefabTargetSelect") {
        applyEnginePresetForTarget(getControlValue(element));
      }
      queueSettingsRefresh();
    });
    return;
  }

  if (kind === "text" || element.isContentEditable || tagName === "INPUT" || tagName === "TEXTAREA") {
    element.addEventListener("input", queueSettingsRefresh);
    element.addEventListener("blur", queueSettingsRefresh);
    return;
  }

  element.addEventListener("change", queueSettingsRefresh);
}

function queueSettingsRefresh() {
  if (settingsRefreshTimer) {
    clearTimeout(settingsRefreshTimer);
  }

  settingsRefreshTimer = setTimeout(async () => {
    settingsRefreshTimer = null;
    try {
      state.settings = readSettingsFromUi();
      persistSettings();
      await refreshCandidates();
    } catch (error) {
      setStatus(`更新設定失敗：${error.message}`, "error");
    }
  }, 120);
}

function applyEnginePresetForTarget(target) {
  const normalizedTarget = String(target || "").trim().toLowerCase();
  if (!ENGINE_PRESET_TARGETS.has(normalizedTarget)) {
    return;
  }

  setControlValue(ui.modeSelect, "normal");
  setControlValue(ui.layerOrderSelect, "auto");
  setControlChecked(ui.prefabPackageCheckbox, true);
  setControlChecked(ui.selectedOnlyCheckbox, false);
  setControlChecked(ui.exportHiddenCheckbox, false);
  setControlChecked(ui.includeEffectsCheckbox, true);
  setControlChecked(ui.fullPathCheckbox, true);
  setControlChecked(ui.metadataCheckbox, true);
  setControlChecked(ui.recursiveNormalCheckbox, true);
  setControlChecked(ui.spineFormatCheckbox, false);
  setControlChecked(ui.spineAtlasCheckbox, false);
}

function cycleSelectControl(control) {
  const options = SELECT_OPTIONS[control.id] || [];
  if (!options.length) {
    return;
  }

  const current = getControlValue(control);
  const index = options.findIndex((item) => item.value === current);
  const nextIndex = index >= 0 ? (index + 1) % options.length : 0;
  setControlValue(control, options[nextIndex].value);
}

function assertUiBindings(keys) {
  const missing = keys.filter((key) => !ui[key]);
  if (missing.length) {
    throw new Error(`Missing UI elements: ${missing.join(", ")}`);
  }
}

function loadSettings() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return { ...DEFAULT_SETTINGS };
    }
    return { ...DEFAULT_SETTINGS, ...JSON.parse(raw) };
  } catch (error) {
    console.warn("Failed to load settings", error);
    return { ...DEFAULT_SETTINGS };
  }
}

function persistSettings() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.settings));
}

function applySettingsToUi() {
  setControlValue(ui.modeSelect, state.settings.mode);
  setControlValue(ui.layerOrderSelect, state.settings.layerOrder);
  setControlValue(ui.prefixInput, state.settings.prefix);
  setControlValue(ui.platformKeywordInput, state.settings.platformFolderKeyword);
  setControlValue(ui.spineKeywordInput, state.settings.spineFolderKeyword);
  setControlChecked(ui.spineFormatCheckbox, state.settings.writeSpineFormat);
  setControlChecked(ui.spineAtlasCheckbox, state.settings.writeSpineAtlas);
  setControlValue(ui.spineJsonFileNameInput, state.settings.spineJsonFileName);
  setControlValue(ui.spineImagesPathInput, state.settings.spineImagesPath);
  setControlValue(ui.prefabTargetSelect, state.settings.prefabTarget);
  setControlChecked(ui.prefabPackageCheckbox, state.settings.writePrefabPackage);
  setControlChecked(ui.selectedOnlyCheckbox, state.settings.selectedOnly);
  setControlChecked(ui.exportHiddenCheckbox, state.settings.exportHidden);
  setControlChecked(ui.includeEffectsCheckbox, state.settings.includeEffects);
  setControlChecked(ui.fullPathCheckbox, state.settings.useFullPathNames);
  setControlChecked(ui.metadataCheckbox, state.settings.writeMetadata);
  setControlChecked(ui.recursiveNormalCheckbox, state.settings.recursiveNormal);
}

function readSettingsFromUi() {
  return {
    mode: getControlValue(ui.modeSelect),
    layerOrder: getControlValue(ui.layerOrderSelect),
    prefix: getControlValue(ui.prefixInput).trim(),
    platformFolderKeyword: getControlValue(ui.platformKeywordInput).trim(),
    spineFolderKeyword: getControlValue(ui.spineKeywordInput).trim(),
    writeSpineFormat: getControlChecked(ui.spineFormatCheckbox),
    writeSpineAtlas: getControlChecked(ui.spineAtlasCheckbox),
    spineJsonFileName: getControlValue(ui.spineJsonFileNameInput).trim(),
    spineImagesPath: getControlValue(ui.spineImagesPathInput).trim(),
    prefabTarget: getControlValue(ui.prefabTargetSelect),
    writePrefabPackage: getControlChecked(ui.prefabPackageCheckbox),
    selectedOnly: getControlChecked(ui.selectedOnlyCheckbox),
    exportHidden: getControlChecked(ui.exportHiddenCheckbox),
    includeEffects: getControlChecked(ui.includeEffectsCheckbox),
    useFullPathNames: getControlChecked(ui.fullPathCheckbox),
    writeMetadata: getControlChecked(ui.metadataCheckbox),
    recursiveNormal: getControlChecked(ui.recursiveNormalCheckbox),
  };
}

function getControlValue(control) {
  if (!control) {
    return "";
  }

  if (control.isContentEditable) {
    return String(control.textContent || "");
  }

  if (control.dataset && control.dataset.kind === "select") {
    return String(control.dataset.value || "");
  }

  if (typeof control.value !== "undefined" && control.value !== null) {
    return String(control.value);
  }

  const selectedItem = control.querySelector ? control.querySelector("sp-menu-item[selected]") : null;
  if (selectedItem) {
    return String(selectedItem.getAttribute("value") || selectedItem.textContent || "").trim();
  }

  return "";
}

function setControlValue(control, value) {
  if (!control) {
    return;
  }

  const text = String(value ?? "");

  if (control.isContentEditable) {
    control.textContent = text;
    return;
  }

  if (control.dataset && control.dataset.kind === "select") {
    const options = SELECT_OPTIONS[control.id] || [];
    let item = options.find((opt) => opt.value === text);
    if (!item && options.length) {
      item = options[0];
    }

    control.dataset.value = item ? item.value : text;
    control.textContent = item ? item.label : text || "-";
    return;
  }

  if (typeof control.value !== "undefined") {
    try {
      control.value = text;
    } catch (error) {
      // ignore and fall through to menu-item selection
    }
  }

  if (!control.querySelectorAll) {
    return;
  }

  const menuItems = Array.from(control.querySelectorAll("sp-menu-item"));
  if (!menuItems.length) {
    return;
  }

  let matched = false;
  menuItems.forEach((item, index) => {
    const candidate = String(item.getAttribute("value") || item.textContent || "").trim();
    const isMatch = candidate === text || (!text && index === 0);
    if (isMatch) {
      item.setAttribute("selected", "");
      matched = true;
    } else {
      item.removeAttribute("selected");
    }
  });

  if (!matched && menuItems[0]) {
    menuItems[0].setAttribute("selected", "");
  }
}

function getControlChecked(control) {
  if (control && control.dataset && control.dataset.kind === "check") {
    return control.dataset.checked === "true";
  }
  return Boolean(control && control.checked);
}

function setControlChecked(control, checked) {
  if (!control) {
    return;
  }

  if (control.dataset && control.dataset.kind === "check") {
    const next = Boolean(checked);
    control.dataset.checked = next ? "true" : "false";
    control.textContent = `${next ? "☑" : "☐"} ${control.dataset.label || ""}`;
    return;
  }

  control.checked = Boolean(checked);
}

async function restoreOutputFolder() {
  const token = localStorage.getItem(FOLDER_TOKEN_KEY);
  if (!token) {
    renderFolderState();
    return;
  }

  try {
    state.outputFolder = await fs.getEntryForPersistentToken(token);
  } catch (error) {
    console.warn("Failed to restore output folder", error);
    localStorage.removeItem(FOLDER_TOKEN_KEY);
    state.outputFolder = null;
  }

  renderFolderState();
}

async function chooseOutputFolder() {
  try {
    const folder = await fs.getFolder();
    if (!folder) {
      return;
    }

    state.outputFolder = folder;
    const token = await fs.createPersistentToken(folder);
    localStorage.setItem(FOLDER_TOKEN_KEY, token);
    renderFolderState();
    setStatus(`輸出資料夾已設定:\n${folder.nativePath}`, "ok");
  } catch (error) {
    setStatus(`選擇資料夾失敗：${error.message}`, "error");
  }
}

function renderFolderState() {
  if (state.outputFolder) {
    ui.folderState.textContent = "已選擇資料夾";
    ui.folderPath.textContent = state.outputFolder.nativePath || state.outputFolder.name;
  } else {
    ui.folderState.textContent = "尚未選擇資料夾";
    ui.folderPath.textContent = "請先選擇匯出資料夾。";
  }
}

async function refreshCandidates() {
  try {
    state.settings = readSettingsFromUi();
    persistSettings();
    const snapshot = buildDocumentSnapshot(state.settings);
    state.candidates = snapshot.candidates;
    state.docInfo = snapshot.docInfo;
    renderSnapshot();
    setStatus(snapshot.message, "ok");
  } catch (error) {
    state.candidates = [];
    state.docInfo = null;
    renderSnapshot();
    setStatus(error.message, "error");
  }
}

function buildDocumentSnapshot(settings) {
  const doc = getActiveDocument();
  const docWidth = toNumber(doc.width);
  const docHeight = toNumber(doc.height);
  const resolution = toNumber(doc.resolution);
  const keywordHint = buildKeywordHint(settings);
  const layers = collectLayers(doc, settings);
  const unorderedCandidates = layers
    .map((layer) => makeCandidate(doc, layer, settings))
    .filter(Boolean);
  const orderedCandidates = orderAssetsForChannel(unorderedCandidates, settings, settings.mode === "spine" ? "spine" : "platform");
  const candidates = assignSequentialExportNames(orderedCandidates);

  return {
    docInfo: {
      id: doc.id,
      title: doc.title || doc.name || "Untitled",
      width: docWidth,
      height: docHeight,
      resolution,
      mode: doc.mode,
    },
    candidates,
    message: candidates.length
      ? `已掃描 ${candidates.length} 個可輸出項目。${keywordHint}`
      : `目前找不到可輸出的圖層。${keywordHint} 請確認 PSD 有可見像素內容，或調整選項後重新掃描。`,
  };
}

function orderAssetsForChannel(items, settings, channel) {
  const ordered = [...(items || [])].sort(compareAssetStackPath);
  const mode = resolveLayerOrderMode(settings);

  if (mode === "reverse") {
    ordered.reverse();
    return ordered;
  }

  if (mode === "psd") {
    return ordered;
  }

  // auto: keep PSD order for platform exports, reverse for Spine stacking.
  if (channel === "spine") {
    ordered.reverse();
  }
  return ordered;
}

function compareAssetStackPath(left, right) {
  return compareNumberPath(getAssetStackPath(left), getAssetStackPath(right));
}

function getAssetStackPath(asset) {
  if (asset && Array.isArray(asset.stackPath) && asset.stackPath.length) {
    return asset.stackPath;
  }
  return [Number.MAX_SAFE_INTEGER];
}

function compareNumberPath(left, right) {
  const a = Array.isArray(left) ? left : [];
  const b = Array.isArray(right) ? right : [];
  const length = Math.max(a.length, b.length);
  for (let index = 0; index < length; index += 1) {
    const aValue = typeof a[index] === "number" ? a[index] : -1;
    const bValue = typeof b[index] === "number" ? b[index] : -1;
    if (aValue !== bValue) {
      return aValue - bValue;
    }
  }
  return 0;
}

function resolveLayerOrderMode(settings) {
  const mode = String((settings && settings.layerOrder) || "").trim().toLowerCase();
  if (mode === "psd" || mode === "reverse") {
    return mode;
  }
  return "auto";
}

function renderSnapshot() {
  if (!state.docInfo) {
    ui.docName.textContent = "尚未掃描";
    ui.docSize.textContent = "-";
    ui.assetCount.textContent = "0";
    ui.namePreview.textContent = "-";
    ui.previewList.innerHTML = "";
    return;
  }

  ui.docName.textContent = state.docInfo.title;
  ui.docSize.textContent = `${state.docInfo.width} x ${state.docInfo.height} @ ${state.docInfo.resolution}dpi`;
  ui.assetCount.textContent = String(state.candidates.length);
  ui.namePreview.textContent = state.candidates[0] ? `${state.candidates[0].exportName}.png` : "-";

  const preview = state.candidates.slice(0, 12);
  ui.previewList.innerHTML = preview.length
    ? preview.map(renderPreviewItem).join("")
    : `<div class="pe-preview-item"><strong>沒有候選圖層</strong><span class="pe-pathline">請確認目前 Photoshop 文件有圖層可供輸出。</span><span class="pe-meta">Tips: 關掉「只輸出目前選取圖層」或打開可見圖層。</span></div>`;
}

function renderPreviewItem(item) {
  const size = `${item.bounds.width} x ${item.bounds.height}`;
  return `<div class="pe-preview-item"><strong>${escapeHtml(item.exportName)}.png</strong><span class="pe-pathline">${escapeHtml(item.sourcePath)}</span><span class="pe-meta">${escapeHtml(item.kind)} | ${size} | (${item.bounds.left}, ${item.bounds.top})</span></div>`;
}

function getActiveDocument() {
  const doc = app.activeDocument;
  if (!doc) {
    throw new Error("請先在 Photoshop 開啟一個 PSD，再使用這個面板。");
  }
  return doc;
}

function collectLayers(doc, settings) {
  const selectedIds = settings.selectedOnly ? getSelectedLayerIds(doc) : null;
  const topLayers = toArray(doc.layers);
  const modeKeywords = getModeFolderKeywords(settings);
  const result = [];

  if (settings.mode === "spine") {
    topLayers.forEach((layer) => {
      walkLeafLayers(layer, result, settings, selectedIds, false, modeKeywords, false);
    });
    return result;
  }

  if (settings.recursiveNormal) {
    topLayers.forEach((layer) => {
      walkRecursiveExportLayers(layer, result, settings, selectedIds, false, modeKeywords, false);
    });
    return result;
  }

  topLayers.forEach((layer) => {
    if (!shouldIncludeLayer(layer, settings, selectedIds)) {
      return;
    }
    if (!matchesFolderKeyword(layer, modeKeywords)) {
      return;
    }
    if (isBackgroundLayer(layer)) {
      return;
    }
    result.push(layer);
  });

  return result;
}

function walkLeafLayers(layer, result, settings, selectedIds, ancestorSelected, keywords, ancestorKeywordMatched) {
  const selectedHere = Boolean(selectedIds && selectedIds.has(layer.id));
  const branchSelected = ancestorSelected || selectedHere;
  const branchKeywordMatched = ancestorKeywordMatched || matchesFolderKeyword(layer, keywords);

  if (!branchSelected && !shouldIncludeLayer(layer, settings, selectedIds, true)) {
    return;
  }

  if (isGroupLayer(layer)) {
    toArray(layer.layers).forEach((child) => walkLeafLayers(child, result, settings, selectedIds, branchSelected, keywords, branchKeywordMatched));
    return;
  }

  if (isBackgroundLayer(layer)) {
    return;
  }

  if (keywords.length && !branchKeywordMatched) {
    return;
  }

  result.push(layer);
}

function walkRecursiveExportLayers(layer, result, settings, selectedIds, ancestorSelected, keywords, ancestorKeywordMatched) {
  const selectedHere = Boolean(selectedIds && selectedIds.has(layer.id));
  const branchSelected = ancestorSelected || selectedHere;
  const branchKeywordMatched = ancestorKeywordMatched || matchesFolderKeyword(layer, keywords);

  if (!branchSelected && !shouldIncludeLayer(layer, settings, selectedIds, true)) {
    return;
  }

  if (isBackgroundLayer(layer)) {
    return;
  }

  const withinKeywordScope = !keywords.length || branchKeywordMatched;
  if (isGroupLayer(layer)) {
    toArray(layer.layers).forEach((child) => walkRecursiveExportLayers(child, result, settings, selectedIds, branchSelected, keywords, branchKeywordMatched));
    return;
  }

  if (!withinKeywordScope) {
    return;
  }

  result.push(layer);
}

function getModeFolderKeywords(settings) {
  const source = settings.mode === "spine" ? settings.spineFolderKeyword : settings.platformFolderKeyword;
  return String(source || "")
    .split(/[,\n，]/g)
    .map((part) => part.trim().toLowerCase())
    .filter(Boolean);
}

function buildKeywordHint(settings) {
  const source = settings.mode === "spine" ? settings.spineFolderKeyword : settings.platformFolderKeyword;
  const label = settings.mode === "spine" ? "Spine" : "平台";
  const keywords = getModeFolderKeywords(settings);
  if (!keywords.length) {
    return `${label}資料夾關鍵字：未設定（不過濾）`;
  }
  return `${label}資料夾關鍵字：${String(source).trim()}`;
}

function matchesFolderKeyword(layer, keywords) {
  if (!keywords || !keywords.length) {
    return true;
  }

  let current = layer;
  while (current && current.typename !== "Document") {
    if (isGroupLayer(current)) {
      const name = String(current.name || "").toLowerCase();
      if (keywords.some((keyword) => name.includes(keyword))) {
        return true;
      }
    }
    current = current.parent;
  }

  return false;
}

function getSelectedLayerIds(doc) {
  const selected = toArray(doc.activeLayers);
  return new Set(selected.map((layer) => layer.id));
}

function shouldIncludeLayer(layer, settings, selectedIds, allowSelectedDescendants = false) {
  if (!settings.exportHidden && !layer.visible) {
    return false;
  }

  if (!selectedIds || !selectedIds.size) {
    return true;
  }

  if (selectedIds.has(layer.id)) {
    return true;
  }

  if (!allowSelectedDescendants || !isGroupLayer(layer)) {
    return false;
  }

  return hasSelectedDescendant(layer, selectedIds);
}

function hasSelectedDescendant(layer, selectedIds) {
  return toArray(layer.layers).some((child) => selectedIds.has(child.id) || (isGroupLayer(child) && hasSelectedDescendant(child, selectedIds)));
}

function makeCandidate(doc, layer, settings) {
  const exportBounds = resolveCandidateBounds(doc, layer, settings);
  if (!exportBounds) {
    return null;
  }
  const { bounds, boundsNoEffects, emptySource } = exportBounds;
  const cocosTrimHint = buildCocosTrimHint(bounds, boundsNoEffects);

  const sourceSegments = buildLayerPath(layer);
  const naming = buildBaseNameDescriptor(sourceSegments, layer, settings);
  const exportFolderSegments = buildExportFolderSegments(sourceSegments, layer, settings);
  const sourcePath = sourceSegments.join("/");
  const centerX = bounds.left + bounds.width / 2;
  const centerY = bounds.top + bounds.height / 2;
  const engineX = roundNumber(centerX - toNumber(doc.width) / 2);
  const engineY = roundNumber(toNumber(doc.height) / 2 - centerY);

  return {
    id: layer.id,
    layer,
    sourceName: layer.name,
    sourcePath,
    sourceSegments,
    sanitizedSourcePath: naming.sanitizedSegments.join("/"),
    exportFolderSegments,
    baseExportName: naming.baseName,
    baseNameAdjusted: naming.sourceChanged,
    exportName: naming.baseName,
    kind: describeLayerKind(layer),
    renderProfile: buildLayerRenderProfile(layer),
    bounds,
    boundsNoEffects,
    emptySource,
    cocosTrimHint,
    stackPath: buildLayerStackPath(layer),
    position: {
      photoshopTopLeft: { x: bounds.left, y: bounds.top },
      photoshopCenter: { x: roundNumber(centerX), y: roundNumber(centerY) },
      unity: { x: engineX, y: engineY },
      cocos: { x: engineX, y: engineY },
      spine: { x: engineX, y: engineY },
    },
  };
}

function buildLayerRenderProfile(layer) {
  if (!layer) {
    return {
      blendMode: "",
      opacity: 100,
      fillOpacity: 100,
      isClippingMask: false,
      locks: {},
    };
  }

  return {
    blendMode: String(layer.blendMode || "").toLowerCase(),
    opacity: roundNumber(toNumber(layer.opacity)),
    fillOpacity: roundNumber(toNumber(layer.fillOpacity)),
    isClippingMask: Boolean(layer.isClippingMask),
    locks: {
      allLocked: Boolean(layer.allLocked),
      pixelsLocked: Boolean(layer.pixelsLocked),
      positionLocked: Boolean(layer.positionLocked),
      transparentPixelsLocked: Boolean(layer.transparentPixelsLocked),
    },
  };
}

function resolveCandidateBounds(doc, layer, settings) {
  const preferredBounds = getLayerBounds(layer, settings.includeEffects);
  const fallbackVisibleBounds = getLayerBounds(layer, false);

  if (hasRenderableBounds(preferredBounds)) {
    return {
      bounds: preferredBounds,
      boundsNoEffects: hasRenderableBounds(fallbackVisibleBounds) ? fallbackVisibleBounds : preferredBounds,
      emptySource: false,
    };
  }

  if (hasRenderableBounds(fallbackVisibleBounds)) {
    return {
      bounds: fallbackVisibleBounds,
      boundsNoEffects: fallbackVisibleBounds,
      emptySource: false,
    };
  }

  return {
    bounds: buildTransparentFallbackBounds(doc, preferredBounds || fallbackVisibleBounds),
    boundsNoEffects: buildTransparentFallbackBounds(doc, fallbackVisibleBounds || preferredBounds),
    emptySource: true,
  };
}

function hasRenderableBounds(bounds) {
  return Boolean(bounds && bounds.width > 0 && bounds.height > 0);
}

function buildTransparentFallbackBounds(doc, seedBounds) {
  const docWidth = Math.max(1, Math.round(toNumber(doc && doc.width)));
  const docHeight = Math.max(1, Math.round(toNumber(doc && doc.height)));
  const rawLeft = roundNumber(toNumber(seedBounds && seedBounds.left));
  const rawTop = roundNumber(toNumber(seedBounds && seedBounds.top));
  const left = clampNumber(Number.isFinite(rawLeft) ? rawLeft : 0, 0, Math.max(0, docWidth - 1));
  const top = clampNumber(Number.isFinite(rawTop) ? rawTop : 0, 0, Math.max(0, docHeight - 1));

  return {
    left,
    top,
    right: left + 1,
    bottom: top + 1,
    width: 1,
    height: 1,
  };
}

function assignSequentialExportNames(candidates) {
  const counters = new Map();
  return (candidates || []).map((candidate) => {
    const key = String((candidate && candidate.baseExportName) || "Layer").trim() || "Layer";
    const next = (counters.get(key) || 0) + 1;
    counters.set(key, next);
    return {
      ...candidate,
      exportName: `${key}_${String(next).padStart(3, "0")}`,
    };
  });
}

function getLayerBounds(layer, includeEffects) {
  const raw = includeEffects ? layer.bounds : (layer.boundsNoEffects || layer.bounds);
  if (!raw) {
    return null;
  }

  const left = roundNumber(toNumber(raw.left));
  const top = roundNumber(toNumber(raw.top));
  const right = roundNumber(toNumber(raw.right));
  const bottom = roundNumber(toNumber(raw.bottom));
  return {
    left,
    top,
    right,
    bottom,
    width: roundNumber(Math.max(0, right - left)),
    height: roundNumber(Math.max(0, bottom - top)),
  };
}

function buildCocosTrimHint(rawBounds, visibleBounds) {
  if (!rawBounds || !visibleBounds) {
    return null;
  }

  const rawWidth = Math.max(1, Math.round(toNumber(rawBounds.width)));
  const rawHeight = Math.max(1, Math.round(toNumber(rawBounds.height)));
  const width = Math.max(1, Math.round(toNumber(visibleBounds.width)));
  const height = Math.max(1, Math.round(toNumber(visibleBounds.height)));
  const trimX = Math.max(0, Math.round(toNumber(visibleBounds.left) - toNumber(rawBounds.left)));
  const trimY = Math.max(0, Math.round(toNumber(visibleBounds.top) - toNumber(rawBounds.top)));

  if (width > rawWidth || height > rawHeight) {
    return null;
  }

  if (trimX + width > rawWidth || trimY + height > rawHeight) {
    return null;
  }

  return {
    rawWidth,
    rawHeight,
    trimX,
    trimY,
    width,
    height,
    offsetX: roundNumber(trimX + width / 2 - rawWidth / 2),
    offsetY: roundNumber(rawHeight / 2 - (trimY + height / 2)),
    hasVisiblePixels: width > 0 && height > 0,
  };
}

function buildBaseNameDescriptor(sourceSegments, layer, settings) {
  const namingRule = buildFolderDrivenNamingRule(sourceSegments, layer);
  const sanitizedSegments = namingRule.sanitizedSegments;
  const base = namingRule.baseName || `layer_${layer.id}`;
  const prefix = sanitizeSegment(settings.prefix);
  return {
    baseName: prefix ? `${prefix}_${base}` : base,
    sanitizedSegments,
    sourceChanged: namingRule.sourceChanged,
  };
}

function buildFolderDrivenNamingRule(sourceSegments, layer) {
  const segments = Array.isArray(sourceSegments) ? sourceSegments.filter(Boolean) : [];
  const ancestors = segments.slice(0, -1);
  const trimmedAncestors = trimGenericRootSegments(ancestors);
  const localeIndex = findLocaleSegmentIndex(trimmedAncestors);
  const locale = localeIndex >= 0 ? normalizeLocaleSegment(trimmedAncestors[localeIndex]) : "";
  const mainFolder = localeIndex > 0
    ? trimmedAncestors[localeIndex - 1]
    : (trimmedAncestors[0] || trimmedAncestors[trimmedAncestors.length - 1] || layer.name || "Layer");
  const fullFolderName = formatFolderName(mainFolder, false);
  const shortFolderName = formatFolderName(mainFolder, true);
  const rawBaseName = locale
    ? [fullFolderName, locale].filter(Boolean).join("_")
    : shortFolderName;
  const baseName = sanitizeSegment(rawBaseName) || sanitizeSegment(layer && layer.name) || "Layer";
  const sanitizedSegments = locale
    ? [fullFolderName, locale].filter(Boolean)
    : [shortFolderName].filter(Boolean);

  return {
    baseName,
    sanitizedSegments,
    sourceChanged: rawBaseName !== baseName,
  };
}

function trimGenericRootSegments(segments) {
  const items = Array.isArray(segments) ? [...segments] : [];
  while (items.length > 1 && isGenericRootSegment(items[0])) {
    items.shift();
  }
  return items;
}

function isGenericRootSegment(segment) {
  const value = String(segment || "").trim().toUpperCase();
  return value === "UI" || value === "ROOT";
}

function findLocaleSegmentIndex(segments) {
  const localeSet = new Set(["JA", "CHT", "CHS", "EN"]);
  for (let index = segments.length - 1; index >= 0; index -= 1) {
    if (localeSet.has(String(segments[index] || "").trim().toUpperCase())) {
      return index;
    }
  }
  return -1;
}

function normalizeLocaleSegment(segment) {
  return String(segment || "").trim().toUpperCase();
}

function formatFolderName(segment, shortOnly) {
  const tokens = splitNamingTokens(segment);
  if (!tokens.length) {
    return "Layer";
  }
  const selected = shortOnly ? tokens.slice(0, 1) : tokens;
  return selected.map((token) => token.charAt(0).toUpperCase() + token.slice(1).toLowerCase()).join("");
}

function splitNamingTokens(segment) {
  return String(segment || "")
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .split(/[^A-Za-z0-9]+/g)
    .map((token) => token.trim())
    .filter(Boolean);
}

function buildLayerPath(layer) {
  const parts = [];
  let current = layer;
  while (current) {
    if (current.name) {
      parts.unshift(current.name);
    }
    current = current.parent;
    if (!current || current.typename === "Document") {
      break;
    }
  }
  return parts;
}

function buildLayerStackPath(layer) {
  const stackPath = [];
  let current = layer;

  while (current && current.parent) {
    const parent = current.parent;
    const index = findLayerIndex(parent ? parent.layers : [], current);
    stackPath.unshift(index);
    if (!parent || parent.typename === "Document") {
      break;
    }
    current = parent;
  }

  return stackPath.filter((value) => value >= 0);
}

function buildExportFolderSegments(sourceSegments, layer, settings) {
  if (!settings || !settings.useFullPathNames) {
    return [];
  }

  const rawSegments = isGroupLayer(layer)
    ? (Array.isArray(sourceSegments) ? sourceSegments : [])
    : (Array.isArray(sourceSegments) ? sourceSegments.slice(0, -1) : []);

  return trimGenericRootSegments(rawSegments)
    .map((segment) => sanitizeSegment(segment))
    .filter(Boolean);
}

function findLayerIndex(collection, targetLayer) {
  const layers = toArray(collection);
  const index = layers.findIndex((layer) => layer && targetLayer && layer.id === targetLayer.id);
  return index >= 0 ? index : Number.MAX_SAFE_INTEGER;
}

function reserveName(baseName, registry) {
  const clean = baseName || "layer";
  let name = clean;
  let index = 2;
  while (registry.has(name.toLowerCase())) {
    name = `${clean}_${index}`;
    index += 1;
  }
  registry.add(name.toLowerCase());
  return name;
}

function sanitizeSegment(input) {
  return String(input || "")
    .replace(/[\\/:*?"<>|]+/g, "_")
    .replace(/\s+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 80);
}

function sanitizeFileStem(input) {
  const fileName = String(input || "").replace(/\.[^.]+$/, "");
  return sanitizeSegment(fileName) || "PSD";
}

function describeLayerKind(layer) {
  if (isGroupLayer(layer)) {
    return "group";
  }

  if (isTextLayer(layer)) {
    return "text";
  }

  return String(layer.kind || "layer").toLowerCase();
}

function isGroupLayer(layer) {
  return layer.kind === constants.LayerKind.GROUP || toArray(layer.layers).length > 0;
}

function isBackgroundLayer(layer) {
  return Boolean(layer.isBackgroundLayer);
}

function isTextLayer(layer) {
  if (!layer) {
    return false;
  }

  if (constants && constants.LayerKind && layer.kind === constants.LayerKind.TEXT) {
    return true;
  }

  if (layer.textItem) {
    return true;
  }

  return String(layer.kind || "").toLowerCase() === "text";
}

function shouldUseMinimalPrefabOutput(target, settings) {
  if (!settings || !settings.writePrefabPackage) {
    return false;
  }
  return target === "unity-6.3" || target === "cocos-3.8.8";
}

async function runExport() {
  if (state.busy) {
    return;
  }

  try {
    if (!state.outputFolder) {
      throw new Error("請先選擇輸出資料夾。");
    }

    const doc = getActiveDocument();
    const snapshot = buildDocumentSnapshot(state.settings);
    state.candidates = snapshot.candidates;
    state.docInfo = snapshot.docInfo;
    renderSnapshot();

    if (!state.candidates.length) {
      throw new Error("沒有可輸出的圖層。請先重新掃描或調整選項。");
    }

    state.busy = true;
    state.quickExportCapability = ENABLE_SELECTION_QUICK_EXPORT ? "unknown" : "unavailable";
    syncBusyState();
    setStatus(`開始匯出 ${state.candidates.length} 個項目...`, "");

    const imagesFolder = await ensureFolder(state.outputFolder, "images");
    const metadataFolder = await ensureFolder(state.outputFolder, "metadata");
    const results = [];
    let reportImagesUseAssetPaths = true;

    for (let batchStart = 0; batchStart < state.candidates.length; batchStart += EXPORT_MODAL_BATCH_SIZE) {
      const batchEnd = Math.min(state.candidates.length, batchStart + EXPORT_MODAL_BATCH_SIZE);
      await core.executeAsModal(async () => {
        for (let index = batchStart; index < batchEnd; index += 1) {
          const item = state.candidates[index];
          const relativeImagePath = buildRelativeAssetImagePath(item);
          setStatus(`正在匯出 ${index + 1} / ${state.candidates.length}\n${relativeImagePath}`, "");
          const fileFolder = await ensureNestedFolders(imagesFolder, item.exportFolderSegments);
          const file = await fileFolder.createFile(`${item.exportName}.png`, { overwrite: true });
          const exportDebug = await exportLayerAsPng(doc, item, file, imagesFolder);
          results.push(makeMetadataRecord(item, exportDebug));
        }
      }, { commandName: `PSD Export Pipeline ${batchStart + 1}-${batchEnd}` });

      if (batchEnd < state.candidates.length) {
        await delay(EXPORT_MODAL_BATCH_COOLDOWN_MS);
      }
    }

    const minimalPrefabOutput = shouldUseMinimalPrefabOutput(state.settings.prefabTarget, state.settings);
    const writeRootLayoutMetadata = state.settings.writeMetadata && !minimalPrefabOutput;

    if (writeRootLayoutMetadata) {
      await writeMetadataFile(metadataFolder, results);
    }

    let spineJsonPath = "";
    let spineAtlasPath = "";
    let spineJsonNameUsed = "";
    let spineFolder = null;
    let engineFolder = null;
    let prefabPackageFolder = null;
    if (state.settings.writeSpineFormat) {
      spineFolder = await ensureFolder(state.outputFolder, "spine");
      const spineJsonName = normalizeSpineJsonFileName(state.settings.spineJsonFileName);
      spineJsonNameUsed = spineJsonName;
      const imagesPath = resolveSpineImagesPathForExport(state.settings.spineImagesPath);
      if (imagesPath === "images/" || imagesPath === "./images/") {
        const spineImagesFolder = await ensureFolder(spineFolder, "images");
        await mirrorExportedImages(imagesFolder, spineImagesFolder, results);
      }
      await writeSpineSkeletonFile(spineFolder, results, spineJsonName, imagesPath);
      spineJsonPath = joinNativePath(state.outputFolder, `spine\\${spineJsonName}`);

      if (state.settings.writeSpineAtlas) {
        const atlasName = replaceExtension(spineJsonName, ".atlas");
        await writeSpineAtlasFile(spineFolder, results, atlasName, imagesPath);
        spineAtlasPath = joinNativePath(state.outputFolder, `spine\\${atlasName}`);
      }
    }

    let prefabPackagePath = "";
    let reportImagesFolder = imagesFolder;
    let reportImagesPath = joinNativePath(state.outputFolder, "images");
    let reportImagesLabel = "images";
    if (state.settings.writePrefabPackage && state.settings.prefabTarget !== "none") {
      engineFolder = await ensureFolder(state.outputFolder, "engine");
      const packageResult = await writeEnginePrefabPackage(engineFolder, results, state.settings.prefabTarget, imagesFolder);
      prefabPackagePath = packageResult.path;
      prefabPackageFolder = packageResult.folder || null;
      reportImagesFolder = packageResult.reportImagesFolder || reportImagesFolder;
      reportImagesPath = packageResult.reportImagesPath || reportImagesPath;
      reportImagesLabel = packageResult.reportImagesLabel || reportImagesLabel;
      reportImagesUseAssetPaths = Boolean(packageResult.reportImagesUseAssetPaths);
      if (packageResult.cleanupRootImages) {
        await safeDeleteEntry(imagesFolder);
      }
    }

    const exportReportResult = await writeExportReport(metadataFolder, {
      sourceItems: state.candidates,
      exportedAssets: results,
      imagesFolder: reportImagesFolder,
      imagesPath: reportImagesPath,
      imagesLabel: reportImagesLabel,
      imagesUseAssetPaths: reportImagesUseAssetPaths,
      metadataFolder,
      writeLayoutMetadata: writeRootLayoutMetadata,
      spineFolder,
      prefabPackageFolder,
      prefabTarget: state.settings.prefabTarget,
      prefabPackagePath,
      spineJsonPath,
      spineAtlasPath,
    });
    const exportReportPath = exportReportResult.path;
    const transparentAssetCount = exportReportResult.payload
      && exportReportResult.payload.checks
      && Array.isArray(exportReportResult.payload.checks.transparentAssets)
      ? exportReportResult.payload.checks.transparentAssets.length
      : 0;

    setStatus([
      `匯出完成，共 ${results.length} 個檔案。`,
      `圖片：${reportImagesPath}`,
      writeRootLayoutMetadata ? `Metadata：${joinNativePath(state.outputFolder, "metadata\\layout.json")}` : (state.settings.writeMetadata ? "Metadata：精簡輸出模式，已省略 layout.json" : "Metadata：已停用"),
      `自檢報告：${exportReportPath}`,
      transparentAssetCount ? `透明異常：${transparentAssetCount} 個 PNG 匯出後為全透明，請先檢查 metadata/export_report.txt` : "",
      spineJsonPath ? `Spine JSON：${spineJsonPath}` : "Spine JSON：已停用",
      spineAtlasPath ? `Spine Atlas：${spineAtlasPath}` : (state.settings.writeSpineFormat ? "Spine Atlas：已停用" : ""),
      spineJsonPath ? `Spine 匯入請使用 ${spineJsonNameUsed}（同層需有 atlas，且 atlas 會指向圖片路徑）。` : "",
      prefabPackagePath ? `Prefab 建置包：${prefabPackagePath}` : (state.settings.writePrefabPackage ? "Prefab 建置包：未輸出" : "Prefab 建置包：已停用"),
    ].filter(Boolean).join("\n"), "ok");
  } catch (error) {
    console.error(error);
    if (state.outputFolder) {
      try {
        const metadataFolder = await ensureFolder(state.outputFolder, "metadata");
        const errorFile = await metadataFolder.createFile("last_export_error.txt", { overwrite: true });
        await errorFile.write([
          `time=${new Date().toISOString()}`,
          `version=${RELEASE_INFO.version}`,
          `build=${RELEASE_INFO.build}`,
          `message=${formatErrorMessage(error)}`,
        ].join("\n"));
      } catch (writeError) {
        console.warn("Unable to write last_export_error.txt", writeError);
      }
    }
    setStatus(`匯出失敗：${formatErrorMessage(error)}`, "error");
  } finally {
    state.busy = false;
    syncBusyState();
  }
}

function syncBusyState() {
  const disabled = state.busy;
  ui.chooseFolderBtn.disabled = disabled;
  ui.refreshBtn.disabled = disabled;
  ui.exportBtn.disabled = disabled;
}

async function ensureFolder(parentFolder, folderName) {
  const existing = (await parentFolder.getEntries()).find((entry) => entry.isFolder && entry.name === folderName);
  if (existing) {
    return existing;
  }
  return parentFolder.createFolder(folderName);
}

async function ensureNestedFolders(parentFolder, folderSegments) {
  let current = parentFolder;
  for (const segment of folderSegments || []) {
    current = await ensureFolder(current, segment);
  }
  return current;
}

async function mirrorExportedImages(sourceImagesFolder, targetImagesFolder, assets) {
  if (!sourceImagesFolder || !targetImagesFolder || !assets || !assets.length) {
    return;
  }

  const imageMap = await collectRelativeFileMap(sourceImagesFolder);

  for (const asset of assets) {
    const sourceInfo = imageMap.get(getAssetImageRelativePath(asset).toLowerCase()) || imageMap.get(`${asset.name}.png`.toLowerCase());
    if (!sourceInfo || !sourceInfo.entry) {
      continue;
    }
    await copyFileEntry(sourceInfo.entry, targetImagesFolder, `${asset.name}.png`);
  }
}

async function copyFileEntry(sourceFile, targetFolder, targetFileName) {
  if (!sourceFile || !targetFolder) {
    return;
  }

  if (typeof sourceFile.copyTo === "function") {
    try {
      await sourceFile.copyTo(targetFolder, { overwrite: true });
      return;
    } catch (error) {
      // fallback to read/write
    }
  }

  const targetFile = await targetFolder.createFile(targetFileName || sourceFile.name, { overwrite: true });
  const binaryFormat = storage.formats && storage.formats.binary ? storage.formats.binary : null;

  if (binaryFormat) {
    const data = await sourceFile.read({ format: binaryFormat });
    await targetFile.write(data, { format: binaryFormat });
    return;
  }

  const data = await sourceFile.read();
  await targetFile.write(data);
}

async function exportLayerAsPng(sourceDoc, item, outputFile, exportFolder) {
  if (!ENABLE_SELECTION_QUICK_EXPORT) {
    if (!ENABLE_DUPLICATE_SAVEAS_FALLBACK) {
      throw new Error(`Selection quick export disabled and duplicate/saveAs fallback disabled: ${item && item.exportName ? item.exportName : "unknown"}`);
    }
    const fallbackDebug = await exportLayerViaDuplicateSaveAs(sourceDoc, item, outputFile);
    return {
      ...fallbackDebug,
      quickExportError: "Selection quick export disabled by config.",
      quickExportSkipped: true,
    };
  }

  if (state.quickExportCapability === "unavailable") {
    if (ENABLE_DOCUMENT_QUICK_EXPORT_FALLBACK) {
      try {
        const documentQuickExportDebug = await exportLayerViaDuplicateQuickExportDocument(sourceDoc, item, outputFile, exportFolder);
        return {
          ...documentQuickExportDebug,
          quickExportError: "Skipped selection quick export after prior host-level failure.",
          quickExportSkipped: true,
        };
      } catch (documentQuickExportError) {
        if (!ENABLE_DUPLICATE_SAVEAS_FALLBACK) {
          throw documentQuickExportError;
        }
        const fallbackDebug = await exportLayerViaDuplicateSaveAs(sourceDoc, item, outputFile);
        return {
          ...fallbackDebug,
          quickExportError: `Selection quick export previously unavailable; document quick export failed: ${formatErrorMessage(documentQuickExportError)}`,
          quickExportSkipped: true,
        };
      }
    }

    if (!ENABLE_DUPLICATE_SAVEAS_FALLBACK) {
      throw new Error(`Quick export unavailable and duplicate/saveAs fallback disabled: ${item && item.exportName ? item.exportName : "unknown"}`);
    }
    const fallbackDebug = await exportLayerViaDuplicateSaveAs(sourceDoc, item, outputFile);
    return {
      ...fallbackDebug,
      quickExportError: "Skipped selection quick export after prior host-level failure.",
      quickExportSkipped: true,
    };
  }

  try {
    const quickExportDebug = await exportLayerViaQuickExport(sourceDoc, item, outputFile, exportFolder);
    state.quickExportCapability = "available";
    return quickExportDebug;
  } catch (selectionError) {
    const hostUnavailable = shouldFallbackFromQuickExportError(selectionError);
    const shouldFallback = ENABLE_DUPLICATE_SAVEAS_FALLBACK && hostUnavailable;
    state.quickExportCapability = hostUnavailable ? "unavailable" : (ENABLE_SELECTION_QUICK_EXPORT ? "unknown" : "unavailable");
    const selectionMessage = formatErrorMessage(selectionError);

    if (ENABLE_DOCUMENT_QUICK_EXPORT_FALLBACK) {
      try {
        const documentQuickExportDebug = await exportLayerViaDuplicateQuickExportDocument(sourceDoc, item, outputFile, exportFolder);
        return {
          ...documentQuickExportDebug,
          quickExportError: selectionMessage,
        };
      } catch (documentQuickExportError) {
        if (!shouldFallback) {
          throw new Error(`Photoshop 內建匯出失敗：${item && item.exportName ? item.exportName : "unknown"} | selection=${selectionMessage} | document=${formatErrorMessage(documentQuickExportError)}`);
        }

        console.warn(`Document quick export failed for ${item && item.exportName ? item.exportName : "unknown"}, falling back to duplicate/saveAs path`, documentQuickExportError);
        const fallbackDebug = await exportLayerViaDuplicateSaveAs(sourceDoc, item, outputFile);
        return {
          ...fallbackDebug,
          quickExportError: `${selectionMessage}; documentQuickExport=${formatErrorMessage(documentQuickExportError)}`,
        };
      }
    }

    if (!shouldFallback) {
      throw selectionError;
    }
    const fallbackDebug = await exportLayerViaDuplicateSaveAs(sourceDoc, item, outputFile);
    return {
      ...fallbackDebug,
      quickExportError: selectionMessage,
    };
  }
}

async function exportLayerViaQuickExport(sourceDoc, item, outputFile, exportFolder) {
  if (!sourceDoc || !item || !outputFile || !exportFolder) {
    throw new Error("Quick export 缺少必要參數。");
  }

  const originalSelection = captureSelectedLayerIds(sourceDoc);
  const attemptSummaries = [];

  try {
    await selectLayerByIdForExport(item.id);

    for (const commandSpec of buildQuickExportCommandSpecs()) {
      const preparedCommandSpec = {
        ...commandSpec,
        layerId: item.id,
      };
      for (const destinationSpec of buildQuickExportDestinationSpecs()) {
        const attempt = await tryQuickExportWithDestination(item, exportFolder, outputFile, preparedCommandSpec, destinationSpec);
        attemptSummaries.push(attempt.summary);
        if (!attempt.success) {
          continue;
        }

        try {
          await writeFileEntryToFile(attempt.file, outputFile);
          return {
            strategy: "quick-export-selection",
            prepareMethod: "photoshop-export-selection",
            prepareSuccess: true,
            batchPlay: attempt.summary.batchPlay,
            attempts: attemptSummaries,
            command: attempt.summary.command,
            destFolder: {
              mode: attempt.summary.destFolder.mode,
              descriptorMode: attempt.summary.destFolder.descriptorMode,
              valueType: attempt.summary.destFolder.valueType,
            },
            selectedLayerIds: [item.id],
            outputFile: attempt.summary.outputFile,
          };
        } finally {
          if (attempt.tempFolder) {
            await safeDeleteEntry(attempt.tempFolder);
          }
        }
      }
    }

    if (areQuickExportAttemptsAllUnavailable(attemptSummaries)) {
      throw new Error(`Quick export command unavailable: ${formatQuickExportAttemptCompact(attemptSummaries)}`);
    }
    throw new Error(`Quick export did not create a usable PNG in any destination mode: ${formatQuickExportAttemptCompact(attemptSummaries)} | ${safeJsonStringify(attemptSummaries)}`);
  } catch (error) {
    console.warn(`Quick export selection failed for ${item && item.exportName ? item.exportName : "unknown"}`, error);
    throw new Error(`Photoshop 內建匯出失敗：${item && item.exportName ? item.exportName : "unknown"} | ${formatErrorMessage(error)} | attempts=${formatQuickExportAttemptCompact(attemptSummaries)} | attemptsJson=${safeJsonStringify(attemptSummaries)}`);
  } finally {
    await restoreSelectedLayersForExport(originalSelection);
  }
}

async function exportLayerViaDuplicateQuickExportDocument(sourceDoc, item, outputFile, exportFolder) {
  if (!sourceDoc || !item || !outputFile || !exportFolder) {
    throw new Error("Document quick export 缺少必要參數。");
  }

  const exportContext = await duplicateDocumentForLayerExport(sourceDoc, item);
  const exportDoc = exportContext && exportContext.doc;
  let resolved = null;
  const attemptSummaries = [];

  try {
    resolved = resolveDuplicatedLayer(exportContext, item);
    if (!resolved || !resolved.layer) {
      throw new Error(`Duplicate export could not resolve target layer for ${item && item.exportName ? item.exportName : "unknown"}`);
    }

    await isolateDocumentToLayerBranch(exportContext, resolved.stackPath, true);
    forceVisible(resolved.layer);
    const cropBounds = buildCropBoundsForExport(item);
    if (cropBounds) {
      await exportDoc.crop(cropBounds);
    }

    for (const commandSpec of buildDocumentQuickExportCommandSpecs()) {
      for (const destinationSpec of buildQuickExportDestinationSpecs()) {
        const attempt = await tryQuickExportDocumentWithDestination(exportDoc, item, exportFolder, outputFile, commandSpec, destinationSpec);
        attemptSummaries.push(attempt.summary);
        if (!attempt.success) {
          continue;
        }

        try {
          await writeFileEntryToFile(attempt.file, outputFile);
          return {
            strategy: "quick-export-document-duplicate",
            prepareMethod: "duplicate-isolate-crop-quick-export-document",
            prepareSuccess: true,
            fallbackUsed: true,
            batchPlay: attempt.summary.batchPlay,
            attempts: attemptSummaries,
            command: attempt.summary.command,
            destFolder: {
              mode: attempt.summary.destFolder.mode,
              descriptorMode: attempt.summary.destFolder.descriptorMode,
              valueType: attempt.summary.destFolder.valueType,
            },
            outputFile: attempt.summary.outputFile,
            fallbackBounds: cropBounds || null,
          };
        } finally {
          if (attempt.tempFolder) {
            await safeDeleteEntry(attempt.tempFolder);
          }
        }
      }
    }

    if (areQuickExportAttemptsAllUnavailable(attemptSummaries)) {
      throw new Error(`Document quick export command unavailable: ${formatQuickExportAttemptCompact(attemptSummaries)}`);
    }
    throw new Error(`Document quick export did not create a usable PNG: ${formatQuickExportAttemptCompact(attemptSummaries)} | ${safeJsonStringify(attemptSummaries)}`);
  } catch (error) {
    throw new Error(`Photoshop 文件快速匯出失敗：${item && item.exportName ? item.exportName : "unknown"} | ${formatErrorMessage(error)} | attempts=${formatQuickExportAttemptCompact(attemptSummaries)} | attemptsJson=${safeJsonStringify(attemptSummaries)}`);
  } finally {
    await closeDocumentWithoutSaving(exportDoc);
  }
}

async function tryQuickExportWithDestination(item, exportFolder, outputFile, commandSpec, destinationSpec) {
  let tempFolder = null;
  let destFolderRef = null;
  let exportBatchPlayResult = null;
  let outputInfo = null;
  let errorMessage = "";
  let rootSnapshot = null;
  let outputFileSnapshot = null;

  try {
    tempFolder = await createQuickExportTempFolder(exportFolder);
    rootSnapshot = await captureFolderRootSnapshot(exportFolder);
    outputFileSnapshot = await captureEntrySnapshot(outputFile);
    destFolderRef = buildBatchPlayFolderReference(tempFolder, destinationSpec);
    exportBatchPlayResult = await runQuickExportBatchPlay(destFolderRef.descriptorValue, commandSpec);
    assertBatchPlaySucceeded(exportBatchPlayResult, `Quick export failed for ${item.exportName}`);

    const exportedFile = await waitForQuickExportOutput(tempFolder, exportFolder, outputFile, rootSnapshot, outputFileSnapshot, QUICK_EXPORT_TIMEOUT_MS);
    outputInfo = exportedFile ? exportedFile.debug : null;
    const completedTempFolder = tempFolder;
    tempFolder = null;
    return {
      success: true,
      file: exportedFile.file,
      tempFolder: completedTempFolder,
      summary: await buildQuickExportAttemptSummary(commandSpec, destFolderRef, exportBatchPlayResult, outputInfo, completedTempFolder, ""),
    };
  } catch (error) {
    errorMessage = formatErrorMessage(error);
    return {
      success: false,
      file: null,
      tempFolder: null,
      summary: await buildQuickExportAttemptSummary(commandSpec, destFolderRef, exportBatchPlayResult, outputInfo, tempFolder, errorMessage),
    };
  } finally {
    if (tempFolder) {
      await safeDeleteEntry(tempFolder);
    }
  }
}

async function tryQuickExportDocumentWithDestination(exportDoc, item, exportFolder, outputFile, commandSpec, destinationSpec) {
  let tempFolder = null;
  let destFolderRef = null;
  let exportBatchPlayResult = null;
  let outputInfo = null;
  let errorMessage = "";
  let rootSnapshot = null;
  let outputFileSnapshot = null;

  try {
    tempFolder = await createQuickExportTempFolder(exportFolder);
    rootSnapshot = await captureFolderRootSnapshot(exportFolder);
    outputFileSnapshot = await captureEntrySnapshot(outputFile);
    destFolderRef = buildBatchPlayFolderReference(tempFolder, destinationSpec);
    exportBatchPlayResult = await runDocumentQuickExportBatchPlay(exportDoc && exportDoc.id, destFolderRef.descriptorValue, commandSpec);
    assertBatchPlaySucceeded(exportBatchPlayResult, `Document quick export failed for ${item.exportName}`);

    const exportedFile = await waitForQuickExportOutput(tempFolder, exportFolder, outputFile, rootSnapshot, outputFileSnapshot, QUICK_EXPORT_TIMEOUT_MS);
    outputInfo = exportedFile ? exportedFile.debug : null;
    const completedTempFolder = tempFolder;
    tempFolder = null;
    return {
      success: true,
      file: exportedFile.file,
      tempFolder: completedTempFolder,
      summary: await buildQuickExportAttemptSummary(commandSpec, destFolderRef, exportBatchPlayResult, outputInfo, completedTempFolder, ""),
    };
  } catch (error) {
    errorMessage = formatErrorMessage(error);
    return {
      success: false,
      file: null,
      tempFolder: null,
      summary: await buildQuickExportAttemptSummary(commandSpec, destFolderRef, exportBatchPlayResult, outputInfo, tempFolder, errorMessage),
    };
  } finally {
    if (tempFolder) {
      await safeDeleteEntry(tempFolder);
    }
  }
}

async function createQuickExportTempFolder(exportFolder) {
  if (!exportFolder || typeof exportFolder.createFolder !== "function") {
    throw new Error("缺少 quick export 輸出資料夾。");
  }
  return exportFolder.createFolder(`_psd_export_tmp_${Date.now()}_${Math.floor(Math.random() * 100000)}`);
}

function shouldFallbackFromQuickExportError(error) {
  const message = formatErrorMessage(error);
  return /-1715|-25920|-128|command unavailable|程式錯誤|指令無法使用|did not create a usable PNG|未找到可用輸出/i.test(message);
}

async function exportLayerViaDuplicateSaveAs(sourceDoc, item, outputFile) {
  const exportContext = await duplicateDocumentForLayerExport(sourceDoc, item);
  const exportDoc = exportContext && exportContext.doc;
  let resolved = null;

  try {
    resolved = resolveDuplicatedLayer(exportContext, item);
    if (!resolved || !resolved.layer) {
      throw new Error(`Duplicate export could not resolve target layer for ${item && item.exportName ? item.exportName : "unknown"}`);
    }

    await unlockTargetBranchForExport(exportDoc, resolved.stackPath);
    await isolateDocumentToLayerBranch(exportContext, resolved.stackPath, true);
    forceVisible(resolved.layer);
    const smartObjectPrep = await maybeConvertLayerToSmartObjectForFallback(exportDoc, item, resolved.layer, resolved.stackPath);
    const prepareInfo = await materializeFallbackPixelsForExport(exportDoc, item, smartObjectPrep.layer || resolved.layer);
    const exportPrepareInfo = {
      ...prepareInfo,
      smartObjectConverted: Boolean(smartObjectPrep.converted),
      smartObjectError: smartObjectPrep.error || "",
      smartObjectSkipReason: smartObjectPrep.skipReason || "",
    };

    const cropBounds = buildCropBoundsForExport(item);
    if (cropBounds) {
      await exportDoc.crop(cropBounds);
    }

    const saveOptions = buildPngSaveOptionsForFallback();
    await exportDoc.saveAs.png(outputFile, saveOptions, true);
    const optimizeInfo = await slimPngFileLossless(outputFile);

    return {
      strategy: "duplicate-saveas-png",
      prepareMethod: buildFallbackPrepareMethod(exportPrepareInfo),
      prepareSuccess: Boolean(exportPrepareInfo && exportPrepareInfo.success !== false),
      fallbackUsed: true,
      bakedVisiblePixels: Boolean(exportPrepareInfo && exportPrepareInfo.bakedVisiblePixels),
      prepareInfo: exportPrepareInfo,
      fallbackBounds: cropBounds || null,
      outputOptimize: optimizeInfo,
    };
  } catch (error) {
    throw new Error(`Duplicate PNG fallback failed: ${formatErrorMessage(error)}`);
  } finally {
    await closeDocumentWithoutSaving(exportDoc);
  }
}

function buildCropBoundsForExport(item) {
  const bounds = item && item.bounds ? item.bounds : null;
  if (!hasRenderableBounds(bounds)) {
    return null;
  }
  return {
    left: toNumber(bounds.left),
    top: toNumber(bounds.top),
    right: toNumber(bounds.left) + toNumber(bounds.width),
    bottom: toNumber(bounds.top) + toNumber(bounds.height),
  };
}

function buildPngSaveOptionsForFallback() {
  const options = {
    compression: clampNumber(PNG_SAVE_COMPRESSION, 0, 9),
    interlaced: false,
  };
  if (constants && constants.PNGMethod) {
    if (constants.PNGMethod.QUICK) {
      options.method = constants.PNGMethod.QUICK;
    }
  }
  return options;
}

async function slimPngFileLossless(fileEntry) {
  if (!ENABLE_PNG_LOSSLESS_SLIMMING || !fileEntry) {
    return {
      enabled: ENABLE_PNG_LOSSLESS_SLIMMING,
      applied: false,
      bytesSaved: 0,
    };
  }

  const binaryFormat = storage.formats && storage.formats.binary ? storage.formats.binary : null;
  if (!binaryFormat) {
    return {
      enabled: true,
      applied: false,
      bytesSaved: 0,
      error: "binary-format-unavailable",
    };
  }

  try {
    const metadata = await readEntryMetadata(fileEntry);
    const metadataSize = metadata && typeof metadata.size === "number" ? metadata.size : 0;
    if (metadataSize > 0 && metadataSize < PNG_LOSSLESS_SLIMMING_MIN_BYTES) {
      return {
        enabled: true,
        applied: false,
        bytesSaved: 0,
        originalSize: metadataSize,
        optimizedSize: metadataSize,
        skipped: "below-min-size",
      };
    }
    if (PNG_LOSSLESS_SLIMMING_MAX_BYTES > 0 && metadataSize > PNG_LOSSLESS_SLIMMING_MAX_BYTES) {
      return {
        enabled: true,
        applied: false,
        bytesSaved: 0,
        originalSize: metadataSize,
        optimizedSize: metadataSize,
        skipped: "above-max-size",
      };
    }

    const sourceData = await fileEntry.read({ format: binaryFormat });
    const sourceBytes = toUint8Array(sourceData);
    const originalSize = sourceBytes ? sourceBytes.byteLength : 0;
    const slimmedBytes = buildSlimmedPngBinary(sourceBytes);
    if (!slimmedBytes || slimmedBytes.byteLength >= originalSize) {
      return {
        enabled: true,
        applied: false,
        bytesSaved: 0,
        originalSize,
        optimizedSize: originalSize,
      };
    }

    await fileEntry.write(slimmedBytes, { format: binaryFormat });
    return {
      enabled: true,
      applied: true,
      bytesSaved: originalSize - slimmedBytes.byteLength,
      originalSize,
      optimizedSize: slimmedBytes.byteLength,
    };
  } catch (error) {
    console.warn(`Unable to slim PNG file: ${fileEntry && fileEntry.name ? fileEntry.name : "unknown"}`, error);
    return {
      enabled: true,
      applied: false,
      bytesSaved: 0,
      error: formatErrorMessage(error),
    };
  }
}

function buildSlimmedPngBinary(bytes) {
  const data = toUint8Array(bytes);
  if (!data || data.byteLength < 12 || !hasPngSignature(data)) {
    return null;
  }

  const signatureSize = 8;
  const keptChunks = [];
  let offset = signatureSize;
  let hasIend = false;

  while (offset + 12 <= data.byteLength) {
    const chunkLength = readUint32Be(data, offset);
    const chunkType = readPngChunkType(data, offset + 4);
    const chunkTotalSize = 12 + chunkLength;
    const chunkEnd = offset + chunkTotalSize;
    if (chunkEnd > data.byteLength) {
      return null;
    }

    if (shouldKeepPngChunkForSlimming(chunkType)) {
      keptChunks.push(data.slice(offset, chunkEnd));
    }

    offset = chunkEnd;
    if (chunkType === "IEND") {
      hasIend = true;
      break;
    }
  }

  if (!hasIend || !keptChunks.length) {
    return null;
  }

  const optimizedSize = signatureSize + keptChunks.reduce((sum, chunk) => sum + chunk.byteLength, 0);
  if (optimizedSize >= data.byteLength) {
    return null;
  }

  const result = new Uint8Array(optimizedSize);
  result.set(data.slice(0, signatureSize), 0);
  let writeOffset = signatureSize;
  keptChunks.forEach((chunk) => {
    result.set(chunk, writeOffset);
    writeOffset += chunk.byteLength;
  });
  return result;
}

function shouldKeepPngChunkForSlimming(chunkType) {
  if (isPngCriticalChunk(chunkType)) {
    return true;
  }
  return chunkType === "tRNS"
    || chunkType === "sRGB"
    || chunkType === "gAMA"
    || chunkType === "cHRM";
}

function isPngCriticalChunk(chunkType) {
  if (!chunkType || chunkType.length !== 4) {
    return false;
  }
  const firstCode = chunkType.charCodeAt(0);
  return firstCode >= 65 && firstCode <= 90;
}

function readUint32Be(bytes, offset) {
  return (bytes[offset] * 16777216)
    + (bytes[offset + 1] << 16)
    + (bytes[offset + 2] << 8)
    + bytes[offset + 3];
}

function readPngChunkType(bytes, offset) {
  return String.fromCharCode(
    bytes[offset],
    bytes[offset + 1],
    bytes[offset + 2],
    bytes[offset + 3]
  );
}

async function closeDocumentWithoutSaving(doc) {
  if (!doc) {
    return;
  }

  try {
    if (typeof doc.closeWithoutSaving === "function") {
      await doc.closeWithoutSaving();
      return;
    }
  } catch (error) {
    console.warn("Unable to close document with closeWithoutSaving", error);
  }

  try {
    if (typeof doc.close === "function" && constants && constants.SaveOptions && constants.SaveOptions.DONOTSAVECHANGES) {
      await doc.close(constants.SaveOptions.DONOTSAVECHANGES);
      return;
    }
  } catch (error) {
    console.warn("Unable to close document with SaveOptions.DONOTSAVECHANGES", error);
  }
}

async function runQuickExportBatchPlay(destFolderValue, commandSpec) {
  const { batchPlay } = require("photoshop").action;
  const commands = [];
  if (commandSpec && commandSpec.includeSelect && typeof commandSpec.layerId === "number") {
    commands.push({
      _obj: "select",
      _target: [{ _ref: "layer", _id: commandSpec.layerId }],
      makeVisible: false,
      layerID: [commandSpec.layerId],
      _isCommand: false,
      _options: { dialogOptions: "dontDisplay" },
    });
  }

  const descriptor = {
    _obj: "exportSelectionAsFileTypePressed",
    fileType: "png",
    quality: 32,
    metadata: 0,
    destFolder: destFolderValue,
    sRGB: true,
  };
  if (commandSpec && commandSpec.name === "layer-id-target-basic" && typeof commandSpec.layerId === "number") {
    descriptor._target = [{ _ref: "layer", _id: commandSpec.layerId }];
  } else if (commandSpec && commandSpec.target) {
    descriptor._target = commandSpec.target;
  }
  if (commandSpec && typeof commandSpec.openWindow === "boolean") {
    descriptor.openWindow = commandSpec.openWindow;
  }
  if (commandSpec && typeof commandSpec.isCommand === "boolean") {
    descriptor._isCommand = commandSpec.isCommand;
  }
  if (commandSpec && commandSpec.dialogOptions) {
    descriptor._options = { dialogOptions: commandSpec.dialogOptions };
  }
  commands.push(descriptor);

  return batchPlay(commands, {
    synchronousExecution: Boolean(commandSpec && commandSpec.synchronousExecution),
    modalBehavior: commandSpec && commandSpec.modalBehavior ? commandSpec.modalBehavior : "execute",
    propagateErrorToDefaultHandler: false,
  });
}

async function runDocumentQuickExportBatchPlay(documentId, destFolderValue, commandSpec) {
  const { batchPlay } = require("photoshop").action;
  const commands = [];
  if (typeof documentId === "number") {
    commands.push({
      _obj: "select",
      _target: [{ _ref: "document", _id: documentId }],
      _isCommand: false,
      _options: { dialogOptions: "dontDisplay" },
    });
  }

  const descriptor = {
    _obj: "exportDocumentAsFileTypePressed",
    fileType: "png",
    quality: 32,
    metadata: 0,
    destFolder: destFolderValue,
    sRGB: true,
  };
  if (commandSpec && typeof commandSpec.openWindow === "boolean") {
    descriptor.openWindow = commandSpec.openWindow;
  }
  if (commandSpec && typeof commandSpec.isCommand === "boolean") {
    descriptor._isCommand = commandSpec.isCommand;
  }
  if (commandSpec && commandSpec.dialogOptions) {
    descriptor._options = { dialogOptions: commandSpec.dialogOptions };
  }
  commands.push(descriptor);

  return batchPlay(commands, {
    synchronousExecution: Boolean(commandSpec && commandSpec.synchronousExecution),
    modalBehavior: commandSpec && commandSpec.modalBehavior ? commandSpec.modalBehavior : "execute",
    propagateErrorToDefaultHandler: false,
  });
}

async function buildQuickExportAttemptSummary(commandSpec, destFolderRef, exportBatchPlayResult, outputInfo, tempFolder, errorMessage) {
  return {
    command: {
      name: commandSpec && commandSpec.name ? commandSpec.name : "unknown",
      modalBehavior: commandSpec && commandSpec.modalBehavior ? commandSpec.modalBehavior : "unknown",
      hasTarget: Boolean(commandSpec && commandSpec.target && (Array.isArray(commandSpec.target) ? commandSpec.target.length : true)),
    },
    destFolder: {
      mode: destFolderRef && destFolderRef.mode ? destFolderRef.mode : "unknown",
      descriptorMode: destFolderRef && destFolderRef.descriptorMode ? destFolderRef.descriptorMode : "unknown",
      valueType: destFolderRef ? typeof destFolderRef.value : "unknown",
    },
    batchPlay: summarizeBatchPlayResult(exportBatchPlayResult),
    outputFile: outputInfo || null,
    tempFolder: tempFolder ? await describeFolderEntries(tempFolder) : null,
    error: errorMessage || "",
  };
}

function captureSelectedLayerIds(doc) {
  try {
    return toArray(doc && doc.activeLayers).map((layer) => layer.id).filter((id) => typeof id === "number");
  } catch (error) {
    console.warn("Unable to capture selected layer ids", error);
    return [];
  }
}

async function selectLayerByIdForExport(layerId) {
  if (typeof layerId !== "number") {
    return;
  }

  const { batchPlay } = require("photoshop").action;
  await batchPlay(
    [
      {
        _obj: "select",
        _target: [{ _ref: "layer", _id: layerId }],
        makeVisible: false,
        layerID: [layerId],
        _isCommand: false,
        _options: { dialogOptions: "dontDisplay" },
      },
    ],
    {
      synchronousExecution: true,
      modalBehavior: "execute",
    }
  );
}

async function restoreSelectedLayersForExport(layerIds) {
  const ids = Array.isArray(layerIds) ? layerIds.filter((id) => typeof id === "number") : [];
  if (!ids.length) {
    return;
  }

  const { batchPlay } = require("photoshop").action;
  const commands = ids.map((id, index) => ({
    _obj: "select",
    _target: [{ _ref: "layer", _id: id }],
    makeVisible: false,
    layerID: [id],
    selectionModifier: index === 0 ? undefined : { _enum: "selectionModifierType", _value: "addToSelection" },
    _isCommand: false,
    _options: { dialogOptions: "dontDisplay" },
  })).map((command) => {
    if (!command.selectionModifier) {
      delete command.selectionModifier;
    }
    return command;
  });

  try {
    await batchPlay(commands, {
      synchronousExecution: true,
      modalBehavior: "execute",
    });
  } catch (error) {
    console.warn("Unable to restore selected layers after export", error);
  }
}

async function findFirstPngFile(folder) {
  if (!folder) {
    return null;
  }

  const entries = await folder.getEntries();
  return entries.find((entry) => !entry.isFolder && /\.png$/i.test(String(entry.name || ""))) || null;
}

async function waitForFirstPngFile(folder, timeoutMs) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    const file = await findFirstPngFile(folder);
    if (file) {
      return file;
    }
    await delay(80);
  }
  return findFirstPngFile(folder);
}

async function waitForCompletedPngFile(folder, timeoutMs) {
  const startedAt = Date.now();
  let lastInspection = null;

  while (Date.now() - startedAt < timeoutMs) {
    const file = await findFirstPngFile(folder);
    if (file) {
      lastInspection = await inspectPngFileEntry(file);
      if (lastInspection.complete) {
        return {
          file,
          debug: {
            ...lastInspection,
            waitedMs: Date.now() - startedAt,
          },
        };
      }
    }
    await delay(100);
  }

  const fallbackFile = await findFirstPngFile(folder);
  if (!fallbackFile) {
    throw new Error("Quick export 未在暫存資料夾產生 PNG。");
  }

  lastInspection = lastInspection || await inspectPngFileEntry(fallbackFile);
  throw new Error(`Quick export 產生的 PNG 尚未寫入完成：${safeJsonStringify(lastInspection)}`);
}

async function waitForQuickExportOutput(tempFolder, exportFolder, outputFile, rootSnapshot, outputFileSnapshot, timeoutMs) {
  const startedAt = Date.now();
  let lastObservation = null;
  const completionTracker = new Map();

  while (Date.now() - startedAt < timeoutMs) {
    const tempCandidate = await findCompletedPngInFolderRecursive(tempFolder, `${tempFolder && tempFolder.name ? tempFolder.name : "_temp"}`, completionTracker);
    if (tempCandidate) {
      return {
        file: tempCandidate.file,
        debug: {
          ...tempCandidate.debug,
          source: "temp-folder",
          waitedMs: Date.now() - startedAt,
        },
      };
    }

    const outputCandidate = await inspectOutputFileCandidate(outputFile, outputFileSnapshot, completionTracker);
    if (outputCandidate) {
      return {
        file: outputCandidate.file,
        debug: {
          ...outputCandidate.debug,
          source: "output-file",
          waitedMs: Date.now() - startedAt,
        },
      };
    }

    const rootCandidate = await findCompletedChangedPngInFolderRoot(exportFolder, rootSnapshot, tempFolder, outputFile, completionTracker);
    if (rootCandidate) {
      return {
        file: rootCandidate.file,
        debug: {
          ...rootCandidate.debug,
          source: "images-root",
          waitedMs: Date.now() - startedAt,
        },
      };
    }

    lastObservation = {
      outputFile: await inspectPngFileEntryMetadataOnly(outputFile),
      rootChanges: await describeChangedRootEntries(exportFolder, rootSnapshot, tempFolder),
      tempFolder: tempFolder ? await describeFolderEntriesRecursive(tempFolder, 2) : null,
    };
    await delay(100);
  }

  throw new Error(`Quick export 未找到可用輸出：${safeJsonStringify(lastObservation)}`);
}

async function findCompletedPngInFolderRecursive(folder, prefix, completionTracker) {
  if (!folder || typeof folder.getEntries !== "function") {
    return null;
  }

  const entries = await folder.getEntries();
  for (const entry of entries) {
    const label = prefix ? `${prefix}/${entry.name}` : String(entry.name || "");
    if (entry.isFolder) {
      const nested = await findCompletedPngInFolderRecursive(entry, label, completionTracker);
      if (nested) {
        return nested;
      }
      continue;
    }
    if (!/\.png$/i.test(String(entry.name || ""))) {
      continue;
    }
    const debug = await inspectPngFileEntryWhenStable(entry, completionTracker, label);
    if (debug.complete) {
      return {
        file: entry,
        debug: {
          ...debug,
          relativePath: label,
        },
      };
    }
  }

  return null;
}

async function inspectOutputFileCandidate(outputFile, snapshot, completionTracker) {
  if (!outputFile) {
    return null;
  }
  const current = await captureEntrySnapshot(outputFile);
  if (!hasSnapshotChanged(snapshot, current)) {
    return null;
  }
  const debug = await inspectPngFileEntryWhenStable(outputFile, completionTracker, outputFile.name || "output-file");
  if (!debug.complete) {
    return null;
  }
  return {
    file: outputFile,
    debug: {
      ...debug,
      relativePath: outputFile.name || "",
    },
  };
}

async function findCompletedChangedPngInFolderRoot(folder, snapshot, tempFolder, outputFile, completionTracker) {
  if (!folder || typeof folder.getEntries !== "function") {
    return null;
  }

  const entries = await folder.getEntries();
  for (const entry of entries) {
    if (entry.isFolder) {
      continue;
    }
    if (!/\.png$/i.test(String(entry.name || ""))) {
      continue;
    }
    if (tempFolder && entry.name === tempFolder.name) {
      continue;
    }

    const current = await captureEntrySnapshot(entry);
    const previous = snapshot ? snapshot.get(String(entry.name || "").toLowerCase()) : null;
    if (!hasSnapshotChanged(previous, current)) {
      continue;
    }

    const debug = await inspectPngFileEntryWhenStable(entry, completionTracker, entry.name || "root-file");
    if (!debug.complete) {
      continue;
    }
    return {
      file: entry,
      debug: {
        ...debug,
        relativePath: entry.name || "",
        isOutputFile: Boolean(outputFile && outputFile.name === entry.name),
      },
    };
  }

  return null;
}

async function captureFolderRootSnapshot(folder) {
  const snapshot = new Map();
  if (!folder || typeof folder.getEntries !== "function") {
    return snapshot;
  }

  const entries = await folder.getEntries();
  for (const entry of entries) {
    snapshot.set(String(entry.name || "").toLowerCase(), await captureEntrySnapshot(entry));
  }
  return snapshot;
}

async function captureEntrySnapshot(entry) {
  if (!entry) {
    return null;
  }
  const metadata = await readEntryMetadata(entry);
  return {
    name: String(entry.name || ""),
    isFolder: Boolean(entry.isFolder),
    size: metadata && typeof metadata.size === "number" ? metadata.size : null,
    lastModified: metadata && metadata.dateModified ? String(metadata.dateModified) : "",
  };
}

function hasSnapshotChanged(before, after) {
  if (!after) {
    return false;
  }
  if (!before) {
    return true;
  }
  return before.size !== after.size || before.lastModified !== after.lastModified;
}

async function describeChangedRootEntries(folder, snapshot, tempFolder) {
  if (!folder || typeof folder.getEntries !== "function") {
    return { entries: [] };
  }

  const entries = await folder.getEntries();
  const changed = [];
  for (const entry of entries) {
    if (tempFolder && entry.name === tempFolder.name) {
      continue;
    }
    const current = await captureEntrySnapshot(entry);
    const previous = snapshot ? snapshot.get(String(entry.name || "").toLowerCase()) : null;
    if (!hasSnapshotChanged(previous, current)) {
      continue;
    }
    changed.push(current);
  }
  return { entries: changed };
}

async function inspectPngFileEntry(file) {
  if (!file) {
    return {
      name: "",
      size: 0,
      lastModified: "",
      readSucceeded: false,
      byteLength: 0,
      signatureOk: false,
      iendOk: false,
      complete: false,
    };
  }
  const metadata = await readEntryMetadata(file);
  const data = await readFileEntryBinary(file);
  const byteLength = getBinaryByteLength(data);
  const signatureOk = hasPngSignature(data);
  const iendOk = hasPngIendChunk(data);
  return {
    name: file && file.name ? file.name : "",
    size: metadata && typeof metadata.size === "number" ? metadata.size : byteLength,
    lastModified: metadata && metadata.dateModified ? String(metadata.dateModified) : "",
    readSucceeded: byteLength > 0,
    byteLength,
    signatureOk,
    iendOk,
    complete: byteLength > 0 && signatureOk && iendOk,
  };
}

async function inspectPngFileEntryWhenStable(file, completionTracker, key) {
  if (!file) {
    return inspectPngFileEntry(file);
  }

  const metadata = await readEntryMetadata(file);
  const size = metadata && typeof metadata.size === "number" ? metadata.size : 0;
  const lastModified = metadata && metadata.dateModified ? String(metadata.dateModified) : "";
  const trackerKey = key || file.name || "png";
  const previous = completionTracker && completionTracker.get(trackerKey);
  const stablePolls = previous && previous.size === size && previous.lastModified === lastModified
    ? previous.stablePolls + 1
    : 1;

  if (completionTracker) {
    completionTracker.set(trackerKey, { size, lastModified, stablePolls });
  }

  if (size <= 0 || stablePolls < PNG_COMPLETION_STABLE_POLLS) {
    return {
      name: file && file.name ? file.name : "",
      size,
      lastModified,
      readSucceeded: false,
      byteLength: 0,
      signatureOk: false,
      iendOk: false,
      complete: false,
      stablePolls,
    };
  }

  return inspectPngFileEntry(file);
}

async function inspectPngFileEntryMetadataOnly(file) {
  if (!file) {
    return inspectPngFileEntry(file);
  }

  const metadata = await readEntryMetadata(file);
  const size = metadata && typeof metadata.size === "number" ? metadata.size : 0;
  return {
    name: file && file.name ? file.name : "",
    size,
    lastModified: metadata && metadata.dateModified ? String(metadata.dateModified) : "",
    readSucceeded: false,
    byteLength: size,
    signatureOk: false,
    iendOk: false,
    complete: false,
  };
}

async function readEntryMetadata(entry) {
  if (!entry || typeof entry.getMetadata !== "function") {
    return null;
  }
  try {
    return await entry.getMetadata();
  } catch (error) {
    return null;
  }
}

async function readFileEntryBinary(file) {
  if (!file) {
    return null;
  }

  const binaryFormat = storage.formats && storage.formats.binary ? storage.formats.binary : null;
  if (!binaryFormat) {
    return null;
  }

  try {
    return await file.read({ format: binaryFormat });
  } catch (error) {
    return null;
  }
}

function getBinaryByteLength(data) {
  if (!data) {
    return 0;
  }
  if (typeof data.byteLength === "number") {
    return data.byteLength;
  }
  if (typeof data.length === "number") {
    return data.length;
  }
  return 0;
}

function hasPngSignature(data) {
  const bytes = toUint8Array(data);
  if (!bytes || bytes.length < 8) {
    return false;
  }
  const signature = [137, 80, 78, 71, 13, 10, 26, 10];
  return signature.every((value, index) => bytes[index] === value);
}

function hasPngIendChunk(data) {
  const bytes = toUint8Array(data);
  if (!bytes || bytes.length < 12) {
    return false;
  }
  const iend = [0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130];
  const offset = bytes.length - iend.length;
  return iend.every((value, index) => bytes[offset + index] === value);
}

function toUint8Array(data) {
  if (!data) {
    return null;
  }
  if (data instanceof Uint8Array) {
    return data;
  }
  if (typeof ArrayBuffer !== "undefined" && data instanceof ArrayBuffer) {
    return new Uint8Array(data);
  }
  if (ArrayBuffer.isView && ArrayBuffer.isView(data)) {
    return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  }
  return null;
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function writeFileEntryToFile(sourceFile, targetFile) {
  if (!sourceFile || !targetFile) {
    return;
  }

  if (sourceFile === targetFile) {
    return;
  }
  if (sourceFile.nativePath && targetFile.nativePath && sourceFile.nativePath === targetFile.nativePath) {
    return;
  }

  const binaryFormat = storage.formats && storage.formats.binary ? storage.formats.binary : null;
  if (binaryFormat) {
    const data = await sourceFile.read({ format: binaryFormat });
    await targetFile.write(data, { format: binaryFormat });
    return;
  }

  const data = await sourceFile.read();
  await targetFile.write(data);
}

function buildQuickExportDestinationSpecs() {
  return [
    { tokenMode: "native-path", descriptorMode: "raw" },
    { tokenMode: "session-token", descriptorMode: "raw" },
  ];
}

function buildQuickExportCommandSpecs() {
  return [
    {
      name: "layer-id-target-basic",
      includeSelect: false,
      layerId: null,
      target: null,
      openWindow: false,
      isCommand: true,
      dialogOptions: "dontDisplay",
      modalBehavior: "execute",
      synchronousExecution: true,
    },
    {
      name: "select-export-enum-basic",
      includeSelect: true,
      layerId: null,
      target: { _ref: "layer", _enum: "ordinal", _value: "targetEnum" },
      openWindow: false,
      isCommand: true,
      dialogOptions: "dontDisplay",
      modalBehavior: "execute",
      synchronousExecution: true,
    },
    {
      name: "select-export-no-target-basic",
      includeSelect: true,
      layerId: null,
      target: null,
      openWindow: false,
      isCommand: true,
      dialogOptions: "dontDisplay",
      modalBehavior: "execute",
      synchronousExecution: true,
    },
    {
      name: "select-export-enum-legacy",
      includeSelect: true,
      layerId: null,
      target: { _ref: "layer", _enum: "ordinal", _value: "targetEnum" },
      openWindow: false,
      isCommand: false,
      dialogOptions: "dontDisplay",
      modalBehavior: "execute",
      synchronousExecution: false,
    },
  ];
}

function buildDocumentQuickExportCommandSpecs() {
  return [
    {
      name: "document-export-basic",
      openWindow: false,
      isCommand: true,
      dialogOptions: "dontDisplay",
      modalBehavior: "execute",
      synchronousExecution: true,
    },
    {
      name: "document-export-legacy",
      openWindow: false,
      isCommand: false,
      dialogOptions: "dontDisplay",
      modalBehavior: "execute",
      synchronousExecution: false,
    },
  ];
}

function buildBatchPlayFolderReference(entry, destinationSpec) {
  if (!entry) {
    throw new Error("缺少 quick export 暫存資料夾。");
  }

  const preferredTokenMode = destinationSpec && destinationSpec.tokenMode ? destinationSpec.tokenMode : "session-token";
  const descriptorMode = destinationSpec && destinationSpec.descriptorMode ? destinationSpec.descriptorMode : "raw";

  let rawValue = entry.nativePath;
  let mode = "native-path";

  if (preferredTokenMode !== "native-path" && typeof fs.createSessionToken === "function") {
    try {
      rawValue = fs.createSessionToken(entry);
      mode = "session-token";
    } catch (error) {
      console.warn("Unable to create session token for quick export folder", error);
    }
  }

  const descriptorValue = descriptorMode === "path-object"
    ? { _path: rawValue, _kind: "local" }
    : rawValue;

  return {
    mode,
    value: rawValue,
    descriptorMode,
    descriptorValue,
  };
}

async function describeFolderEntries(folder) {
  if (!folder || typeof folder.getEntries !== "function") {
    return { entries: [] };
  }

  try {
    const entries = await folder.getEntries();
    return {
      entries: entries.map((entry) => ({
        name: entry && entry.name ? entry.name : "",
        isFolder: Boolean(entry && entry.isFolder),
      })),
    };
  } catch (error) {
    return {
      entries: [],
      error: formatErrorMessage(error),
    };
  }
}

async function describeFolderEntriesRecursive(folder, maxDepth = 2, prefix = "") {
  if (!folder || typeof folder.getEntries !== "function") {
    return { entries: [] };
  }

  try {
    const entries = await folder.getEntries();
    const result = [];
    for (const entry of entries) {
      const label = prefix ? `${prefix}/${entry.name}` : String(entry.name || "");
      result.push({
        name: label,
        isFolder: Boolean(entry && entry.isFolder),
      });
      if (entry && entry.isFolder && maxDepth > 0) {
        const nested = await describeFolderEntriesRecursive(entry, maxDepth - 1, label);
        result.push(...toArray(nested && nested.entries));
      }
    }
    return { entries: result };
  } catch (error) {
    return {
      entries: [],
      error: formatErrorMessage(error),
    };
  }
}

function assertBatchPlaySucceeded(result, context) {
  const firstError = toArray(result).find((item) => item && item._obj === "error");
  if (!firstError) {
    return;
  }
  throw new Error(`${context}: ${formatErrorMessage(firstError)}`);
}

function summarizeBatchPlayResult(result) {
  return toArray(result).map((item) => {
    if (!item || typeof item !== "object") {
      return { type: typeof item };
    }
    const summary = {};
    ["_obj", "message", "result"].forEach((key) => {
      if (key in item) {
        summary[key] = item[key];
      }
    });
    if (!Object.keys(summary).length) {
      summary.status = "ok";
    }
    return summary;
  });
}

function safeJsonStringify(value) {
  try {
    return JSON.stringify(value);
  } catch (error) {
    return String(value);
  }
}

function formatQuickExportAttemptCompact(attempts) {
  return toArray(attempts).map((attempt) => {
    const command = attempt && attempt.command && attempt.command.name ? attempt.command.name : "cmd?";
    const dest = attempt && attempt.destFolder && attempt.destFolder.mode ? attempt.destFolder.mode : "dest?";
    const shape = attempt && attempt.destFolder && attempt.destFolder.descriptorMode ? attempt.destFolder.descriptorMode : "shape?";
    const batchPlay = toArray(attempt && attempt.batchPlay);
    const errorResult = batchPlay.find((item) => item && item._obj === "error");
    const resultCode = errorResult && typeof errorResult.result !== "undefined" ? String(errorResult.result) : (attempt && attempt.error ? "ERR" : "OK");
    return `${command}/${dest}/${shape}:${resultCode}`;
  }).join(", ");
}

function isQuickExportCommandUnavailableAttempt(attempt) {
  const batchPlay = toArray(attempt && attempt.batchPlay);
  const errorResult = batchPlay.find((item) => item && item._obj === "error");
  if (!errorResult) {
    return false;
  }
  const resultCode = Number(errorResult.result);
  if (resultCode === -128 || resultCode === -1715 || resultCode === -25920) {
    return true;
  }
  const message = String(errorResult.message || attempt.error || "").toLowerCase();
  return message.includes("not available")
    || message.includes("程式錯誤")
    || message.includes("指令無法使用");
}

function areQuickExportAttemptsAllUnavailable(attempts) {
  const attemptList = toArray(attempts);
  if (!attemptList.length) {
    return false;
  }
  return attemptList.every((attempt) => isQuickExportCommandUnavailableAttempt(attempt));
}

async function prepareTargetLayerForExport(item, layer) {
  if (!item || !layer) {
    return { method: "none", success: true };
  }

  if (String(item.kind || "").toLowerCase() === "smartobject") {
    const converted = await convertSmartObjectToLayersForExport(layer);
    if (converted) {
      return {
        method: "smartobject-convert-to-layers",
        strategy: "smartobject-convert",
        success: true,
      };
    }
  }

  const rasterized = await rasterizeTargetLayerForExport(item, layer);
  return {
    method: String(item.kind || "").toLowerCase() === "smartobject" ? "smartobject-rasterize" : "none",
    strategy: String(item.kind || "").toLowerCase() === "smartobject" ? "smartobject-rasterize" : "standard",
    success: rasterized !== false,
  };
}

async function convertSmartObjectToLayersForExport(layer) {
  if (!layer) {
    return false;
  }

  try {
    const { batchPlay } = require("photoshop").action;
    await batchPlay(
      [
        {
          _obj: "select",
          _target: [{ _ref: "layer", _id: layer.id }],
          makeVisible: false,
          layerID: [layer.id],
          _isCommand: false,
          _options: { dialogOptions: "dontDisplay" },
        },
        {
          _obj: "placedLayerConvertToLayers",
          _options: { dialogOptions: "dontDisplay" },
        },
      ],
      {
        synchronousExecution: true,
        modalBehavior: "execute",
      }
    );
    return true;
  } catch (error) {
    console.warn(`Unable to convert smart object to layers before export: ${layer.name || layer.id}`, error);
    return false;
  }
}

async function rasterizeTargetLayerForExport(item, layer) {
  if (!item || !layer) {
    return false;
  }

  if (String(item.kind || "").toLowerCase() !== "smartobject") {
    return true;
  }

  if (typeof layer.rasterize !== "function") {
    return false;
  }

  try {
    await layer.rasterize(constants.RasterizeType.ENTIRELAYER);
    return true;
  } catch (error) {
    console.warn(`Unable to rasterize smart object before export: ${item.exportName}`, error);
    return false;
  }
}

async function duplicateDocumentForLayerExport(sourceDoc, item) {
  if (!sourceDoc || typeof sourceDoc.duplicate !== "function") {
    throw new Error("目前 Photoshop UXP 版本不支援文件複製 API。");
  }

  const exportDoc = await sourceDoc.duplicate(`export_${sanitizeFileStem(item.exportName)}`, false);
  return {
    doc: exportDoc,
  };
}

async function unlockTargetBranchForExport(doc, stackPath) {
  if (!doc) {
    return;
  }

  const path = Array.isArray(stackPath) ? stackPath : [];
  const visited = new Set();

  for (let depth = 1; depth <= path.length; depth += 1) {
    const layer = getLayerByStackPath(doc, path.slice(0, depth));
    if (!layer || visited.has(layer.id)) {
      continue;
    }
    visited.add(layer.id);
    await unlockSingleLayerForExport(layer);
  }

  const targetLayer = getLayerByStackPath(doc, path);
  if (!targetLayer) {
    return;
  }

  const branch = [];
  collectLayerBranchForExport(targetLayer, branch);
  for (const layer of branch) {
    if (!layer || visited.has(layer.id)) {
      continue;
    }
    visited.add(layer.id);
    await unlockSingleLayerForExport(layer);
  }
}

async function unlockDocumentLayersForExport(doc) {
  if (!doc) {
    return;
  }

  const layers = collectAllLayersForExport(doc);
  for (const layer of layers) {
    await unlockSingleLayerForExport(layer);
  }
}

function collectAllLayersForExport(doc) {
  const result = [];
  toArray(doc && doc.layers).forEach((layer) => collectLayerBranchForExport(layer, result));
  return result;
}

function collectLayerBranchForExport(layer, result) {
  if (!layer) {
    return;
  }

  result.push(layer);
  if (!isGroupLayer(layer)) {
    return;
  }

  toArray(layer.layers).forEach((child) => collectLayerBranchForExport(child, result));
}

async function unlockSingleLayerForExport(layer) {
  if (!layer) {
    return;
  }

  if (!hasAnyLayerLock(layer)) {
    return;
  }

  clearLayerLocksForExport(layer);
  if (!hasAnyLayerLock(layer)) {
    return;
  }

  try {
    const { batchPlay } = require("photoshop").action;
    await batchPlay(
      [
        {
          _obj: "select",
          _target: [{ _ref: "layer", _id: layer.id }],
          makeVisible: false,
          layerID: [layer.id],
          _isCommand: false,
          _options: { dialogOptions: "dontDisplay" },
        },
        {
          _obj: "applyLocking",
          _target: [{ _ref: "layer", _enum: "ordinal", _value: "targetEnum" }],
          layerLocking: {
            _obj: "layerLocking",
            protectAll: false,
            protectComposite: false,
            protectPosition: false,
            protectTransparency: false,
          },
          _options: { dialogOptions: "dontDisplay" },
        },
      ],
      {
        synchronousExecution: true,
        modalBehavior: "execute",
      }
    );
  } catch (error) {
    console.warn(`Unable to batch unlock layer: ${layer.name || layer.id}`, error);
  }

  clearLayerLocksForExport(layer);
}

function hasAnyLayerLock(layer) {
  if (!layer) {
    return false;
  }
  return Boolean(
    layer.allLocked
    || layer.pixelsLocked
    || layer.positionLocked
    || layer.transparentPixelsLocked
    || layer.locked
  );
}

function clearLayerLocksForExport(layer) {
  const lockProps = [
    "allLocked",
    "pixelsLocked",
    "positionLocked",
    "transparentPixelsLocked",
    "locked",
  ];

  lockProps.forEach((prop) => {
    if (!(prop in layer) || !layer[prop]) {
      return;
    }

    try {
      layer[prop] = false;
    } catch (error) {
      console.warn(`Unable to clear layer lock: ${prop}`, error);
    }
  });
}

function getLayerByStackPath(doc, stackPath) {
  let current = null;
  let collection = toArray(doc && doc.layers);

  for (const index of Array.isArray(stackPath) ? stackPath : []) {
    current = collection[index];
    if (!current) {
      return null;
    }
    collection = toArray(current.layers);
  }

  return current;
}

function getLayerBySourceSegments(doc, sourceSegments) {
  const segments = Array.isArray(sourceSegments) ? sourceSegments : [];
  if (!segments.length) {
    return null;
  }

  let collection = toArray(doc && doc.layers);
  let current = null;
  const stackPath = [];

  for (const segment of segments) {
    const index = collection.findIndex((layer) => String(layer && layer.name || "") === String(segment || ""));
    if (index < 0) {
      return null;
    }
    current = collection[index];
    stackPath.push(index);
    collection = toArray(current.layers);
  }

  return current ? { layer: current, stackPath } : null;
}

function resolveDuplicatedLayer(exportContext, item) {
  const doc = exportContext && exportContext.doc;
  const direct = getLayerByStackPath(doc, item && item.stackPath);
  if (isResolvedLayerCompatible(item, direct)) {
    return {
      layer: direct,
      stackPath: Array.isArray(item && item.stackPath) ? [...item.stackPath] : [],
    };
  }

  const fallback = getLayerBySourceSegments(doc, item && item.sourceSegments);
  if (fallback && isResolvedLayerCompatible(item, fallback.layer)) {
    return fallback;
  }

  const byName = findBestLayerMatchByName(doc, item);
  if (byName) {
    return byName;
  }

  return null;
}

function isResolvedLayerCompatible(item, layer) {
  if (!item || !layer) {
    return false;
  }

  const expectedKind = String(item.kind || "").toLowerCase();
  if (expectedKind === "text" && !isTextLayer(layer)) {
    return false;
  }
  if (expectedKind === "group" && !isGroupLayer(layer)) {
    return false;
  }
  if (expectedKind && expectedKind !== "group" && expectedKind !== "text" && isGroupLayer(layer)) {
    return false;
  }

  if (String(item.sourceName || "") && String(layer.name || "") !== String(item.sourceName || "")) {
    return false;
  }

  return true;
}

function findBestLayerMatchByName(doc, item) {
  if (!doc || !item || !String(item.sourceName || "")) {
    return null;
  }

  const allMatches = [];
  collectLayersByNameWithPath(toArray(doc.layers), [String(item.sourceName || "")], [], allMatches);

  const compatible = allMatches.filter((entry) => isResolvedLayerCompatible(item, entry.layer));
  if (!compatible.length) {
    return null;
  }

  const ranked = compatible
    .map((entry) => ({
      ...entry,
      score: calculateLayerMatchScore(item, entry.layer),
    }))
    .sort((a, b) => a.score - b.score);

  return ranked.length
    ? {
        layer: ranked[0].layer,
        stackPath: ranked[0].stackPath,
      }
    : null;
}

function collectLayersByNameWithPath(layers, names, parentPath, result) {
  toArray(layers).forEach((layer, index) => {
    const layerPath = [...parentPath, index];
    if (names.includes(String(layer && layer.name || ""))) {
      result.push({
        layer,
        stackPath: layerPath,
      });
    }
    if (isGroupLayer(layer)) {
      collectLayersByNameWithPath(toArray(layer.layers), names, layerPath, result);
    }
  });
}

function calculateLayerMatchScore(item, layer) {
  if (!item || !layer || !item.bounds) {
    return Number.MAX_SAFE_INTEGER;
  }
  const layerBounds = getLayerBounds(layer, true) || getLayerBounds(layer, false);
  if (!layerBounds) {
    return Number.MAX_SAFE_INTEGER;
  }

  return Math.abs(roundNumber(toNumber(layerBounds.left)) - roundNumber(toNumber(item.bounds.left)))
    + Math.abs(roundNumber(toNumber(layerBounds.top)) - roundNumber(toNumber(item.bounds.top)))
    + Math.abs(roundNumber(toNumber(layerBounds.width)) - roundNumber(toNumber(item.bounds.width)))
    + Math.abs(roundNumber(toNumber(layerBounds.height)) - roundNumber(toNumber(item.bounds.height)));
}

async function isolateDocumentToLayerBranch(exportContext, stackPath, forceVisibleBranch) {
  const doc = exportContext && exportContext.doc;
  if (!doc) {
    return;
  }
  const keepPaths = await collectVisibilityKeepPaths(doc, stackPath);
  applyVisibilityToPaths(toArray(doc.layers), keepPaths, [], Boolean(forceVisibleBranch));
}

async function collectVisibilityKeepPaths(doc, stackPath) {
  const targetPath = Array.isArray(stackPath) ? [...stackPath] : [];
  const keepPaths = [targetPath];
  const targetLayer = getLayerByStackPath(doc, targetPath);
  if (!targetLayer) {
    return keepPaths;
  }

  const targetGrouped = await isLayerGroupedForExport(targetLayer);
  if (targetGrouped) {
    const clippingBasePath = await findClippingBaseStackPath(doc, targetPath);
    if (clippingBasePath) {
      keepPaths.push(clippingBasePath);
    }
  }

  return keepPaths;
}

async function findClippingBaseStackPath(doc, stackPath) {
  if (!Array.isArray(stackPath) || !stackPath.length) {
    return null;
  }

  const parentPath = stackPath.slice(0, -1);
  const targetIndex = stackPath[stackPath.length - 1];
  const siblings = getLayerCollectionByStackPath(doc, parentPath);

  for (let index = targetIndex + 1; index < siblings.length; index += 1) {
    const layer = siblings[index];
    if (!layer) {
      continue;
    }
    const grouped = await isLayerGroupedForExport(layer);
    if (!grouped) {
      return [...parentPath, index];
    }
  }

  return null;
}

async function isLayerGroupedForExport(layer) {
  if (!layer) {
    return false;
  }

  if (layer.isClippingMask === true) {
    return true;
  }

  try {
    const { batchPlay } = require("photoshop").action;
    const result = await batchPlay(
      [
        {
          _obj: "get",
          _target: [
            { _property: "group" },
            { _ref: "layer", _id: layer.id },
          ],
          _options: { dialogOptions: "dontDisplay" },
        },
      ],
      {
        synchronousExecution: true,
        modalBehavior: "execute",
      }
    );
    return Boolean(result && result[0] && result[0].group);
  } catch (error) {
    console.warn(`Unable to query clipping/group state for layer: ${layer.name || layer.id}`, error);
    return false;
  }
}

function getLayerCollectionByStackPath(doc, parentPath) {
  if (!Array.isArray(parentPath) || !parentPath.length) {
    return toArray(doc && doc.layers);
  }

  const parentLayer = getLayerByStackPath(doc, parentPath);
  return toArray(parentLayer && parentLayer.layers);
}

function applyVisibilityToPaths(layers, keepPaths, parentPath, forceVisibleBranch) {
  const keepKeys = new Set((keepPaths || []).map(makeStackPathKey));
  const branchKeys = new Set();
  (keepPaths || []).forEach((path) => {
    for (let length = 1; length < path.length; length += 1) {
      branchKeys.add(makeStackPathKey(path.slice(0, length)));
    }
  });

  toArray(layers).forEach((layer, index) => {
    const layerPath = [...parentPath, index];
    const key = makeStackPathKey(layerPath);
    const keep = keepKeys.has(key);
    const keepBranch = keep || branchKeys.has(key);

    try {
      if (keepBranch && forceVisibleBranch) {
        layer.visible = true;
      } else if (!keepBranch) {
        layer.visible = false;
      }
    } catch (error) {
      console.warn("Unable to update export visibility", error);
    }

    if (isGroupLayer(layer)) {
      if (keepBranch) {
        applyVisibilityToPaths(toArray(layer.layers), keepPaths, layerPath, forceVisibleBranch);
      } else {
        hideLayerBranch(layer);
      }
    }
  });
}

function makeStackPathKey(path) {
  return (Array.isArray(path) ? path : []).join(".");
}

function hideLayerBranch(layer) {
  try {
    layer.visible = false;
  } catch (error) {
    console.warn("Unable to hide layer branch", error);
  }

  if (!isGroupLayer(layer)) {
    return;
  }

  toArray(layer.layers).forEach(hideLayerBranch);
}

async function bakeVisiblePixelsForExport(doc) {
  if (!doc || typeof doc.mergeVisibleLayers !== "function") {
    return;
  }

  await doc.mergeVisibleLayers();
}

async function materializeFallbackPixelsForExport(doc, item, layer) {
  if (!shouldBakeVisiblePixelsForFallback(item)) {
    return {
      method: "none",
      success: true,
      bakedVisiblePixels: false,
    };
  }

  const rasterized = await rasterizeLayerForFallback(item, layer);
  if (rasterized.success) {
    return {
      ...rasterized,
      bakedVisiblePixels: true,
    };
  }

  await bakeVisiblePixelsForExport(doc);
  return {
    method: "merge-visible-layers",
    success: true,
    bakedVisiblePixels: true,
    rasterizeError: rasterized.error || "",
  };
}

function buildFallbackPrepareMethod(prepareInfo) {
  const method = prepareInfo && prepareInfo.method ? prepareInfo.method : "none";
  const smartObjectConverted = Boolean(prepareInfo && prepareInfo.smartObjectConverted);
  if (smartObjectConverted && method === "rasterize-text-layer") {
    return "duplicate-isolate-smartobject-text-rasterize-crop-saveAs";
  }
  if (smartObjectConverted && method === "rasterize-target-layer") {
    return "duplicate-isolate-smartobject-rasterize-crop-saveAs";
  }
  if (smartObjectConverted && method === "merge-visible-layers") {
    return "duplicate-isolate-smartobject-bake-crop-saveAs";
  }
  if (smartObjectConverted) {
    return "duplicate-isolate-smartobject-crop-saveAs";
  }
  if (method === "rasterize-target-layer") {
    return "duplicate-isolate-rasterize-crop-saveAs";
  }
  if (method === "rasterize-text-layer") {
    return "duplicate-isolate-text-rasterize-crop-saveAs";
  }
  if (method === "merge-visible-layers") {
    return "duplicate-isolate-bake-crop-saveAs";
  }
  return "duplicate-isolate-crop-saveAs";
}

function shouldBakeVisiblePixelsForFallback(item) {
  if (!item) {
    return false;
  }

  if (shouldBakeVisiblePixelsForItem(item) || isCompositeRiskAsset(item)) {
    return true;
  }

  return hasEffectBoundsExpansion(item.bounds, item.boundsNoEffects);
}

async function rasterizeLayerForFallback(item, layer) {
  if (!item || !layer || isGroupLayer(layer)) {
    return {
      method: "none",
      success: false,
      error: "",
    };
  }

  const layerKind = String(item.kind || "").toLowerCase();
  const isTextCandidate = isTextLayer(layer) || layerKind === "text";
  const shouldRasterizeText = isTextCandidate && shouldRasterizeTextLayerForFallback(item);
  const shouldRasterize = shouldRasterizeText || layerKind === "smartobject";
  const rasterizeMethod = shouldRasterizeText ? "rasterize-text-layer" : "rasterize-target-layer";

  if (!shouldRasterize) {
    return {
      method: "none",
      success: false,
      error: "",
    };
  }

  if (typeof layer.rasterize !== "function") {
    return {
      method: rasterizeMethod,
      success: false,
      error: "rasterize API unavailable",
    };
  }

  try {
    await layer.rasterize(constants.RasterizeType.ENTIRELAYER);
    return {
      method: rasterizeMethod,
      success: true,
      error: "",
    };
  } catch (error) {
    console.warn(`Unable to rasterize layer before fallback export: ${item.exportName}`, error);
    return {
      method: rasterizeMethod,
      success: false,
      error: formatErrorMessage(error),
    };
  }
}

function shouldRasterizeTextLayerForFallback(item) {
  if (!item) {
    return false;
  }

  const profile = item.renderProfile || {};
  const blendMode = String(profile.blendMode || "").toLowerCase();
  if (Number(profile.fillOpacity) < 100) {
    return true;
  }
  if (Number(profile.opacity) < 100) {
    return true;
  }
  if (blendMode && blendMode !== "normal") {
    return true;
  }
  return hasEffectBoundsExpansion(item.bounds, item.boundsNoEffects);
}

function getSmartObjectConversionSkipReason(item, layer) {
  if (!item) {
    return "missing-item";
  }
  if (!layer) {
    return "missing-layer";
  }
  if (isGroupLayer(layer)) {
    return "group-layer-not-supported";
  }
  if (isTextLayer(layer) || String(item.kind || "").toLowerCase() === "text") {
    if (!shouldRasterizeTextLayerForFallback(item)) {
      return "text-low-risk-skip-smartobject";
    }
    return "";
  }
  return "not-text-layer";
}

async function maybeConvertLayerToSmartObjectForFallback(doc, item, layer, stackPath) {
  const skipReason = getSmartObjectConversionSkipReason(item, layer);
  if (skipReason) {
    return {
      converted: false,
      layer,
      error: "",
      skipReason,
    };
  }

  try {
    const { batchPlay } = require("photoshop").action;
    await batchPlay(
      [
        {
          _obj: "select",
          _target: [{ _ref: "layer", _id: layer.id }],
          makeVisible: false,
          layerID: [layer.id],
          _isCommand: false,
          _options: { dialogOptions: "dontDisplay" },
        },
        {
          _obj: "newPlacedLayer",
          _options: { dialogOptions: "dontDisplay" },
        },
      ],
      {
        synchronousExecution: true,
        modalBehavior: "execute",
      }
    );

    const convertedLayer = getLayerByStackPath(doc, stackPath) || toArray(doc && doc.activeLayers)[0] || layer;
    forceVisible(convertedLayer);
    return {
      converted: true,
      layer: convertedLayer,
      error: "",
      skipReason: "",
    };
  } catch (error) {
    console.warn(`Unable to convert layer to smart object before fallback export: ${item && item.exportName ? item.exportName : "unknown"}`, error);
    return {
      converted: false,
      layer,
      error: formatErrorMessage(error),
      skipReason: "",
    };
  }
}

function shouldBakeVisiblePixelsForItem(item) {
  if (!item) {
    return false;
  }

  if (String(item.kind || "").toLowerCase() === "text") {
    return shouldRasterizeTextLayerForFallback(item);
  }

  if (isTextLayer(item.layer)) {
    return shouldRasterizeTextLayerForFallback(item);
  }

  const layer = item.layer;
  if (layer && isGroupLayer(layer)) {
    return branchContainsTextLayer(layer);
  }

  return false;
}

function branchContainsTextLayer(layer) {
  if (!layer) {
    return false;
  }

  if (isTextLayer(layer)) {
    return true;
  }

  return toArray(layer.layers).some((child) => branchContainsTextLayer(child));
}

function hasEffectBoundsExpansion(bounds, boundsNoEffects) {
  if (!hasRenderableBounds(bounds) || !hasRenderableBounds(boundsNoEffects)) {
    return false;
  }

  return (
    roundNumber(toNumber(bounds.left)) !== roundNumber(toNumber(boundsNoEffects.left))
    || roundNumber(toNumber(bounds.top)) !== roundNumber(toNumber(boundsNoEffects.top))
    || roundNumber(toNumber(bounds.width)) !== roundNumber(toNumber(boundsNoEffects.width))
    || roundNumber(toNumber(bounds.height)) !== roundNumber(toNumber(boundsNoEffects.height))
  );
}

function formatErrorMessage(error) {
  if (!error) {
    return "未知錯誤";
  }

  if (typeof error.message === "string" && error.message.trim()) {
    return error.message.trim();
  }

  if (typeof error.reason === "string" && error.reason.trim()) {
    return error.reason.trim();
  }

  if (typeof error === "string" && error.trim()) {
    return error.trim();
  }

  try {
    return JSON.stringify(error);
  } catch (jsonError) {
    return String(error);
  }
}

function forceVisible(layer) {
  try {
    layer.visible = true;
  } catch (error) {
    console.warn("Unable to force visibility", error);
  }

  if (!isGroupLayer(layer)) {
    return;
  }

  toArray(layer.layers).forEach(forceVisible);
}

function makeMetadataRecord(item, exportDebug) {
  const cocosContentSize = item.cocosTrimHint
    ? { width: item.cocosTrimHint.width, height: item.cocosTrimHint.height }
    : { width: item.bounds.width, height: item.bounds.height };

  return {
    id: item.id,
    name: item.exportName,
    baseExportName: item.baseExportName,
    sourceName: item.sourceName,
    sourcePath: item.sourcePath,
    sanitizedSourcePath: item.sanitizedSourcePath,
    exportFolderSegments: item.exportFolderSegments || [],
    file: buildMetadataImagePath(item),
    kind: item.kind,
    emptySource: Boolean(item.emptySource),
    renderProfile: item.renderProfile || buildLayerRenderProfile(item.layer),
    exportDebug: exportDebug || null,
    stackPath: item.stackPath,
    bounds: item.bounds,
    boundsNoEffects: item.boundsNoEffects,
    pivot: { x: 0.5, y: 0.5 },
    photoshop: {
      topLeft: item.position.photoshopTopLeft,
      center: item.position.photoshopCenter,
    },
    unity: {
      anchor: "center",
      anchoredPosition: item.position.unity,
      sizeDelta: { x: item.bounds.width, y: item.bounds.height },
    },
    cocos: {
      anchor: "center",
      position: item.position.cocos,
      contentSize: cocosContentSize,
    },
    spine: {
      slotName: item.exportName,
      attachmentName: item.exportName,
      bonePosition: item.position.spine,
      size: { width: item.bounds.width, height: item.bounds.height },
    },
  };
}

async function writeMetadataFile(metadataFolder, assets) {
  const payload = {
    version: 1,
    generatedAt: new Date().toISOString(),
    document: {
      id: state.docInfo.id,
      name: state.docInfo.title,
      width: state.docInfo.width,
      height: state.docInfo.height,
      resolution: state.docInfo.resolution,
      mode: String(state.docInfo.mode),
    },
    settings: {
      ...state.settings,
      coordinateSystem: {
        photoshop: "top-left origin, +y downward",
        unity: "canvas-center origin, +y upward",
        cocos: "canvas-center origin, +y upward",
        spine: "canvas-center origin, +y upward",
      },
    },
    assets,
  };

  const file = await metadataFolder.createFile("layout.json", { overwrite: true });
  await file.write(JSON.stringify(payload, null, 2));
}

async function writeExportReport(metadataFolder, context) {
  const payload = await buildExportReportPayload(context);
  const jsonFile = await metadataFolder.createFile("export_report.json", { overwrite: true });
  await jsonFile.write(JSON.stringify(payload, null, 2));

  const textFile = await metadataFolder.createFile("export_report.txt", { overwrite: true });
  await textFile.write(buildExportReportText(payload));

  return {
    path: joinNativePath(state.outputFolder, "metadata\\export_report.json"),
    payload,
  };
}

async function buildExportReportPayload(context) {
  const imagesFileSet = context.imagesUseAssetPaths ? await collectRelativeFilePathSet(context.imagesFolder) : await collectFileNameSet(context.imagesFolder);
  const spineImageFolder = context.spineFolder ? await getOptionalFolder(context.spineFolder, "images") : null;
  const spineImageFileSet = await collectFileNameSet(spineImageFolder);
  const unityEditorFolder = context.prefabTarget === "unity-6.0" && context.prefabPackageFolder
    ? await getOptionalFolder(context.prefabPackageFolder, "Editor")
    : null;
  const unityImageFileSet = context.prefabTarget === "unity-6.0"
    ? await collectFileNameSet(context.prefabPackageFolder ? await getOptionalFolder(context.prefabPackageFolder, "images") : null)
    : new Set();
  const unity63AssetsFolder = context.prefabTarget === "unity-6.3" && context.prefabPackageFolder
    ? await getOptionalFolder(context.prefabPackageFolder, "Assets")
    : null;
  const unity63PsdExportFolder = context.prefabTarget === "unity-6.3"
    ? await getNestedFolder(unity63AssetsFolder, ["PSDExport", sanitizeFileStem(state.docInfo.title)])
    : null;
  const unity63ImagesFolder = context.prefabTarget === "unity-6.3"
    ? await getOptionalFolder(unity63PsdExportFolder, "Images")
    : null;
  const unity63EditorFolder = context.prefabTarget === "unity-6.3"
    ? await getOptionalFolder(unity63PsdExportFolder, "Editor")
    : null;
  const cocosResourcesFolder = context.prefabTarget === "cocos-3.8.8" && context.prefabPackageFolder
    ? await getOptionalFolder(context.prefabPackageFolder, "resources")
    : null;
  const cocosImageFileSet = context.prefabTarget === "cocos-3.8.8"
    ? await collectFileNameSet(cocosResourcesFolder ? await getOptionalFolder(cocosResourcesFolder, "psd_images") : null)
    : new Set();
  const cocosExtensionFolder = context.prefabTarget === "cocos-3.8.8" && context.prefabPackageFolder
    ? await getNestedFolder(context.prefabPackageFolder, ["extensions", "psd-export-pipeline-cocos388"])
    : null;

  const missingFiles = [];
  missingFiles.push(...findMissingExpectedFiles(
    context.exportedAssets.map((asset) => getExpectedImageLookupPath(asset, Boolean(context.imagesUseAssetPaths))),
    imagesFileSet,
    context.imagesLabel || "images"
  ));

  if (context.writeLayoutMetadata) {
    missingFiles.push(...findMissingExpectedFiles(["layout.json"], await collectFileNameSet(context.metadataFolder), "metadata"));
  }

  if (state.settings.writeSpineFormat && context.spineFolder) {
    missingFiles.push(...findMissingExpectedFiles([normalizeSpineJsonFileName(state.settings.spineJsonFileName)], await collectFileNameSet(context.spineFolder), "spine"));
    if (state.settings.writeSpineAtlas) {
      missingFiles.push(...findMissingExpectedFiles([replaceExtension(normalizeSpineJsonFileName(state.settings.spineJsonFileName), ".atlas")], await collectFileNameSet(context.spineFolder), "spine"));
    }
    const imagesPath = resolveSpineImagesPathForExport(state.settings.spineImagesPath);
    if (imagesPath === "images/" || imagesPath === "./images/") {
      missingFiles.push(...findMissingExpectedFiles(context.exportedAssets.map((asset) => `${asset.name}.png`), spineImageFileSet, "spine/images"));
    }
  }

  if (context.prefabTarget === "unity-6.0" && context.prefabPackageFolder) {
    missingFiles.push(...findMissingExpectedFiles(context.exportedAssets.map((asset) => `${asset.name}.png`), unityImageFileSet, "engine/unity6_0/images"));
    missingFiles.push(...findMissingExpectedFiles(["layout_unity6.json", "README_Unity6.txt"], await collectFileNameSet(context.prefabPackageFolder), "engine/unity6_0"));
    missingFiles.push(...findMissingExpectedFiles(["PsdUnity6PrefabAutoBuilder.cs"], await collectFileNameSet(unityEditorFolder), "engine/unity6_0/Editor"));
  }

  if (context.prefabTarget === "unity-6.3" && context.prefabPackageFolder) {
    missingFiles.push(...findMissingExpectedFiles(["layout_unity6_3.json", "layout_unity6_3.json.meta"], await collectFileNameSet(unity63PsdExportFolder), `engine/unity6_3/Assets/PSDExport/${sanitizeFileStem(state.docInfo.title)}`));
    missingFiles.push(...findMissingExpectedFiles(context.exportedAssets.flatMap((asset) => [`${asset.name}.png`, `${asset.name}.png.meta`]), await collectFileNameSet(unity63ImagesFolder), `engine/unity6_3/Assets/PSDExport/${sanitizeFileStem(state.docInfo.title)}/Images`));
    missingFiles.push(...findMissingExpectedFiles(["PsdUnity63PrefabAutoBuilder.cs", "PsdUnity63PrefabAutoBuilder.cs.meta"], await collectFileNameSet(unity63EditorFolder), `engine/unity6_3/Assets/PSDExport/${sanitizeFileStem(state.docInfo.title)}/Editor`));
  }

  if (context.prefabTarget === "cocos-3.8.8" && context.prefabPackageFolder) {
    const directRootName = `${sanitizeFileStem(state.docInfo.title)}_PSDAssets`;
    const directRootFolder = await getOptionalFolder(context.prefabPackageFolder, directRootName);
    const directPrefabFolder = await getOptionalFolder(directRootFolder, "Prefab");
    const directTextureFolder = await getOptionalFolder(directRootFolder, "Texture");
    const directExportTextureFolder = await getOptionalFolder(directTextureFolder, "psd_export");
    missingFiles.push(...findMissingExpectedFiles(
      [`${directRootName}.meta`],
      await collectFileNameSet(context.prefabPackageFolder),
      "engine/cocos3_8_8"
    ));
    missingFiles.push(...findMissingExpectedFiles(
      ["Prefab.meta", "Texture.meta"],
      await collectFileNameSet(directRootFolder),
      `engine/cocos3_8_8/${directRootName}`
    ));
    missingFiles.push(...findMissingExpectedFiles(
      ["psd_export.meta"],
      await collectFileNameSet(directTextureFolder),
      `engine/cocos3_8_8/${directRootName}/Texture`
    ));
    missingFiles.push(...findMissingExpectedFiles(
      [`${sanitizeFileStem(state.docInfo.title)}_PSD.prefab`, `${sanitizeFileStem(state.docInfo.title)}_PSD.prefab.meta`],
      await collectFileNameSet(directPrefabFolder),
      `engine/cocos3_8_8/${directRootName}/Prefab`
    ));
    missingFiles.push(...findMissingExpectedFiles(
      context.exportedAssets.flatMap((asset) => [`${asset.name}.png`, `${asset.name}.png.meta`]),
      await collectFileNameSet(directExportTextureFolder),
      `engine/cocos3_8_8/${directRootName}/Texture/psd_export`
    ));
  }

  const duplicateBaseNames = [];
  const duplicateSourceNames = summarizeDuplicateItems(context.sourceItems, (item) => item.sourceName, (item) => ({
    exportName: item.exportName,
    sourcePath: item.sourcePath,
  }));
  const pathWarnings = buildPathWarnings(context.sourceItems);
  const transparentAssets = ENABLE_DEEP_EXPORT_TRANSPARENCY_CHECK
    ? await analyzeExportedAssetTransparency(context.imagesFolder, context.exportedAssets, Boolean(context.imagesUseAssetPaths))
    : [];
  const probableTextLayerFailures = transparentAssets.filter((item) => String(item.kind || "").toLowerCase() === "text");
  const compositeRiskAssets = (context.exportedAssets || []).filter((asset) => isCompositeRiskAsset(asset));
  const renameWarnings = [];
  const spineRuleTests = runSpineOrderRuleSelfTests();

  const warnings = [
    ...pathWarnings.map((item) => `[path] ${item.message}`),
    ...duplicateSourceNames.map((item) => `[duplicate-source] ${item.key} x${item.count}`),
    ...transparentAssets.map((item) => `[transparent] ${item.exportName}.png exported with no visible pixels (${item.kind})`),
    ...probableTextLayerFailures.map((item) => `[text-render] ${item.exportName}.png is a text layer export with no visible pixels; check missing/substituted fonts or rasterize before export`),
    ...compositeRiskAssets.map((item) => `[composite-risk] ${item.name}.png blend=${item.renderProfile && item.renderProfile.blendMode ? item.renderProfile.blendMode : "unknown"} clip=${Boolean(item.renderProfile && item.renderProfile.isClippingMask)} fill=${item.renderProfile && typeof item.renderProfile.fillOpacity === "number" ? item.renderProfile.fillOpacity : "?"}`),
  ];
  const errors = missingFiles.map((item) => `[missing] ${item}`);

  return {
    release: RELEASE_INFO,
    generatedAt: new Date().toISOString(),
    document: {
      name: state.docInfo.title,
      width: state.docInfo.width,
      height: state.docInfo.height,
    },
    summary: {
      exportedCount: context.exportedAssets.length,
      warningCount: warnings.length,
      errorCount: errors.length,
      spineRuleTestsPassed: spineRuleTests.failed.length === 0,
    },
    outputs: {
      imagesPath: context.imagesPath || joinNativePath(state.outputFolder, "images"),
      metadataPath: joinNativePath(state.outputFolder, "metadata"),
      spineJsonPath: context.spineJsonPath || "",
      spineAtlasPath: context.spineAtlasPath || "",
      prefabPackagePath: context.prefabPackagePath || "",
    },
    assets: context.exportedAssets || [],
    warnings,
    errors,
    checks: {
      missingFiles,
      duplicateBaseNames,
      duplicateSourceNames,
      renameWarnings,
      pathWarnings,
      transparentAssets,
      probableTextLayerFailures,
      compositeRiskAssets,
      spineRuleTests,
    },
  };
}

function isCompositeRiskAsset(asset) {
  if (!asset || !asset.renderProfile) {
    return false;
  }

  const blendMode = String(asset.renderProfile.blendMode || "").toLowerCase();
  return Boolean(
    asset.renderProfile.isClippingMask
    || (blendMode && blendMode !== "normal")
    || Number(asset.renderProfile.fillOpacity) < 100
    || Number(asset.renderProfile.opacity) < 100
  );
}

function buildExportReportText(payload) {
  const lines = [
    `PSD Export Pipeline ${payload.release.version} (${payload.release.build})`,
    `stamp: ${payload.release.stamp}`,
    `generatedAt: ${payload.generatedAt}`,
    `document: ${payload.document.name} (${payload.document.width} x ${payload.document.height})`,
    `exportedCount: ${payload.summary.exportedCount}`,
    `warningCount: ${payload.summary.warningCount}`,
    `errorCount: ${payload.summary.errorCount}`,
    "",
    "[warnings]",
    ...(payload.warnings.length ? payload.warnings : ["none"]),
    "",
    "[errors]",
    ...(payload.errors.length ? payload.errors : ["none"]),
    "",
    "[transparent-assets]",
    ...(payload.checks.transparentAssets && payload.checks.transparentAssets.length
      ? payload.checks.transparentAssets.map((item) => `${item.file || `${item.exportName}.png`} | ${item.kind} | ${item.sourcePath}`)
      : ["none"]),
    "",
    "[text-render-risks]",
    ...(payload.checks.probableTextLayerFailures && payload.checks.probableTextLayerFailures.length
      ? payload.checks.probableTextLayerFailures.map((item) => `${item.file || `${item.exportName}.png`} | ${item.sourcePath}`)
      : ["none"]),
    "",
    "[composite-risk-assets]",
    ...(payload.checks.compositeRiskAssets && payload.checks.compositeRiskAssets.length
      ? payload.checks.compositeRiskAssets.map((item) => `${item.name}.png | ${item.sourcePath} | blend=${item.renderProfile && item.renderProfile.blendMode ? item.renderProfile.blendMode : "unknown"} | clip=${Boolean(item.renderProfile && item.renderProfile.isClippingMask)} | opacity=${item.renderProfile && typeof item.renderProfile.opacity === "number" ? item.renderProfile.opacity : "?"} | fill=${item.renderProfile && typeof item.renderProfile.fillOpacity === "number" ? item.renderProfile.fillOpacity : "?"}`)
      : ["none"]),
    "",
    "[export-debug]",
    ...(payload.assets && payload.assets.length
      ? payload.assets.map((item) => `${item.file || `${item.name}.png`} | ${item.sourcePath} | strategy=${item.exportDebug && item.exportDebug.strategy ? item.exportDebug.strategy : "standard"} | prepare=${item.exportDebug && item.exportDebug.prepareMethod ? item.exportDebug.prepareMethod : "none"} | success=${item.exportDebug && item.exportDebug.prepareSuccess === false ? "false" : "true"} | fallback=${item.exportDebug && item.exportDebug.fallbackUsed ? "true" : "false"} | so=${item.exportDebug && item.exportDebug.prepareInfo && item.exportDebug.prepareInfo.smartObjectConverted ? "true" : "false"} | soErr=${item.exportDebug && item.exportDebug.prepareInfo && item.exportDebug.prepareInfo.smartObjectError ? item.exportDebug.prepareInfo.smartObjectError : "-"} | soSkip=${item.exportDebug && item.exportDebug.prepareInfo && item.exportDebug.prepareInfo.smartObjectSkipReason ? item.exportDebug.prepareInfo.smartObjectSkipReason : "-"} | command=${item.exportDebug && item.exportDebug.command && item.exportDebug.command.name ? item.exportDebug.command.name : "unknown"} | dest=${item.exportDebug && item.exportDebug.destFolder && item.exportDebug.destFolder.mode ? item.exportDebug.destFolder.mode : "unknown"} | destShape=${item.exportDebug && item.exportDebug.destFolder && item.exportDebug.destFolder.descriptorMode ? item.exportDebug.destFolder.descriptorMode : "unknown"} | source=${item.exportDebug && item.exportDebug.outputFile && item.exportDebug.outputFile.source ? item.exportDebug.outputFile.source : "unknown"} | file=${item.exportDebug && item.exportDebug.outputFile && item.exportDebug.outputFile.relativePath ? item.exportDebug.outputFile.relativePath : "-"} | bytes=${item.exportDebug && item.exportDebug.outputFile && typeof item.exportDebug.outputFile.byteLength === "number" ? item.exportDebug.outputFile.byteLength : 0} | waitedMs=${item.exportDebug && item.exportDebug.outputFile && typeof item.exportDebug.outputFile.waitedMs === "number" ? item.exportDebug.outputFile.waitedMs : 0}`)
      : ["none"]),
    "",
    "[spine-rule-tests]",
    ...payload.checks.spineRuleTests.cases.map((item) => `${item.name}: ${item.passed ? "PASS" : "FAIL"} (${item.expected.join(",")} => ${item.actual.join(",")})`),
  ];
  return lines.join("\n");
}

async function collectFileNameSet(folder) {
  if (!folder) {
    return new Set();
  }
  const entries = await folder.getEntries();
  return new Set(entries.filter((entry) => !entry.isFolder).map((entry) => String(entry.name || "").toLowerCase()));
}

async function collectRelativeFileEntries(folder, prefix = "") {
  if (!folder) {
    return [];
  }

  const entries = await folder.getEntries();
  const files = [];
  for (const entry of entries) {
    const relativePath = prefix ? `${prefix}/${entry.name}` : String(entry.name || "");
    if (entry.isFolder) {
      files.push(...await collectRelativeFileEntries(entry, relativePath));
      continue;
    }
    files.push({
      entry,
      relativePath,
    });
  }
  return files;
}

async function collectRelativeFileMap(folder) {
  const files = await collectRelativeFileEntries(folder);
  const map = new Map();
  files.forEach((item) => {
    map.set(String(item.relativePath || "").toLowerCase(), item);
  });
  return map;
}

async function collectRelativeFilePathSet(folder) {
  const map = await collectRelativeFileMap(folder);
  return new Set(Array.from(map.keys()));
}

async function analyzeExportedAssetTransparency(imagesFolder, exportedAssets, useAssetPaths = true) {
  if (!imagesFolder || !Array.isArray(exportedAssets) || !exportedAssets.length) {
    return [];
  }

  const imageMap = useAssetPaths
    ? await collectRelativeFileMap(imagesFolder)
    : new Map((await imagesFolder.getEntries())
      .filter((entry) => !entry.isFolder)
      .map((entry) => [String(entry.name || "").toLowerCase(), { entry, relativePath: entry.name }]));
  const transparentAssets = [];

  for (const asset of exportedAssets) {
    if (asset && asset.emptySource) {
      continue;
    }

    const fileInfo = imageMap.get(getExpectedImageLookupPath(asset, useAssetPaths).toLowerCase()) || imageMap.get(`${asset.name}.png`.toLowerCase());
    if (!fileInfo || !fileInfo.entry) {
      continue;
    }

    const trimInfo = await safeAnalyzePngAlphaBounds(fileInfo.entry, asset);
    if (!trimInfo || trimInfo.hasVisiblePixels !== false) {
      continue;
    }

    transparentAssets.push({
      exportName: asset.name,
      file: asset.file || `images/${getExpectedImageLookupPath(asset, useAssetPaths)}`,
      kind: asset.kind,
      sourcePath: asset.sourcePath,
      rawWidth: Math.max(1, Math.round(toNumber(trimInfo.rawWidth))),
      rawHeight: Math.max(1, Math.round(toNumber(trimInfo.rawHeight))),
    });
  }

  return transparentAssets;
}

async function getOptionalFolder(parentFolder, folderName) {
  if (!parentFolder) {
    return null;
  }
  const entries = await parentFolder.getEntries();
  return entries.find((entry) => entry.isFolder && entry.name === folderName) || null;
}

async function getOptionalFile(parentFolder, fileName) {
  if (!parentFolder) {
    return null;
  }
  const entries = await parentFolder.getEntries();
  return entries.find((entry) => !entry.isFolder && entry.name === fileName) || null;
}

async function getNestedFolder(parentFolder, segments) {
  let current = parentFolder;
  for (const segment of segments || []) {
    current = await getOptionalFolder(current, segment);
    if (!current) {
      return null;
    }
  }
  return current;
}

function findMissingExpectedFiles(expectedFileNames, fileNameSet, prefix) {
  const existing = fileNameSet || new Set();
  return (expectedFileNames || [])
    .filter(Boolean)
    .filter((fileName) => !existing.has(String(fileName).toLowerCase()))
    .map((fileName) => `${prefix}/${fileName}`);
}

function buildRelativeAssetImagePath(item) {
  const fileName = `${item && item.exportName ? item.exportName : "layer"}.png`;
  const folderSegments = item && Array.isArray(item.exportFolderSegments) ? item.exportFolderSegments.filter(Boolean) : [];
  return folderSegments.length ? `${folderSegments.join("/")}/${fileName}` : fileName;
}

function buildMetadataImagePath(item) {
  return `images/${buildRelativeAssetImagePath(item)}`;
}

function getAssetImageRelativePath(asset) {
  const rawFile = asset && asset.file ? String(asset.file) : "";
  if (rawFile.toLowerCase().startsWith("images/")) {
    return rawFile.slice(7);
  }
  return rawFile || `${asset && asset.name ? asset.name : "layer"}.png`;
}

function getExpectedImageLookupPath(asset, useAssetPaths) {
  return useAssetPaths ? getAssetImageRelativePath(asset) : `${asset.name}.png`;
}

function summarizeDuplicateItems(items, keyGetter, itemMapper) {
  const registry = new Map();
  (items || []).forEach((item) => {
    const key = String(keyGetter(item) || "").trim();
    if (!key) {
      return;
    }
    if (!registry.has(key.toLowerCase())) {
      registry.set(key.toLowerCase(), {
        key,
        items: [],
      });
    }
    registry.get(key.toLowerCase()).items.push(itemMapper(item));
  });

  return Array.from(registry.values())
    .filter((entry) => entry.items.length > 1)
    .map((entry) => ({
      key: entry.key,
      count: entry.items.length,
      items: entry.items,
    }));
}

function buildPathWarnings(items) {
  const warnings = [];
  (items || []).forEach((item) => {
    if (!item.sourcePath) {
      warnings.push({
        exportName: item.exportName,
        message: `${item.exportName}.png has empty sourcePath`,
      });
    }

    const outputRelativePath = `images/${item.exportName}.png`;
    const outputFullPath = joinNativePath(state.outputFolder, outputRelativePath.replace(/\//g, "\\"));
    if (outputFullPath.length > 240) {
      warnings.push({
        exportName: item.exportName,
        message: `${outputRelativePath} path length ${outputFullPath.length} exceeds 240`,
      });
    }

    if (item.baseNameAdjusted) {
      warnings.push({
        exportName: item.exportName,
        message: `${item.sourcePath} sanitized to ${item.sanitizedSourcePath || item.baseExportName}`,
      });
    }
  });
  return warnings;
}

function runSpineOrderRuleSelfTests() {
  const createAsset = (name, stackPath) => ({
    name,
    stackPath,
    spine: {
      slotName: name,
      attachmentName: name,
      bonePosition: { x: 0, y: 0 },
    },
    bounds: { width: 1, height: 1 },
  });
  const assets = [
    createAsset("front", [0]),
    createAsset("mid_group_a", [1, 0]),
    createAsset("mid_group_b", [1, 1]),
    createAsset("back", [2]),
  ];
  const cases = [
    {
      name: "auto uses reverse PSD order for Spine",
      settings: { layerOrder: "auto" },
      channel: "spine",
      expected: ["back", "mid_group_b", "mid_group_a", "front"],
    },
    {
      name: "psd keeps explicit top-to-bottom order",
      settings: { layerOrder: "psd" },
      channel: "spine",
      expected: ["front", "mid_group_a", "mid_group_b", "back"],
    },
    {
      name: "reverse flips nested hierarchy deterministically",
      settings: { layerOrder: "reverse" },
      channel: "platform",
      expected: ["back", "mid_group_b", "mid_group_a", "front"],
    },
  ].map((testCase) => {
    const actual = orderAssetsForChannel(assets, testCase.settings, testCase.channel).map((item) => item.name);
    return {
      name: testCase.name,
      expected: testCase.expected,
      actual,
      passed: testCase.expected.join("|") === actual.join("|"),
    };
  });

  return {
    passed: cases.filter((item) => item.passed).length,
    failed: cases.filter((item) => !item.passed),
    cases,
  };
}

async function writeSpineSkeletonFile(spineFolder, assets, fileName, imagesPath) {
  const payload = buildSpineSkeletonPayload(assets, imagesPath);
  const file = await spineFolder.createFile(fileName, { overwrite: true });
  await file.write(JSON.stringify(payload, null, 2));
}

async function writeSpineAtlasFile(spineFolder, assets, fileName, imagesPath) {
  const atlas = buildSpineAtlasText(assets, imagesPath);
  const file = await spineFolder.createFile(fileName, { overwrite: true });
  await file.write(atlas);
}

async function writeEnginePrefabPackage(engineFolder, assets, target, imagesFolder) {
  if (target === "unity-6.0") {
    const unityFolder = await ensureFolder(engineFolder, "unity6_0");
    await writeUnity6PrefabPackage(unityFolder, assets, imagesFolder);
    return {
      folder: unityFolder,
      path: joinNativePath(state.outputFolder, "engine\\unity6_0"),
      reportImagesFolder: await getOptionalFolder(unityFolder, "images"),
      reportImagesPath: joinNativePath(state.outputFolder, "engine\\unity6_0\\images"),
      reportImagesLabel: "engine/unity6_0/images",
      reportImagesUseAssetPaths: false,
      cleanupRootImages: false,
    };
  }

  if (target === "unity-6.3") {
    const unityFolder = await ensureFolder(engineFolder, "unity6_3");
    const unityInfo = await writeUnity63PrefabPackage(unityFolder, assets, imagesFolder);
    return {
      folder: unityFolder,
      path: joinNativePath(state.outputFolder, "engine\\unity6_3"),
      reportImagesFolder: unityInfo.reportImagesFolder,
      reportImagesPath: unityInfo.reportImagesPath,
      reportImagesLabel: unityInfo.reportImagesLabel,
      reportImagesUseAssetPaths: false,
      cleanupRootImages: true,
    };
  }

  if (target === "cocos-3.8.8") {
    const cocosFolder = await ensureFolder(engineFolder, "cocos3_8_8");
    const cocosInfo = await writeCocosPrefabPackage(cocosFolder, assets, imagesFolder);
    return {
      folder: cocosFolder,
      path: joinNativePath(state.outputFolder, "engine\\cocos3_8_8"),
      reportImagesFolder: cocosInfo.reportImagesFolder,
      reportImagesPath: cocosInfo.reportImagesPath,
      reportImagesLabel: cocosInfo.reportImagesLabel,
      reportImagesUseAssetPaths: false,
      cleanupRootImages: true,
    };
  }

  return {
    folder: null,
    path: "",
    reportImagesUseAssetPaths: true,
  };
}

async function writeUnity6PrefabPackage(folder, assets, imagesFolder) {
  const packageImagesFolder = await ensureFolder(folder, "images");
  await mirrorExportedImages(imagesFolder, packageImagesFolder, assets);

  const payload = buildEnginePrefabPayload(assets, "unity-6.0", "./images/");
  const layoutFile = await folder.createFile("layout_unity6.json", { overwrite: true });
  await layoutFile.write(JSON.stringify(payload, null, 2));

  const editorFolder = await ensureFolder(folder, "Editor");
  const scriptFile = await editorFolder.createFile("PsdUnity6PrefabAutoBuilder.cs", { overwrite: true });
  await scriptFile.write(buildUnity6BuilderScript());

  const readme = await folder.createFile("README_Unity6.txt", { overwrite: true });
  await readme.write([
    "1) 將 engine/unity6_0 整包複製到 Unity 專案 Assets 底下。",
    "2) Unity 會自動偵測 layout_unity6.json 並在同資料夾建立 <PSD名稱>_PSD.prefab。",
    "3) 圖片來源使用 unity6_0/images，不需要再手動搬圖或重排。",
    "4) 如需重建，使用選單：Tools > PSD Export > Rebuild All layout_unity6 Prefabs。",
  ].join("\n"));
}

async function writeUnity63PrefabPackage(folder, assets, imagesFolder) {
  await clearFolderContents(folder);
  const docFolderName = sanitizeFileStem(state.docInfo.title);
  const assetsRoot = await ensureFolder(folder, "Assets");
  const psdExportRoot = await ensureFolder(assetsRoot, "PSDExport");
  const docFolder = await ensureFolder(psdExportRoot, docFolderName);
  const docImagesFolder = await ensureFolder(docFolder, "Images");
  const editorFolder = await ensureFolder(docFolder, "Editor");

  await mirrorExportedImages(imagesFolder, docImagesFolder, assets);

  const payload = buildEnginePrefabPayload(assets, "unity-6.3", "./Images/");
  const layoutFile = await docFolder.createFile("layout_unity6_3.json", { overwrite: true });
  await layoutFile.write(JSON.stringify(payload, null, 2));

  const builderFile = await editorFolder.createFile("PsdUnity63PrefabAutoBuilder.cs", { overwrite: true });
  await builderFile.write(buildUnity63BuilderScript());

  await writeUnity63AssetTreeMetas(assetsRoot, psdExportRoot, docFolder, docImagesFolder, editorFolder, assets);

  return {
    reportImagesFolder: docImagesFolder,
    reportImagesPath: joinNativePath(state.outputFolder, `engine\\unity6_3\\Assets\\PSDExport\\${docFolderName}\\Images`),
    reportImagesLabel: `engine/unity6_3/Assets/PSDExport/${docFolderName}/Images`,
  };
}

async function writeUnity63AssetTreeMetas(assetsRoot, psdExportRoot, docFolder, docImagesFolder, editorFolder, assets) {
  const psdExportRootMeta = await assetsRoot.createFile(`${psdExportRoot.name}.meta`, { overwrite: true });
  await psdExportRootMeta.write(buildUnityFolderMeta(stableUnityGuid(`unity63:folder:${psdExportRoot.name}`)));

  const docFolderMeta = await psdExportRoot.createFile(`${docFolder.name}.meta`, { overwrite: true });
  await docFolderMeta.write(buildUnityFolderMeta(stableUnityGuid(`unity63:folder:${docFolder.name}`)));

  const imagesFolderMeta = await docFolder.createFile(`${docImagesFolder.name}.meta`, { overwrite: true });
  await imagesFolderMeta.write(buildUnityFolderMeta(stableUnityGuid(`unity63:folder:${docFolder.name}:Images`)));

  const editorFolderMeta = await docFolder.createFile(`${editorFolder.name}.meta`, { overwrite: true });
  await editorFolderMeta.write(buildUnityFolderMeta(stableUnityGuid(`unity63:folder:${docFolder.name}:Editor`)));

  const layoutMeta = await docFolder.createFile("layout_unity6_3.json.meta", { overwrite: true });
  await layoutMeta.write(buildUnityTextMeta(stableUnityGuid(`unity63:file:${docFolder.name}:layout_unity6_3.json`)));

  const scriptMeta = await editorFolder.createFile("PsdUnity63PrefabAutoBuilder.cs.meta", { overwrite: true });
  await scriptMeta.write(buildUnityCsMeta(stableUnityGuid(`unity63:file:${docFolder.name}:Editor:PsdUnity63PrefabAutoBuilder.cs`)));

  for (const asset of assets) {
    const imageMeta = await docImagesFolder.createFile(`${asset.name}.png.meta`, { overwrite: true });
    await imageMeta.write(buildUnitySpriteMeta(stableUnityGuid(`unity63:file:${docFolder.name}:Images:${asset.name}.png`)));
  }
}

async function writeCocosPrefabPackage(folder, assets, imagesFolder) {
  await clearFolderContents(folder);
  const directPrefabInfo = await writeCocosDirectPrefabPackage(folder, assets, imagesFolder);
  return {
    reportImagesFolder: directPrefabInfo.reportImagesFolder,
    reportImagesPath: directPrefabInfo.reportImagesPath,
    reportImagesLabel: directPrefabInfo.reportImagesLabel,
  };
}

async function cleanupCocosDirectPrefabArtifacts(folder, rootFolderName) {
  const rootFolder = await getOptionalFolder(folder, rootFolderName);
  const rootMetaFile = await getOptionalFile(folder, `${rootFolderName}.meta`);
  await safeDeleteEntry(rootMetaFile);
  await safeDeleteEntry(rootFolder);
}

async function clearFolderContents(folder) {
  if (!folder || !folder.isFolder) {
    return;
  }

  const entries = await folder.getEntries();
  for (const entry of entries) {
    await safeDeleteEntry(entry);
  }
}

async function safeDeleteEntry(entry) {
  if (!entry) {
    return;
  }

  try {
    if (entry.isFolder) {
      const children = await entry.getEntries();
      for (const child of children) {
        await safeDeleteEntry(child);
      }
    }

    if (typeof entry.delete === "function") {
      await entry.delete();
    }
  } catch (error) {
    console.warn("Failed to delete entry during cleanup", error);
  }
}

async function writeCocosDirectPrefabPackage(folder, assets, imagesFolder) {
  const rootFolderName = `${sanitizeFileStem(state.docInfo.title)}_PSDAssets`;
  const rootFolder = await ensureFolder(folder, rootFolderName);
  const prefabFolder = await ensureFolder(rootFolder, "Prefab");
  const textureFolder = await ensureFolder(rootFolder, "Texture");
  const exportTextureFolder = await ensureFolder(textureFolder, "psd_export");
  const prefabFileName = `${sanitizeFileStem(state.docInfo.title)}_PSD.prefab`;

  await mirrorExportedImages(imagesFolder, exportTextureFolder, assets);

  const folderEntries = [
    {
      parentFolder: folder,
      fileName: `${rootFolderName}.meta`,
      payload: buildCocosDirectoryMeta({
        uuid: stableUuid(`cocos:folder:${rootFolderName}`),
        isBundle: true,
      }),
    },
    {
      parentFolder: rootFolder,
      fileName: "Prefab.meta",
      payload: buildCocosDirectoryMeta({
        uuid: stableUuid(`cocos:folder:${rootFolderName}:Prefab`),
      }),
    },
    {
      parentFolder: rootFolder,
      fileName: "Texture.meta",
      payload: buildCocosDirectoryMeta({
        uuid: stableUuid(`cocos:folder:${rootFolderName}:Texture`),
      }),
    },
    {
      parentFolder: textureFolder,
      fileName: "psd_export.meta",
      payload: buildCocosDirectoryMeta({
        uuid: stableUuid(`cocos:folder:${rootFolderName}:Texture:psd_export`),
        minimal: true,
      }),
    },
  ];

  for (const entry of folderEntries) {
    const file = await entry.parentFolder.createFile(entry.fileName, { overwrite: true });
    await file.write(JSON.stringify(entry.payload, null, 2));
  }

  const spriteMetaMap = new Map();
  for (const asset of assets) {
    const imageUuid = stableUuid(`cocos:image:${rootFolderName}:${asset.name}`);
    const copiedPngFile = await getOptionalFile(exportTextureFolder, `${asset.name}.png`);
    const trimInfo = await safeAnalyzePngAlphaBounds(copiedPngFile, asset);
    spriteMetaMap.set(asset.name, {
      imageUuid,
      trimInfo,
    });
    const file = await exportTextureFolder.createFile(`${asset.name}.png.meta`, { overwrite: true });
    await file.write(JSON.stringify(buildCocosImageMeta(asset, imageUuid, trimInfo), null, 2));
  }

  const prefabFile = await prefabFolder.createFile(prefabFileName, { overwrite: true });
  await prefabFile.write(JSON.stringify(buildCocosDirectPrefabDocument(assets, rootFolderName, prefabFileName, spriteMetaMap), null, 2));

  const prefabMetaFile = await prefabFolder.createFile(`${prefabFileName}.meta`, { overwrite: true });
  await prefabMetaFile.write(JSON.stringify(buildCocosPrefabMeta(prefabFileName), null, 2));

  return {
    rootFolderName,
    prefabFileName,
    reportImagesFolder: exportTextureFolder,
    reportImagesPath: joinNativePath(state.outputFolder, `engine\\cocos3_8_8\\${rootFolderName}\\Texture\\psd_export`),
    reportImagesLabel: `engine/cocos3_8_8/${rootFolderName}/Texture/psd_export`,
  };
}

async function safeAnalyzePngAlphaBounds(fileEntry, asset) {
  if (ENABLE_COCOS_FAST_TRIM_HINT && asset && asset.cocosTrimHint) {
    return buildTrimFallbackFromAsset(asset);
  }

  try {
    const trimInfo = await withTimeout(
      analyzePngAlphaBounds(fileEntry),
      2200,
      `PNG trim analysis: ${asset && asset.name ? asset.name : "unknown"}`
    );
    if (isSuspiciousTinyTrimResult(trimInfo, asset)) {
      return buildTrimFallbackFromAsset(asset);
    }
    if (shouldPreferAssetTrimHint(trimInfo, asset)) {
      return buildTrimFallbackFromAsset(asset);
    }
    return trimInfo;
  } catch (error) {
    console.warn("PNG trim analysis timed out or failed; using fallback trim", error);
    return buildTrimFallbackFromAsset(asset);
  }
}

function buildTrimFallbackFromAsset(asset) {
  const hint = asset && asset.cocosTrimHint ? asset.cocosTrimHint : null;
  const rawWidth = Math.max(1, Math.round(toNumber(hint ? hint.rawWidth : (asset && asset.bounds ? asset.bounds.width : 1))));
  const rawHeight = Math.max(1, Math.round(toNumber(hint ? hint.rawHeight : (asset && asset.bounds ? asset.bounds.height : 1))));
  const width = Math.max(1, Math.round(toNumber(hint ? hint.width : rawWidth)));
  const height = Math.max(1, Math.round(toNumber(hint ? hint.height : rawHeight)));
  const trimX = Math.max(0, Math.round(toNumber(hint ? hint.trimX : 0)));
  const trimY = Math.max(0, Math.round(toNumber(hint ? hint.trimY : 0)));
  const offsetX = roundNumber(toNumber(hint ? hint.offsetX : 0));
  const offsetY = roundNumber(toNumber(hint ? hint.offsetY : 0));
  return {
    rawWidth,
    rawHeight,
    trimX,
    trimY,
    width,
    height,
    offsetX,
    offsetY,
    hasVisiblePixels: null,
  };
}

function isSuspiciousTinyTrimResult(trimInfo, asset) {
  if (!trimInfo || !asset || !asset.bounds) {
    return false;
  }

  const fallbackWidth = Math.max(1, Math.round(toNumber(asset.bounds.width)));
  const fallbackHeight = Math.max(1, Math.round(toNumber(asset.bounds.height)));
  const rawWidth = Math.max(1, Math.round(toNumber(trimInfo.rawWidth)));
  const rawHeight = Math.max(1, Math.round(toNumber(trimInfo.rawHeight)));
  const width = Math.max(1, Math.round(toNumber(trimInfo.width)));
  const height = Math.max(1, Math.round(toNumber(trimInfo.height)));

  return fallbackWidth > 1
    && fallbackHeight > 1
    && rawWidth <= 1
    && rawHeight <= 1
    && width <= 1
    && height <= 1;
}

function shouldPreferAssetTrimHint(trimInfo, asset) {
  if (!asset || !asset.cocosTrimHint) {
    return false;
  }

  if (!trimInfo || trimInfo.hasVisiblePixels === false) {
    return true;
  }

  const hint = asset.cocosTrimHint;
  const width = Math.max(1, Math.round(toNumber(trimInfo.width)));
  const height = Math.max(1, Math.round(toNumber(trimInfo.height)));
  const rawWidth = Math.max(1, Math.round(toNumber(trimInfo.rawWidth)));
  const rawHeight = Math.max(1, Math.round(toNumber(trimInfo.rawHeight)));

  return rawWidth === hint.rawWidth
    && rawHeight === hint.rawHeight
    && width === rawWidth
    && height === rawHeight
    && (hint.width < rawWidth || hint.height < rawHeight);
}

function buildCocosDirectoryMeta(options) {
  const minimal = Boolean(options && options.minimal);
  const userData = options && options.isBundle
    ? { isBundle: true }
    : (minimal ? {} : { compressionType: {}, isRemoteBundle: {} });

  return {
    ver: "1.2.0",
    importer: "directory",
    imported: true,
    uuid: options && options.uuid ? options.uuid : stableUuid(`cocos:directory:${Date.now()}`),
    files: [],
    subMetas: {},
    userData,
  };
}

async function analyzePngAlphaBounds(fileEntry) {
  if (!fileEntry) {
    return {
      rawWidth: 1,
      rawHeight: 1,
      trimX: 0,
      trimY: 0,
      width: 1,
      height: 1,
      offsetX: 0,
      offsetY: 0,
      hasVisiblePixels: false,
    };
  }

  try {
    const bitmap = await loadBitmapFromFileEntry(fileEntry);
    const canvas = document.createElement("canvas");
    canvas.width = bitmap.width;
    canvas.height = bitmap.height;
    const context = canvas.getContext("2d");
    context.drawImage(bitmap, 0, 0);
    const imageData = context.getImageData(0, 0, canvas.width, canvas.height).data;

    let minX = canvas.width;
    let minY = canvas.height;
    let maxX = -1;
    let maxY = -1;

    for (let index = 3; index < imageData.length; index += 4) {
      if (imageData[index] <= 0) {
        continue;
      }
      const pixelIndex = (index - 3) / 4;
      const x = pixelIndex % canvas.width;
      const y = Math.floor(pixelIndex / canvas.width);
      if (x < minX) minX = x;
      if (y < minY) minY = y;
      if (x > maxX) maxX = x;
      if (y > maxY) maxY = y;
    }

    const rawWidth = Math.max(1, canvas.width);
    const rawHeight = Math.max(1, canvas.height);
    if (maxX < minX || maxY < minY) {
      return {
        rawWidth,
        rawHeight,
        trimX: 0,
        trimY: 0,
        width: rawWidth,
        height: rawHeight,
        offsetX: 0,
        offsetY: 0,
        hasVisiblePixels: false,
      };
    }

    const width = Math.max(1, maxX - minX + 1);
    const height = Math.max(1, maxY - minY + 1);
    return {
      rawWidth,
      rawHeight,
      trimX: minX,
      trimY: minY,
      width,
      height,
      offsetX: roundNumber(minX + width / 2 - rawWidth / 2),
      offsetY: roundNumber(rawHeight / 2 - (minY + height / 2)),
      hasVisiblePixels: true,
    };
  } catch (error) {
    console.warn("Failed to analyze PNG alpha bounds", error);
    const fallbackWidth = Math.max(1, Math.round(toNumber((fileEntry && fileEntry.width) || 1)));
    const fallbackHeight = Math.max(1, Math.round(toNumber((fileEntry && fileEntry.height) || 1)));
    return {
      rawWidth: fallbackWidth,
      rawHeight: fallbackHeight,
      trimX: 0,
      trimY: 0,
      width: fallbackWidth,
      height: fallbackHeight,
      offsetX: 0,
      offsetY: 0,
      hasVisiblePixels: false,
    };
  }
}

async function loadBitmapFromFileEntry(fileEntry) {
  const binaryFormat = storage.formats && storage.formats.binary ? storage.formats.binary : null;
  const raw = binaryFormat ? await fileEntry.read({ format: binaryFormat }) : await fileEntry.read();
  const bytes = normalizeBinaryReadResult(raw);
  if (!bytes) {
    throw new Error("PNG read returned empty data");
  }

  if (typeof Blob === "function" && typeof createImageBitmap === "function") {
    try {
      const blob = new Blob([bytes], { type: "image/png" });
      return await withTimeout(Promise.resolve(createImageBitmap(blob)), 1800, "createImageBitmap");
    } catch (error) {
      console.warn("createImageBitmap decode failed, falling back to Image()", error);
    }
  }

  if (typeof Image === "function") {
    const dataUrl = `data:image/png;base64,${arrayBufferToBase64(bytes)}`;
    return loadImageFromDataUrl(dataUrl);
  }

  throw new Error("No supported PNG decode path is available in this UXP runtime");
}

function normalizeBinaryReadResult(raw) {
  if (raw instanceof ArrayBuffer) {
    return raw;
  }

  if (raw && raw.buffer instanceof ArrayBuffer) {
    const offset = typeof raw.byteOffset === "number" ? raw.byteOffset : 0;
    const length = typeof raw.byteLength === "number" ? raw.byteLength : raw.buffer.byteLength;
    return raw.buffer.slice(offset, offset + length);
  }

  return null;
}

function withTimeout(promise, timeoutMs, label) {
  return new Promise((resolve, reject) => {
    let settled = false;
    const timer = setTimeout(() => {
      if (settled) {
        return;
      }
      settled = true;
      reject(new Error(`${label || "operation"} timed out after ${timeoutMs}ms`));
    }, timeoutMs);

    Promise.resolve(promise).then((value) => {
      if (settled) {
        return;
      }
      settled = true;
      clearTimeout(timer);
      resolve(value);
    }).catch((error) => {
      if (settled) {
        return;
      }
      settled = true;
      clearTimeout(timer);
      reject(error);
    });
  });
}

function loadImageFromDataUrl(dataUrl) {
  return withTimeout(new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = (error) => reject(error || new Error("Image decode failed"));
    image.src = dataUrl;
  }), 1800, "Image decode");
}

function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  let binary = "";
  for (let index = 0; index < bytes.length; index += chunkSize) {
    const chunk = bytes.subarray(index, Math.min(index + chunkSize, bytes.length));
    binary += String.fromCharCode.apply(null, Array.from(chunk));
  }
  return btoa(binary);
}

function buildCocosImageMeta(asset, imageUuid, trimInfo) {
  const info = normalizeCocosTrimInfo(asset, trimInfo);
  const width = info.width;
  const height = info.height;
  const rawWidth = info.rawWidth;
  const rawHeight = info.rawHeight;
  const halfWidth = width / 2;
  const halfHeight = height / 2;
  return {
    ver: "1.0.27",
    importer: "image",
    imported: true,
    uuid: imageUuid,
    files: [".json", ".png"],
    subMetas: {
      "6c48a": {
        importer: "texture",
        uuid: `${imageUuid}@6c48a`,
        displayName: asset.name,
        id: "6c48a",
        name: "texture",
        userData: {
          wrapModeS: "clamp-to-edge",
          wrapModeT: "clamp-to-edge",
          minfilter: "linear",
          magfilter: "linear",
          mipfilter: "none",
          anisotropy: 0,
          isUuid: true,
          imageUuidOrDatabaseUri: imageUuid,
          visible: false,
        },
        ver: "1.0.22",
        imported: true,
        files: [".json"],
        subMetas: {},
      },
      f9941: {
        importer: "sprite-frame",
        uuid: `${imageUuid}@f9941`,
        displayName: asset.name,
        id: "f9941",
        name: "spriteFrame",
        userData: {
          trimThreshold: 1,
          rotated: false,
          offsetX: info.offsetX,
          offsetY: info.offsetY,
          trimX: info.trimX,
          trimY: info.trimY,
          width,
          height,
          rawWidth,
          rawHeight,
          borderTop: 0,
          borderBottom: 0,
          borderLeft: 0,
          borderRight: 0,
          packable: true,
          pixelsToUnit: 100,
          pivotX: 0.5,
          pivotY: 0.5,
          meshType: 0,
          vertices: {
            rawPosition: [-halfWidth, -halfHeight, 0, halfWidth, -halfHeight, 0, -halfWidth, halfHeight, 0, halfWidth, halfHeight, 0],
            indexes: [0, 1, 2, 2, 1, 3],
            uv: [info.trimX, rawHeight - info.trimY, info.trimX + width, rawHeight - info.trimY, info.trimX, rawHeight - info.trimY - height, info.trimX + width, rawHeight - info.trimY - height],
            nuv: [
              roundNumber(info.trimX / rawWidth),
              roundNumber((rawHeight - info.trimY) / rawHeight),
              roundNumber((info.trimX + width) / rawWidth),
              roundNumber((rawHeight - info.trimY) / rawHeight),
              roundNumber(info.trimX / rawWidth),
              roundNumber((rawHeight - info.trimY - height) / rawHeight),
              roundNumber((info.trimX + width) / rawWidth),
              roundNumber((rawHeight - info.trimY - height) / rawHeight),
            ],
            minPos: [-halfWidth, -halfHeight, 0],
            maxPos: [halfWidth, halfHeight, 0],
          },
          isUuid: true,
          imageUuidOrDatabaseUri: `${imageUuid}@6c48a`,
          atlasUuid: "",
          trimType: "auto",
        },
        ver: "1.0.12",
        imported: true,
        files: [".json"],
        subMetas: {},
      },
    },
    userData: {
      type: "sprite-frame",
      hasAlpha: true,
      fixAlphaTransparencyArtifacts: false,
      redirect: `${imageUuid}@6c48a`,
    },
  };
}

function normalizeCocosTrimInfo(asset, trimInfo) {
  const fallbackWidth = Math.max(1, Math.round(toNumber(asset.bounds.width)));
  const fallbackHeight = Math.max(1, Math.round(toNumber(asset.bounds.height)));
  const info = trimInfo || {};
  const rawWidth = Math.max(1, Math.round(toNumber(info.rawWidth || fallbackWidth)));
  const rawHeight = Math.max(1, Math.round(toNumber(info.rawHeight || fallbackHeight)));
  const width = Math.max(1, Math.round(toNumber(info.width || rawWidth)));
  const height = Math.max(1, Math.round(toNumber(info.height || rawHeight)));
  const trimX = Math.max(0, Math.round(toNumber(info.trimX || 0)));
  const trimY = Math.max(0, Math.round(toNumber(info.trimY || 0)));
  return {
    rawWidth,
    rawHeight,
    width,
    height,
    trimX,
    trimY,
    offsetX: roundNumber(toNumber(typeof info.offsetX === "number" ? info.offsetX : trimX + width / 2 - rawWidth / 2)),
    offsetY: roundNumber(toNumber(typeof info.offsetY === "number" ? info.offsetY : rawHeight / 2 - (trimY + height / 2))),
  };
}

function buildCocosPrefabMeta(prefabFileName) {
  return {
    ver: "1.1.57",
    importer: "prefab",
    imported: true,
    uuid: stableUuid(`cocos:prefab:${prefabFileName}`),
    files: [".json"],
    subMetas: {},
    userData: {
      syncNodeName: replaceExtension(prefabFileName, ""),
    },
  };
}

function buildCocosDirectPrefabDocument(assets, rootFolderName, prefabFileName, spriteMetaMap) {
  const prefabName = replaceExtension(prefabFileName, "");
  const records = [];
  const rootNodeIndex = 1;
  const uiRootNodeIndex = 2;
  const docWidth = roundNumber(toNumber(state.docInfo.width));
  const docHeight = roundNumber(toNumber(state.docInfo.height));
  const cocosRenderOrderAssets = [...(assets || [])].reverse();

  records.push({
    __type__: "cc.Prefab",
    _name: prefabName,
    _objFlags: 0,
    __editorExtras__: {},
    _native: "",
    data: { __id__: rootNodeIndex },
    optimizationPolicy: 0,
    persistent: false,
  });

  records.push({
    __type__: "cc.Node",
    _name: prefabName,
    _objFlags: 0,
    __editorExtras__: {},
    _parent: null,
    _children: [{ __id__: uiRootNodeIndex }],
    _active: true,
    _components: [{ __id__: 0 }], // patched later
    _prefab: { __id__: 0 }, // patched later
    _lpos: makeCocosVec3(0, 0, 0),
    _lrot: makeCocosQuat(),
    _lscale: makeCocosVec3(1, 1, 1),
    _mobility: 0,
    _layer: 1073741824,
    _euler: makeCocosVec3(0, 0, 0),
    _id: "",
  });

  records.push({
    __type__: "cc.Node",
    _name: "UIRoot",
    _objFlags: 0,
    __editorExtras__: {},
    _parent: { __id__: rootNodeIndex },
    _children: [],
    _active: true,
    _components: [{ __id__: 0 }], // patched later
    _prefab: { __id__: 0 }, // patched later
    _lpos: makeCocosVec3(0, 0, 0),
    _lrot: makeCocosQuat(),
    _lscale: makeCocosVec3(1, 1, 1),
    _mobility: 0,
    _layer: 1073741824,
    _euler: makeCocosVec3(0, 0, 0),
    _id: "",
  });

  cocosRenderOrderAssets.forEach((asset) => {
    const spriteMeta = spriteMetaMap.get(asset.name) || {
      imageUuid: stableUuid(`cocos:image:${rootFolderName}:${asset.name}`),
      trimInfo: null,
    };
    const trimInfo = normalizeCocosTrimInfo(asset, spriteMeta.trimInfo);
    const nodeIndex = records.length;
    const uiTransformIndex = nodeIndex + 1;
    const uiTransformCompIndex = nodeIndex + 2;
    const spriteIndex = nodeIndex + 3;
    const spriteCompIndex = nodeIndex + 4;
    const prefabInfoIndex = nodeIndex + 5;
    const contentSize = {
      width: Math.max(1, roundNumber(toNumber(asset && asset.bounds ? asset.bounds.width : trimInfo.rawWidth))),
      height: Math.max(1, roundNumber(toNumber(asset && asset.bounds ? asset.bounds.height : trimInfo.rawHeight))),
    };
    const position = asset.cocos && asset.cocos.position ? asset.cocos.position : { x: 0, y: 0 };

    records[uiRootNodeIndex]._children.push({ __id__: nodeIndex });
    records.push({
      __type__: "cc.Node",
      _name: asset.name,
      _objFlags: 0,
      __editorExtras__: {},
      _parent: { __id__: uiRootNodeIndex },
      _children: [],
      _active: true,
      _components: [{ __id__: uiTransformIndex }, { __id__: spriteIndex }],
      _prefab: { __id__: prefabInfoIndex },
      _lpos: makeCocosVec3(position.x || 0, position.y || 0, 0),
      _lrot: makeCocosQuat(),
      _lscale: makeCocosVec3(1, 1, 1),
      _mobility: 0,
      _layer: 1073741824,
      _euler: makeCocosVec3(0, 0, 0),
      _id: "",
    });
    records.push({
      __type__: "cc.UITransform",
      _name: "",
      _objFlags: 0,
      __editorExtras__: {},
      node: { __id__: nodeIndex },
      _enabled: true,
      __prefab: { __id__: uiTransformCompIndex },
      _contentSize: makeCocosSize(contentSize.width, contentSize.height),
      _anchorPoint: makeCocosVec2(0.5, 0.5),
      _id: "",
    });
    records.push({
      __type__: "cc.CompPrefabInfo",
      fileId: stableFileId(`cocos:prefab:${rootFolderName}:${asset.name}:ui-transform`),
    });
    records.push({
      __type__: "cc.Sprite",
      _name: "",
      _objFlags: 0,
      __editorExtras__: {},
      node: { __id__: nodeIndex },
      _enabled: true,
      __prefab: { __id__: spriteCompIndex },
      _customMaterial: null,
      _srcBlendFactor: 2,
      _dstBlendFactor: 4,
      _color: makeCocosColor(),
      _spriteFrame: {
        __uuid__: `${spriteMeta.imageUuid}@f9941`,
        __expectedType__: "cc.SpriteFrame",
      },
      _type: 0,
      _fillType: 0,
      _sizeMode: 0,
      _fillCenter: makeCocosVec2(0, 0),
      _fillStart: 0,
      _fillRange: 0,
      _isTrimmedMode: false,
      _useGrayscale: false,
      _atlas: null,
      _id: "",
    });
    records.push({
      __type__: "cc.CompPrefabInfo",
      fileId: stableFileId(`cocos:prefab:${rootFolderName}:${asset.name}:sprite`),
    });
    records.push(makeCocosPrefabInfo(rootNodeIndex, stableFileId(`cocos:prefab:${rootFolderName}:${asset.name}:prefab-info`)));
  });

  const uiRootTransformIndex = records.length;
  const uiRootTransformCompIndex = uiRootTransformIndex + 1;
  const uiRootPrefabInfoIndex = uiRootTransformIndex + 2;
  records.push({
    __type__: "cc.UITransform",
    _name: "",
    _objFlags: 0,
    __editorExtras__: {},
    node: { __id__: uiRootNodeIndex },
    _enabled: true,
    __prefab: { __id__: uiRootTransformCompIndex },
    _contentSize: makeCocosSize(docWidth, docHeight),
    _anchorPoint: makeCocosVec2(0.5, 0.5),
    _id: "",
  });
  records.push({
    __type__: "cc.CompPrefabInfo",
    fileId: stableFileId(`cocos:prefab:${rootFolderName}:ui-root-transform`),
  });
  records.push(makeCocosPrefabInfo(rootNodeIndex, stableFileId(`cocos:prefab:${rootFolderName}:ui-root-prefab-info`)));

  const rootTransformIndex = records.length;
  const rootTransformCompIndex = rootTransformIndex + 1;
  const rootPrefabInfoIndex = rootTransformIndex + 2;
  records.push({
    __type__: "cc.UITransform",
    _name: "",
    _objFlags: 0,
    __editorExtras__: {},
    node: { __id__: rootNodeIndex },
    _enabled: true,
    __prefab: { __id__: rootTransformCompIndex },
    _contentSize: makeCocosSize(docWidth, docHeight),
    _anchorPoint: makeCocosVec2(0.5, 0.5),
    _id: "",
  });
  records.push({
    __type__: "cc.CompPrefabInfo",
    fileId: stableFileId(`cocos:prefab:${rootFolderName}:root-transform`),
  });
  records.push(makeCocosPrefabInfo(rootNodeIndex, stableFileId(`cocos:prefab:${rootFolderName}:root-prefab-info`), true));

  records[rootNodeIndex]._components = [{ __id__: rootTransformIndex }];
  records[rootNodeIndex]._prefab = { __id__: rootPrefabInfoIndex };
  records[uiRootNodeIndex]._components = [{ __id__: uiRootTransformIndex }];
  records[uiRootNodeIndex]._prefab = { __id__: uiRootPrefabInfoIndex };

  return records;
}

function makeCocosPrefabInfo(rootNodeIndex, fileId, rootTail = false) {
  const payload = {
    __type__: "cc.PrefabInfo",
    root: { __id__: rootNodeIndex },
    asset: { __id__: 0 },
    fileId,
  };

  if (rootTail) {
    payload.targetOverrides = null;
    return payload;
  }

  payload.instance = null;
  payload.targetOverrides = null;
  payload.nestedPrefabInstanceRoots = null;
  return payload;
}

function makeCocosVec3(x, y, z) {
  return {
    __type__: "cc.Vec3",
    x: roundNumber(toNumber(x)),
    y: roundNumber(toNumber(y)),
    z: roundNumber(toNumber(z)),
  };
}

function makeCocosVec2(x, y) {
  return {
    __type__: "cc.Vec2",
    x: roundNumber(toNumber(x)),
    y: roundNumber(toNumber(y)),
  };
}

function makeCocosSize(width, height) {
  return {
    __type__: "cc.Size",
    width: roundNumber(toNumber(width)),
    height: roundNumber(toNumber(height)),
  };
}

function makeCocosQuat() {
  return {
    __type__: "cc.Quat",
    x: 0,
    y: 0,
    z: 0,
    w: 1,
  };
}

function makeCocosColor() {
  return {
    __type__: "cc.Color",
    r: 255,
    g: 255,
    b: 255,
    a: 255,
  };
}

function stableUuid(seed) {
  const hex = buildStableHex(seed, 32);
  return [
    hex.slice(0, 8),
    hex.slice(8, 12),
    hex.slice(12, 16),
    hex.slice(16, 20),
    hex.slice(20, 32),
  ].join("-");
}

function stableUnityGuid(seed) {
  return buildStableHex(`unity-guid:${seed}`, 32);
}

function stableFileId(seed) {
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  const hex = buildStableHex(`file:${seed}`, 44);
  let result = "";
  for (let index = 0; index < 44; index += 2) {
    const value = parseInt(hex.slice(index, index + 2), 16);
    result += alphabet[value % alphabet.length];
  }
  return result.slice(0, 22);
}

function buildStableHex(seed, length) {
  let hex = "";
  let salt = 0;
  while (hex.length < length) {
    hex += hashSeedToHex(`${seed}:${salt}`);
    salt += 1;
  }
  return hex.slice(0, length);
}

function hashSeedToHex(seed) {
  let h1 = 0x811c9dc5;
  let h2 = 0x9e3779b9;
  let h3 = 0x85ebca6b;
  let h4 = 0xc2b2ae35;
  const text = String(seed || "");
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index);
    h1 = Math.imul(h1 ^ code, 0x01000193);
    h2 = Math.imul(h2 ^ code, 0x27d4eb2d);
    h3 = Math.imul(h3 ^ code, 0x165667b1);
    h4 = Math.imul(h4 ^ code, 0xd3a2646c);
  }
  return [h1, h2, h3, h4]
    .map((value) => (value >>> 0).toString(16).padStart(8, "0"))
    .join("");
}

function buildEnginePrefabPayload(assets, target, imagesRelativeToEnginePackage) {
  const prefabFileName = `${sanitizeFileStem(state.docInfo.title)}_PSD.prefab`;
  const layoutAssets = shouldReverseLayoutAssetsForTarget(target) ? [...(assets || [])].reverse() : [...(assets || [])];
  return {
    version: 1,
    release: RELEASE_INFO,
    target,
    generatedAt: new Date().toISOString(),
    document: {
      name: state.docInfo.title,
      width: state.docInfo.width,
      height: state.docInfo.height,
      safeName: sanitizeFileStem(state.docInfo.title),
    },
    paths: {
      imagesRelativeToEnginePackage: imagesRelativeToEnginePackage || "./images/",
      prefabRelativeToLayout: `./${prefabFileName}`,
    },
    assets: layoutAssets.map((asset) => ({
      name: asset.name,
      texture: asset.name,
      textureFile: `${asset.name}.png`,
      width: asset.bounds.width,
      height: asset.bounds.height,
      unity: asset.unity,
      cocos: asset.cocos,
    })),
  };
}

function shouldReverseLayoutAssetsForTarget(target) {
  const normalized = String(target || "").trim().toLowerCase();
  return normalized === "unity-6.0" || normalized === "unity-6.3";
}

function buildUnityFolderMeta(guid) {
  return [
    "fileFormatVersion: 2",
    `guid: ${guid}`,
    "folderAsset: yes",
    "DefaultImporter:",
    "  externalObjects: {}",
    "  userData: ",
    "  assetBundleName: ",
    "  assetBundleVariant: ",
    "",
  ].join("\n");
}

function buildUnityTextMeta(guid) {
  return [
    "fileFormatVersion: 2",
    `guid: ${guid}`,
    "TextScriptImporter:",
    "  externalObjects: {}",
    "  userData: ",
    "  assetBundleName: ",
    "  assetBundleVariant: ",
    "",
  ].join("\n");
}

function buildUnityCsMeta(guid) {
  return [
    "fileFormatVersion: 2",
    `guid: ${guid}`,
    "MonoImporter:",
    "  externalObjects: {}",
    "  serializedVersion: 2",
    "  defaultReferences: []",
    "  executionOrder: 0",
    "  icon: {instanceID: 0}",
    "  userData: ",
    "  assetBundleName: ",
    "  assetBundleVariant: ",
    "",
  ].join("\n");
}

function buildUnitySpriteMeta(guid) {
  return [
    "fileFormatVersion: 2",
    `guid: ${guid}`,
    "TextureImporter:",
    "  internalIDToNameTable: []",
    "  externalObjects: {}",
    "  serializedVersion: 13",
    "  mipmaps:",
    "    mipMapMode: 0",
    "    enableMipMap: 0",
    "    sRGBTexture: 1",
    "    linearTexture: 0",
    "    fadeOut: 0",
    "    borderMipMap: 0",
    "    mipMapsPreserveCoverage: 0",
    "    alphaTestReferenceValue: 0.5",
    "    mipMapFadeDistanceStart: 1",
    "    mipMapFadeDistanceEnd: 3",
    "  bumpmap:",
    "    convertToNormalMap: 0",
    "    externalNormalMap: 0",
    "    heightScale: 0.25",
    "    normalMapFilter: 0",
    "    flipGreenChannel: 0",
    "  isReadable: 0",
    "  streamingMipmaps: 0",
    "  streamingMipmapsPriority: 0",
    "  vTOnly: 0",
    "  ignoreMipmapLimit: 0",
    "  grayScaleToAlpha: 0",
    "  generateCubemap: 6",
    "  cubemapConvolution: 0",
    "  seamlessCubemap: 0",
    "  textureFormat: 1",
    "  maxTextureSize: 2048",
    "  textureSettings:",
    "    serializedVersion: 2",
    "    filterMode: 1",
    "    aniso: 1",
    "    mipBias: 0",
    "    wrapU: 1",
    "    wrapV: 1",
    "    wrapW: 0",
    "  nPOTScale: 0",
    "  lightmap: 0",
    "  compressionQuality: 50",
    "  spriteMode: 1",
    "  spriteExtrude: 1",
    "  spriteMeshType: 1",
    "  alignment: 0",
    "  spritePivot: {x: 0.5, y: 0.5}",
    "  spritePixelsToUnits: 100",
    "  spriteBorder: {x: 0, y: 0, z: 0, w: 0}",
    "  spriteGenerateFallbackPhysicsShape: 1",
    "  alphaUsage: 1",
    "  alphaIsTransparency: 1",
    "  spriteTessellationDetail: -1",
    "  textureType: 8",
    "  textureShape: 1",
    "  singleChannelComponent: 0",
    "  flipbookRows: 1",
    "  flipbookColumns: 1",
    "  maxTextureSizeSet: 0",
    "  compressionQualitySet: 0",
    "  textureFormatSet: 0",
    "  ignorePngGamma: 0",
    "  applyGammaDecoding: 0",
    "  swizzle: 50462976",
    "  cookieLightType: 0",
    "  platformSettings:",
    "  - serializedVersion: 3",
    "    buildTarget: DefaultTexturePlatform",
    "    maxTextureSize: 2048",
    "    resizeAlgorithm: 0",
    "    textureFormat: -1",
    "    textureCompression: 1",
    "    compressionQuality: 50",
    "    crunchedCompression: 0",
    "    allowsAlphaSplitting: 0",
    "    overridden: 0",
    "    ignorePlatformSupport: 0",
    "    androidETC2FallbackOverride: 0",
    "    forceMaximumCompressionQuality_BC6H_BC7: 0",
    "  - serializedVersion: 3",
    "    buildTarget: Standalone",
    "    maxTextureSize: 2048",
    "    resizeAlgorithm: 0",
    "    textureFormat: -1",
    "    textureCompression: 1",
    "    compressionQuality: 50",
    "    crunchedCompression: 0",
    "    allowsAlphaSplitting: 0",
    "    overridden: 0",
    "    ignorePlatformSupport: 0",
    "    androidETC2FallbackOverride: 0",
    "    forceMaximumCompressionQuality_BC6H_BC7: 0",
    "  - serializedVersion: 3",
    "    buildTarget: iPhone",
    "    maxTextureSize: 2048",
    "    resizeAlgorithm: 0",
    "    textureFormat: -1",
    "    textureCompression: 1",
    "    compressionQuality: 50",
    "    crunchedCompression: 0",
    "    allowsAlphaSplitting: 0",
    "    overridden: 0",
    "    ignorePlatformSupport: 0",
    "    androidETC2FallbackOverride: 0",
    "    forceMaximumCompressionQuality_BC6H_BC7: 0",
    "  - serializedVersion: 3",
    "    buildTarget: Android",
    "    maxTextureSize: 2048",
    "    resizeAlgorithm: 0",
    "    textureFormat: -1",
    "    textureCompression: 1",
    "    compressionQuality: 50",
    "    crunchedCompression: 0",
    "    allowsAlphaSplitting: 0",
    "    overridden: 0",
    "    ignorePlatformSupport: 0",
    "    androidETC2FallbackOverride: 0",
    "    forceMaximumCompressionQuality_BC6H_BC7: 0",
    "  spriteSheet:",
    "    serializedVersion: 2",
    "    sprites: []",
    "    outline: []",
    "    physicsShape: []",
    "    bones: []",
    "    spriteID: 5e97eb03825dee720800000000000000",
    "    internalID: 0",
    "    vertices: []",
    "    indices: ",
    "    edges: []",
    "    weights: []",
    "    secondaryTextures: []",
    "    nameFileIdTable: {}",
    "  mipmapLimitGroupName: ",
    "  pSDRemoveMatte: 0",
    "  userData: ",
    "  assetBundleName: ",
    "  assetBundleVariant: ",
    "",
  ].join("\n");
}

function buildUnity6BuilderScript() {
  return buildUnityBuilderScript({
    className: "PsdUnity6PrefabAutoBuilder",
    postprocessorClassName: "PsdUnity6PrefabAutoPostprocessor",
    layoutFileName: "layout_unity6.json",
    layoutSearchToken: "layout_unity6",
    menuPath: "Tools/PSD Export/Rebuild All layout_unity6 Prefabs",
  });
}

function buildUnity63BuilderScript() {
  return buildUnityBuilderScript({
    className: "PsdUnity63PrefabAutoBuilder",
    postprocessorClassName: "PsdUnity63PrefabAutoPostprocessor",
    layoutFileName: "layout_unity6_3.json",
    layoutSearchToken: "layout_unity6_3",
    menuPath: "Tools/PSD Export/Rebuild All layout_unity6_3 Prefabs",
  });
}

function buildUnityBuilderScript(options) {
  const className = options && options.className ? options.className : "PsdUnity6PrefabAutoBuilder";
  const postprocessorClassName = options && options.postprocessorClassName ? options.postprocessorClassName : "PsdUnity6PrefabAutoPostprocessor";
  const layoutFileName = options && options.layoutFileName ? options.layoutFileName : "layout_unity6.json";
  const layoutSearchToken = options && options.layoutSearchToken ? options.layoutSearchToken : "layout_unity6";
  const menuPath = options && options.menuPath ? options.menuPath : "Tools/PSD Export/Rebuild All layout_unity6 Prefabs";

  return [
    "using System;",
    "using System.Collections.Generic;",
    "using System.IO;",
    "using System.Linq;",
    "using UnityEditor;",
    "using UnityEngine;",
    "using UnityEngine.UI;",
    "",
    `public static class ${className}`,
    "{",
    `    internal const string LayoutFileName = \"${layoutFileName}\";`,
    "",
    "    [Serializable] private class LayoutRoot { public LayoutDocument document; public LayoutPaths paths; public List<LayoutAsset> assets; }",
    "    [Serializable] private class LayoutDocument { public string name; public float width; public float height; }",
    "    [Serializable] private class LayoutPaths { public string imagesRelativeToEnginePackage; }",
    "    [Serializable] private class LayoutAsset { public string name; public string texture; public string textureFile; public float width; public float height; public UnityLayout unity; }",
    "    [Serializable] private class UnityLayout { public string anchor; public Vec2 anchoredPosition; public Vec2 sizeDelta; }",
    "    [Serializable] private class Vec2 { public float x; public float y; }",
    "",
    "    [InitializeOnLoadMethod]",
    "    private static void Initialize()",
    "    {",
    "        EditorApplication.delayCall += () => BuildAllLayouts(false);",
    "    }",
    "",
    `    [MenuItem(\"${menuPath}\")]`,
    "    public static void RebuildAllLayoutsFromMenu()",
    "    {",
    "        BuildAllLayouts(true);",
    "    }",
    "",
    "    internal static void BuildFromLayoutAssetPath(string assetJsonPath, bool logResult)",
    "    {",
    "        if (string.IsNullOrEmpty(assetJsonPath)) return;",
    "        if (!assetJsonPath.EndsWith(LayoutFileName, StringComparison.OrdinalIgnoreCase)) return;",
    "",
    "        var jsonFullPath = Path.GetFullPath(Path.Combine(GetProjectRoot(), assetJsonPath));",
    "        if (!File.Exists(jsonFullPath)) return;",
    "",
    "        LayoutRoot data = null;",
    "        try",
    "        {",
    "            var json = File.ReadAllText(jsonFullPath, System.Text.Encoding.UTF8);",
    "            data = JsonUtility.FromJson<LayoutRoot>(json);",
    "        }",
    "        catch (Exception error)",
    "        {",
    "            Debug.LogError($\"[PSD Export] Failed to parse layout file: {assetJsonPath}\\n{error}\");",
    "            return;",
    "        }",
    "",
    "        if (data == null || data.assets == null)",
    "        {",
    "            Debug.LogWarning($\"[PSD Export] Invalid layout file: {assetJsonPath}\");",
    "            return;",
    "        }",
    "",
    "        var docName = SanitizeName(data.document?.name);",
    "        var root = new GameObject($\"{docName}_PSD\", typeof(RectTransform));",
    "        var rootRect = root.GetComponent<RectTransform>();",
    "        rootRect.anchorMin = rootRect.anchorMax = rootRect.pivot = new Vector2(0.5f, 0.5f);",
    "        rootRect.anchoredPosition = Vector2.zero;",
    "        rootRect.anchoredPosition3D = Vector3.zero;",
    "        rootRect.sizeDelta = new Vector2(data.document.width, data.document.height);",
    "        rootRect.localRotation = Quaternion.identity;",
    "        rootRect.localScale = Vector3.one;",
    "",
    "        var layoutDir = Path.GetDirectoryName(jsonFullPath) ?? string.Empty;",
    "        var imageDir = Path.GetFullPath(Path.Combine(layoutDir, data.paths?.imagesRelativeToEnginePackage ?? \"./images/\"));",
    "",
    "        foreach (var asset in data.assets)",
    "        {",
    "            var go = new GameObject(asset.name, typeof(RectTransform), typeof(Image));",
    "            go.transform.SetParent(root.transform, false);",
    "            var rect = go.GetComponent<RectTransform>();",
    "            rect.anchorMin = rect.anchorMax = rect.pivot = new Vector2(0.5f, 0.5f);",
    "            rect.anchoredPosition3D = new Vector3(asset.unity.anchoredPosition.x, asset.unity.anchoredPosition.y, 0f);",
    "            rect.sizeDelta = new Vector2(asset.unity.sizeDelta.x, asset.unity.sizeDelta.y);",
    "            rect.localRotation = Quaternion.identity;",
    "            rect.localScale = Vector3.one;",
    "",
    "            var textureFileName = string.IsNullOrEmpty(asset.textureFile) ? $\"{asset.name}.png\" : asset.textureFile;",
    "            var textureFullPath = Path.GetFullPath(Path.Combine(imageDir, textureFileName));",
    "            var textureAssetPath = ToAssetPath(textureFullPath);",
    "            EnsureSpriteImport(textureAssetPath);",
    "            var sprite = string.IsNullOrEmpty(textureAssetPath) ? null : AssetDatabase.LoadAssetAtPath<Sprite>(textureAssetPath);",
    "            var image = go.GetComponent<Image>();",
    "            image.sprite = sprite;",
    "            if (sprite == null) image.color = new Color(1f, 0.25f, 0.25f, 0.35f);",
    "            rect.sizeDelta = new Vector2(asset.unity.sizeDelta.x, asset.unity.sizeDelta.y);",
    "        }",
    "",
    "        var layoutAssetDir = Path.GetDirectoryName(assetJsonPath)?.Replace(\"\\\\\", \"/\") ?? \"Assets\";",
    "        if (string.IsNullOrEmpty(layoutAssetDir)) layoutAssetDir = \"Assets\";",
    "        var savePath = Path.Combine(layoutAssetDir, $\"{docName}_PSD.prefab\").Replace(\"\\\\\", \"/\");",
    "        PrefabUtility.SaveAsPrefabAsset(root, savePath);",
    "",
    "        UnityEngine.Object.DestroyImmediate(root);",
    "        if (logResult) Debug.Log($\"[PSD Export] Prefab built: {savePath}\");",
    "    }",
    "",
    "    private static void BuildAllLayouts(bool logResult)",
    "    {",
    `        var layoutGuids = AssetDatabase.FindAssets(\"${layoutSearchToken} t:TextAsset\");`,
    "        foreach (var guid in layoutGuids)",
    "        {",
    "            var assetPath = AssetDatabase.GUIDToAssetPath(guid);",
    "            if (!Path.GetFileName(assetPath).Equals(LayoutFileName, StringComparison.OrdinalIgnoreCase)) continue;",
    "            BuildFromLayoutAssetPath(assetPath, logResult);",
    "        }",
    "",
    "        AssetDatabase.SaveAssets();",
    "        AssetDatabase.Refresh();",
    "    }",
    "",
    "    private static string GetProjectRoot()",
    "    {",
    "        var assetsPath = Application.dataPath.Replace(\"\\\\\", \"/\");",
    "        return assetsPath.Substring(0, assetsPath.Length - \"Assets\".Length);",
    "    }",
    "",
    "    private static string ToAssetPath(string fullPath)",
    "    {",
    "        var normalized = Path.GetFullPath(fullPath).Replace(\"\\\\\", \"/\");",
    "        var projectRoot = GetProjectRoot().Replace(\"\\\\\", \"/\");",
    "        if (!normalized.StartsWith(projectRoot, StringComparison.OrdinalIgnoreCase)) return null;",
    "        return normalized.Substring(projectRoot.Length).TrimStart('/');",
    "    }",
    "",
    "    private static string SanitizeName(string value)",
    "    {",
    "        var raw = string.IsNullOrWhiteSpace(value) ? \"PSD\" : value.Trim();",
    "        var invalid = Path.GetInvalidFileNameChars();",
    "        var chars = raw.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray();",
    "        var clean = new string(chars).Trim();",
    "        return string.IsNullOrEmpty(clean) ? \"PSD\" : clean;",
    "    }",
    "",
    "    private static void EnsureSpriteImport(string assetPath)",
    "    {",
    "        if (string.IsNullOrEmpty(assetPath)) return;",
    "        var importer = AssetImporter.GetAtPath(assetPath) as TextureImporter;",
    "        if (importer == null) return;",
    "",
    "        var changed = false;",
    "        if (importer.textureType != TextureImporterType.Sprite) { importer.textureType = TextureImporterType.Sprite; changed = true; }",
    "        if (importer.spriteImportMode != SpriteImportMode.Single) { importer.spriteImportMode = SpriteImportMode.Single; changed = true; }",
    "        if (changed) importer.SaveAndReimport();",
    "    }",
    "}",
    "",
    `public sealed class ${postprocessorClassName} : AssetPostprocessor`,
    "{",
    "    static void OnPostprocessAllAssets(string[] importedAssets, string[] deletedAssets, string[] movedAssets, string[] movedFromAssetPaths)",
    "    {",
    "        var rebuilt = false;",
    "        foreach (var path in importedAssets)",
    "        {",
    `            if (!Path.GetFileName(path).Equals(${className}.LayoutFileName, StringComparison.OrdinalIgnoreCase)) continue;`,
    `            ${className}.BuildFromLayoutAssetPath(path, true);`,
    "            rebuilt = true;",
    "        }",
    "",
    "        if (rebuilt)",
    "        {",
    "            AssetDatabase.SaveAssets();",
    "            AssetDatabase.Refresh();",
    "        }",
    "    }",
    "}",
    "",
  ].join("\n");
}

function buildCocos388BuilderScript() {
  return [
    "import { _decorator, Component, JsonAsset, Node, Sprite, SpriteFrame, UITransform, Vec3, Layers, director, resources } from 'cc';",
    "import { EDITOR } from 'cc/env';",
    "const { ccclass, property, executeInEditMode } = _decorator;",
    "",
    "@ccclass('PsdCocos388PrefabBuilder')",
    "@executeInEditMode",
    "export class PsdCocos388PrefabBuilder extends Component {",
    "  @property({ type: JsonAsset })",
    "  layoutJson: JsonAsset | null = null;",
    "",
    "  @property",
    "  imageResourceDir = 'psd_images';",
    "",
    "  @property",
    "  autoBuildInEditor = true;",
    "",
    "  @property",
    "  buildOnLoad = true;",
    "",
    "  private building = false;",
    "  private pendingBuild = false;",
    "  private lastBuildKey = '';",
    "",
    "  onLoad() {",
    "    if (this.buildOnLoad) this.scheduleRebuild('onLoad');",
    "  }",
    "",
    "  onEnable() {",
    "    if (this.buildOnLoad) this.scheduleRebuild('onEnable');",
    "  }",
    "",
    "  update() {",
    "    if (!EDITOR || !this.autoBuildInEditor || this.building) return;",
    "    const currentKey = this.computeBuildKey();",
    "    if (!currentKey || currentKey === this.lastBuildKey) return;",
    "    this.scheduleRebuild('layout changed');",
    "  }",
    "",
    "  async rebuild() {",
    "    if (this.building) return;",
    "    if (!this.layoutJson) {",
    "      console.warn('[PSD Export] layoutJson is not assigned.');",
    "      return;",
    "    }",
    "",
    "    this.building = true;",
    "    this.pendingBuild = false;",
    "    try {",
    "      const data = this.layoutJson.json as any;",
    "      const assets = (data.assets || []) as any[];",
    "      const documentInfo = data.document || {};",
    "      this.clearChildren();",
    "      this.node.layer = Layers.Enum.UI_2D;",
    "      const rootTransform = this.node.getComponent(UITransform) || this.node.addComponent(UITransform);",
    "      rootTransform.setContentSize(documentInfo.width || 0, documentInfo.height || 0);",
    "",
    "      const renderAssets = [...assets].reverse();",
    "      for (const asset of renderAssets) {",
    "        const child = new Node(asset.name);",
    "        child.setParent(this.node);",
    "        child.layer = Layers.Enum.UI_2D;",
    "        const cocos = asset.cocos || {};",
    "        const position = cocos.position || { x: 0, y: 0 };",
    "        const bounds = asset.bounds || {};",
    "        const contentSize = { width: bounds.width || asset.width || 0, height: bounds.height || asset.height || 0 };",
    "        child.setPosition(new Vec3(position.x || 0, position.y || 0, 0));",
    "",
    "        const transform = child.addComponent(UITransform);",
    "        transform.setContentSize(contentSize.width || 0, contentSize.height || 0);",
    "",
    "        const sprite = child.addComponent(Sprite);",
    "        sprite.sizeMode = Sprite.SizeMode.CUSTOM;",
    "        sprite.trim = false;",
    "        const textureFile = asset.textureFile || `${asset.texture || asset.name}.png`;",
    "        const frame = await this.loadSpriteFrame(textureFile);",
    "        if (frame) sprite.spriteFrame = frame;",
    "      }",
    "",
    "      this.lastBuildKey = this.computeBuildKey();",
    "      if (EDITOR) console.info(`[PSD Export] Cocos prefab rebuilt: ${assets.length} nodes.`);",
    "    } catch (error) {",
    "      console.error('[PSD Export] Cocos rebuild failed:', error);",
    "    } finally {",
    "      this.building = false;",
    "    }",
    "  }",
    "",
    "  private scheduleRebuild(reason: string) {",
    "    if (this.pendingBuild || this.building) return;",
    "    this.pendingBuild = true;",
    "    director.once(director.EVENT_AFTER_UPDATE, () => {",
    "      this.rebuild().catch((error) => console.error(`[PSD Export] Rebuild error (${reason}):`, error));",
    "    });",
    "  }",
    "",
    "  private computeBuildKey(): string {",
    "    if (!this.layoutJson) return '';",
    "    const data = this.layoutJson.json as any;",
    "    const generatedAt = String(data?.generatedAt || '');",
    "    const docName = String(data?.document?.name || '');",
    "    const assetCount = Array.isArray(data?.assets) ? data.assets.length : 0;",
    "    const uuid = this.layoutJson.uuid || (this.layoutJson as any)._uuid || '';",
    "    return `${uuid}|${this.imageResourceDir}|${generatedAt}|${docName}|${assetCount}`;",
    "  }",
    "",
    "  private clearChildren() {",
    "    const children = [...this.node.children];",
    "    for (const child of children) {",
    "      child.removeFromParent();",
    "      child.destroy();",
    "    }",
    "  }",
    "",
    "  private loadSpriteFrame(textureFileName: string): Promise<SpriteFrame | null> {",
    "    const file = String(textureFileName || '').trim();",
    "    const noExt = file.replace(/\\.[^/.]+$/, '');",
    "    const path = `${this.imageResourceDir}/${noExt}/spriteFrame`;",
    "    return new Promise((resolve) => {",
    "      resources.load(path, SpriteFrame, (error, asset) => {",
    "        if (error) {",
    "          console.warn(`[PSD Export] Missing sprite frame: ${path}`);",
    "          resolve(null);",
    "          return;",
    "        }",
    "        resolve(asset);",
    "      });",
    "    });",
    "  }",
    "}",
    "",
  ].join("\n");
}

function buildCocos388ExtensionPackageJson() {
  return JSON.stringify({
    name: "psd-export-pipeline-cocos388",
    version: RELEASE_INFO.version,
    package_version: 2,
    description: "Auto build PSD prefabs from layout_cocos3_8_8.json",
    main: "./dist/main.js",
    contributions: {
      scene: {
        script: "./dist/scene.js",
      },
    },
  }, null, 2);
}

function buildCocos388ExtensionMainScript() {
  return [
    "const fs = require('fs');",
    "const path = require('path');",
    "",
    "const EXTENSION_NAME = 'psd-export-pipeline-cocos388';",
    "const LAYOUT_FILE_NAME = 'layout_cocos3_8_8.json';",
    "const POLL_INTERVAL_MS = 1800;",
    "",
    "let timer = null;",
    "let running = false;",
    "let pending = false;",
    "const fingerprints = new Map();",
    "",
    "exports.load = function load() {",
    "  queueScan(250);",
    "  timer = setInterval(() => queueScan(0), POLL_INTERVAL_MS);",
    "};",
    "",
    "exports.unload = function unload() {",
    "  if (timer) {",
    "    clearInterval(timer);",
    "    timer = null;",
    "  }",
    "};",
    "",
    "function queueScan(delayMs) {",
    "  setTimeout(() => {",
    "    scanLayouts().catch((error) => console.error('[PSD Export] Cocos auto prefab scan failed:', error));",
    "  }, delayMs);",
    "}",
    "",
    "async function scanLayouts() {",
    "  if (running) {",
    "    pending = true;",
    "    return;",
    "  }",
    "",
    "  running = true;",
    "  try {",
    "    const assetsRoot = path.join(Editor.Project.path, 'assets');",
    "    const layoutFiles = collectLayoutFiles(assetsRoot);",
    "    const active = new Set(layoutFiles);",
    "",
    "    for (const file of layoutFiles) {",
    "      const stat = fs.statSync(file);",
    "      const fingerprint = `${stat.size}:${stat.mtimeMs}`;",
    "      if (fingerprints.get(file) === fingerprint) continue;",
    "",
    "      const layoutData = JSON.parse(fs.readFileSync(file, 'utf8'));",
    "      const result = await Editor.Message.request('scene', 'execute-scene-script', {",
    "        name: EXTENSION_NAME,",
    "        method: 'buildLayoutPrefab',",
    "        args: [layoutData],",
    "      });",
    "",
    "      if (!result || !result.ok) {",
    "        const message = result && result.error ? result.error : 'unknown error';",
    "        console.warn(`[PSD Export] Auto prefab build skipped: ${file} -> ${message}`);",
    "        continue;",
    "      }",
    "",
    "      const prefabFullPath = path.resolve(path.dirname(file), result.prefabRelativePath || `./${safeFileName(layoutData?.document?.safeName || layoutData?.document?.name || 'PSD')}_PSD.prefab`);",
    "      fs.mkdirSync(path.dirname(prefabFullPath), { recursive: true });",
    "      fs.writeFileSync(prefabFullPath, String(result.serializedPrefab || ''), 'utf8');",
    "      fingerprints.set(file, fingerprint);",
    "",
    "      if (Array.isArray(result.missingFrames) && result.missingFrames.length) {",
    "        console.warn(`[PSD Export] Prefab saved with missing sprite frames (${result.missingFrames.length}): ${prefabFullPath}`);",
    "      } else {",
    "        console.info(`[PSD Export] Prefab saved: ${prefabFullPath}`);",
    "      }",
    "    }",
    "",
    "    for (const key of Array.from(fingerprints.keys())) {",
    "      if (!active.has(key)) fingerprints.delete(key);",
    "    }",
    "  } finally {",
    "    running = false;",
    "    if (pending) {",
    "      pending = false;",
    "      queueScan(0);",
    "    }",
    "  }",
    "}",
    "",
    "function collectLayoutFiles(rootDir) {",
    "  if (!fs.existsSync(rootDir)) return [];",
    "  const results = [];",
    "  const stack = [rootDir];",
    "  while (stack.length) {",
    "    const current = stack.pop();",
    "    const entries = fs.readdirSync(current, { withFileTypes: true });",
    "    for (const entry of entries) {",
    "      const fullPath = path.join(current, entry.name);",
    "      if (entry.isDirectory()) {",
    "        stack.push(fullPath);",
    "        continue;",
    "      }",
    "      if (entry.isFile() && entry.name === LAYOUT_FILE_NAME) {",
    "        results.push(fullPath);",
    "      }",
    "    }",
    "  }",
    "  return results.sort();",
    "}",
    "",
    "function safeFileName(input) {",
    "  return String(input || 'PSD').replace(/[\\\\/:*?\"<>|]+/g, '_').replace(/\\s+/g, '_').replace(/_+/g, '_').replace(/^_+|_+$/g, '') || 'PSD';",
    "}",
    "",
  ].join("\n");
}

function buildCocos388ExtensionSceneScript() {
  return [
    "const path = require('path');",
    "module.paths.push(path.join(Editor.App.path, 'node_modules'));",
    "const { Layers, Node, Sprite, SpriteFrame, UITransform, Vec3, resources } = require('cc');",
    "",
    "exports.methods = {",
    "  async buildLayoutPrefab(layoutData) {",
    "    if (!layoutData || !Array.isArray(layoutData.assets)) {",
    "      return { ok: false, error: 'layout data is invalid' };",
    "    }",
    "",
    "    const missingFrames = [];",
    "    const documentInfo = layoutData.document || {};",
    "    const rootName = `${safeFileName(documentInfo.safeName || documentInfo.name || 'PSD')}_PSD`;",
    "    const root = new Node(rootName);",
    "    root.layer = Layers.Enum.UI_2D;",
    "    const rootTransform = root.addComponent(UITransform);",
    "    rootTransform.setContentSize(documentInfo.width || 0, documentInfo.height || 0);",
    "",
    "    try {",
    "      const renderAssets = [...layoutData.assets].reverse();",
    "      for (const asset of renderAssets) {",
    "        const child = new Node(asset.name);",
    "        child.layer = Layers.Enum.UI_2D;",
    "        child.setParent(root);",
    "",
    "        const cocos = asset.cocos || {};",
    "        const position = cocos.position || { x: 0, y: 0 };",
    "        const bounds = asset.bounds || {};",
    "        const contentSize = { width: bounds.width || asset.width || 0, height: bounds.height || asset.height || 0 };",
    "        child.setPosition(new Vec3(position.x || 0, position.y || 0, 0));",
    "",
    "        const transform = child.addComponent(UITransform);",
    "        transform.setContentSize(contentSize.width || 0, contentSize.height || 0);",
    "",
    "        const sprite = child.addComponent(Sprite);",
    "        sprite.sizeMode = Sprite.SizeMode.CUSTOM;",
    "        sprite.trim = false;",
    "        const textureFile = asset.textureFile || `${asset.texture || asset.name}.png`;",
    "        const frame = await loadSpriteFrame(layoutData, textureFile);",
    "        if (frame) {",
    "          sprite.spriteFrame = frame;",
    "        } else {",
    "          missingFrames.push(textureFile);",
    "        }",
    "      }",
    "",
    "      const prefabAsset = createPrefabAsset(root);",
    "      const serializedPrefab = serializePrefab(prefabAsset);",
    "      if (!serializedPrefab) {",
    "        return { ok: false, error: 'prefab serializer is unavailable', missingFrames };",
    "      }",
    "",
    "      return {",
    "        ok: true,",
    "        prefabRelativePath: layoutData?.paths?.prefabRelativeToLayout || `./${rootName}.prefab`,",
    "        serializedPrefab,",
    "        missingFrames,",
    "        assetCount: layoutData.assets.length,",
    "      };",
    "    } catch (error) {",
    "      return { ok: false, error: error && error.message ? error.message : String(error), missingFrames };",
    "    } finally {",
    "      root.destroyAllChildren();",
    "      root.destroy();",
    "    }",
    "  },",
    "};",
    "",
    "async function loadSpriteFrame(layoutData, textureFileName) {",
    "  const imageDir = String(layoutData?.paths?.imagesRelativeToEnginePackage || './resources/psd_images/').replace(/\\\\/g, '/');",
    "  const resourceRoot = imageDir.replace(/^\\.\\//, '').replace(/^resources\\//, '').replace(/\\/$/, '');",
    "  const noExt = String(textureFileName || '').trim().replace(/\\.[^/.]+$/, '');",
    "  const resourcePath = `${resourceRoot}/${noExt}/spriteFrame`;",
    "  return new Promise((resolve) => {",
    "    resources.load(resourcePath, SpriteFrame, (error, asset) => {",
    "      if (error) {",
    "        resolve(null);",
    "        return;",
    "      }",
    "      resolve(asset);",
    "    });",
    "  });",
    "}",
    "",
    "function createPrefabAsset(rootNode) {",
    "  try {",
    "    if (globalThis.Editor && typeof Editor.require === 'function') {",
    "      const prefabUtils = Editor.require('scene://utils/prefab');",
    "      if (prefabUtils) {",
    "        if (typeof prefabUtils.createPrefabFrom === 'function') return prefabUtils.createPrefabFrom(rootNode);",
    "        if (typeof prefabUtils.createPrefab === 'function') return prefabUtils.createPrefab(rootNode);",
    "      }",
    "    }",
    "  } catch (error) {",
    "    console.warn('[PSD Export] scene://utils/prefab unavailable:', error);",
    "  }",
    "",
    "  if (globalThis.EditorExtends) {",
    "    const candidates = [EditorExtends.PrefabUtils, EditorExtends.Prefab, EditorExtends.prefab];",
    "    for (const item of candidates) {",
    "      if (!item) continue;",
    "      if (typeof item.createPrefabFrom === 'function') return item.createPrefabFrom(rootNode);",
    "      if (typeof item.createPrefab === 'function') return item.createPrefab(rootNode);",
    "    }",
    "  }",
    "",
    "  return null;",
    "}",
    "",
    "function serializePrefab(prefabAsset) {",
    "  if (!prefabAsset) return '';",
    "  if (typeof prefabAsset === 'string') return prefabAsset;",
    "  if (globalThis.EditorExtends && typeof EditorExtends.serialize === 'function') {",
    "    return EditorExtends.serialize(prefabAsset);",
    "  }",
    "  if (globalThis.Editor && typeof Editor.serialize === 'function') {",
    "    return Editor.serialize(prefabAsset);",
    "  }",
    "  return '';",
    "}",
    "",
    "function safeFileName(input) {",
    "  return String(input || 'PSD').replace(/[\\\\/:*?\"<>|]+/g, '_').replace(/\\s+/g, '_').replace(/_+/g, '_').replace(/^_+|_+$/g, '') || 'PSD';",
    "}",
    "",
  ].join("\n");
}

function buildSpineSkeletonPayload(assets, imagesPath) {
  const spineItems = buildSpineExportItems(assets);
  const bones = [{ name: "root" }];
  const slots = [];
  const skinAttachments = {};

  spineItems.forEach((item) => {
    bones.push({
      name: item.boneName,
      parent: "root",
      x: item.boneX,
      y: item.boneY,
      scaleX: 1,
      scaleY: 1,
    });

    slots.push({
      name: item.slotName,
      bone: item.boneName,
      attachment: item.attachmentName,
    });

    skinAttachments[item.slotName] = {
      [item.attachmentName]: {
        type: "region",
        path: item.attachmentName,
        x: 0,
        y: 0,
        width: item.width,
        height: item.height,
      },
    };
  });

  return {
    skeleton: {
      hash: "",
      spine: "4.2.0",
      x: roundNumber(-toNumber(state.docInfo.width) / 2),
      y: roundNumber(-toNumber(state.docInfo.height) / 2),
      width: roundNumber(toNumber(state.docInfo.width)),
      height: roundNumber(toNumber(state.docInfo.height)),
      images: imagesPath,
      audio: "",
    },
    bones,
    slots,
    skins: [
      {
        name: "default",
        attachments: skinAttachments,
      },
    ],
    animations: {},
  };
}

function buildSpineAtlasText(assets, imagesPath) {
  const spineItems = buildSpineExportItems(assets);
  const normalizedImagesPath = normalizeSpineImagesPath(imagesPath);
  const chunks = [];

  spineItems.forEach((item) => {
    const width = Math.max(1, Math.round(item.width));
    const height = Math.max(1, Math.round(item.height));
    const textureFile = `${normalizedImagesPath}${item.textureName}.png`;
    chunks.push([
      textureFile,
      `size: ${width}, ${height}`,
      "format: RGBA8888",
      "filter: Linear,Linear",
      "repeat: none",
      item.attachmentName,
      "  rotate: false",
      "  xy: 0, 0",
      `  size: ${width}, ${height}`,
      `  orig: ${width}, ${height}`,
      "  offset: 0, 0",
      "  index: -1",
      "",
    ].join("\n"));
  });

  return chunks.join("\n");
}

function normalizeSpineJsonFileName(input) {
  const fallback = "skeleton.json";
  const name = String(input || "").trim();
  if (!name) {
    return fallback;
  }
  return name.toLowerCase().endsWith(".json") ? name : `${name}.json`;
}

function normalizeSpineImagesPath(input) {
  const value = String(input || "").trim().replace(/\\/g, "/");
  if (!value) {
    return "images/";
  }
  return value.endsWith("/") ? value : `${value}/`;
}

function resolveSpineImagesPathForExport(input) {
  const normalized = normalizeSpineImagesPath(input);
  if (normalized === "../images/") {
    return "images/";
  }
  return normalized;
}

function replaceExtension(fileName, extension) {
  const name = String(fileName || "");
  const idx = name.lastIndexOf(".");
  if (idx === -1) {
    return `${name}${extension}`;
  }
  return `${name.slice(0, idx)}${extension}`;
}

function buildSpineExportItems(assets) {
  const usedBoneNames = new Set(["root"]);
  const usedSlotNames = new Set();
  const usedAttachmentNames = new Set();

  const orderedAssets = orderAssetsForChannel(assets, state.settings, "spine");

  return orderedAssets.map((asset, index) => {
    const base = sanitizeSpineName(asset.name || `item_${index + 1}`);
    const boneBase = `bone_${base}`;
    const slotBase = sanitizeSpineName(asset.spine.slotName || asset.name || `slot_${index + 1}`);
    const attachmentBase = sanitizeSpineName(asset.spine.attachmentName || asset.name || `attachment_${index + 1}`);
    return {
      textureName: asset.name,
      boneName: reserveUniqueName(boneBase, usedBoneNames),
      slotName: reserveUniqueName(slotBase, usedSlotNames),
      attachmentName: reserveUniqueName(attachmentBase, usedAttachmentNames),
      boneX: roundNumber(toNumber(asset.spine.bonePosition.x)),
      boneY: roundNumber(toNumber(asset.spine.bonePosition.y)),
      width: roundNumber(toNumber(asset.bounds.width)),
      height: roundNumber(toNumber(asset.bounds.height)),
    };
  });
}

function sanitizeSpineName(input) {
  const normalized = String(input || "")
    .trim()
    .replace(/\s+/g, "_")
    .replace(/[\\:*?"<>|]+/g, "_")
    .replace(/^_+|_+$/g, "");
  return normalized || "item";
}

function reserveUniqueName(base, registry) {
  const safe = base || "item";
  let name = safe;
  let suffix = 2;
  while (registry.has(name)) {
    name = `${safe}_${suffix}`;
    suffix += 1;
  }
  registry.add(name);
  return name;
}

function setStatus(message, tone) {
  ui.statusBox.textContent = message;
  ui.statusBox.className = `pe-status${tone ? ` ${tone}` : ""}`;
}

function toArray(collection) {
  if (!collection) {
    return [];
  }
  if (Array.isArray(collection)) {
    return collection;
  }
  return Array.from(collection);
}

function toNumber(value) {
  if (typeof value === "number") {
    return value;
  }
  if (value && typeof value === "object") {
    if (typeof value._value === "number") {
      return value._value;
    }
    if (typeof value.value === "number") {
      return value.value;
    }
  }
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function roundNumber(value) {
  return Math.round(value * 1000) / 1000;
}

function clampNumber(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function joinNativePath(folder, child) {
  const base = folder.nativePath || folder.name || "";
  if (!base) {
    return child;
  }
  return `${base}\\${child}`;
}

function escapeHtml(input) {
  return String(input)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}


