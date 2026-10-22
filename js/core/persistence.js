(function (global) {
  "use strict";

  const FILE_HANDLE_DB = "poleCalculatorFileHandles";
  const FILE_HANDLE_STORE = "handles";
  const SAVE_HANDLE_KEY = "currentSaveFile";
  const JSON_PICKER_ID = "pole-calculator-json";

  function safeJobFilePart(value) {
    const raw = String(value || "pole_job")
      .replace(/\.(?:json|xlsx?|xlsm|csv)$/i, "")
      .replace(/_(?:Pole_Calculator|AutoProposed|Debug)$/i, "")
      .replace(/^excel[_\-\s]*/i, "")
      .replace(/[_\-\s]*20\d{2}[-_]\d{2}[-_]\d{2}$/i, "")
      .trim();
    return (raw || "pole_job")
      .replace(/[<>:"/\\|?*\u0000-\u001F]+/g, "_")
      .replace(/\s+/g, "_");
  }

  function create({ S, Calculations, ProjectExport, ExcelImport, ExcelReview, els, render, recordUndoSnapshot, toast }) {
    let saveFileHandle = null;
    let lastSavedSerialized = "";
    let hasUnsavedChanges = false;

    function saveFileName() {
      const state = S.getState();
      return `${safeJobFilePart(state.jobName || state.importedFileName || "pole_job")}_Pole_Calculator.json`;
    }

    function openFileHandleDb() {
      return new Promise((resolve, reject) => {
        if (!global.indexedDB) return resolve(null);
        const request = global.indexedDB.open(FILE_HANDLE_DB, 1);
        request.onupgradeneeded = () => request.result.createObjectStore(FILE_HANDLE_STORE);
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
    }

    async function storeFileHandle(handle) {
      if (!handle) return;
      const db = await openFileHandleDb();
      if (!db) return;
      await new Promise((resolve, reject) => {
        const tx = db.transaction(FILE_HANDLE_STORE, "readwrite");
        tx.objectStore(FILE_HANDLE_STORE).put(handle, SAVE_HANDLE_KEY);
        tx.oncomplete = resolve;
        tx.onerror = () => reject(tx.error);
      });
      db.close();
    }

    async function readStoredFileHandle() {
      const db = await openFileHandleDb();
      if (!db) return null;
      const handle = await new Promise((resolve, reject) => {
        const tx = db.transaction(FILE_HANDLE_STORE, "readonly");
        const request = tx.objectStore(FILE_HANDLE_STORE).get(SAVE_HANDLE_KEY);
        request.onsuccess = () => resolve(request.result || null);
        request.onerror = () => reject(request.error);
      });
      db.close();
      return handle;
    }

    function savePayload() {
      Calculations.recalculateAll();
      const state = JSON.parse(JSON.stringify(S.getState()));
      return {
        app: "pole-calculator",
        exportedAt: new Date().toISOString(),
        version: state.version || S.CURRENT_VERSION,
        sourceFile: state.importedFileName || "",
        state
      };
    }

    function serializedSavePayload() {
      return JSON.stringify(savePayload(), null, 2);
    }

    function updateSaveButtonState() {
      if (!els.saveLocalBtn) return;
      els.saveLocalBtn.textContent = hasUnsavedChanges ? "Save *" : "Save";
      els.saveLocalBtn.title = saveFileHandle
        ? "Save changes to the selected JSON file"
        : "Choose where to save the job JSON file";
    }

    function markDirty() {
      hasUnsavedChanges = true;
      updateSaveButtonState();
    }

    function markClean(serialized = "") {
      lastSavedSerialized = serialized || JSON.stringify(savePayload());
      hasUnsavedChanges = false;
      updateSaveButtonState();
    }

    function captureDirtyState() {
      return { lastSavedSerialized, hasUnsavedChanges };
    }

    function restoreDirtyState(snapshot) {
      if (!snapshot) return;
      lastSavedSerialized = snapshot.lastSavedSerialized || "";
      hasUnsavedChanges = snapshot.hasUnsavedChanges === true;
      updateSaveButtonState();
    }

    async function writeSaveFile(handle, text) {
      const writable = await handle.createWritable();
      await writable.write(text);
      await writable.close();
    }

    async function saveLocalFile() {
      const text = serializedSavePayload();
      const suggestedName = saveFileName();
      if (global.showSaveFilePicker) {
        if (!saveFileHandle) {
          saveFileHandle = await global.showSaveFilePicker({
            id: JSON_PICKER_ID,
            suggestedName,
            types: [{ description: "Pole Calculator JSON", accept: { "application/json": [".json"] } }]
          });
        }
        await writeSaveFile(saveFileHandle, text);
        await storeFileHandle(saveFileHandle);
        markClean(text);
        toast(`Saved ${saveFileHandle.name || suggestedName}.`, "success");
        return true;
      }
      ProjectExport.downloadJson(suggestedName, JSON.parse(text));
      markClean(text);
      toast("Browser file overwrite is not available; JSON was downloaded.", "warning");
      return true;
    }

    async function loadLocalFile() {
      let file = null;
      if (global.showOpenFilePicker) {
        const pickerOptions = {
          id: JSON_PICKER_ID,
          multiple: false,
          types: [{ description: "Pole Calculator JSON", accept: { "application/json": [".json"] } }]
        };
        const storedHandle = await readStoredFileHandle();
        if (storedHandle) pickerOptions.startIn = storedHandle;
        let handles;
        try {
          handles = await global.showOpenFilePicker(pickerOptions);
        } catch (error) {
          if (!pickerOptions.startIn || error?.name !== "TypeError") throw error;
          delete pickerOptions.startIn;
          handles = await global.showOpenFilePicker(pickerOptions);
        }
        const [handle] = handles;
        saveFileHandle = handle;
        await storeFileHandle(saveFileHandle);
        file = handle ? await handle.getFile() : null;
      } else {
        file = await new Promise(resolve => {
          const input = document.createElement("input");
          input.type = "file";
          input.accept = ".json,application/json";
          input.addEventListener("change", () => resolve(input.files?.[0] || null), { once: true });
          input.click();
        });
      }
      if (!file) return false;
      recordUndoSnapshot();
      await ExcelImport.importJsonFile(file);
      ExcelReview.clearResults();
      if (S.getState().excelReviewSource?.collection?.rows?.length) ExcelReview.runReview();
      render();
      markClean(serializedSavePayload());
      toast("JSON loaded.", "success");
      return true;
    }

    return {
      savePayload,
      serializedSavePayload,
      saveLocalFile,
      loadLocalFile,
      markDirty,
      markClean,
      updateSaveButtonState,
      hasUnsavedChanges: () => hasUnsavedChanges,
      resetSaveFileHandle: () => { saveFileHandle = null; updateSaveButtonState(); },
      captureDirtyState,
      restoreDirtyState
    };
  }

  global.AppPersistence = Object.freeze({ create, safeJobFilePart });
})(window);
