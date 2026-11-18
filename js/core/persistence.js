(function (global) {
  "use strict";

  const FILE_HANDLE_DB = "poleCalculatorFileHandles";
  const FILE_HANDLE_STORE = "handles";
  const SAVE_HANDLE_KEY = "currentSaveFile";
  const JSON_PICKER_ID = "pole-calculator-json";

  function emptySaveDiagnostics() {
    return {
      status: "not_attempted",
      attemptedAt: null,
      completedAt: null,
      fileName: "",
      method: "",
      location: "",
      locationKnown: false,
      stage: "",
      reason: "No Save attempt has been made for this job in the current session.",
      fallbackFrom: null,
      warning: null
    };
  }

  function emptyLoadDiagnostics() {
    return {
      status: "not_attempted",
      attemptedAt: null,
      completedAt: null,
      fileName: "",
      method: "",
      location: "",
      locationKnown: false,
      stage: "",
      reason: "No Load attempt has been made for this job in the current session.",
      fallbackFrom: null,
      warning: null
    };
  }

  function errorDetails(error, stage) {
    return {
      stage,
      name: String(error?.name || "Error"),
      message: String(error?.message || error || "Unknown save error")
    };
  }

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

  function create({ S, Calculations, ProjectExport, ExcelImport, ExcelReview, els, render, recordUndoSnapshot, toast, confirmFallback = async () => false }) {
    let saveFileHandle = null;
    let lastSavedSerialized = "";
    let hasUnsavedChanges = false;
    let lastSaveDiagnostics = emptySaveDiagnostics();
    let lastLoadDiagnostics = emptyLoadDiagnostics();
    let restoredFileHandle = null;

    readStoredFileHandle()
      .then(handle => { restoredFileHandle = handle; })
      .catch(error => console.warn("Could not restore the remembered JSON file location.", error));

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
      restoredFileHandle = handle;
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

    function chooseJsonFileWithInput() {
      return new Promise(resolve => {
        const input = global.document.createElement("input");
        let settled = false;
        const finish = file => {
          if (settled) return;
          settled = true;
          input.remove?.();
          resolve(file || null);
        };
        input.type = "file";
        input.accept = ".json,application/json";
        input.addEventListener("change", () => finish(input.files?.[0] || null), { once: true });
        input.addEventListener("cancel", () => finish(null), { once: true });
        input.click();
      });
    }

    async function saveLocalFile() {
      const attemptedAt = new Date().toISOString();
      let stage = "serialize";
      let suggestedName = "";
      lastSaveDiagnostics = {
        ...emptySaveDiagnostics(),
        status: "in_progress",
        attemptedAt,
        reason: "Save is in progress."
      };

      try {
        const text = serializedSavePayload();
        stage = "prepare_filename";
        suggestedName = saveFileName();
        lastSaveDiagnostics.fileName = suggestedName;

        if (global.showSaveFilePicker) {
          let handle;
          let nativeSaveError = null;
          stage = saveFileHandle ? "write_selected_file" : "choose_save_location";
          try {
            handle = saveFileHandle || await global.showSaveFilePicker({
              id: JSON_PICKER_ID,
              suggestedName,
              types: [{ description: "Pole Calculator JSON", accept: { "application/json": [".json"] } }]
            });
            stage = "write_selected_file";
            await writeSaveFile(handle, text);
          } catch (error) {
            if (error?.name === "AbortError") {
              const abortedPicker = errorDetails(error, stage);
              lastSaveDiagnostics = {
                ...lastSaveDiagnostics,
                status: "picker_aborted",
                stage,
                reason: "The Save dialog ended before a file was selected and written.",
                error: abortedPicker
              };
              const useDownloadFallback = await confirmFallback(
                "Save dialog did not finish",
                "The browser closed the file picker before the app received a destination. Download a copy to your browser's Downloads location instead?",
                "Download JSON"
              );
              if (!useDownloadFallback) {
                lastSaveDiagnostics = {
                  ...lastSaveDiagnostics,
                  status: "cancelled",
                  completedAt: new Date().toISOString(),
                  reason: "The Save dialog did not return a file, and the browser-download fallback was declined."
                };
                toast("Save cancelled; no file was downloaded.", "info");
                return false;
              }
              nativeSaveError = abortedPicker;
              saveFileHandle = null;
              console.warn("The Save picker was aborted; using the confirmed browser download fallback.", error);
            } else {
              nativeSaveError = errorDetails(error, stage);
              saveFileHandle = null;
              console.warn("Saving through the selected file handle failed; downloading JSON instead.", error);
            }
          }

          if (!nativeSaveError) {
            saveFileHandle = handle;
            markClean(text);
            lastSaveDiagnostics = {
              ...lastSaveDiagnostics,
              status: "saved",
              completedAt: new Date().toISOString(),
              fileName: handle.name || suggestedName,
              method: "file-system-access-api",
              location: "The location chosen in the Save dialog. The browser does not expose its full path to the app.",
              locationKnown: false,
              stage: "write_selected_file",
              reason: "The selected file was written and its writable stream closed successfully."
            };
            try {
              await storeFileHandle(saveFileHandle);
            } catch (error) {
              // Remembering the destination is convenient, but an IndexedDB
              // failure must not turn a completed file write into a failed save.
              const warning = errorDetails(error, "remember_save_location");
              lastSaveDiagnostics.warning = warning;
              console.warn("Could not remember the JSON save location.", error);
            }
            toast(`Saved ${saveFileHandle.name || suggestedName}.`, "success");
            return true;
          }

          stage = "fallback_download";
          lastSaveDiagnostics = {
            ...lastSaveDiagnostics,
            status: "in_progress",
            stage,
            fileName: suggestedName,
            method: "browser-download",
            location: "The browser's configured download location (usually Downloads). The app cannot confirm the final disk write.",
            locationKnown: false,
            fallbackFrom: nativeSaveError,
            reason: "The selected file could not be written; trying a browser download instead."
          };
          ProjectExport.downloadJson(suggestedName, JSON.parse(text));
          markClean(text);
          lastSaveDiagnostics = {
            ...lastSaveDiagnostics,
            status: "download_started",
            completedAt: new Date().toISOString(),
            fileName: suggestedName,
            method: "browser-download",
            location: "The browser's configured download location (usually Downloads). The app cannot confirm the final disk write.",
            locationKnown: false,
            stage,
            reason: "A browser download was initiated after the selected-file write failed."
          };
          toast("The selected location could not be written; the JSON download was started instead.", "warning");
          return true;
        }

        stage = "browser_download";
        lastSaveDiagnostics = {
          ...lastSaveDiagnostics,
          status: "in_progress",
          fileName: suggestedName,
          method: "browser-download",
          location: "The browser's configured download location (usually Downloads). The app cannot confirm the final disk write.",
          locationKnown: false,
          stage,
          reason: "Starting a browser download because the file picker is unavailable."
        };
        ProjectExport.downloadJson(suggestedName, JSON.parse(text));
        markClean(text);
        lastSaveDiagnostics = {
          ...lastSaveDiagnostics,
          status: "download_started",
          completedAt: new Date().toISOString(),
          fileName: suggestedName,
          method: "browser-download",
          location: "The browser's configured download location (usually Downloads). The app cannot confirm the final disk write.",
          locationKnown: false,
          stage,
          reason: "The browser file picker is unavailable; a browser download was initiated instead."
        };
        toast("Browser file overwrite is not available; the JSON download was started.", "warning");
        return true;
      } catch (error) {
        if (error?.name !== "AbortError" && lastSaveDiagnostics.status !== "cancelled") {
          lastSaveDiagnostics = {
            ...lastSaveDiagnostics,
            status: "failed",
            completedAt: new Date().toISOString(),
            fileName: suggestedName || lastSaveDiagnostics.fileName,
            stage,
            reason: String(error?.message || error || "Unknown save error"),
            error: errorDetails(error, stage)
          };
        }
        throw error;
      }
    }

    async function loadLocalFile() {
      const attemptedAt = new Date().toISOString();
      let file = null;
      let stage = "choose_json_file";
      let method = "";
      let pickerError = null;
      lastLoadDiagnostics = {
        ...emptyLoadDiagnostics(),
        status: "in_progress",
        attemptedAt,
        reason: "Load is in progress."
      };

      try {
        if (global.showOpenFilePicker) {
          const pickerOptions = {
            id: JSON_PICKER_ID,
            multiple: false,
            types: [{ description: "Pole Calculator JSON", accept: { "application/json": [".json"] } }]
          };
          const storedHandle = saveFileHandle || restoredFileHandle;
          if (storedHandle) pickerOptions.startIn = storedHandle;

          stage = "choose_json_file";
          try {
            let handles;
            try {
              handles = await global.showOpenFilePicker(pickerOptions);
            } catch (error) {
              if (!pickerOptions.startIn || error?.name !== "TypeError") throw error;
              delete pickerOptions.startIn;
              handles = await global.showOpenFilePicker(pickerOptions);
            }
            const [handle] = handles;
            if (!handle) {
              lastLoadDiagnostics = {
                ...lastLoadDiagnostics,
                status: "cancelled",
                completedAt: new Date().toISOString(),
                stage,
                reason: "The Open dialog returned no selected file."
              };
              return false;
            }
            stage = "read_selected_file";
            file = await handle.getFile();
            saveFileHandle = handle;
            method = "file-system-access-api";
            try {
              await storeFileHandle(saveFileHandle);
            } catch (error) {
              lastLoadDiagnostics.warning = errorDetails(error, "remember_save_location");
              console.warn("Could not remember the JSON file location after loading.", error);
            }
          } catch (error) {
            if (error?.name !== "AbortError" && error?.name !== "SecurityError") throw error;
            pickerError = errorDetails(error, stage);
            lastLoadDiagnostics = {
              ...lastLoadDiagnostics,
              status: "picker_aborted",
              stage,
              reason: "The Open dialog ended before the app received the selected JSON file.",
              error: pickerError
            };
            const useCompatibilityPicker = await confirmFallback(
              "Open dialog did not finish",
              "The browser ended the file picker before the app received the JSON file. Try the compatibility file chooser instead?",
              "Choose JSON"
            );
            if (!useCompatibilityPicker) {
              lastLoadDiagnostics = {
                ...lastLoadDiagnostics,
                status: "cancelled",
                completedAt: new Date().toISOString(),
                reason: "The Open dialog did not return a file, and the compatibility chooser was declined."
              };
              toast("Load cancelled; no file was opened.", "info");
              return false;
            }
            stage = "compatibility_file_input";
            method = "html-file-input";
            saveFileHandle = null;
            file = await chooseJsonFileWithInput();
          }
        } else {
          stage = "compatibility_file_input";
          method = "html-file-input";
          file = await chooseJsonFileWithInput();
        }

        if (!file) {
          lastLoadDiagnostics = {
            ...lastLoadDiagnostics,
            status: "cancelled",
            completedAt: new Date().toISOString(),
            stage,
            reason: "No JSON file was selected in the compatibility chooser.",
            fallbackFrom: pickerError
          };
          return false;
        }

        stage = "import_json";
        recordUndoSnapshot();
        await ExcelImport.importJsonFile(file);
        ExcelReview.clearResults();
        if (S.getState().excelReviewSource?.collection?.rows?.length) ExcelReview.runReview();
        render();
        lastSaveDiagnostics = emptySaveDiagnostics();
        markClean(serializedSavePayload());
        lastLoadDiagnostics = {
          ...lastLoadDiagnostics,
          status: "loaded",
          completedAt: new Date().toISOString(),
          fileName: file.name || "",
          method,
          location: method === "file-system-access-api"
            ? "The location selected in the Open dialog. The browser does not expose its full path to the app."
            : "Selected through the browser's compatibility file chooser; the full path is not exposed to the app.",
          locationKnown: false,
          stage,
          reason: "The JSON file was selected and imported successfully.",
          fallbackFrom: pickerError
        };
        toast("JSON loaded.", "success");
        return true;
      } catch (error) {
        lastLoadDiagnostics = {
          ...lastLoadDiagnostics,
          status: "failed",
          completedAt: new Date().toISOString(),
          fileName: file?.name || lastLoadDiagnostics.fileName,
          method,
          stage,
          reason: String(error?.message || error || "Unknown load error"),
          error: errorDetails(error, stage),
          fallbackFrom: pickerError
        };
        throw error;
      }
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
      getSaveDiagnostics: () => JSON.parse(JSON.stringify(lastSaveDiagnostics)),
      getLoadDiagnostics: () => JSON.parse(JSON.stringify(lastLoadDiagnostics)),
      resetSaveFileHandle: () => {
        saveFileHandle = null;
        lastSaveDiagnostics = emptySaveDiagnostics();
        lastLoadDiagnostics = emptyLoadDiagnostics();
        updateSaveButtonState();
      },
      captureDirtyState,
      restoreDirtyState
    };
  }

  global.AppPersistence = Object.freeze({ create, safeJobFilePart });
})(window);
