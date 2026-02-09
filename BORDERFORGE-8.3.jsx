/**
 * BorderForge Export for Photoshop (v8.3)
 * --------------------------------------
 * Created By Seventeen
 *
 * ============================================================================
 * USER MANUAL (OPERATING GUIDE)
 * ============================================================================
 *
 * Overview
 * --------
 * BorderForge Export is a batch exporter for Photoshop that takes a folder of images,
 * optionally normalizes and formats them, adds a configurable border, and exports JPEGs
 * into a clean output folder using consistent rules.
 *
 * It’s designed for “make a set look consistent and shareable” workflows (Instagram
 * is a common target), but works anywhere you want:
 * - Consistent aspect ratio canvas (or preserve original proportions with equal padding)
 * - Consistent long-side pixel sizing
 * - Configurable padding
 * - Predictable filename suffixing and directive stripping
 * - Optional JPEG size constraints (min/max KB)
 * - Crash/kill safety + resume
 * - Dry Run preview that uses a hidden temp folder and auto-opens results
 *
 * IMPORTANT: Photoshop dialogs are suppressed
 * ------------------------------------------
 * The script sets:
 *   app.displayDialogs = DialogModes.NO;
 * This prevents Photoshop from interrupting the batch with modal prompts (e.g., profile
 * mismatches). The script attempts safe defaults so runs can complete unattended.
 *
 * ============================================================================
 * SECTION A — QUICK START
 * ============================================================================
 *
 * A1) Normal Batch Export
 * ----------------------
 * 1) Run the script in Photoshop.
 * 2) In the UI:
 *    - Choose Input folder
 *    - Choose Output mode (default auto “With Borders…” or custom output folder)
 *    - Set Aspect ratio behavior (enforce ratio OR ignore ratio for equal-border mode)
 *    - Set Long side and Padding
 *    - Configure Border mode (White/Black/Auto/Average/Custom)
 *    - Configure Output (JPEG quality and optional min/max KB constraints)
 *    - Configure Advanced (sRGB conversion, embed profile, PPI metadata, skip/silent,
 *      chunking, scratch retries)
 * 3) Click Run.
 *
 * A2) Dry Run (Preview)
 * --------------------
 * Dry Run allows parameter testing without touching the real export folder.
 *
 * What Dry Run does:
 * - Picks a random image from the input folder
 * - Runs the same pipeline as batch mode
 * - Exports into a hidden temp folder under system temp
 * - Opens the exported result in Photoshop
 * - Deletes the temp JPEG file from disk after opening (document remains open)
 * - Shows original/export names and file sizes in a dialog
 *
 * Dry Run buttons:
 * - Run Again: repeats with a new random image
 * - Run Batch Now: immediately starts the real batch export (if output folder is set)
 * - Return to UI: goes back to the main dialog
 *
 * ============================================================================
 * SECTION B — UI TABS AND CONTROLS (WHAT EACH OPTION DOES)
 * ============================================================================
 *
 * B1) Folders + Preset tab
 * -----------------------
 * Remember last folder choices:
 * - When enabled, input/output folder paths are saved to SETTINGS_FILE.
 * - This is REQUIRED for crash-resume to work without showing the UI.
 *
 * Input folder:
 * - The source folder containing your images.
 * - Supported extensions: JPG/JPEG, PNG, TIF/TIFF.
 * - Files are gathered from the top level only (no recursion) and sorted by name.
 *
 * Default output:
 * - Auto output folder lives inside the input folder.
 * - It is chosen as the next available numbered folder:
 *     With Borders
 *     With Borders 2
 *     With Borders 3
 *     ...
 *
 * Output folder manager:
 * - “Delete old ‘With Borders…’ folders” removes With Borders and With Borders N
 *   folders inside the input folder (deterministic sweep up to the latest found index).
 * - It refreshes the default output display afterward.
 *
 * Custom output folder:
 * - When enabled, you choose a folder anywhere and exports go there instead of
 *   auto “With Borders…”.
 *
 * Presets:
 * - Convenience presets for common Instagram layouts.
 * - Presets update ratio, long side, and recommended max KB.
 *
 * B2) Aspect + Fit + Sizing tab
 * ----------------------------
 * Aspect dropdown:
 * - Standard ratios plus Custom.
 *
 * Ignore aspect ratio:
 * - When ON, the script DOES NOT force the canvas into the target ratio.
 * - Instead it scales down (if needed) to Long side, then adds equal padding
 *   on all sides (canvas grows by 2*padding in both dimensions).
 *
 * Custom W:H:
 * - Enabled only when “Custom” ratio is selected and Ignore aspect ratio is OFF.
 *
 * Long side (px):
 * - Target maximum dimension in pixels.
 * - Used differently depending on ignoreRatio mode:
 *   - ignoreRatio=true: caps the image’s long side if larger than longSide
 *   - ignoreRatio=false: canvasH is longSide and canvasW is computed from ratio
 *
 * Padding (px):
 * - Border padding around the content, in pixels.
 * - In ratio-enforced mode, padding defines the inner content area size.
 *
 * B3) Border + Output tab
 * ----------------------
 * Ignore border module:
 * - Skips border/canvas expansion logic entirely.
 * - Still resizes to long side (capLongSide) and exports.
 *
 * Border color dropdown:
 * - White, Black, Average (per image), AUTO (brightness), AUTO (filename directives),
 *   Custom (picked color).
 *
 * Custom picker:
 * - “Pick color…” uses Photoshop’s color picker, reads app.foregroundColor,
 *   converts to HEX, and stores it in the UI (customHexUI).
 *
 * Preview swatch:
 * - Displays only for deterministic colors (White/Black/Custom).
 * - Hidden for AUTO / AVERAGE modes because color is content-dependent.
 *
 * Output controls:
 * - JPEG quality start (0–12): initial save quality.
 * - Ignore file size limits: when ON, save once at chosen quality.
 * - Min/Max KB: when enabled, script “steps” quality up/down to best match range.
 * - Filename suffix: appended to the base filename.
 *
 * B4) Advanced tab
 * ---------------
 * sRGB conversion:
 * - OFF: no conversion
 * - AUTO: convert only if document profile does not look like sRGB
 * - FORCE: always convert to sRGB IEC61966-2.1
 *
 * Embed profile:
 * - If ON, embed the profile in the JPEG.
 *
 * Set PPI:
 * - If ON, sets document resolution metadata without resampling pixels.
 *
 * Batch stability:
 * - Skip if output exists: avoids overwriting outputs (except forced redo on resume).
 * - Silent mode: skip per-file alerts; log errors and continue.
 * - Chunk size: process N images then do deeper cleanup.
 * - Cooldown ms: optional sleep between chunks.
 * - Scratch retries: special retry loop for “scratch disks are full”.
 *
 * ============================================================================
 * SECTION C — CRASH/KILL SAFETY + RESUME (HOW TO USE IT)
 * ============================================================================
 *
 * Run-state file:
 * - RUN_STATE_FILE is written in system temp while running.
 * - It tracks lastStartedIdx/Name and lastDoneIdx/Name.
 *
 * If Photoshop crashes or is force-quit:
 * - cleanExit remains false
 * - next run prompts you to Resume
 *
 * If you Resume:
 * - Loads SETTINGS_FILE
 * - Requires rememberFolders was ON
 * - Recreates output folder choice
 * - Resumes at saved index and redoes the most recently STARTED image
 *
 * ============================================================================
 * EXTENDED MANUAL (HOW THE CODE ACHIEVES EACH FUNCTION)
 * ============================================================================
 *
 * Architecture (high level)
 * -------------------------
 * 1) Entry point sets dialog mode to NO.
 * 2) Main loop:
 *    - checkUnexpectedStopAndPrompt() -> possible resume
 *    - showMainDialog() -> user chooses options or cancels
 *    - gatherAndSortFiles() -> deterministic file list
 *    - Dry Run path -> startDryRunSession()
 *    - Run path -> runBatch()
 *
 * Batch pipeline (per file)
 * -------------------------
 * processOne() performs:
 * - safeOpenWithRetry() to mitigate intermittent “open options” errors
 * - RGB conversion, optional sRGB conversion
 * - flatten
 * - border color decision (WHITE/BLACK/CUSTOM/AUTO/AVERAGE/AUTO_FILENAME)
 * - background fill and canvas sizing logic:
 *   - ignoreRatio=true: optional downscale then equal padding canvas expansion
 *   - ignoreRatio=false: ratio canvas computed from longSide + W:H ratio, fit inside padding
 * - saveJpegRespectingIgnore() which routes to:
 *   - saveJpegOnce() if ignoreFileSizeLimits
 *   - saveJpegPossiblyRanged() which steps quality to approach min/max KB targets
 *
 * Cleanup strategy
 * ----------------
 * - lightCleanupFast() runs frequently between files/attempts
 * - deepCleanup() runs between chunks or after scratch errors
 *
 * End of Manual
 * — Seventeen
 * ============================================================================
 */

#target photoshop
app.displayDialogs = DialogModes.NO;

var PROGRESS_FILE = File(Folder.temp + "/ig_batch_progress.txt"); // kept for redundancy
var LOG_FILE = File(Folder.temp + "/ig_batch_log.txt");
var SETTINGS_FILE = File(Folder.userData + "/ig_borderforge_settings_v6_5.txt");

// Crash/kill detection + resume state
var RUN_STATE_FILE = File(Folder.temp + "/ig_borderforge_runstate_v6_5.txt");

function showFatal(e) {
  try {
    alert(
      "BorderForge crashed:\n\n" + e + "\n\n" +
      (e && e.line ? ("Line: " + e.line + "\n") : "") +
      (e && e.fileName ? ("File: " + e.fileName + "\n") : "")
    );
  } catch (_) {}
}

try {

  (function () {
    while (true) {

      // 1) Check for unexpected stop and offer resume
      var resumeDecision = checkUnexpectedStopAndPrompt();
      if (resumeDecision && resumeDecision.doResume) {
        var resumed = runFromSavedSettings(resumeDecision);
        if (resumed) return;
      }

      // 2) Normal UI loop
      var ui = showMainDialog();
      if (!ui) return;

      if (ui.restartUI) continue;

      var inFolder = ui.inFolder;
      var outFolder = ui.outFolder;
      var opts = ui.opts;

      if (!inFolder || !inFolder.exists) {
        alert("Input folder is not set or does not exist.");
        continue;
      }

      var files = gatherAndSortFiles(inFolder);
      if (!files.length) {
        alert("No image files found in the chosen input folder.");
        continue;
      }

      // Dry run path
      if (ui.action === "dryrun") {
        // NOTE: we pass the REAL outFolder in now so Dry Run can run the batch immediately.
        var dryResult = startDryRunSession(files, inFolder, outFolder, opts);

        if (dryResult && dryResult.runBatchNow) {
          // Ensure output folder exists (and create if needed)
          if (!outFolder) {
            alert("Output folder is not set. Return to the main UI and choose an output folder.");
            continue;
          }
          if (!outFolder.exists) {
            try {
              if (!outFolder.create()) throw new Error("Could not create output folder.");
            } catch (e) {
              alert("Could not create output folder:\n" + outFolder.fsName + "\n\n" + e);
              continue;
            }
          }

          clearRunStateSafe();
          var didReturnToUI = runBatch(files, 0, inFolder, outFolder, opts, -1);
          if (didReturnToUI) continue;
          return;
        }

        // Otherwise return to UI loop for tweaks
        continue;
      }

      // Run path
      if (ui.action === "run") {
        if (!outFolder) {
          alert("Output folder is not set.");
          continue;
        }
        if (!outFolder.exists) {
          try {
            if (!outFolder.create()) throw new Error("Could not create output folder.");
          } catch (e) {
            alert("Could not create output folder:\n" + outFolder.fsName + "\n\n" + e);
            continue;
          }
        }

        clearRunStateSafe();
        var didReturnToUI2 = runBatch(files, 0, inFolder, outFolder, opts, -1);
        if (didReturnToUI2) continue;
        return;
      }

      continue;
    }
  })();

} catch (e) {
  showFatal(e);
}

/* ───────────────────────────────────────────────────────────────────
 * Crash/kill detection + resume
 * ─────────────────────────────────────────────────────────────────── */

function checkUnexpectedStopAndPrompt() {
  if (!RUN_STATE_FILE.exists) return null;

  var st = loadKeyValueFileSafe(RUN_STATE_FILE);
  if (!st || !st.hasAny) {
    var yes0 = confirm(
      "An unexpected stop was detected from the last run.\n\n" +
      "Resume from where you left off?\n\n" +
      "(If you choose Yes, it will attempt to resume using your last saved settings.)"
    );
    if (!yes0) {
      clearRunStateSafe();
      return { clearAndContinueToUI: true, doResume: false };
    }
    return { doResume: true, lastStartedIdx: -1, lastStartedName: "" };
  }

  var cleanExit = (String(st.cleanExit || "").toLowerCase() === "true");
  if (cleanExit) {
    clearRunStateSafe();
    return null;
  }

  var lastStartedIdx = parseInt(st.lastStartedIdx, 10);
  if (!isFinite(lastStartedIdx)) lastStartedIdx = -1;
  var lastStartedName = String(st.lastStartedName || "");

  var msg =
    "Unexpected failure detected.\n\n" +
    (lastStartedName ? ("Last started: " + lastStartedName + "\n") : "") +
    (lastStartedIdx >= 0 ? ("Index: " + lastStartedIdx + "\n\n") : "\n") +
    "Would you like to resume from where you left off?\n\n" +
    "YES: Re-run the last started image (overwriting its output if needed) and continue.\n" +
    "NO: Clear recovery state and return to the main UI.";

  var yes = confirm(msg);
  if (!yes) {
    clearRunStateSafe();
    return { clearAndContinueToUI: true, doResume: false };
  }

  return { doResume: true, lastStartedIdx: lastStartedIdx, lastStartedName: lastStartedName };
}

function runFromSavedSettings(resumeDecision) {
  var persisted = loadSettingsSafe(SETTINGS_FILE);
  if (!persisted || !persisted.hasAny) {
    alert("Resume was requested, but no saved settings were found.\n\nReturning to the main UI.");
    return false;
  }

  var rememberFolders = toBoolLocal(persisted.rememberFolders, false);
  if (!rememberFolders) {
    alert(
      "Resume was requested, but 'Remember last folder choices' was OFF in the last saved settings.\n\n" +
      "Enable it and run once, then crash-resume can work without needing the UI."
    );
    return false;
  }

  var inPath = String(persisted.inputPath || "");
  var useCustomOut = toBoolLocal(persisted.useCustomOut, false);
  var outPath = useCustomOut ? String(persisted.customOutPath || "") : "";

  if (!inPath) {
    alert("Resume was requested, but the last input folder path is missing.\n\nReturning to the main UI.");
    return false;
  }

  var inFolder = Folder(inPath);
  if (!inFolder.exists) {
    alert("Resume failed: the last input folder no longer exists:\n\n" + inPath + "\n\nReturning to the main UI.");
    return false;
  }

  var outFolder;
  if (useCustomOut) {
    if (!outPath) {
      alert("Resume failed: the last custom output folder path is missing.\n\nReturning to the main UI.");
      return false;
    }
    outFolder = Folder(outPath);
  } else {
    outFolder = getNextNumberedSubfolder(inFolder, "With Borders");
  }

  if (!outFolder.exists) {
    try {
      if (!outFolder.create()) throw new Error("Could not create output folder.");
    } catch (e) {
      alert("Resume failed: could not create output folder:\n\n" + outFolder.fsName + "\n\n" + e + "\n\nReturning to the main UI.");
      return false;
    }
  }

  var opts = buildOptsFromPersisted(persisted);
  if (!opts) {
    alert("Resume failed: saved settings were incomplete/invalid.\n\nReturning to the main UI.");
    return false;
  }

  var files = gatherAndSortFiles(inFolder);
  if (!files.length) {
    alert("Resume failed: no image files found in the saved input folder.\n\nReturning to the main UI.");
    return false;
  }

  var st = loadKeyValueFileSafe(RUN_STATE_FILE);
  var resumeIdx = -1;

  if (st && st.hasAny) {
    var n = parseInt(st.lastStartedIdx, 10);
    if (isFinite(n)) resumeIdx = n;
  }
  if (resumeIdx < 0 && resumeDecision && isFinite(resumeDecision.lastStartedIdx)) resumeIdx = resumeDecision.lastStartedIdx;
  if (resumeIdx < 0) resumeIdx = 0;
  if (resumeIdx >= files.length) resumeIdx = Math.max(0, files.length - 1);

  var forceRedoIdx = resumeIdx;

  var didReturnToUI = runBatch(files, resumeIdx, inFolder, outFolder, opts, forceRedoIdx);
  if (didReturnToUI) return false;
  return true;
}

function buildOptsFromPersisted(s) {
  try {
    var ratioW = parseInt(s.ratioW, 10);
    var ratioH = parseInt(s.ratioH, 10);
    if (!isFinite(ratioW) || ratioW <= 0) ratioW = 4;
    if (!isFinite(ratioH) || ratioH <= 0) ratioH = 5;

    var longSide = parseInt(s.longSide, 10); if (!isFinite(longSide) || longSide < 200) longSide = 1350;
    var padding = parseInt(s.padding, 10); if (!isFinite(padding) || padding < 0) padding = 40;

    var jpegQuality = parseInt(s.jpegQuality, 10); if (!isFinite(jpegQuality) || jpegQuality < 0 || jpegQuality > 12) jpegQuality = 10;
    var minKB = parseInt(s.minKB, 10); if (!isFinite(minKB) || minKB < 0) minKB = 0;
    var maxKB = parseInt(s.maxKB, 10); if (!isFinite(maxKB) || maxKB < 0) maxKB = 0;

    var chunkSize = parseInt(s.chunkSize, 10); if (!isFinite(chunkSize) || chunkSize < 1) chunkSize = 100;
    var cooldownMs = parseInt(s.cooldownMs, 10); if (!isFinite(cooldownMs) || cooldownMs < 0) cooldownMs = 0;
    var scratchMaxRetries = parseInt(s.scratchMaxRetries, 10); if (!isFinite(scratchMaxRetries) || scratchMaxRetries < 0) scratchMaxRetries = 2;
    var scratchRetryCooldownMs = parseInt(s.scratchRetryCooldownMs, 10); if (!isFinite(scratchRetryCooldownMs) || scratchRetryCooldownMs < 0) scratchRetryCooldownMs = 5000;

    var setPPI = toBoolLocal(s.setPPI, true);
    var ppi = parseInt(s.ppi, 10); if (!isFinite(ppi) || ppi < 1) ppi = 72;

    var borderMode = String(s.borderMode || "WHITE");
    var borderHex = String(s.borderHex || "FFFFFF");

    return {
      ratioW: ratioW,
      ratioH: ratioH,
      longSide: longSide,
      padding: padding,

      ignoreRatio: toBoolLocal(s.ignoreRatio, false),

      ignoreBorder: toBoolLocal(s.ignoreBorder, false),
      ignoreFileSizeLimits: toBoolLocal(s.ignoreFileSizeLimits, false),

      borderMode: borderMode,
      borderHex: borderHex,

      jpegQuality: jpegQuality,
      minKB: minKB,
      maxKB: maxKB,

      suffix: String(s.suffix || "WB"),

      srgbMode: String(s.srgbMode || "OFF"),
      embedProfile: toBoolLocal(s.embedProfile, false),

      setPPI: setPPI,
      ppi: setPPI ? ppi : null,

      skipExisting: toBoolLocal(s.skipExisting, true),
      silentMode: toBoolLocal(s.silentMode, true),

      chunkSize: chunkSize,
      cooldownMs: cooldownMs,
      scratchMaxRetries: scratchMaxRetries,
      scratchRetryCooldownMs: scratchRetryCooldownMs,

      rememberFolders: toBoolLocal(s.rememberFolders, true)
    };

  } catch (_) {
    return null;
  }
}

function gatherAndSortFiles(inFolder) {
  var files = inFolder.getFiles(function (f) {
    return f instanceof File && f.name.match(/\.(jpe?g|png|tiff?)$/i);
  });
  files.sort(function (a, b) {
    var an = a.name.toLowerCase(), bn = b.name.toLowerCase();
    return an < bn ? -1 : (an > bn ? 1 : 0);
  });
  return files;
}

/* ───────────────────────────────────────────────────────────────────
 * Batch runner
 * Returns true if caller should continue to UI loop; false if finished.
 * ─────────────────────────────────────────────────────────────────── */
function runBatch(files, startIdx, inFolder, outFolder, opts, forceRedoIdx) {

  var BORDER_WHITE = hexToSolidColor("FFFFFF");
  var BORDER_BLACK = hexToSolidColor("000000");
  var BORDER_CUSTOM = hexToSolidColor(opts.borderHex);

  var total = files.length;
  var idx = clampInt(startIdx, 0, Math.max(0, total - 1));

  var prog = makeProgressUI(total, idx);
  try { prog.win.show(); prog.win.active = true; } catch (_) {}

  logLine("----- RUN -----");
  logLine("Created By Seventeen");
  logLine("Input: " + inFolder.fsName);
  logLine("Output: " + outFolder.fsName);
  logLine("Start index: " + idx);
  logLine("Force redo idx: " + forceRedoIdx);

  var originalRulerUnits = app.preferences.rulerUnits;
  var originalHistoryStates = app.preferences.numberOfHistoryStates;

  app.preferences.rulerUnits = Units.PIXELS;
  try { app.preferences.numberOfHistoryStates = 1; } catch (_) {}

  writeRunStateBase(inFolder, outFolder, idx, total);

  var skipped = 0;
  updateProgressUI(prog, idx, total, "Starting…");

  var endedNormally = true;

  while (idx < total) {

    var processedThisChunk = 0;

    while (idx < total && processedThisChunk < opts.chunkSize) {

      var fileObj = files[idx];
      var outFile = makeOutputFile(outFolder, fileObj, opts.suffix);

      markFileStarted(idx, fileObj.name);

      updateProgressUI(prog, idx, total, "Processing: " + fileObj.name);

      var skipThis = (opts.skipExisting && outFile.exists && idx !== forceRedoIdx);

      if (skipThis) {
        skipped++;
        logLine("SKIP (exists) " + fileObj.name);
        idx++; processedThisChunk++;
        trySaveProgress(idx);
        lightCleanupFast();
        updateProgressUI(prog, idx, total, "Skipped (exists): " + fileObj.name);
        markFileDone(idx - 1, fileObj.name);
        continue;
      }

      var scratchAttempts = 0;

      while (true) {
        try {
          processOne(fileObj, outFile, opts, BORDER_WHITE, BORDER_BLACK, BORDER_CUSTOM);
          logLine("OK " + fileObj.name);
          markFileDone(idx, fileObj.name);
          break;

        } catch (e) {
          var msg = String(e);
          logLine("ERR " + fileObj.name + " :: " + msg);

          if (/scratch disks are full/i.test(msg)) {
            scratchAttempts++;
            if (scratchAttempts <= opts.scratchMaxRetries) {
              deepCleanup();
              sleepMs(opts.scratchRetryCooldownMs);
              continue;
            }

            finalizePrefs(originalRulerUnits, originalHistoryStates);
            trySaveProgress(idx);
            try { if (prog && prog.win) prog.win.close(); } catch (_) {}

            alert(
              "Scratch disk error. Photoshop could not recover within the script.\n\n" +
              "Quit and relaunch Photoshop, then run again.\n" +
              "The script will detect the interrupted run and offer Resume.\n\n" +
              "Last file attempted: " + fileObj.name + "\n" +
              "Log: " + LOG_FILE.fsName
            );
            return true;
          }

          skipped++;
          if (!opts.silentMode) {
            finalizePrefs(originalRulerUnits, originalHistoryStates);
            trySaveProgress(idx);
            try { if (prog && prog.win) prog.win.close(); } catch (_) {}
            alert("Photoshop error while processing:\n\n" + fileObj.name + "\n\n" + msg);
            endedNormally = false;
            break;
          }
          break;

        } finally {
          lightCleanupFast();
        }
      }

      if (!endedNormally) break;

      idx++; processedThisChunk++;
      trySaveProgress(idx);
      updateProgressUI(prog, idx, total, "Done: " + fileObj.name);

      if (forceRedoIdx === (idx - 1)) forceRedoIdx = -1;
    }

    if (!endedNormally) break;

    deepCleanup();
    if (opts.cooldownMs > 0) sleepMs(opts.cooldownMs);
  }

  finalizePrefs(originalRulerUnits, originalHistoryStates);

  try { if (prog && prog.win) prog.win.close(); } catch (_) {}

  if (!endedNormally) {
    return true;
  }

  try { if (PROGRESS_FILE.exists) PROGRESS_FILE.remove(); } catch (_) {}
  markRunCleanExitAndRemove();

  alert(
    "Completed.\n\n" +
    "Processed: " + total + "\n" +
    "Skipped: " + skipped + "\n" +
    "Output: " + outFolder.fsName + "\n" +
    "Log: " + LOG_FILE.fsName
  );
  return false;
}

function writeRunStateBase(inFolder, outFolder, startIdx, total) {
  var st = {};
  st.running = "true";
  st.cleanExit = "false";
  st.startedAt = (new Date()).toString();
  st.inputPath = inFolder ? inFolder.fsName : "";
  st.outputPath = outFolder ? outFolder.fsName : "";
  st.total = String(total);
  st.lastStartedIdx = String(startIdx);
  st.lastStartedName = "";
  st.lastDoneIdx = "-1";
  st.lastDoneName = "";
  saveKeyValueFileSafe(RUN_STATE_FILE, st);
}

function markFileStarted(idx, name) {
  var st = loadKeyValueFileSafe(RUN_STATE_FILE);
  if (!st) st = {};
  st.running = "true";
  st.cleanExit = "false";
  st.heartbeat = (new Date()).toString();
  st.lastStartedIdx = String(idx);
  st.lastStartedName = String(name || "");
  saveKeyValueFileSafe(RUN_STATE_FILE, st);
}

function markFileDone(idx, name) {
  var st = loadKeyValueFileSafe(RUN_STATE_FILE);
  if (!st) st = {};
  st.running = "true";
  st.cleanExit = "false";
  st.heartbeat = (new Date()).toString();
  st.lastDoneIdx = String(idx);
  st.lastDoneName = String(name || "");
  saveKeyValueFileSafe(RUN_STATE_FILE, st);
}

function markRunCleanExitAndRemove() {
  try {
    if (!RUN_STATE_FILE.exists) return;
    var st = loadKeyValueFileSafe(RUN_STATE_FILE);
    if (!st) st = {};
    st.running = "false";
    st.cleanExit = "true";
    st.endedAt = (new Date()).toString();
    saveKeyValueFileSafe(RUN_STATE_FILE, st);
    try { RUN_STATE_FILE.remove(); } catch (_) {}
  } catch (_) {}
}

function clearRunStateSafe() {
  try { if (RUN_STATE_FILE.exists) RUN_STATE_FILE.remove(); } catch (_) {}
}

function loadKeyValueFileSafe(fileObj) {
  var out = { hasAny: false };
  try {
    if (!fileObj || !fileObj.exists) return out;
    fileObj.open("r");
    var text = fileObj.read();
    fileObj.close();

    var lines = String(text).split(/\r?\n/);
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];
      if (!line) continue;
      if (/^\s*#/.test(line)) continue;
      var eq = line.indexOf("=");
      if (eq < 0) continue;
      var k = line.substring(0, eq);
      var v = line.substring(eq + 1);
      k = String(k).replace(/^\s+|\s+$/g, "");
      v = String(v).replace(/^\s+|\s+$/g, "");
      if (!k) continue;
      out[k] = v;
      out.hasAny = true;
    }
    return out;
  } catch (_) {
    try { if (fileObj && fileObj.opened) fileObj.close(); } catch (__) {}
    return { hasAny: false };
  }
}

function saveKeyValueFileSafe(fileObj, obj) {
  try {
    if (!fileObj) return;
    fileObj.open("w");
    fileObj.writeln("# BorderForge run-state (v8.3)");
    fileObj.writeln("# Created By Seventeen");
    for (var k in obj) {
      if (!obj.hasOwnProperty(k)) continue;
      var v = obj[k];
      if (typeof v === "undefined") continue;
      fileObj.writeln(String(k) + "=" + String(v));
    }
    fileObj.close();
  } catch (_) {
    try { if (fileObj && fileObj.opened) fileObj.close(); } catch (__) {}
  }
}

function toBoolLocal(v, fallback) {
  if (typeof v === "boolean") return v;
  var s = String(v).toLowerCase();
  if (s === "true" || s === "1" || s === "yes") return true;
  if (s === "false" || s === "0" || s === "no") return false;
  return !!fallback;
}

/* ───────────────────────────────────────────────────────────────────
 * Output folder purge (BUTTON-TRIGGERED, DETERMINISTIC)
 * ─────────────────────────────────────────────────────────────────── */

function getNextNumberedSubfolderAndIndex(parentFolder, baseName) {
  var n = 1;
  while (true) {
    var name = (n === 1) ? baseName : (baseName + " " + n);
    var candidate = Folder(parentFolder.fsName + "/" + name);
    if (!candidate.exists) return { folder: candidate, nextIndex: n };
    n++;
  }
}

function deleteFolderTree(folderObj) {
  try {
    if (!folderObj || !folderObj.exists) return false;

    var items = folderObj.getFiles();
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      try {
        if (it instanceof Folder) {
          deleteFolderTree(it);
        } else {
          try { it.remove(); } catch (_) {}
        }
      } catch (_) {}
    }

    try { return folderObj.remove(); } catch (_) { return false; }
  } catch (_) {
    return false;
  }
}

function purgeUsingNextIndex(parentFolder, baseName, nextIndex) {
  var deleted = 0;
  var maxN = (isFinite(nextIndex) ? (nextIndex - 1) : 0);
  if (!parentFolder || !parentFolder.exists) return 0;
  if (maxN < 1) return 0;

  logLine("PURGE deterministic: parent=" + parentFolder.fsName + " base=" + baseName + " maxN=" + maxN);

  var f1 = Folder(parentFolder.fsName + "/" + baseName);
  if (f1.exists) {
    logLine("PURGE deleting: " + f1.fsName);
    if (deleteFolderTree(f1)) deleted++;
    else logLine("PURGE failed delete: " + f1.fsName);
  } else {
    logLine("PURGE missing: " + f1.fsName);
  }

  for (var n = 2; n <= maxN; n++) {
    var fn = Folder(parentFolder.fsName + "/" + baseName + " " + n);
    if (fn.exists) {
      logLine("PURGE deleting: " + fn.fsName);
      if (deleteFolderTree(fn)) deleted++;
      else logLine("PURGE failed delete: " + fn.fsName);
    } else {
      logLine("PURGE missing: " + fn.fsName);
    }
  }

  logLine("PURGE done. Deleted " + deleted + " folder(s).");
  return deleted;
}

/* ───────────────────────────────────────────────────────────────────
 * UI
 * ─────────────────────────────────────────────────────────────────── */
function showMainDialog() {

  var DEFAULTS = getDefaultSettings();
  var persisted = loadSettingsSafe(SETTINGS_FILE);

  var targetW = 700;
  var targetH = 700;

  var w = new Window("dialog", "BorderForge Export (v8.3)");
  w.orientation = "column";
  w.alignChildren = "fill";
  w.spacing = 10;
  w.margins = 12;

  try { w.resizable = true; } catch (_) {}
  w.preferredSize = [targetW, targetH];
  w.minimumSize = [560, 620];

  var tabs = w.add("tabbedpanel");
  tabs.alignChildren = "fill";

  var tabFolders = tabs.add("tab", undefined, "Folders + Preset");
  tabFolders.orientation = "column";
  tabFolders.alignChildren = "fill";

  var tabSizing = tabs.add("tab", undefined, "Aspect + Fit + Sizing");
  tabSizing.orientation = "column";
  tabSizing.alignChildren = "fill";

  var tabBorderOut = tabs.add("tab", undefined, "Border + Output");
  tabBorderOut.orientation = "column";
  tabBorderOut.alignChildren = "fill";

  var tabAdv = tabs.add("tab", undefined, "Advanced");
  tabAdv.orientation = "column";
  tabAdv.alignChildren = "fill";

  tabs.selection = tabFolders;

  // ─────────────────────────────────────────────────────────────
  // INFO TEXT (used by the info buttons)
  // ─────────────────────────────────────────────────────────────
  var INFO = {
    folders: {
      title: "Folders",
      body:
        "What it does:\n" +
        "• Selects the input folder to batch-process.\n" +
        "• Controls where output goes (default auto-numbered “With Borders…” or a custom folder).\n\n" +
        "How to use:\n" +
        "1) Click “Choose…” next to Input and pick a folder containing images.\n" +
        "2) Decide output mode:\n" +
        "   - Default output: Auto-creates the next available “With Borders”, “With Borders 2”, etc.\n" +
        "   - Custom output: Enable the checkbox and pick a folder.\n\n" +
        "Notes:\n" +
        "• Input folder must exist.\n" +
        "• Default output is computed from the input folder.\n" +
        "• Custom output must be selected + chosen before running."
    },
    outmgr: {
      title: "Output folder manager",
      body:
        "What it does:\n" +
        "• Deletes old auto-generated “With Borders…” folders inside the input folder.\n" +
        "• It is deterministic: it only deletes folders it knows belong to the sequence.\n\n" +
        "How it decides what to delete:\n" +
        "• If the next default output would be “With Borders 7”, then it assumes 1..6 exist and are deletable.\n\n" +
        "How to use:\n" +
        "1) Choose an input folder.\n" +
        "2) Click “Delete old “With Borders…” folders”.\n" +
        "3) Confirm the list and it will remove them.\n\n" +
        "Warning:\n" +
        "• This permanently deletes folders and their contents."
    },
    preset: {
      title: "Preset",
      body:
        "What it does:\n" +
        "• Quickly fills in common values for aspect ratio, long side, and file size limits.\n\n" +
        "How to use:\n" +
        "• Pick a preset (Portrait/Landscape/Square/Story).\n" +
        "• If you choose “Custom”, your current settings remain editable.\n\n" +
        "Notes:\n" +
        "• Presets change multiple controls at once (ratio, long side, max KB, etc.)."
    },
    aspectfit: {
      title: "Aspect + Fit",
      body:
        "What it does:\n" +
        "• Chooses the final canvas aspect ratio (e.g., 4:5, 1:1).\n" +
        "• Controls whether the image is fit into that canvas with padding, or kept proportional with equal borders.\n\n" +
        "How to use:\n" +
        "• Pick an aspect from the dropdown (or choose Custom and enter W:H).\n" +
        "• If “Ignore aspect ratio” is ON:\n" +
        "  - The script keeps original proportions.\n" +
        "  - It adds equal padding around the image (no forced crop/fit to a target ratio).\n\n" +
        "Notes:\n" +
        "• Custom W:H must be positive integers."
    },
    sizing: {
      title: "Sizing",
      body:
        "What it does:\n" +
        "• Sets output dimensions before adding the border/canvas.\n\n" +
        "Controls:\n" +
        "• Long side (px): The long edge of the output canvas (or the long edge of the image when “Ignore aspect ratio” is ON).\n" +
        "• Padding (px): Border thickness on each side (added to the canvas).\n\n" +
        "How to use:\n" +
        "• For Instagram Portrait 4:5, a common value is 1350 long-side.\n" +
        "• Increase padding to make the border thicker."
    },
    bordermodule: {
      title: "Border module",
      body:
        "What it does:\n" +
        "• Adds a border by expanding the canvas and filling the background with a chosen color.\n\n" +
        "Ignore border module:\n" +
        "• When enabled, the script does NOT add canvas/border. It still resizes and exports.\n\n" +
        "Border color modes:\n" +
        "• White: pure white border.\n" +
        "• Black: pure black border.\n" +
        "• Average (per image): border color becomes the image’s average color.\n" +
        "• AUTO (brightness): computes a quick luminance estimate; chooses black for dark images, white for bright.\n" +
        "• AUTO (filename -White-/-Black-/-Average-/-Lum-): if the filename starts with one of these prefixes,\n" +
        "  it forces that mode for that image only.\n" +
        "• Custom (picked color): pick any color via the Photoshop color picker; the hex is stored.\n\n" +
        "How to use:\n" +
        "1) Pick a border color mode.\n" +
        "2) If Custom, click “Pick color…” and it updates the preview.\n" +
        "3) Run / Dry Run to see it applied."
    },
    output: {
      title: "Output",
      body:
        "What it does:\n" +
        "• Controls JPEG output quality, optional file size targeting, and the filename suffix.\n\n" +
        "JPEG quality start (0–12):\n" +
        "• The initial JPEG quality used when saving.\n\n" +
        "File size limits (Min/Max KB):\n" +
        "• If limits are enabled, the script may re-save at different JPEG qualities to get closer to the target range.\n" +
        "• 0 means “no limit”.\n\n" +
        "Ignore file size limits:\n" +
        "• Saves once at the starting quality (fastest).\n\n" +
        "Filename suffix:\n" +
        "• Appended to the base filename before “.jpg” (default: WB)."
    },
    colormanagement: {
      title: "Color management",
      body:
        "What it does:\n" +
        "• Controls optional conversion to sRGB and whether the JPEG embeds a color profile.\n\n" +
        "sRGB conversion:\n" +
        "• OFF: no conversion.\n" +
        "• AUTO: converts to sRGB only if the document doesn’t already look like sRGB.\n" +
        "• FORCE: always converts to sRGB.\n\n" +
        "Embed color profile:\n" +
        "• When enabled, the output JPEG includes an ICC profile.\n" +
        "• Leaving it off can reduce file size slightly but may change how some apps interpret color."
    },
    ppi: {
      title: "PPI metadata",
      body:
        "What it does:\n" +
        "• Sets the image resolution metadata (PPI) without resampling pixels.\n\n" +
        "How to use:\n" +
        "• Keep enabled if you want a consistent PPI value (commonly 72 for web).\n" +
        "• Disable if you don’t care about the PPI field.\n\n" +
        "Note:\n" +
        "• This does not change pixel dimensions—only the resolution metadata."
    },
    batch: {
      title: "Batch stability",
      body:
        "What it does:\n" +
        "• Controls behavior for large batches: skipping existing files, silent errors, chunking, and scratch-disk recovery.\n\n" +
        "Controls:\n" +
        "• Skip if output exists: avoids reprocessing files already exported.\n" +
        "• Silent mode: avoids per-file alert popups (better for unattended runs).\n" +
        "• Chunk size: processes N files, then cleans up memory (helps long runs).\n" +
        "• Cooldown ms: optional pause after each chunk.\n" +
        "• Scratch retries + retry cooldown: if Photoshop reports scratch disk issues, the script will try cleanup + retry.\n\n" +
        "Tip:\n" +
        "• If you have instability on huge folders, reduce chunk size (e.g., 25–50)."
    }
  };

  // ─────────────────────────────────────────────────────────────
  // INFO SYSTEM (FIXED): modal dialog on top, not grey, close works
  // + inline circular "i" icon buttons drawn via onDraw
  // ─────────────────────────────────────────────────────────────
  function showInfoDialog(infoKey) {
    var it = INFO[infoKey];
    if (!it) it = { title: String(infoKey), body: ("No info available for: " + infoKey) };

    var dlg = new Window("dialog", "BorderForge Info — " + it.title);
    dlg.orientation = "column";
    dlg.alignChildren = "fill";
    dlg.margins = [12, 12, 12, 12];
    dlg.spacing = 10;

    var et = dlg.add("edittext", undefined, it.body, { multiline: true, scrolling: true });
    et.preferredSize = [460, 280];
    et.enabled = true;

    // lock the text but keep enabled (prevents greyed-out feel)
    et._lockText = String(it.body || "");
    et.onChanging = function () {
      if (this.text !== this._lockText) this.text = this._lockText;
    };

    var row = dlg.add("group");
    row.orientation = "row";
    row.alignment = "right";

    var btnClose = row.add("button", undefined, "Close", { name: "ok" });
    btnClose.onClick = function () { try { dlg.close(0); } catch (_) {} };

    // open centered over main window (on top)
    try {
      var wl = w.location, ws = w.size;
      if (wl && ws) {
        dlg.location = [
          wl[0] + Math.max(0, Math.round((ws[0] - 520) / 2)),
          wl[1] + Math.max(0, Math.round((ws[1] - 360) / 2))
        ];
      } else {
        dlg.center();
      }
    } catch (_) {
      try { dlg.center(); } catch (__) {}
    }

    try { dlg.active = true; } catch (_) {}
    dlg.show();
  }

function addInfoIcon(parentRowOrGroup, infoKey) {
  // Use a small TEXT button with the Unicode circled-i (U+24D8).
  // This avoids ScriptUI iconbutton onDraw rendering issues (white circle).
  var b = parentRowOrGroup.add("button", undefined, "\u24D8", { name: "bf_info_" + infoKey });

  // Make it compact like an icon
  b.preferredSize = [22, 18];
  b.maximumSize   = [22, 18];
  b.minimumSize   = [22, 18];

  b.helpTip = "Info";
  b.onClick = function () { showInfoDialog(infoKey); };

  // Optional: make it feel a bit more like a toolbutton
  try { b.alignment = ["left", "center"]; } catch (_) {}

  return b;
}

  // ---------- Folders
  var gFolders = tabFolders.add("panel", undefined, "Folders");
  gFolders.orientation = "column";
  gFolders.alignChildren = "fill";

  var cbRememberFolders = gFolders.add("checkbox", undefined, "Remember last folder choices");
  cbRememberFolders.value = true;

  var rowIn = gFolders.add("group");
  rowIn.orientation = "row";
  rowIn.alignChildren = ["left", "center"];
  rowIn.add("statictext", undefined, "Input:");
  addInfoIcon(rowIn, "folders");
  var etIn = rowIn.add("edittext", undefined, "");
  etIn.characters = 45;
  var btnIn = rowIn.add("button", undefined, "Choose…");

  var txtDefaultOut = gFolders.add("statictext", undefined, "Default output: (choose input folder)", {multiline:true});
  txtDefaultOut.maximumSize.width = targetW - 40;

  // Output folder manager (button)
  var gOutMgr = gFolders.add("panel", undefined, "Output folder manager");
  gOutMgr.orientation = "column";
  gOutMgr.alignChildren = "left";
  gOutMgr.margins = [10,10,10,10];

  var rowPurge = gOutMgr.add("group");
  rowPurge.orientation = "row";
  rowPurge.alignChildren = ["left","center"];

  var btnPurge = rowPurge.add("button", undefined, "Delete old “With Borders…” folders");
  addInfoIcon(rowPurge, "outmgr");
  btnPurge.enabled = false;

  var rowOutMode = gFolders.add("group");
  rowOutMode.orientation = "row";
  rowOutMode.alignChildren = ["left", "center"];
  var cbCustomOut = rowOutMode.add("checkbox", undefined, "Use custom output folder");
  var btnOut = rowOutMode.add("button", undefined, "Choose…");
  btnOut.enabled = false;

  var txtCustomOut = gFolders.add("statictext", undefined, "Custom output: (not set)", {multiline:true});
  txtCustomOut.maximumSize.width = targetW - 40;
  txtCustomOut.enabled = false;

  var inFolder = null;
  var defaultOutFolder = null;
  var customOutFolder = null;

  // deterministic purge bookkeeping
  var defaultOutNextIndex = 1;

  function refreshDefaultOutLabel() {
    if (!inFolder) {
      txtDefaultOut.text = "Default output: (choose input folder)";
      defaultOutFolder = null;
      defaultOutNextIndex = 1;
      btnPurge.enabled = false;
      return;
    }

    var info = getNextNumberedSubfolderAndIndex(inFolder, "With Borders");
    defaultOutFolder = info.folder;
    defaultOutNextIndex = info.nextIndex;

    txtDefaultOut.text = "Default output: " + defaultOutFolder.fsName;

    // Enable purge only if there's something to purge (i.e., nextIndex >= 2)
    btnPurge.enabled = (defaultOutNextIndex >= 2);
  }

  btnIn.onClick = function () {
    var f = Folder.selectDialog("Choose the INPUT folder");
    if (f) {
      inFolder = f;
      etIn.text = f.fsName;
      refreshDefaultOutLabel();
    }
  };

  // purge button click behavior (deterministic)
  btnPurge.onClick = function () {
    if (!inFolder || !inFolder.exists) {
      alert("Choose a valid input folder first.");
      return;
    }

    // Recompute right now (in case folders changed since selection)
    var info = getNextNumberedSubfolderAndIndex(inFolder, "With Borders");
    defaultOutNextIndex = info.nextIndex;

    var maxN = defaultOutNextIndex - 1;
    if (maxN < 1) {
      alert("No old “With Borders…” folders were detected.");
      refreshDefaultOutLabel();
      return;
    }

    var msg =
      "This will delete the following folders inside:\n\n" +
      inFolder.fsName + "\n\n" +
      "• With Borders\n" +
      (maxN >= 2 ? ("• With Borders 2 … With Borders " + maxN + "\n") : "") +
      "\nContinue?";

    var ok = confirm(msg);
    if (!ok) return;

    clearRunStateSafe();

    var deletedCount = purgeUsingNextIndex(inFolder, "With Borders", defaultOutNextIndex);

    refreshDefaultOutLabel();

    alert("Deleted folders: " + deletedCount + "\n\nDefault output has been refreshed.");
  };

  cbCustomOut.onClick = function () {
    btnOut.enabled = cbCustomOut.value;
    txtCustomOut.enabled = cbCustomOut.value;
  };

  btnOut.onClick = function () {
    var f = Folder.selectDialog("Choose CUSTOM output folder");
    if (f) {
      customOutFolder = f;
      txtCustomOut.text = "Custom output: " + f.fsName;
    }
  };

  // ---------- Presets
  var gPreset = tabFolders.add("panel", undefined, "Preset");
  gPreset.orientation = "column";
  gPreset.alignChildren = "left";

  var rowPreset = gPreset.add("group");
  rowPreset.orientation = "row";
  rowPreset.alignChildren = ["left","center"];
  rowPreset.add("statictext", undefined, "Preset:");
  var ddPreset = rowPreset.add("dropdownlist", undefined, [
    "Custom",
    "Instagram Feed (Portrait 4:5, 1080×1350)",
    "Instagram Feed (Landscape 5:4, 1350×1080)",
    "Instagram Square (1:1, 1080×1080)",
    "Instagram Story/Reel (9:16, 1080×1920)"
  ]);
  addInfoIcon(rowPreset, "preset");
  ddPreset.selection = 1;

  // ---------- Aspect + Fit (Tab 2)
  var gCanvas = tabSizing.add("panel", undefined, "Aspect + Fit");
  gCanvas.orientation = "column";
  gCanvas.alignChildren = "left";

  var rowRatio = gCanvas.add("group");
  rowRatio.orientation = "row";
  rowRatio.alignChildren = ["left","center"];
  rowRatio.add("statictext", undefined, "Aspect:");
  var ddRatio = rowRatio.add("dropdownlist", undefined, [
    "4:5 (Instagram Portrait)",
    "5:4 (Instagram Landscape)",
    "1:1 (Square)",
    "9:16 (Story/Reel)",
    "16:9 (Widescreen)",
    "Custom"
  ]);
  addInfoIcon(rowRatio, "aspectfit");
  ddRatio.selection = 0;

  var cbIgnoreRatio = gCanvas.add("checkbox", undefined,
    "Ignore aspect ratio (equal border mode; keep original proportions)");
  cbIgnoreRatio.value = false;

  var rowCustom = gCanvas.add("group");
  rowCustom.orientation = "row";
  rowCustom.alignChildren = ["left","center"];
  rowCustom.add("statictext", undefined, "Custom W:H");
  var etW = rowCustom.add("edittext", undefined, "4"); etW.characters = 5;
  rowCustom.add("statictext", undefined, ":");
  var etH = rowCustom.add("edittext", undefined, "5"); etH.characters = 5;
  rowCustom.enabled = false;

  ddRatio.onChange = function () {
    rowCustom.enabled = (ddRatio.selection && ddRatio.selection.text === "Custom") && !cbIgnoreRatio.value;
  };

  // ---------- Sizing (Tab 2)
  var gSizing = tabSizing.add("panel", undefined, "Sizing");
  gSizing.orientation = "column";
  gSizing.alignChildren = "left";

  var rowLong = gSizing.add("group");
  rowLong.orientation = "row";
  rowLong.alignChildren = ["left","center"];
  rowLong.add("statictext", undefined, "Long side (px):");
  var etLong = rowLong.add("edittext", undefined, "1350"); etLong.characters = 7;
  addInfoIcon(rowLong, "sizing");

  var rowPad = gSizing.add("group");
  rowPad.orientation = "row";
  rowPad.alignChildren = ["left","center"];
  rowPad.add("statictext", undefined, "Padding (px):");
  var etPad = rowPad.add("edittext", undefined, "40"); etPad.characters = 7;

  function syncRatioControlsEnabled() {
    var enabled = !cbIgnoreRatio.value;
    ddRatio.enabled = enabled;
    rowCustom.enabled = enabled && (ddRatio.selection && ddRatio.selection.text === "Custom");
  }
  cbIgnoreRatio.onClick = syncRatioControlsEnabled;

  // ---------- Border module (Tab 3)
  var gBorder = tabBorderOut.add("panel", undefined, "Border module");
  gBorder.orientation = "column";
  gBorder.alignChildren = "left";

  var cbIgnoreBorder = gBorder.add("checkbox", undefined,
    "Ignore border module (no border, no added canvas)");
  cbIgnoreBorder.value = false;

  var rowBorderMode = gBorder.add("group");
  rowBorderMode.orientation = "row";
  rowBorderMode.alignChildren = ["left", "center"];
  rowBorderMode.add("statictext", undefined, "Border color:");
  var ddBorderMode = rowBorderMode.add("dropdownlist", undefined, [
    "White",
    "Black",
    "Average (per image)",
    "AUTO (brightness)",
    "AUTO (filename -White-/-Black-/-Average-/-Lum-)",
    "Custom (picked color)"
  ]);
  addInfoIcon(rowBorderMode, "bordermodule");
  ddBorderMode.selection = 0;

  var swRow = gBorder.add("group");
  swRow.orientation = "row";
  swRow.alignChildren = ["left", "center"];
  swRow.add("statictext", undefined, "Preview:");
  var swatch = swRow.add("panel", undefined, "");
  swatch.preferredSize = [56, 22];

  var gPick = gBorder.add("panel", undefined, "Custom color picker");
  gPick.orientation = "column";
  gPick.alignChildren = "left";

  var rowPick = gPick.add("group");
  rowPick.orientation = "row";
  rowPick.alignChildren = ["left", "center"];

  var btnPick = rowPick.add("button", undefined, "Pick color…");
  rowPick.add("statictext", undefined, "HEX:");
  var etHex = rowPick.add("edittext", undefined, "FFFFFF");
  etHex.characters = 10;
  etHex.enabled = false;

  var customHexUI = "FFFFFF";

  function hexToRgbObj(hex6) {
    var r = parseInt(hex6.substring(0,2), 16);
    var g = parseInt(hex6.substring(2,4), 16);
    var b = parseInt(hex6.substring(4,6), 16);
    if (!isFinite(r)) r = 255;
    if (!isFinite(g)) g = 255;
    if (!isFinite(b)) b = 255;
    return {r:r,g:g,b:b};
  }

  function normalizeHex6(s) {
    s = String(s || "").replace(/[^0-9a-fA-F]/g, "");
    if (s.length === 3) {
      s = s.charAt(0)+s.charAt(0)+s.charAt(1)+s.charAt(1)+s.charAt(2)+s.charAt(2);
    }
    if (s.length < 6) {
      while (s.length < 6) s += "F";
    }
    return s.substring(0,6).toUpperCase();
  }

  function setSwatchFromHex(hex6) {
    customHexUI = normalizeHex6(hex6);
    etHex.text = customHexUI;
    updateBorderPreviewVisibilityAndColor();
  }

  btnPick.onClick = function () {
    try {
      var ok = app.showColorPicker();
      if (ok) {
        var c = app.foregroundColor;
        var r = clampInt(Math.round(c.rgb.red), 0, 255);
        var g = clampInt(Math.round(c.rgb.green), 0, 255);
        var b = clampInt(Math.round(c.rgb.blue), 0, 255);
        var hex = rgbToHex(r, g, b);
        setSwatchFromHex(hex);
        ddBorderMode.selection = ddBorderMode.items[5];
      }
    } catch (_) {
      alert("Could not open the color picker on this system.");
    }
    syncBorderUIEnabled();
  };

  function borderModeToKey(text) {
    if (text === "White") return "WHITE";
    if (text === "Black") return "BLACK";
    if (text.indexOf("Average") === 0) return "AVERAGE";
    if (text.indexOf("AUTO (brightness)") === 0) return "AUTO";
    if (text.indexOf("AUTO (filename") === 0) return "AUTO_FILENAME";
    return "CUSTOM";
  }

  function drawSwatchFill(rgbObj) {
    var fillRgb = rgbObj || {r:240,g:240,b:240};
    swatch.onDraw = function () {
      try {
        var g = swatch.graphics;
        var w0 = swatch.size[0], h0 = swatch.size[1];

        function mkBrush(rgb) {
          return g.newBrush(g.BrushType.SOLID_COLOR, [rgb.r/255, rgb.g/255, rgb.b/255, 1]);
        }
        function mkPen(rgb, px) {
          return g.newPen(g.PenType.SOLID_COLOR, [rgb.r/255, rgb.g/255, rgb.b/255, 1], px || 1);
        }

        g.rectPath(0, 0, w0, h0);
        g.fillPath(mkBrush(fillRgb));

        g.rectPath(0, 0, w0, h0);
        g.strokePath(mkPen({r:60,g:60,b:60}, 1));
      } catch (_) {}
    };
    try { swatch.notify("onDraw"); } catch (_) {}
    try { swatch.repaint(); } catch (_) {}
  }

  function updateBorderPreviewVisibilityAndColor() {
    var ignore = cbIgnoreBorder.value;
    var key = borderModeToKey(ddBorderMode.selection ? ddBorderMode.selection.text : "White");

    var isAutoishMode = (key === "AUTO" || key === "AUTO_FILENAME" || key === "AVERAGE");
    swRow.visible = !isAutoishMode;

    if (!swRow.visible) {
      try { w.layout.layout(true); } catch (_) {}
      return;
    }

    if (ignore) {
      drawSwatchFill({r:235,g:235,b:235});
    } else if (key === "WHITE") {
      drawSwatchFill({r:255,g:255,b:255});
    } else if (key === "BLACK") {
      drawSwatchFill({r:0,g:0,b:0});
    } else if (key === "CUSTOM") {
      drawSwatchFill(hexToRgbObj(customHexUI));
    } else {
      drawSwatchFill({r:200,g:200,b:200});
    }

    try { w.layout.layout(true); } catch (_) {}
  }

  function syncBorderUIEnabled() {
    var ignore = cbIgnoreBorder.value;
    ddBorderMode.enabled = !ignore;

    var isCustom = (!ignore) && (ddBorderMode.selection && ddBorderMode.selection.text.indexOf("Custom") === 0);
    gPick.enabled = isCustom;
    btnPick.enabled = isCustom;
    etHex.enabled = false;

    updateBorderPreviewVisibilityAndColor();
  }

  cbIgnoreBorder.onClick = syncBorderUIEnabled;
  ddBorderMode.onChange = syncBorderUIEnabled;

  // ---------- Output (Tab 3)
  var gOutput = tabBorderOut.add("panel", undefined, "Output");
  gOutput.orientation = "column";
  gOutput.alignChildren = "left";

  var rowQ = gOutput.add("group");
  rowQ.orientation = "row";
  rowQ.alignChildren = ["left","center"];
  rowQ.add("statictext", undefined, "JPEG quality start (0–12):");
  var etQ = rowQ.add("edittext", undefined, "10"); etQ.characters = 4;
  addInfoIcon(rowQ, "output");

  var cbIgnoreFileSize = gOutput.add("checkbox", undefined,
    "Ignore file size limits (save once; no quality stepping)");
  cbIgnoreFileSize.value = false;

  var rowMin = gOutput.add("group");
  rowMin.orientation = "row";
  rowMin.alignChildren = ["left","center"];
  rowMin.add("statictext", undefined, "Min file size (KB, 0 = none):");
  var etMinKB = rowMin.add("edittext", undefined, "0"); etMinKB.characters = 8;

  var rowMax = gOutput.add("group");
  rowMax.orientation = "row";
  rowMax.alignChildren = ["left","center"];
  rowMax.add("statictext", undefined, "Max file size (KB, 0 = none):");
  var etMaxKB = rowMax.add("edittext", undefined, "900"); etMaxKB.characters = 8;

  function syncFileSizeControlsEnabled() {
    var enabled = !cbIgnoreFileSize.value;
    rowMin.enabled = enabled;
    rowMax.enabled = enabled;
  }
  cbIgnoreFileSize.onClick = syncFileSizeControlsEnabled;

  var rowSuffix = gOutput.add("group");
  rowSuffix.orientation = "row";
  rowSuffix.alignChildren = ["left","center"];
  rowSuffix.add("statictext", undefined, "Filename suffix:");
  var etSuffix = rowSuffix.add("edittext", undefined, "WB"); etSuffix.characters = 10;

  // ---------- Advanced tab
  var gColor = tabAdv.add("panel", undefined, "Color management");
  gColor.orientation = "column";
  gColor.alignChildren = "left";

  var rowSRGB = gColor.add("group");
  rowSRGB.orientation = "row";
  rowSRGB.alignChildren = ["left", "center"];
  rowSRGB.add("statictext", undefined, "sRGB conversion:");
  var ddSRGB = rowSRGB.add("dropdownlist", undefined, ["OFF", "AUTO", "FORCE"]);
  addInfoIcon(rowSRGB, "colormanagement");
  ddSRGB.selection = 0;

  var cbEmbed = gColor.add("checkbox", undefined, "Embed color profile in JPEG");
  cbEmbed.value = false;

  var gPPI = tabAdv.add("panel", undefined, "PPI metadata");
  gPPI.orientation = "column";
  gPPI.alignChildren = "left";

  var rowPPI = gPPI.add("group");
  rowPPI.orientation = "row";
  rowPPI.alignChildren = ["left", "center"];
  var cbSetPPI = rowPPI.add("checkbox", undefined, "Set PPI (no resample) to");
  cbSetPPI.value = true;
  var etPPI = rowPPI.add("edittext", undefined, "72"); etPPI.characters = 6;
  addInfoIcon(rowPPI, "ppi");
  etPPI.enabled = cbSetPPI.value;
  cbSetPPI.onClick = function () { etPPI.enabled = cbSetPPI.value; };

  var gBatch = tabAdv.add("panel", undefined, "Batch stability");
  gBatch.orientation = "column";
  gBatch.alignChildren = "left";

  var cbSkip = gBatch.add("checkbox", undefined, "Skip if output exists");
  cbSkip.value = true;

  var cbSilent = gBatch.add("checkbox", undefined, "Silent mode (no per-file alerts on skips)");
  cbSilent.value = true;

  var rowChunk = gBatch.add("group");
  rowChunk.orientation = "row";
  rowChunk.alignChildren = ["left", "center"];
  rowChunk.add("statictext", undefined, "Chunk size:");
  var etChunk = rowChunk.add("edittext", undefined, "100"); etChunk.characters = 6;
  addInfoIcon(rowChunk, "batch");

  var rowCool = gBatch.add("group");
  rowCool.orientation = "row";
  rowCool.alignChildren = ["left", "center"];
  rowCool.add("statictext", undefined, "Cooldown ms (0 = none):");
  var etCool = rowCool.add("edittext", undefined, "0"); etCool.characters = 8;

  var rowScr = gBatch.add("group");
  rowScr.orientation = "row";
  rowScr.alignChildren = ["left", "center"];
  rowScr.add("statictext", undefined, "Scratch retries:");
  var etScr = rowScr.add("edittext", undefined, "2"); etScr.characters = 4;
  rowScr.add("statictext", undefined, "retry cooldown ms:");
  var etScrCool = rowScr.add("edittext", undefined, "5000"); etScrCool.characters = 8;

  var rowReset = gBatch.add("group");
  rowReset.orientation = "row";
  rowReset.alignChildren = ["left", "center"];
  var btnReset = rowReset.add("button", undefined, "Reset settings to defaults");

  btnReset.onClick = function () {
    applySettingsToUI(DEFAULTS, true);
    try { if (SETTINGS_FILE.exists) SETTINGS_FILE.remove(); } catch (_) {}
    refreshDefaultOutLabel();
  };

  // ---------- Preset logic
  function applyPreset(name) {
    if (name.indexOf("Instagram Feed (Portrait") === 0) {
      ddRatio.selection = 0;
      etW.text = "4"; etH.text = "5";
      etLong.text = "1350";
      etMaxKB.text = "900";
      etMinKB.text = "0";
      etQ.text = "10";
      cbIgnoreRatio.value = false;

    } else if (name.indexOf("Instagram Feed (Landscape") === 0) {
      ddRatio.selection = 1;
      etW.text = "5"; etH.text = "4";
      etLong.text = "1080";
      etMaxKB.text = "900";
      etMinKB.text = "0";
      etQ.text = "10";
      cbIgnoreRatio.value = false;

    } else if (name.indexOf("Instagram Square") === 0) {
      ddRatio.selection = 2;
      etW.text = "1"; etH.text = "1";
      etLong.text = "1080";
      etMaxKB.text = "900";
      etMinKB.text = "0";
      etQ.text = "10";
      cbIgnoreRatio.value = false;

    } else if (name.indexOf("Instagram Story/Reel") === 0) {
      ddRatio.selection = 3;
      etW.text = "9"; etH.text = "16";
      etLong.text = "1920";
      etMaxKB.text = "1200";
      etMinKB.text = "0";
      etQ.text = "10";
      cbIgnoreRatio.value = false;
    }

    syncRatioControlsEnabled();
    syncFileSizeControlsEnabled();
    syncBorderUIEnabled();
  }

  ddPreset.onChange = function () {
    var t = ddPreset.selection ? ddPreset.selection.text : "Custom";
    if (t !== "Custom") applyPreset(t);
  };

  applyPreset(ddPreset.selection.text);

  // Footer
  var footer = w.add("group");
  footer.orientation = "row";
  footer.alignChildren = ["left", "center"];
  footer.add("statictext", undefined, "Created By Seventeen");

  // Buttons (no Cancel — per request)
  var gBtn = w.add("group");
  gBtn.alignment = "right";

  var btnClose = gBtn.add("button", undefined, "Close", {name: "cancel"});
  var btnDry = gBtn.add("button", undefined, "Dry Run");
  var btnRun = gBtn.add("button", undefined, "Run", {name: "ok"});

  var chosenAction = null;

  btnClose.onClick = function () { chosenAction = "close"; w.close(0); };
  btnDry.onClick   = function () { chosenAction = "dryrun"; w.close(1); };
  btnRun.onClick   = function () { chosenAction = "run"; w.close(1); };

  if (persisted && persisted.hasAny) {
    applySettingsToUI(persisted, false);
  }

  syncRatioControlsEnabled();
  syncFileSizeControlsEnabled();
  syncBorderUIEnabled();
  refreshDefaultOutLabel();

  w.onResizing = w.onResize = function () { try { this.layout.resize(); } catch (_) {} };

  try { w.layout.layout(true); } catch (_) {}
  try { w.center(); } catch (_) {}

  w.show();
  if (chosenAction === "close" || !chosenAction) return null;

  if (etIn.text && (!inFolder || inFolder.fsName !== etIn.text)) {
    var f2 = Folder(etIn.text);
    if (f2.exists) inFolder = f2;
  }
  if (!inFolder || !inFolder.exists) {
    alert("Please choose a valid input folder.");
    return { restartUI: true };
  }

  defaultOutFolder = getNextNumberedSubfolder(inFolder, "With Borders");

  var useCustomOut = cbCustomOut.value;
  var outFolder = useCustomOut ? customOutFolder : defaultOutFolder;

  // For Dry Run we still compute an outFolder (so Dry Run can optionally run batch)
  // If custom output is selected but missing, we allow Dry Run (but block Run).
  if (chosenAction === "run") {
    if (useCustomOut && !outFolder) {
      alert("You selected a custom output folder but did not choose one.");
      return { restartUI: true };
    }
  }

  var borderMode = borderModeToKey(ddBorderMode.selection ? ddBorderMode.selection.text : "White");
  var borderHex = customHexUI;

  var ratioW = 4, ratioH = 5;
  var preset2 = ddRatio.selection ? ddRatio.selection.text : "4:5 (Instagram Portrait)";
  if (preset2.indexOf("4:5") === 0) { ratioW = 4; ratioH = 5; }
  else if (preset2.indexOf("5:4") === 0) { ratioW = 5; ratioH = 4; }
  else if (preset2.indexOf("1:1") === 0) { ratioW = 1; ratioH = 1; }
  else if (preset2.indexOf("9:16") === 0) { ratioW = 9; ratioH = 16; }
  else if (preset2.indexOf("16:9") === 0) { ratioW = 16; ratioH = 9; }
  else {
    ratioW = parseInt(etW.text, 10);
    ratioH = parseInt(etH.text, 10);
    if (!isFinite(ratioW) || !isFinite(ratioH) || ratioW <= 0 || ratioH <= 0) {
      alert("Custom aspect ratio must be positive integers.");
      return { restartUI: true };
    }
  }

  var longSide = parseInt(etLong.text, 10);
  var padding = parseInt(etPad.text, 10);
  var jpegQuality = parseInt(etQ.text, 10);
  var minKB = parseInt(etMinKB.text, 10);
  var maxKB = parseInt(etMaxKB.text, 10);

  var setPPI = cbSetPPI.value;
  var ppi = parseInt(etPPI.text, 10);

  var chunkSize = parseInt(etChunk.text, 10);
  var cooldownMs = parseInt(etCool.text, 10);
  var scratchMaxRetries = parseInt(etScr.text, 10);
  var scratchRetryCooldownMs = parseInt(etScrCool.text, 10);

  if (!isFinite(longSide) || longSide < 200) { alert("Long side must be >= 200."); return { restartUI: true }; }
  if (!isFinite(padding) || padding < 0) { alert("Padding must be >= 0."); return { restartUI: true }; }
  if (!isFinite(jpegQuality) || jpegQuality < 0 || jpegQuality > 12) { alert("JPEG quality must be 0–12."); return { restartUI: true }; }
  if (!isFinite(minKB) || minKB < 0) { alert("Min file size KB must be >= 0."); return { restartUI: true }; }
  if (!isFinite(maxKB) || maxKB < 0) { alert("Max file size KB must be >= 0."); return { restartUI: true }; }
  if (!cbIgnoreFileSize.value && minKB > 0 && maxKB > 0 && minKB > maxKB) {
    alert("Min file size KB must be <= Max file size KB.");
    return { restartUI: true };
  }

  if (setPPI && (!isFinite(ppi) || ppi < 1)) { alert("PPI must be >= 1."); return { restartUI: true }; }
  if (!isFinite(chunkSize) || chunkSize < 1) { alert("Chunk size must be >= 1."); return { restartUI: true }; }
  if (!isFinite(cooldownMs) || cooldownMs < 0) { alert("Cooldown must be >= 0."); return { restartUI: true }; }
  if (!isFinite(scratchMaxRetries) || scratchMaxRetries < 0) { alert("Scratch retries must be >= 0."); return { restartUI: true }; }
  if (!isFinite(scratchRetryCooldownMs) || scratchRetryCooldownMs < 0) { alert("Scratch retry cooldown must be >= 0."); return { restartUI: true }; }

  var suffix = etSuffix.text;
  if (!suffix || !suffix.length) suffix = "WB";

  var opts = {
    ratioW: ratioW,
    ratioH: ratioH,
    longSide: longSide,
    padding: padding,

    ignoreRatio: cbIgnoreRatio.value,

    ignoreBorder: cbIgnoreBorder.value,
    ignoreFileSizeLimits: cbIgnoreFileSize.value,

    borderMode: borderMode,
    borderHex: borderHex,

    jpegQuality: jpegQuality,
    minKB: minKB,
    maxKB: maxKB,

    suffix: suffix,

    srgbMode: ddSRGB.selection.text,
    embedProfile: cbEmbed.value,

    setPPI: setPPI,
    ppi: setPPI ? ppi : null,

    skipExisting: cbSkip.value,
    silentMode: cbSilent.value,

    chunkSize: chunkSize,
    cooldownMs: cooldownMs,
    scratchMaxRetries: scratchMaxRetries,
    scratchRetryCooldownMs: scratchRetryCooldownMs,

    rememberFolders: cbRememberFolders.value
  };

  saveSettingsSafe(SETTINGS_FILE, buildSettingsForSave(opts, {
    rememberFolders: cbRememberFolders.value,
    inputPath: inFolder ? inFolder.fsName : "",
    useCustomOut: cbCustomOut.value,
    customOutPath: customOutFolder ? customOutFolder.fsName : ""
  }));

  return {
    action: chosenAction,
    inFolder: inFolder,
    outFolder: outFolder,
    opts: opts
  };

  function applySettingsToUI(s, isResetToDefaults) {
    if (isResetToDefaults) ddPreset.selection = 1;

    if (typeof s.rememberFolders !== "undefined") cbRememberFolders.value = toBool(s.rememberFolders, DEFAULTS.rememberFolders);
    if (typeof s.ignoreRatio !== "undefined") cbIgnoreRatio.value = toBool(s.ignoreRatio, DEFAULTS.ignoreRatio);

    if (s.ratioPresetText) {
      var t = String(s.ratioPresetText);
      if (t.indexOf("4:5") === 0) setDropdownByText(ddRatio, "4:5 (Instagram Portrait)", ddRatio.selection);
      else if (t.indexOf("5:4") === 0) setDropdownByText(ddRatio, "5:4 (Instagram Landscape)", ddRatio.selection);
      else if (t.indexOf("1:1") === 0) setDropdownByText(ddRatio, "1:1 (Square)", ddRatio.selection);
      else if (t.indexOf("9:16") === 0) setDropdownByText(ddRatio, "9:16 (Story/Reel)", ddRatio.selection);
      else if (t.indexOf("16:9") === 0) setDropdownByText(ddRatio, "16:9 (Widescreen)", ddRatio.selection);
    }
    if (typeof s.ratioW !== "undefined") etW.text = String(s.ratioW);
    if (typeof s.ratioH !== "undefined") etH.text = String(s.ratioH);

    if (typeof s.longSide !== "undefined") etLong.text = String(s.longSide);
    if (typeof s.padding !== "undefined") etPad.text = String(s.padding);

    if (typeof s.ignoreBorder !== "undefined") cbIgnoreBorder.value = toBool(s.ignoreBorder, DEFAULTS.ignoreBorder);

    var bm = String(s.borderMode || DEFAULTS.borderMode);
    if (bm === "WHITE") setDropdownByText(ddBorderMode, "White", ddBorderMode.selection);
    else if (bm === "BLACK") setDropdownByText(ddBorderMode, "Black", ddBorderMode.selection);
    else if (bm === "AVERAGE") setDropdownByText(ddBorderMode, "Average (per image)", ddBorderMode.selection);
    else if (bm === "AUTO") setDropdownByText(ddBorderMode, "AUTO (brightness)", ddBorderMode.selection);
    else if (bm === "AUTO_FILENAME") setDropdownByText(ddBorderMode, "AUTO (filename -White-/-Black-/-Average-/-Lum-)", ddBorderMode.selection);
    else setDropdownByText(ddBorderMode, "Custom (picked color)", ddBorderMode.selection);

    if (typeof s.borderHex !== "undefined") setSwatchFromHex(String(s.borderHex));

    if (typeof s.jpegQuality !== "undefined") etQ.text = String(s.jpegQuality);
    if (typeof s.ignoreFileSizeLimits !== "undefined") cbIgnoreFileSize.value = toBool(s.ignoreFileSizeLimits, DEFAULTS.ignoreFileSizeLimits);
    if (typeof s.minKB !== "undefined") etMinKB.text = String(s.minKB);
    if (typeof s.maxKB !== "undefined") etMaxKB.text = String(s.maxKB);
    if (typeof s.suffix !== "undefined") etSuffix.text = String(s.suffix);

    if (s.srgbMode) setDropdownByText(ddSRGB, s.srgbMode, ddSRGB.selection);
    if (typeof s.embedProfile !== "undefined") cbEmbed.value = toBool(s.embedProfile, DEFAULTS.embedProfile);

    if (typeof s.setPPI !== "undefined") cbSetPPI.value = toBool(s.setPPI, DEFAULTS.setPPI);
    if (typeof s.ppi !== "undefined") etPPI.text = String(s.ppi);
    etPPI.enabled = cbSetPPI.value;

    if (typeof s.skipExisting !== "undefined") cbSkip.value = toBool(s.skipExisting, DEFAULTS.skipExisting);
    if (typeof s.silentMode !== "undefined") cbSilent.value = toBool(s.silentMode, DEFAULTS.silentMode);

    if (typeof s.chunkSize !== "undefined") etChunk.text = String(s.chunkSize);
    if (typeof s.cooldownMs !== "undefined") etCool.text = String(s.cooldownMs);
    if (typeof s.scratchMaxRetries !== "undefined") etScr.text = String(s.scratchMaxRetries);
    if (typeof s.scratchRetryCooldownMs !== "undefined") etScrCool.text = String(s.scratchRetryCooldownMs);

    if (toBool(s.rememberFolders, false)) {
      if (s.inputPath) {
        var fIn = Folder(s.inputPath);
        if (fIn.exists) {
          inFolder = fIn;
          etIn.text = fIn.fsName;
          refreshDefaultOutLabel();
        }
      }
      if (typeof s.useCustomOut !== "undefined") cbCustomOut.value = toBool(s.useCustomOut, false);
      btnOut.enabled = cbCustomOut.value;
      txtCustomOut.enabled = cbCustomOut.value;

      if (cbCustomOut.value && s.customOutPath) {
        var fOut = Folder(s.customOutPath);
        if (fOut.exists) {
          customOutFolder = fOut;
          txtCustomOut.text = "Custom output: " + fOut.fsName;
        }
      }
    }

    syncRatioControlsEnabled();
    syncFileSizeControlsEnabled();
    syncBorderUIEnabled();
    refreshDefaultOutLabel();
  }

  function setDropdownByText(dd, text, fallbackSelection) {
    if (!dd || !dd.items || !dd.items.length) return;
    for (var i = 0; i < dd.items.length; i++) {
      if (dd.items[i].text === text) { dd.selection = dd.items[i]; return; }
    }
    dd.selection = fallbackSelection;
  }

  function toBool(v, fallback) {
    if (typeof v === "boolean") return v;
    var s = String(v).toLowerCase();
    if (s === "true" || s === "1" || s === "yes") return true;
    if (s === "false" || s === "0" || s === "no") return false;
    return !!fallback;
  }
}

/* ───────────────────────────────────────────────────────────────────
 * Dry Run (hidden temp export + auto open + delete)
 * Returns: { runBatchNow: boolean }
 * ─────────────────────────────────────────────────────────────────── */

function getHiddenTestFolder() {
  var f = Folder(Folder.temp + "/.borderforge_test");
  if (!f.exists) { try { f.create(); } catch (_) {} }
  return f;
}

function formatBytes(bytes) {
  if (!bytes || bytes <= 0) return "0 B";
  var kb = bytes / 1024.0;
  if (kb < 1024) return (Math.round(kb * 10) / 10) + " KB";
  var mb = kb / 1024.0;
  return (Math.round(mb * 10) / 10) + " MB";
}

function pickRandomIndex(total) {
  if (total <= 1) return 0;
  return Math.floor(Math.random() * total);
}

function startDryRunSession(files, inFolder, outFolder, opts) {
  var BORDER_WHITE = hexToSolidColor("FFFFFF");
  var BORDER_BLACK = hexToSolidColor("000000");
  var BORDER_CUSTOM = hexToSolidColor(opts.borderHex);

  var runBatchNow = false;

  var dlg = new Window("dialog", "BorderForge Dry Run (v8.3)");
  dlg.orientation = "column";
  dlg.alignChildren = "fill";
  dlg.margins = [12,12,12,12];
  dlg.spacing = 10;

  var txtTitle = dlg.add("statictext", undefined, "Random test image using current parameters:");
  txtTitle.characters = 62;

  var gInfo = dlg.add("panel", undefined, "Details");
  gInfo.orientation = "column";
  gInfo.alignChildren = "fill";
  gInfo.margins = [10,10,10,10];

  var stInName = gInfo.add("statictext", undefined, "Original file: (none)"); stInName.characters = 72;
  var stInSize = gInfo.add("statictext", undefined, "Original size: (none)"); stInSize.characters = 72;
  var stOutName = gInfo.add("statictext", undefined, "Exported file name: (none)"); stOutName.characters = 72;
  var stOutSize = gInfo.add("statictext", undefined, "Exported size: (none)"); stOutSize.characters = 72;
  var stMode = gInfo.add("statictext", undefined, "Border mode: (none)"); stMode.characters = 72;
  var stPath = gInfo.add("statictext", undefined, "Temp export: (hidden)"); stPath.characters = 72;

  var stRealOut = gInfo.add("statictext", undefined, "Batch output: (not set)"); stRealOut.characters = 72;
  try {
    stRealOut.text = "Batch output: " + (outFolder ? outFolder.fsName : "(not set)");
  } catch (_) {
    stRealOut.text = "Batch output: (not set)";
  }

  var gBtns = dlg.add("group");
  gBtns.orientation = "row";
  gBtns.alignChildren = ["left","center"];

  var btnAgain = gBtns.add("button", undefined, "Run Again");
  var btnRunBatch = gBtns.add("button", undefined, "Run Batch Now");
  var btnBack = gBtns.add("button", undefined, "Return to UI");

  if (!outFolder) {
    btnRunBatch.enabled = false;
    btnRunBatch.helpTip = "Output folder is not set. Return to UI and choose output folder.";
  }

  var footer = dlg.add("statictext", undefined, "Created By Seventeen");
  footer.alignment = "right";

  var lastTestDoc = null;

  function closeLastTestDoc() {
    try {
      if (lastTestDoc && lastTestDoc instanceof Document) {
        lastTestDoc.close(SaveOptions.DONOTSAVECHANGES);
      }
    } catch (_) {}
    lastTestDoc = null;
  }

  function doOneDryRun() {
    btnAgain.enabled = false;
    try { dlg.update(); } catch (_) {}

    closeLastTestDoc();

    var idx = pickRandomIndex(files.length);
    var fileObj = files[idx];

    var inBytes = 0;
    try { inBytes = fileObj.length; } catch (_) { inBytes = 0; }

    var testFolder = getHiddenTestFolder();
    var outFile = new File(testFolder.fsName + "/bf_test_" + (new Date().getTime()) + "_" + idx + ".jpg");
    var displayOutName = outFile.name;

    try {
      processOne(fileObj, outFile, opts, BORDER_WHITE, BORDER_BLACK, BORDER_CUSTOM);

      var outBytes = 0;
      try { outBytes = outFile.length; } catch (_) { outBytes = 0; }

      try {
        lastTestDoc = app.open(outFile);
        try { app.refresh(); } catch (_) {}
      } catch (_) {
        lastTestDoc = null;
      }

      try { outFile.remove(); } catch (_) {}

      stInName.text = "Original file: " + fileObj.name;
      stInSize.text = "Original size: " + formatBytes(inBytes);
      stOutName.text = "Exported file name: " + displayOutName + " (temp)";
      stOutSize.text = "Exported size: " + formatBytes(outBytes);
      stMode.text = "Border mode: " + String(opts.borderMode || "");
      stPath.text = "Temp export: " + testFolder.fsName + " (hidden)";
      try { stRealOut.text = "Batch output: " + (outFolder ? outFolder.fsName : "(not set)"); } catch (_) {}

    } catch (e) {
      try { if (outFile && outFile.exists) outFile.remove(); } catch (_) {}

      stInName.text = "Original file: " + fileObj.name;
      stInSize.text = "Original size: " + formatBytes(inBytes);
      stOutName.text = "Exported file name: (failed)";
      stOutSize.text = "Exported size: (failed)";
      stMode.text = "Border mode: " + String(opts.borderMode || "");
      stPath.text = "Temp export: " + getHiddenTestFolder().fsName + " (hidden)";
      try { stRealOut.text = "Batch output: " + (outFolder ? outFolder.fsName : "(not set)"); } catch (_) {}

      alert("Dry run failed on:\n\n" + fileObj.name + "\n\n" + e);
    } finally {
      btnAgain.enabled = true;
      try { dlg.layout.layout(true); } catch (_) {}
      try { dlg.update(); } catch (_) {}
    }
  }

  btnAgain.onClick = function () { doOneDryRun(); };

  btnRunBatch.onClick = function () {
    runBatchNow = true;
    dlg.close(1);
  };

  btnBack.onClick = function () {
    runBatchNow = false;
    dlg.close(0);
  };

  doOneDryRun();

  dlg.center();
  dlg.show();

  return { runBatchNow: runBatchNow };
}

/* ───────────────────────────────────────────────────────────────────
 * Processing
 * ─────────────────────────────────────────────────────────────────── */
function processOne(fileObj, outFile, opts, BORDER_WHITE, BORDER_BLACK, BORDER_CUSTOM) {
  var doc = null;

  try {
    doc = safeOpenWithRetry(fileObj, 2, 200);

    try { doc.changeMode(ChangeMode.RGB); } catch (_) {}

    if (opts.srgbMode === "FORCE") {
      try { doc.convertProfile("sRGB IEC61966-2.1", Intent.PERCEPTUAL, true, true); } catch (_) {}
    } else if (opts.srgbMode === "AUTO") {
      var alreadySRGB = false;
      try {
        var prof = (doc.colorProfileName || "");
        alreadySRGB = /srgb/i.test(prof);
      } catch (_) {}
      if (!alreadySRGB) {
        try { doc.convertProfile("sRGB IEC61966-2.1", Intent.PERCEPTUAL, true, true); } catch (_) {}
      }
    }

    try { doc.flatten(); } catch (_) {}

    if (opts.ignoreBorder) {
      capLongSide(doc, opts.longSide);
      if (opts.setPPI && opts.ppi) {
        try { setResolutionNoResample(opts.ppi); } catch (_) {}
      }
      saveJpegRespectingIgnore(doc, outFile, opts);
      return;
    }

    var borderColor = BORDER_WHITE;

    if (opts.borderMode === "BLACK") {
      borderColor = BORDER_BLACK;

    } else if (opts.borderMode === "CUSTOM") {
      borderColor = BORDER_CUSTOM;

    } else if (opts.borderMode === "AVERAGE") {
      borderColor = getAverageColorFast(doc, BORDER_WHITE);

    } else if (opts.borderMode === "AUTO") {
      borderColor = pickAutoBorderColor(doc, BORDER_WHITE, BORDER_BLACK);

    } else if (opts.borderMode === "AUTO_FILENAME") {
      var directive = detectFilenameBorderDirective(fileObj);
      if (directive === "WHITE") borderColor = BORDER_WHITE;
      else if (directive === "BLACK") borderColor = BORDER_BLACK;
      else if (directive === "AVERAGE") borderColor = getAverageColorFast(doc, BORDER_WHITE);
      else if (directive === "LUM") borderColor = pickAutoBorderColor(doc, BORDER_WHITE, BORDER_BLACK);
      else borderColor = pickAutoBorderColor(doc, BORDER_WHITE, BORDER_BLACK);
    }

    app.backgroundColor = borderColor;

    try { doc.backgroundLayer; } catch (_) {
      var fill0 = doc.artLayers.add();
      fill0.name = "BORDER_BG";
      doc.selection.selectAll();
      doc.selection.fill(app.backgroundColor);
      doc.selection.deselect();
      doc.flatten();
    }

    var wPx = doc.width.as('px');
    var hPx = doc.height.as('px');

    if (opts.ignoreRatio) {

      var contentMax = Math.max(1, Math.round(opts.longSide));
      var contentLong = Math.max(wPx, hPx);

      if (contentLong > contentMax) {
        var s0 = contentMax / contentLong;
        doc.resizeImage(
          UnitValue(Math.round(wPx * s0), 'px'),
          UnitValue(Math.round(hPx * s0), 'px'),
          null,
          ResampleMethod.BICUBIC
        );
        wPx = doc.width.as('px');
        hPx = doc.height.as('px');
      }

      if (opts.setPPI && opts.ppi) {
        try { setResolutionNoResample(opts.ppi); } catch (_) {}
      }

      var canvasW0 = Math.round(wPx + 2 * opts.padding);
      var canvasH0 = Math.round(hPx + 2 * opts.padding);

      doc.resizeCanvas(
        UnitValue(canvasW0, 'px'),
        UnitValue(canvasH0, 'px'),
        AnchorPosition.MIDDLECENTER
      );

    } else {
      var ratioW = opts.ratioW;
      var ratioH = opts.ratioH;

      var canvasH = Math.round(opts.longSide);
      var canvasW = Math.round(opts.longSide * ratioW / ratioH);

      var innerW = Math.max(1, canvasW - 2 * opts.padding);
      var innerH = Math.max(1, canvasH - 2 * opts.padding);

      var scale = Math.min(innerW / wPx, innerH / hPx);
      if (scale !== 1) {
        doc.resizeImage(
          UnitValue(Math.round(wPx * scale), 'px'),
          UnitValue(Math.round(hPx * scale), 'px'),
          null,
          ResampleMethod.BICUBIC
        );
      }

      if (opts.setPPI && opts.ppi) {
        try { setResolutionNoResample(opts.ppi); } catch (_) {}
      }

      doc.resizeCanvas(
        UnitValue(canvasW, 'px'),
        UnitValue(canvasH, 'px'),
        AnchorPosition.MIDDLECENTER
      );
    }

    saveJpegRespectingIgnore(doc, outFile, opts);

  } finally {
    try { if (doc) doc.close(SaveOptions.DONOTSAVECHANGES); } catch (_) {}
  }
}

function detectFilenameBorderDirective(fileObj) {
  try {
    var name = (fileObj && fileObj.name) ? String(fileObj.name) : "";
    if (name.indexOf("-White-") === 0) return "WHITE";
    if (name.indexOf("-Black-") === 0) return "BLACK";
    if (name.indexOf("-Average-") === 0) return "AVERAGE";
    if (name.indexOf("-Lum-") === 0) return "LUM";
  } catch (_) {}
  return "";
}

function stripBorderPrefixFromBase(baseName) {
  try {
    var b = String(baseName || "");
    var prefixes = ["-White-", "-Black-", "-Average-", "-Lum-"];
    for (var i = 0; i < prefixes.length; i++) {
      var p = prefixes[i];
      if (b.indexOf(p) === 0) return b.substring(p.length);
    }
    return b;
  } catch (_) {
    return baseName;
  }
}

function capLongSide(doc, maxLongSide) {
  var wPx = doc.width.as('px');
  var hPx = doc.height.as('px');
  var contentLong = Math.max(wPx, hPx);
  var contentMax = Math.max(1, Math.round(maxLongSide));
  if (contentLong > contentMax) {
    var s = contentMax / contentLong;
    doc.resizeImage(
      UnitValue(Math.round(wPx * s), 'px'),
      UnitValue(Math.round(hPx * s), 'px'),
      null,
      ResampleMethod.BICUBIC
    );
  }
}

function saveJpegRespectingIgnore(doc, outFile, opts) {
  if (opts.ignoreFileSizeLimits) {
    saveJpegOnce(doc, outFile, opts);
    return;
  }
  saveJpegPossiblyRanged(doc, outFile, opts);
}

function saveJpegOnce(doc, outFile, opts) {
  var j = new JPEGSaveOptions();
  j.formatOptions = FormatOptions.OPTIMIZEDBASELINE;
  j.embedColorProfile = opts.embedProfile;
  j.quality = clampInt(opts.jpegQuality, 0, 12);
  doc.saveAs(outFile, j, true);
}

function saveJpegPossiblyRanged(doc, outFile, opts) {
  var j = new JPEGSaveOptions();
  j.formatOptions = FormatOptions.OPTIMIZEDBASELINE;
  j.embedColorProfile = opts.embedProfile;

  var startQ = clampInt(opts.jpegQuality, 0, 12);

  var minBytes = (opts.minKB > 0) ? (opts.minKB * 1024) : 0;
  var maxBytes = (opts.maxKB > 0) ? (opts.maxKB * 1024) : 0;

  if (minBytes <= 0 && maxBytes <= 0) {
    j.quality = startQ;
    doc.saveAs(outFile, j, true);
    return;
  }

  function saveAtQuality(q) {
    j.quality = q;
    doc.saveAs(outFile, j, true);
    try { $.sleep(20); } catch (_) {}
    var bytes = 0;
    try { bytes = outFile.length; } catch (_) { bytes = 0; }
    return bytes;
  }

  function inRange(bytes) {
    if (bytes <= 0) return false;
    if (minBytes > 0 && bytes < minBytes) return false;
    if (maxBytes > 0 && bytes > maxBytes) return false;
    return true;
  }

  function score(bytes) {
    if (bytes <= 0) return 1e18;
    if (inRange(bytes)) return 0;
    if (minBytes > 0 && bytes < minBytes) return (minBytes - bytes);
    if (maxBytes > 0 && bytes > maxBytes) return (bytes - maxBytes);
    return 1e18;
  }

  var bestScore = 1e18;

  var bytes0 = saveAtQuality(startQ);
  bestScore = score(bytes0);
  if (bestScore === 0) return;

  var q;
  if (maxBytes > 0 && bytes0 > maxBytes) {
    for (q = startQ - 1; q >= 0; q--) {
      var b = saveAtQuality(q);
      var sc = score(b);
      if (sc < bestScore) bestScore = sc;
      if (sc === 0) return;
      if (minBytes > 0 && b > 0 && b < minBytes) break;
    }

  } else if (minBytes > 0 && bytes0 > 0 && bytes0 < minBytes) {
    for (q = startQ + 1; q <= 12; q++) {
      var b2 = saveAtQuality(q);
      var sc2 = score(b2);
      if (sc2 < bestScore) bestScore = sc2;
      if (sc2 === 0) return;
      if (maxBytes > 0 && b2 > maxBytes) break;
    }

  } else {
    for (q = 0; q <= 12; q++) {
      var b3 = saveAtQuality(q);
      var sc3 = score(b3);
      if (sc3 < bestScore) bestScore = sc3;
      if (sc3 === 0) return;
    }
  }
}

function pickAutoBorderColor(doc, WHITE, BLACK) {
  var lum = getLuminanceFast(doc);
  if (lum < 0) return WHITE;
  if (lum <= 105) return BLACK;
  return WHITE;
}

function getLuminanceFast(doc) {
  var tmp = null;
  try {
    tmp = doc.duplicate("LUM_TMP", false);
    try { tmp.flatten(); } catch (_) {}

    try { tmp.resizeImage(UnitValue(50, 'px'), UnitValue(50, 'px'), null, ResampleMethod.BICUBIC); } catch (_) {}

    try { executeAction(charIDToTypeID("Avrg"), undefined, DialogModes.NO); }
    catch (_) { try { tmp.close(SaveOptions.DONOTSAVECHANGES); } catch (__) {} return -1; }

    var sampler = tmp.colorSamplers.add([0, 0]);
    var c = sampler.color;
    sampler.remove();

    var y = (0.2126 * c.rgb.red) + (0.7152 * c.rgb.green) + (0.0722 * c.rgb.blue);

    tmp.close(SaveOptions.DONOTSAVECHANGES);
    tmp = null;

    return y;

  } catch (_) {
    try { if (tmp) tmp.close(SaveOptions.DONOTSAVECHANGES); } catch (__) {}
    return -1;
  }
}

function getAverageColorFast(doc, fallbackSolidColor) {
  var tmp = null;
  try {
    tmp = doc.duplicate("AVG_TMP", false);
    try { tmp.flatten(); } catch (_) {}

    try { executeAction(charIDToTypeID("Avrg"), undefined, DialogModes.NO); }
    catch (_) { try { tmp.close(SaveOptions.DONOTSAVECHANGES); } catch (__) {} return fallbackSolidColor; }

    var sampler = tmp.colorSamplers.add([0, 0]);
    var c = sampler.color;
    sampler.remove();

    var out = new SolidColor();
    out.rgb.red = c.rgb.red;
    out.rgb.green = c.rgb.green;
    out.rgb.blue = c.rgb.blue;

    tmp.close(SaveOptions.DONOTSAVECHANGES);
    tmp = null;

    return out;

  } catch (_) {
    try { if (tmp) tmp.close(SaveOptions.DONOTSAVECHANGES); } catch (__) {}
    return fallbackSolidColor;
  }
}

function setResolutionNoResample(ppi) {
  var idImgS = charIDToTypeID("ImgS");
  var desc = new ActionDescriptor();
  desc.putBoolean(charIDToTypeID("Rsmpl"), false);
  desc.putUnitDouble(charIDToTypeID("Rslt"), charIDToTypeID("#Rsl"), ppi);
  executeAction(idImgS, desc, DialogModes.NO);
}

/* ───────────────────────────────────────────────────────────────────
 * Progress UI (no cancel)
 * ─────────────────────────────────────────────────────────────────── */
function makeProgressUI(total, startIdx) {
  var win = new Window("palette", "BorderForge Progress");
  win.orientation = "column";
  win.alignChildren = "fill";
  win.margins = [12, 12, 12, 12];
  try { win.alwaysOnTop = true; } catch (_) {}

  var txt = win.add("statictext", undefined, "");
  txt.characters = 60;

  var bar = win.add("progressbar", undefined, 0, Math.max(1, total));
  bar.preferredSize = [540, 18];

  var txt2 = win.add("statictext", undefined, "");
  txt2.characters = 60;

  var row = win.add("group");
  row.orientation = "row";
  row.alignChildren = ["right", "center"];
  row.add("statictext", undefined, "Created By Seventeen");

  bar.value = Math.max(0, Math.min(total, startIdx));
  txt.text = startIdx + " / " + total;
  txt2.text = "";

  try { win.layout.layout(true); } catch (_) {}
  return { win: win, bar: bar, txt: txt, txt2: txt2, _tick: 0 };
}

function updateProgressUI(prog, idx, total, statusLine) {
  if (!prog || !prog.win) return;
  try {
    prog.bar.maxvalue = Math.max(1, total);
    prog.bar.value = Math.max(0, Math.min(total, idx));
    prog.txt.text = idx + " / " + total;
    prog.txt2.text = statusLine || "";

    try { $.sleep(10); } catch (_) {}

    prog._tick = (prog._tick || 0) + 1;
    if ((prog._tick % 25) === 0) {
      try { prog.win.show(); } catch (_) {}
    }
    try { prog.win.active = true; } catch (_) {}
    try { prog.win.update(); } catch (_) {}
  } catch (_) {}
}

/* ───────────────────────────────────────────────────────────────────
 * Persistent settings
 * ─────────────────────────────────────────────────────────────────── */
function getDefaultSettings() {
  return {
    hasAny: true,

    rememberFolders: true,
    inputPath: "",
    useCustomOut: false,
    customOutPath: "",

    ratioPresetText: "4:5 (Instagram Portrait)",
    ratioW: 4,
    ratioH: 5,
    longSide: 1350,
    padding: 40,
    ignoreRatio: false,

    ignoreBorder: false,
    ignoreFileSizeLimits: false,

    borderMode: "WHITE",
    borderHex: "FFFFFF",

    jpegQuality: 10,
    minKB: 0,
    maxKB: 900,
    suffix: "WB",

    srgbMode: "OFF",
    embedProfile: false,

    setPPI: true,
    ppi: 72,

    skipExisting: true,
    silentMode: true,
    chunkSize: 100,
    cooldownMs: 0,
    scratchMaxRetries: 2,
    scratchRetryCooldownMs: 5000
  };
}

function buildSettingsForSave(opts, folderBits) {
  return {
    hasAny: true,

    rememberFolders: !!folderBits.rememberFolders,
    inputPath: folderBits.inputPath || "",
    useCustomOut: !!folderBits.useCustomOut,
    customOutPath: folderBits.customOutPath || "",

    ratioPresetText: opts.ratioW + ":" + opts.ratioH,
    ratioW: opts.ratioW,
    ratioH: opts.ratioH,
    longSide: opts.longSide,
    padding: opts.padding,
    ignoreRatio: opts.ignoreRatio,

    ignoreBorder: opts.ignoreBorder,
    ignoreFileSizeLimits: opts.ignoreFileSizeLimits,

    borderMode: opts.borderMode,
    borderHex: opts.borderHex,

    jpegQuality: opts.jpegQuality,
    minKB: opts.minKB,
    maxKB: opts.maxKB,
    suffix: opts.suffix,

    srgbMode: opts.srgbMode,
    embedProfile: opts.embedProfile,

    setPPI: opts.setPPI,
    ppi: opts.ppi || 72,

    skipExisting: opts.skipExisting,
    silentMode: opts.silentMode,

    chunkSize: opts.chunkSize,
    cooldownMs: opts.cooldownMs,
    scratchMaxRetries: opts.scratchMaxRetries,
    scratchRetryCooldownMs: opts.scratchRetryCooldownMs
  };
}

function loadSettingsSafe(fileObj) {
  var out = { hasAny: false };
  try {
    if (!fileObj || !fileObj.exists) return out;

    fileObj.open("r");
    var text = fileObj.read();
    fileObj.close();

    var lines = String(text).split(/\r?\n/);
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];
      if (!line) continue;
      if (/^\s*#/.test(line)) continue;
      var eq = line.indexOf("=");
      if (eq < 0) continue;
      var k = line.substring(0, eq);
      var v = line.substring(eq + 1);
      k = String(k).replace(/^\s+|\s+$/g, "");
      v = String(v).replace(/^\s+|\s+$/g, "");
      if (!k) continue;

      out[k] = v;
      out.hasAny = true;
    }

    var numKeys = ["ratioW","ratioH","longSide","padding","jpegQuality","minKB","maxKB","ppi",
      "chunkSize","cooldownMs","scratchMaxRetries","scratchRetryCooldownMs"];
    for (var j = 0; j < numKeys.length; j++) {
      var nk = numKeys[j];
      if (typeof out[nk] !== "undefined") {
        var n = parseInt(out[nk], 10);
        if (isFinite(n)) out[nk] = n;
      }
    }

    return out;

  } catch (_) {
    try { if (fileObj && fileObj.opened) fileObj.close(); } catch (__) {}
    return { hasAny: false };
  }
}

function saveSettingsSafe(fileObj, settingsObj) {
  try {
    if (!fileObj) return;

    fileObj.open("w");

    fileObj.writeln("# BorderForge Export settings (v8.3)");
    fileObj.writeln("# Created By Seventeen");
    fileObj.writeln("# Simple key=value format for reliability.");
    fileObj.writeln("");

    for (var k in settingsObj) {
      if (!settingsObj.hasOwnProperty(k)) continue;
      var v = settingsObj[k];
      if (typeof v === "undefined") continue;
      fileObj.writeln(String(k) + "=" + String(v));
    }

    fileObj.close();
  } catch (_) {
    try { if (fileObj && fileObj.opened) fileObj.close(); } catch (__) {}
  }
}

/* ───────────────────────────────────────────────────────────────────
 * Output folder auto-numbering
 * ─────────────────────────────────────────────────────────────────── */
function getNextNumberedSubfolder(parentFolder, baseName) {
  var n = 1;
  while (true) {
    var name = (n === 1) ? baseName : (baseName + " " + n);
    var candidate = Folder(parentFolder.fsName + "/" + name);
    if (!candidate.exists) return candidate;
    n++;
  }
}

/* ───────────────────────────────────────────────────────────────────
 * Helpers
 * ─────────────────────────────────────────────────────────────────── */
function safeOpenWithRetry(fileObj, retries, sleepMsVal) {
  var lastErr;
  for (var i = 0; i <= retries; i++) {
    try {
      if (!fileObj.exists) throw new Error("File does not exist: " + fileObj.fsName);
      return app.open(fileObj);
    } catch (e) {
      lastErr = e;
      sleepMsFn(sleepMsVal);
    }
  }
  throw lastErr;
}

function makeOutputFile(outFolder, fileObj, suffix) {
  var dot = fileObj.name.lastIndexOf('.');
  var base = dot > -1 ? fileObj.name.substring(0, dot) : fileObj.name;
  base = stripBorderPrefixFromBase(base);
  return new File(outFolder.fsName + "/" + base + suffix + ".jpg");
}

function trySaveProgress(idx) {
  try {
    PROGRESS_FILE.open("w");
    PROGRESS_FILE.write(idx);
    PROGRESS_FILE.close();
  } catch (_) {}
}

function logLine(s) {
  try {
    LOG_FILE.open("a");
    LOG_FILE.writeln(timestampSafe() + " " + s);
    LOG_FILE.close();
  } catch (_) {}
}

function timestampSafe() {
  var d = new Date();
  function pad2(n){ return (n < 10) ? ("0"+n) : String(n); }
  return d.getFullYear() + "-" + pad2(d.getMonth()+1) + "-" + pad2(d.getDate()) + " " +
    pad2(d.getHours()) + ":" + pad2(d.getMinutes()) + ":" + pad2(d.getSeconds());
}

function sleepMs(ms) { sleepMsFn(ms); }
function sleepMsFn(ms) { try { $.sleep(ms); } catch (_) {} }

function lightCleanupFast() {
  try { app.purge(PurgeTarget.CLIPBOARD); } catch (_) {}
  try { $.gc(); } catch (_) {}
}

function deepCleanup() {
  try { app.purge(PurgeTarget.CLIPBOARD); } catch (_) {}
  try { app.purge(PurgeTarget.HISTORYCACHES); } catch (_) {}
  try { $.gc(); } catch (_) {}
}

function finalizePrefs(originalRulerUnits, originalHistoryStates) {
  try { app.preferences.rulerUnits = originalRulerUnits; } catch (_) {}
  try { app.preferences.numberOfHistoryStates = originalHistoryStates; } catch (_) {}
}

function hexToSolidColor(hex6) {
  var c = new SolidColor();
  var r = parseInt(hex6.substring(0,2), 16);
  var g = parseInt(hex6.substring(2,4), 16);
  var b = parseInt(hex6.substring(4,6), 16);
  c.rgb.red = r; c.rgb.green = g; c.rgb.blue = b;
  return c;
}

function clampInt(n, lo, hi) {
  n = parseInt(n, 10);
  if (!isFinite(n)) return lo;
  if (n < lo) return lo;
  if (n > hi) return hi;
  return n;
}

function rgbToHex(r, g, b) {
  function h2(n) {
    var s = n.toString(16).toUpperCase();
    return (s.length === 1) ? ("0" + s) : s;
  }
  return h2(r) + h2(g) + h2(b);
}
