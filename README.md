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
