# PR Maker 1.1.0 — 2026-09-17

## Reliability

- RCC recognizes the observed `Remote Control Center - [ARMGC Monitor]` title.
- PR follows a separate Purchase Requisition window as well as MDI views. After
  Add, the description form must be detected in the PR window or its owned editor.
- Every PR text field must be a readable edit/date control before typing; entered
  description, date, selected account, enabled unit-price flag and optional part
  number are read back. An unrecognized control or mismatch stops subsequent input.
- Unreadable custom controls deliberately report incomplete verification. A live
  PR entry is still needed to verify every control variant in the installation.
- Foreground changes and cancellation stop further input. Legacy generic exception
  handlers cannot swallow the task-stop signal. An interrupted PR is not retried
  automatically: previously entered fields must be reviewed in HI-TOPS.
- Password updates use an atomic file replacement. Existing credentials remain in
  their existing AppData location. The current password is displayed read-only in
  settings as requested; new password entries are masked and can be revealed.

## UI

- Labeled dock buttons with tooltips and a separate expanded PR form.
- Current stage and elapsed time, stop button, log shortcut and safe retry for
  M&C/RCC/login. Controls are locked during a task; worker events are delivered
  through a queue consumed by the Tk thread.
- Ctrl+Shift+P shows the widget when idle. Ctrl+Shift+F12 requests cancellation;
  Escape cancels when the widget has focus, otherwise it hides the widget.
- Opening the widget during automation does not steal ERP focus. Tray actions
  include cancellation and graceful exit after the current worker stops.
- Reject empty/whitespace descriptions, unknown account codes and nonempty part
  numbers shorter than four characters before any ERP action.
- Filter account choices by code/name; keep the five most recent selections in
  `preferences.json`, separate from credentials. ERP dropdown navigation continues
  to use the original ordered `ACCOUNT_CODES`, never the reordered recent list.
- Settings display the current password, mask new passwords, require confirmation and reuse an existing settings
  dialog. Window placement clamps to the cursor monitor's work area, including
  negative monitor origins. A named mutex prevents duplicate 1.1 instances.

## Performance and distribution

- PR initializes HI-TOPS once, and waits for its window instead of a fixed 2.5 s.
- Matching requests the target window bounds and caches decoded/resized templates.
  Matching prioritizes the target monitor scale, remembers successful template scales,
  and searches ratios for assets captured at a different DPI (100–300%). Per-monitor
  V2 awareness is enabled before GUI imports; OCR is normalized to monitor DPI.
  OCR is bounded by the remaining navigation deadline and a maximum 1.5 s subprocess
  call. Cancellation waits for any current image/OCR operation to return.
- The Tesseract availability check is cached. Tesseract remains an external OCR
  prerequisite; the observed development machine has it installed.
- `requirements-build.lock` pins the clean build environment, including transitive
  dependencies. `build.bat` stops on install, test, build or smoke-check failure.
- `dist/release/PRMakerWidget.exe` is built separately from previous executables.
- `install.ps1` smoke-tests and hashes the copy before installing into
  `%LOCALAPPDATA%\PRMaker\app\PRMakerWidget.exe`. Existing installations are backed
  up during replacement. An existing startup shortcut is updated to this fixed path.
  Startup registration is not silently enabled if it was previously absent.
- Old pre-1.1 widgets do not participate in the new single-instance mutex. Exit the
  old widget through its tray menu before launching 1.1 for the first time.

## Validation

```powershell
build\venv\Scripts\python.exe -B -m unittest discover -s tests -p '*_regression.py' -v
build\venv\Scripts\python.exe -B tests\ui_smoke.py
build.bat
powershell -ExecutionPolicy Bypass -File install.ps1
```

Tests do not create ERP documents or send input to HI-TOPS. The hidden UI smoke test
checks layout at the current DPI, validation, focus protection, busy state and
cancellation recovery. Read-only checks confirmed the actual separate PR window and
that the installed WindowsForms date controls support bounded text readback.
No end-to-end PR creation was performed as part of these checks.

## 1.1.1 follow-up

- Wait for the owned Purchase Requisition Detail window before Description entry;
  the list's Description filter no longer qualifies as editor readiness.
- Log cancellation/focus-loss reasons to distinguish stopped tasks from failures.
- Open Berthing Schedule with Alt+V, Home, nine Down keys, then Enter, matching
  the supplied Vessel menu. Each key retains focus/cancellation protection.
  Confirm the resulting window without repeating selection on timeout.
  This assumes the supplied menu order; image/OCR matching is no longer used for
  this step. Existing open schedules are reused.

## 1.1.2 follow-up

- Reuse and activate an existing Maintenance & Repair window before searching
  Inventory; do not return to the HI-TOPS tile when M&R is already running.
- When M&R is absent, explicitly activate HI-TOPS before searching/clicking its
  tile. Record each menu stage and activation failure in the persistent log.
