# HI-TOPS menu navigation reliability (2026-09-14)

## Confirmed defects

- Window detection previously selected any title containing `MONITOR`, and could
  exclude the real M&C window when its title also contained `HI-TOPS`.
- Failed foreground activation did not stop keyboard input. Menu detection failure
  fell through to nine Down keys and Enter without checking which menu was open.
- `assets/berthing_schedule.png` is Base64 text rather than a PNG. Its decoded bytes
  are also unreadable by Pillow. The original file is retained as evidence; the new
  navigator detects and skips it. Replace it with a real menu capture if image
  matching is needed. Native menu commands and OCR work without this image.
- Monitoring was hovered only once; subsequent retries only searched the screen.
- OCR compared individual words against a multiword target, so `Berthing Schedule`
  and `Purchase Request` generally could not match.
- `.gitignore` was UTF-16, preventing Git from applying its patterns correctly.

## Changes

- Identify windows by executable path under the configured HI-TOPS installation,
  then classify main, M&C, RCC and Maintenance separately. A query failure does not
  fall back to unrelated programs. `roi_helpers.INSTALL_DIR` is the installation root.
- Verify foreground ownership after activation and again before menu clicks.
- Search images/OCR within the target window, preserving negative monitor origins.
- Retry Monitoring entry up to three times. After sending a launch click, wait for
  the target window instead of launching duplicate instances.
- Select RCC directly by its image/text rather than an unscaled 40px offset.
- Select Berthing Schedule by the actual enabled native menu command when exposed;
  otherwise reopen Vessel and use image/OCR. Remove blind Enter and index navigation.
- Confirm a visible Berthing Schedule window or child-window caption before success.
  A custom-drawn view without that caption may open successfully but report that
  confirmation failed; capture its actual accessible identity before extending this check.
- Recognized HI-TOPS error dialogs stop navigation and remain available to inspect.
  The old broad popup watchdog is not enabled.
- Log task results, exceptions and navigation stages to
  `%LOCALAPPDATA%\PRMaker\logs\navigation.log` (2 MB rotation, three backups).
  Logs do not record passwords, form inputs or desktop screenshots.
- Show failures in the widget. Preserve PR early failure results, stop after a failed
  Add lookup, and report password write errors instead of closing as if saved.
- Detect login by title/password edit controls instead of window width; stop when
  login cannot be confirmed. Custom login controls may need an additional identifier.

## Validation

```powershell
py -3.12 -B -m unittest discover -s tests -p '*_regression.py' -v
py -3.12 -m PyInstaller --noconfirm --distpath dist/navigation-fix --workpath build/navigation-fix PRMakerWidget.spec
```

The tests mock input and window APIs: they do not drive the ERP. They cover title
ambiguity, unrelated processes, focus failures, hover retries, timeout behavior,
native command selection, multiword OCR, DPI fallback and negative monitor origins.
The installed runtime also passed module imports and a native menu-caption read.
Tesseract was found on the development machine. It remains an external prerequisite
for OCR fallback on another machine.

The new EXE is `dist/navigation-fix/PRMakerWidget.exe`; the previous
`dist/PRMakerWidget.exe` is preserved. Exit the old widget through its tray menu
before opening the new EXE. HI-TOPS menu navigation itself still needs a live check:
test M&C initially closed, already open and minimized, and then RCC/PR. Check that
the expected view opens and that no unrelated window receives input. If it stops,
consult the stage and window details in the log.

## Backup

The configured origin is `https://github.com/yarokim83/HPNTHitops.git`.
The pre-change local/remote commit was `0aaeffd5f739b5b8e3a5e524f27b7c1736518f38`.
Changes are kept on a separate `codex/` branch; existing local screenshots and
bytecode changes are excluded from the source backup.
