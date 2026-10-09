---
name: compare-dwg-revisions
description: Compare two old/new AutoCAD DWG revisions without exporting DXF, ignore content that only moved, semantically consolidate raw differences into a near-deliverable Excel change register, and produce native AutoCAD comparison snapshots with matching item callouts. Use for structural, architectural, MEP, shop, or other CAD drawings when the user says phrases such as「比較DWG圖面並整理EXCEL」、「圖面比較」、「比較新舊CAD」、「整理修改內容到Excel」or asks to mark Excel item numbers on a comparison snapshot. Do not use for editing design geometry or making engineering approval decisions.
---

# CAD DWG revision comparison

Compare DWG revisions on copies with AutoCAD Core Console. Do not require the AutoCAD GUI and do not export DXF.

## Inputs

Resolve exactly two `.dwg` files and identify which is old and which is new. Infer revision order from explicit labels, filenames, dates, or title-block text. Ask one concise question only when the order is genuinely ambiguous.

Default outputs (written into a new, empty output directory):

- `DWG圖面修改彙整.xlsx`
- `CAD圖面比較快照_原始.dwg`
- `CAD圖面比較快照_項次標示.dwg`
- `changes.json`, `change_candidates.json`, and execution logs for traceability

## Safety rules

- Copy both inputs before processing. Never save, close, or modify the originals.
- Prefer `accoreconsole.exe`; do not attach to the user's active AutoCAD session or the `mcp__autocad__*` MCP session.
- Require the AutoCAD GUI to be closed before starting. Never close it automatically — ask the user to close it.
- Run every Core Console process through the bundled `AutoCADProfileGuard.ps1`. It preserves and verifies `FILEDIA`/`SECURELOAD`; abort when the settings cannot be snapshotted or restored.
- Never terminate an AutoCAD process that this workflow did not start.
- Treat generated changes as review candidates. Do not claim structural, code, construction, or design approval.
- Preserve existing outputs by using a new output directory or obtaining approval before replacement.

## Prerequisites (this environment)

- AutoCAD Core Console: auto-detected under `C:\Program Files\Autodesk\AutoCAD <year>\accoreconsole.exe` for years 2020–2027. Pass `-CoreConsole` only if it's installed elsewhere.
- Python 3 on PATH with `openpyxl` installed (`python -c "import openpyxl"` to verify). No Node.js dependency — the Excel builder is pure Python.

## Workflow

1. Read [references/autocad-batch.md](references/autocad-batch.md) before running AutoCAD.
2. Read [references/change-review.md](references/change-review.md) before reviewing or rewriting the candidate descriptions.
3. Confirm AutoCAD's GUI (`acad.exe`) is not running, and that the AutoCAD MCP server (if connected) is not holding a session against either drawing.
4. Run:

   ```powershell
   powershell.exe -ExecutionPolicy Bypass -File <skill>/scripts/Invoke-DwgRevisionCompare.ps1 `
     -OldDrawing <old.dwg> `
     -NewDrawing <new.dwg> `
     -OutputDirectory <new-output-folder>
   ```

   `-PythonExe` and `-CoreConsole` are optional overrides; both are auto-detected otherwise.

5. Treat `change_candidates.json` as internal evidence only. Never deliver it as the modification register.
6. Confirm `changes.json` has `meta.delivery_ready=true`. The command stops automatically when floor/area recognition, text decoding, consolidation, or row-count quality checks fail.
7. Inspect the final Excel register directly (open it, or read it with the `xlsx` skill / `openpyxl`/`pandas`). This version does not render workbook preview images — spot-check specific old-to-new values and all rows marked for geometry review by hand.
8. If `changes.json` needs correction, edit the curated rows and rerun `scripts/New-DwgRevisionDeliverables.ps1` with the same Python and AutoCAD Core Console paths.
9. Report the three primary deliverables and any unresolved professional-review notes.

## Interpretation rules

- Match identical content before considering position. Identical content with only a translated position is a pure move and must not appear in the Excel register.
- Compare block definitions as well as model/paper-space inserts. Use matching block inserts as callout locations for changes inside block definitions.
- Treat text, dimensions, block names, entity type, layer, relative geometry, and bounding-box dimensions as evidence; do not rely on entity handles.
- Detect drawing/floor regions from plan-title text and title text height before assigning changes. Split repeated evidence by region, then merge it by one design decision and discipline category.
- Consolidate dimensions, grids, columns, beams, slabs/foundations, walls/openings, elevation notes, revisions, and general notes into concise area-level decisions.
- Exclude anonymous/system block internals and representation-only block replacement from the deliverable register while retaining the raw candidate evidence.
- Never leave generic rows such as `DIMENSION changed`, `geometry entities changed`, or `model space` when enough evidence exists to state a floor/area and design meaning.
- Use the generic columns 項次、圖面／樓層、位置、類別、更改內容簡述、備註. Adapt 圖面／樓層 to the discipline rather than forcing structural-floor terminology.

## Acceptance gate

- Deliver only the curated `changes.json`/Excel output, not the raw candidate list.
- Require decoded text with no replacement characters, meaningful drawing-area labels, concise old-to-new values where available, and one row per design decision within each area.
- Keep uncertain geometry in a consolidated row with a review note instead of expanding it into many entity-count rows.
- If the quality gate fails, continue reviewing and regenerating. Do not tell the user the task is complete.

## Failure handling

- If AutoCAD Core Console is unavailable, stop and report the missing AutoCAD installation. Do not silently fall back to DXF.
- If Core Console waits or fails, inspect the log for the DWG-version prompt, `SECURELOAD`, missing SHX warnings, profile access, or an existing output file.
- If direct DWG fact extraction is incomplete because of proxy/custom objects, keep the native comparison snapshot, flag the limitation, and request human review of those regions.
- If `meta.delivery_ready` is false, use `change_candidates.json`, the native snapshot, and the review rules to repair the curated register before delivery.

## Known limitations vs. the original Codex version

This skill was ported from `.agents/skills/compare-dwg-revisions/` (built for the Codex CLI) so it runs standalone under Claude Code:

- The original Excel builder (`build_excel.mjs`) imported `@oai/artifact-tool`, a package bundled only inside Codex's sandbox. It is not installable via npm and does not exist in this environment. It has been replaced with `scripts/build_excel.py` (openpyxl), which reproduces the same sheet layout, styling, and columns.
- Because `@oai/artifact-tool` also rendered PNG previews of the workbook, that preview step is gone. Step 7 above (manual inspection) replaces it.
- The Node.js runtime is no longer required anywhere in this workflow; `-NodeExe`/`-NodeModules` parameters were removed from both PowerShell scripts.
- `codex_app__load_workspace_dependencies` (a Codex-only tool for locating a bundled Python/Node) was removed; the scripts now auto-detect `python` on PATH and fall back to an explicit `-PythonExe` override.
- `agents/openai.yaml` (Codex's slash-command metadata) was not carried over — it has no equivalent in Claude Code's skill format.
- Everything else — the LISP fact exporter, `compare_facts.py`, `curate_changes.py`, `generate_annotation_lsp.py`, the AutoCAD Core Console command sequence, and `AutoCADProfileGuard.ps1` — is unchanged and was not Codex-specific.
