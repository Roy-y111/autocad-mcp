# AutoCAD batch workflow

Use `accoreconsole.exe` without attaching to the GUI. Require the AutoCAD GUI to be closed before starting because Core Console shares the same registry-backed user profile.

## Required command order

For native comparison:

```text
(save FILEDIA and CMDECHO with `getvar`)
(set FILEDIA to 0 and CMDECHO to 1 with `setvar`)
-COMPARE
<closed old drawing copy>
COMPAREEXPORT
<blank line accepts the default DWG version>
<snapshot output path>
(restore FILEDIA and CMDECHO)
QUIT
N
```

The blank line after `COMPAREEXPORT` is required: with `FILEDIA=0`, AutoCAD asks for the target DWG version before the filename.

For LISP loaded from the generated work directory, save `SECURELOAD`, set it to `0` only for the Core Console work, and restore it before quitting.

Always invoke Core Console through `AutoCADProfileGuard.ps1`. It snapshots the registry-backed `FileDialog` and `SECURELOAD` values, restores them in a PowerShell `finally` block after every Core Console process, and verifies the restored values. The in-script `getvar`/`setvar` restoration is the first layer; the registry guard is the failure-path layer.

## Operational constraints

- Stage copies as `old.dwg` and `new.dwg`.
- Stop before launching Core Console when an `acad.exe` GUI process is running. Ask the user to close AutoCAD; do not close it automatically.
- Use forward slashes inside `.scr` and LISP strings.
- Prefer Core Console over COM `SendCommand`. If COM is ever used, send `\n`, not `\r\n`.
- Allow user-profile access when the sandbox asks; Core Console initializes the AutoCAD profile.
- Missing SHX messages are warnings unless they prevent opening or materially affect title/text interpretation.
- Export DWG strings as numeric character-code arrays, then decode them in Python. Do not JSON-escape raw MBCS text because a Big5 trail byte can be mistaken for a backslash and corrupt Chinese labels.
- Keep the raw comparison snapshot unchanged and annotate a copy.
- `Invoke-DwgRevisionCompare.ps1` auto-detects `accoreconsole.exe` by scanning `C:\Program Files\Autodesk\AutoCAD <year>\` for years 2020–2027. Pass `-CoreConsole` explicitly if AutoCAD is installed elsewhere or under a non-default year.
