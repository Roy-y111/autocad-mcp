# AutoCAD MCP Server

[繁體中文](README.md) | **English**

Lets MCP clients such as Claude drive a running AutoCAD directly through the Windows COM interface: open and save drawings, draw entities, manage layers, and run commands. It also ships a script that reads beam/slab labels and dimension tables from structural plan drawings, plus three Claude skills (structural dimension table to Excel, DWG revision comparison, silicone usage estimation).

## Requirements

- **Windows** (uses `pywin32` COM automation; will not run on macOS / Linux)
- **AutoCAD** desktop installed and already running
- Python 3.10 or newer (the code uses `match` statements)

## Installation

```bash
git clone https://github.com/Roy-y111/autocad-mcp.git
cd autocad-mcp
pip install -r requirements.txt
```

## Configuring your MCP client

Copy `claude_mcp_config.example.json`, replace `<專案資料夾完整路徑>` with the full path of the folder containing `main.py` on your machine, then paste it into your MCP client's config file (for example Claude Desktop's `claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "autocad": {
      "command": "python",
      "args": ["C:\\path\\to\\autocad-mcp\\main.py"],
      "env": {}
    }
  }
}
```

Note: backslashes inside JSON must be written as `\\`. After saving, restart the MCP client and make sure AutoCAD is open.

## Tools

| Category | Tools |
|---|---|
| Status and files | `acad_status`, `new_drawing`, `open_drawing`, `save_drawing` |
| Drawing | `draw_line`, `draw_circle`, `draw_rectangle`, `draw_polyline`, `add_text`, `add_mtext` |
| Layers and color | `list_layers`, `set_layer`, `rename_layer`, `set_color` |
| View and counting | `zoom_extents`, `count_entities` |
| Advanced (**disabled by default**) | `run_command` (runs arbitrary commands on the AutoCAD command line; must be enabled manually, see below) |

## Helper script

`read_beams_slabs.py`: connects to the currently open structural plan, reads beam and slab labels and AcDbTable dimension tables per floor range, and writes `beams_slabs_report.txt` (in the same folder as the script). The Y-coordinate range of each floor is defined in `FLOORS` at the top of the script; adjust it to match your drawing.

## Claude Skills

`skills/` contains three Claude skills for use with AutoCAD. Each has a `SKILL.md` (trigger conditions and workflow) and a `scripts/` folder:

| Skill | Purpose | Needs AutoCAD MCP |
|---|---|---|
| [structural-drawing-to-excel](skills/structural-drawing-to-excel/SKILL.md) | Reads the dimension tables from structural plans and produces an Excel member list with one sheet per floor (main beams, secondary beams, slabs, columns) for creating family types in Revit / Dynamo | Yes |
| [compare-dwg-revisions](skills/compare-dwg-revisions/SKILL.md) | Compares old and new DWG revisions, ignores content that only moved, produces an Excel change list, and generates an AutoCAD comparison drawing with numbered annotations | No (uses AutoCAD Core Console) |
| [silicone-usage-estimate](skills/silicone-usage-estimate/SKILL.md) | Gets the size of each type from door/window drawings, counts doors and windows per floor, and estimates silicone usage for frames and glass | Yes (or direct win32com) |

Usage: copy the skill folders you want into `~/.claude/skills/` (or your project's `.claude/skills/`), restart Claude, then say one of the trigger phrases listed in the `SKILL.md`. The skills and their trigger phrases are written in Traditional Chinese.

Notes:
- Scripts read and write intermediate files, placed by default in `<project root>/outputs/`. The `.lsp` files of `silicone-usage-estimate` can be redirected with the `SILICONE_OUT_DIR` environment variable; see that skill's `SKILL.md`.
- The example commands in `structural-drawing-to-excel` and `silicone-usage-estimate` contain `{USER}` or illustrative paths that you must replace with your own.
- The `.lsp` and PowerShell scripts have only been tested in the author's environment. If your AutoCAD version differs, try them on a test drawing first.

## ⚠️ Security notes

This server effectively hands your AutoCAD and local file access to an AI. Please understand the following before use:

- **`run_command` can execute arbitrary AutoCAD commands**, including loading LISP and running scripts, which is comparable to running programs on your computer. **This tool is disabled by default** and does not appear in the tool list. If you are sure you need it, add `"AUTOCAD_MCP_ENABLE_RUN_COMMAND": "1"` to `env` in the MCP config and restart:

  ```json
  "env": { "AUTOCAD_MCP_ENABLE_RUN_COMMAND": "1" }
  ```
- **`open_drawing` and `save_drawing` accept arbitrary paths**, so they can read or overwrite any location you have permission for.
- **Indirect prompt injection**: text, layer names, or attributes inside a DWG may be treated by the AI as instructions. **Only open drawings from sources you trust.**
- Keep your MCP client's per-action approval enabled before risky operations, and back up important drawings first.
- The server runs locally over stdio only and does not open network connections. Do not wrap it into a public network service.
