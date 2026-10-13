# Changelog

All notable changes to this project are documented in this file.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [0.1.0] - 2026-10-13

First public release.

### Added
- AutoCAD MCP server over Windows COM: `acad_status`, `new_drawing`, `open_drawing`, `save_drawing`, `draw_line`, `draw_circle`, `draw_rectangle`, `draw_polyline`, `add_text`, `add_mtext`, `list_layers`, `set_layer`, `rename_layer`, `set_color`, `zoom_extents`, `count_entities`.
- `run_command` tool, disabled by default and enabled only with `AUTOCAD_MCP_ENABLE_RUN_COMMAND=1`.
- `read_beams_slabs.py` helper for reading beam/slab labels and dimension tables from structural plans.
- Claude skills: `structural-drawing-to-excel`, `compare-dwg-revisions`, `silicone-usage-estimate`.
- Traditional Chinese and English README, MIT license, example MCP client config.

[0.1.0]: https://github.com/Roy-y111/autocad-mcp/releases/tag/v0.1.0
