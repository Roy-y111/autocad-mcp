# Change review rules

## Candidate confidence

- High: same nearby entity changed text, dimension, block name, or a defining numeric property.
- Medium: unmatched entity additions/removals with stable layer and drawing-area context.
- Low: large grouped geometry count changes, proxy objects, anonymous blocks, hatches, or drawings with major coordinate-system changes.

## Discipline adaptation

- Structural: prefer floor, grid, beam, column, wall, slab, foundation, opening, and elevation terms.
- Architectural: prefer sheet, floor, room/zone, wall, door, window, stair, finish, dimension, and note terms.
- MEP: prefer system, floor/zone, equipment tag, route, size, elevation, circuit, and note terms.
- Generic: retain drawing/layout, layer, block, entity type, and coordinate location.

## Consolidation

- Merge repetitions caused by one design decision, but retain multiple callout coordinates under one item.
- Do not merge unrelated edits merely because they share a layer.
- Describe old-to-new values when both are available.
- Keep pure moves out of the register.
- Leave 備註 blank unless uncertainty, proxy-object limitations, or human verification is required.

## Delivery quality gate

- Treat raw entity differences as evidence, never as the Excel deliverable.
- Resolve model-space coordinates to a drawing, floor, layout, or named region before delivery.
- Remove anonymous blocks, dimension-arrow blocks, representation-only block swaps, and entity-count-only rows from the final register.
- Combine related grid and dimension edits, related member-label edits, and related wall/opening geometry within each drawing region.
- Retain a consolidated geometry row only when the native comparison snapshot shows a meaningful affected area; mark it for review in 備註.
- Reject the result when descriptions contain replacement characters, generic entity names without meaning, or an implausible row explosion relative to the number of recognized drawing regions.
