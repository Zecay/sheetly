# Known bugs

Open issues found while testing. Not fixed yet.

## 1. Borders do not show on a freshly merged cell
- **Seen:** merge E:G (several rows) with a yellow fill, then apply "All Borders" (or any border) from the Borders menu: no border is drawn around the merged cell.
- **Expected:** the border is drawn around the merged area.
- **Notes:** the merged-cell draw code in `src/ui/grid.ts` (merged cells loop) was changed to read bottom/right borders from the other cells of the merged area. Borders applied through the Borders menu on an already-merged area probably only land on some cells, or the merge/border order matters. Check `applyStyle` for merged ranges in `src/core/ops.ts` and `drawBorders` input.
- **Found:** 2026-11-20
