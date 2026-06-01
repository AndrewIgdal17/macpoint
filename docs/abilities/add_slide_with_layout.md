# Ability spec: `add_slide_with_layout`

## Purpose

Insert a **new slide** after slide index **`after_slide`** using a **layout** identified by **`layout_name`** within a **`template_name`** context — so agents can grow a deck (new section slides, agenda repeats, etc.) without manual duplicate in PowerPoint.


## Parameters (contract)

Registered in [`macpoint/server.py`](../../macpoint/server.py) (~147–150). **Mac v0 ignores all three** and returns `_not_impl`.

| Parameter | Type | Required | Notes |
|-----------|------|----------|--------|
| `template_name` | string | Yes | Intended: theme / attached template / deck identifier (**#verify** vs upstream — file path vs display name). |
| `layout_name` | string | Yes | Intended: slide layout name to apply (often discovered via [`analyze_template`](analyze_template.md)). |
| `after_slide` | int | Yes | Intended: insertion anchor slide index (**1-based vs 0-based** **#verify** vs reference; inclusive/exclusive semantics). |

## MacPoint today


## Reference parity audit

- [ ] Read upstream `add_slide_with_layout` in **`repos/powerpoint-mcp/`** when available: exact meaning of `template_name` and `layout_name` (case sensitivity, fuzzy match).
- [ ] `after_slide` when deck has 0 slides, at end, or out of range.
- [ ] Return string / errors on duplicate layout names across masters.
- [ ] Update [mac-deltas.md](../mac-deltas.md) when Mac semantics are fixed.

## Implementation plan

1. **Resolve deck** — Active `.pptx` from [`state`](../../macpoint/state.py) / AppleScript (same family as [`populate_placeholder`](populate_placeholder.md)); map `template_name` to master or external template only after parity audit.
2. **python-pptx MVP** — Find `SlideLayout` by name on the correct `SlideMaster`; `Presentation.slides.add_slide(layout)`; insert at index derived from `after_slide` per documented convention; `prs.save(...)`.
3. **Bounds and errors** — Clear errors for unknown layout, bad index, read-only file.
4. **File lock** — Same operator guidance as **populate_placeholder** if PowerPoint holds the file open during save.
5. **AppleScript fallback** — Spike only if OOXML insertion is unreliable on a given Office build.
6. **Verification** — [verification-checklist.md](../verification-checklist.md) when implemented; update [reverse engineering checklist](../../reverse_engineering_checklist.md) row from **Stub** to **Partial** when code ships.

## Related

- [Reverse engineering checklist](../../reverse_engineering_checklist.md)
- [mac-deltas.md](../mac-deltas.md)
- [Verification checklist](../verification-checklist.md)
- [analyze_template.md](analyze_template.md) (layout / master names)
- [manage_presentation.md](manage_presentation.md) (open / save deck)
- [populate_placeholder.md](populate_placeholder.md) (disk write + locks)
- Code: [`macpoint/server.py`](../../macpoint/server.py)
