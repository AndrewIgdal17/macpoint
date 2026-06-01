# Ability spec: `manage_slide`

## Purpose

Single MCP surface for **structural slide changes** in a deck: **duplicate**, **delete**, **reorder**, or **move** slides by index—so agents can edit deck shape without manual PowerPoint steps or raw OOXML.


## Parameters (contract)

Registered in [`macpoint/server.py`](../../macpoint/server.py) (~191–197). **Mac v0 ignores all parameters** and returns `_not_impl`.

| Parameter | Type | Required | Notes |
|-----------|------|----------|--------|
| `operation` | string | Yes | Intended: discriminator for duplicate / delete / reorder / move (exact spellings and case **#verify** vs upstream). |
| `slide_number` | `str` \| `int` | Yes | Intended: primary slide index for the operation (**1-based vs 0-based #verify**; align with [`switch_slide`](switch_slide.md) / [`populate_placeholder`](populate_placeholder.md) when implemented). |
| `target_position` | int \| null | No | Intended: destination index for reorder/move operations; ignored or error for duplicate/delete (**#verify**). |

## MacPoint today


## Reference parity audit

Audit upstream (GitHub or clone under `repos/powerpoint-mcp/` when present) before freezing Mac semantics:

- [ ] Exact **`operation`** values (case, aliases, unknown-operation errors).
- [ ] Whether each operation is **idempotent** or has side effects (e.g. duplicate twice).
- [ ] **`slide_number`** indexing and validation (empty deck, out of range).
- [ ] **`target_position`**: inclusive/exclusive placement, 1-based vs 0-based, required vs forbidden per operation.
- [ ] **Last slide delete**, **single-slide deck**, **no slides**.
- [ ] **Multiple open presentations** — reference may target by name; Mac v0 tools generally use **active presentation** only (**#verify**).
- [ ] Whether reference updates **current slide selection** after mutation.
- [ ] Record deltas in [mac-deltas.md](../mac-deltas.md) when behavior ships.

## Implementation plan

1. **Resolve `.pptx` path** — Same family as [`populate_placeholder`](populate_placeholder.md): [`state.get_last_active()`](../../macpoint/state.py) and fallback [`active_presentation_path()`](../../macpoint/backends/applescript_ppt.py) when the app reports a saved path.
2. **python-pptx MVP** — Use library-supported patterns for **delete** (drop slide from `presentation.slides`) and **reorder** (slide id / part order per `python-pptx` capabilities and OOXML constraints). **Duplicate** may require cloning slide parts and relationships (higher complexity—document limits in mac-deltas if subset-only).
3. **AppleScript spike** — If PowerPoint’s dictionary exposes duplicate/delete/move with stable behavior across Office builds, prefer for **live deck** consistency; else **disk-only** path with operator **close/reopen** note.
4. **File lock / stale UI** — Same caution as **populate_placeholder**: writing while PowerPoint holds the file may fail; editing on disk with the deck open may not refresh the UI until reload (**#verify** per build).
5. **Errors** — Clear messages for unknown `operation`, bad indices, missing path, locked file.
6. **Verification** — [verification-checklist.md](../verification-checklist.md) when implemented; update [reverse engineering checklist](../../reverse_engineering_checklist.md) row from **Stub** to **Partial** when code ships.

## Related

- [Reverse engineering checklist](../../reverse_engineering_checklist.md)
- [mac-deltas.md](../mac-deltas.md)
- [Verification checklist](../verification-checklist.md)
- [manage_presentation.md](manage_presentation.md)
- [switch_slide.md](switch_slide.md)
- [populate_placeholder.md](populate_placeholder.md)
- [add_slide_with_layout.md](add_slide_with_layout.md) (insert vs structural edit)
- Code: [`macpoint/server.py`](../../macpoint/server.py)
