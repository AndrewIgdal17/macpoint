# Ability spec: `slide_snapshot`

## Purpose

Intended to give agents **visual or structured feedback** about a slide: e.g. a **screenshot** path, optional **bounding region**, or **text/shape extraction** so they can debug templates, confirm layout after `switch_slide`, or recover from [`populate_placeholder`](populate_placeholder.md) errors (“no placeholder matched … use **slide_snapshot** when available”).

Behavioral reference: **[powerpoint-mcp](https://github.com/Ayushmaniar/powerpoint-mcp)** — tool name and parameters stay aligned for cross-platform prompts; Mac implementation is **not started** (stub).

## Parameters (contract)

Registered on the MCP tool today ([`macpoint/server.py`](../../macpoint/server.py) ~107–111); **Mac v0 ignores all of them** and always returns the not-implemented string.

| Parameter | Type | Default | Notes |
|-----------|------|---------|--------|
| `slide_number` | `str` \| `int` \| omitted | `None` | Intended: which slide to capture (indexing TBD vs reference). |
| `include_screenshot` | `bool` \| omitted | `False` | Intended: whether to emit image data or a file path. |
| `screenshot_filename` | `str` \| omitted | `None` | Intended: suggested output path or basename for a screenshot file. |

When implementation lands, document **1-based vs 0-based** slide indexing and default slide when `slide_number` is null, after upstream parity audit.

## MacPoint today

**Stub only.** The handler assigns all arguments to `_` and returns [`_not_impl("slide_snapshot")`](../../macpoint/server.py) — the shared helper text:

`Not implemented on Mac v0: slide_snapshot. See Tools/MacPoint/docs/mac-deltas.md and Projects/MacPoint/Capability matrix.`

No AppleScript, no `python-pptx` read path, no file output.

## Reference parity audit

- [ ] Read upstream `slide_snapshot` in **`repos/powerpoint-mcp/`** when available: return type (file path, base64, JSON with dimensions?), required vs optional args, behavior when `slide_number` is missing.
- [ ] Document `include_screenshot` / `screenshot_filename` semantics from reference (naming, extension, overwrite rules).
- [ ] Decide Mac MVP return shape (e.g. POSIX path to a PNG under `/tmp` or vault scratch) for agent ergonomics.
- [ ] Update [mac-deltas.md](../mac-deltas.md) when any slice is implemented (not for this doc-only add).

## Implementation plan


1. **MVP — text subset (no image)** — Read-only `python-pptx`: enumerate slide `shapes` / placeholder names and text for slide `n`; return structured string or JSON-friendly text. Unblocks `populate_placeholder` debugging without display export.
2. **Export path** — AppleScript or UI automation: export slide or deck to **PDF/PNG** if PowerPoint’s dictionary supports it; respect `screenshot_filename` where possible (mkdir parents, safe paths).
3. **Screenshot flag** — Honor `include_screenshot` with clear behavior when `False` (text-only) vs `True` (require successful image write).
4. **Bbox / region** — Later phase if reference supports crop or region metadata; may depend on raster export pipeline.
5. **Verification** — Add rows to [verification-checklist.md](../verification-checklist.md) per shipped slice (Office version, output file exists, slide index edge cases).
6. **Rollup** — Update [reverse engineering checklist](../../reverse_engineering_checklist.md) table status and this doc’s **MacPoint today** when code replaces the stub.

## Related

- [Reverse engineering checklist](../../reverse_engineering_checklist.md)
- [mac-deltas.md](../mac-deltas.md)
- [Verification checklist](../verification-checklist.md)
- [populate_placeholder.md](populate_placeholder.md) (references `slide_snapshot` in error guidance)
- [manage_presentation.md](manage_presentation.md) (open deck before capture)
- Code: [`macpoint/server.py`](../../macpoint/server.py) (`slide_snapshot`, `_not_impl`)
