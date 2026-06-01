# Ability spec: `populate_placeholder`

## Purpose

Set **plain text** on a slide by matching a **placeholder name** (substring against shape or placeholder names) and writing the **`.pptx` on disk`** with **python-pptx**. This is **not** live in-app COM editing: PowerPoint may still have the file open, which interacts badly with **file locks** (see below).

Behavioral reference: **[powerpoint-mcp](https://github.com/Ayushmaniar/powerpoint-mcp)** — same tool surface for agents; Mac v0 supports a subset of `content_type` and matching behavior.

## Parameters (contract)

| Parameter | Type | Required | Notes |
|-----------|------|----------|--------|
| `placeholder_name` | string | Yes | Matched **case-insensitively** as a **substring** of `shape.name` or placeholder name (see [`pptx_backend.populate_plain_text`](../../macpoint/backends/pptx_backend.py)). |
| `content` | string | Yes | Written after [`_strip_simple_tags`](../../macpoint/backends/pptx_backend.py) (removes literal `<b>`, `</b>`, `<i>`, `</i>`, `<u>`, `</u>` only — not a full HTML sanitizer). |
| `content_type` | string | No (default `auto`) | Mac v0: only **`auto`** or **`text`** accepted; anything else returns an error from [`server.py`](../../macpoint/server.py). |
| `slide_number` | `str` \| `int` | No | Defaults to **1** if omitted. Coerced with `int(...)`; **1-based** slide index. Must be in range for the deck or `populate_plain_text` raises `ValueError`. |

## MacPoint today

Implementation: [`macpoint/server.py`](../../macpoint/server.py) (`populate_placeholder`, ~154–187).

**Path resolution (which `.pptx` is edited)**

1. [`state.get_last_active()`](../../macpoint/state.py) — set by **`manage_presentation`** (`open`, `create` with template + `file_path`, `save_as`) on success.
2. If `None`, [`applescript_ppt.active_presentation_path()`](../../macpoint/backends/applescript_ppt.py).
3. If path is still missing or `not path.exists()`, return an error instructing the operator to `open` / `save_as` first and to **close** the deck if a lock prevents saving.

**`content_type` gate** — Rejects values other than `auto` and `text` with a clear MCP error string.

**Backend** — [`macpoint/backends/pptx_backend.py`](../../macpoint/backends/pptx_backend.py) `populate_plain_text`:

- Loads `Presentation(pptx_path)`; rejects `slide_number` outside `1 … len(prs.slides)`.
- Scans `slide.shapes` for names **containing** `placeholder_name` (lowered), with `has_text_frame`; sets `text_frame.text`.
- If no matches, scans `slide.placeholders` the same way.
- If still zero matches, raises `ValueError` (surfaced as `Error: …`) suggesting `slide_snapshot` when available or inspecting names in PowerPoint.
- Calls **`prs.save(str(pptx_path))`** — overwrites the file on disk.

**Cross-tool dependency:** Callers should establish path context with **[manage_presentation](manage_presentation.md)** before relying on `populate_placeholder` in automation.

### Operational risks (file lock)

If PowerPoint holds a **write lock** on the `.pptx`, `Presentation.save` may fail or corrupt behavior may occur. Operator pattern: **save**, **close** the presentation in PowerPoint, run `populate_placeholder`, then **open** again. See [mac-deltas.md](../mac-deltas.md) (template / lock notes) and [Operational risks (Mac v0)](../../reverse_engineering_checklist.md#operational-risks-mac-v0) in the reverse-engineering checklist.

## Reference parity audit

- [ ] Read upstream `populate_placeholder` in **`repos/powerpoint-mcp/`** when available: supported `content_type` values, image/plot paths, HTML/LaTeX, exact vs substring name matching.
- [ ] List reference error messages vs Mac strings; align where it helps agents without breaking scripts.
- [ ] Document intentional Mac subset in [mac-deltas.md](../mac-deltas.md) when parity scope changes (not for this doc-only add).

## Implementation plan

1. **Upstream diff** — Map each `content_type` to implement vs stub vs document-only.
2. **Richer content** — Images/plots likely need different APIs than plain `text_frame.text`; design before exposing.
3. **Matching** — Optional exact-name mode or shape index to reduce accidental multi-match (today: counts shapes updated and returns `Updated {n} shape(s)…`).
4. **Locks** — Consider documenting-only vs optional “close via AppleScript before write” (risky UX); prefer explicit operator workflow until designed.
5. **Verification** — Extend [verification-checklist.md](../verification-checklist.md) with a `populate_placeholder` row (path known, deck closed, known placeholder name).

**Acceptance criteria (when changing code)**

- `populate_placeholder.md`, [mac-deltas.md](../mac-deltas.md), and verification checklist stay aligned with `server.py` / `pptx_backend.py`.

## Related

- [Reverse engineering checklist](../../reverse_engineering_checklist.md)
- [mac-deltas.md](../mac-deltas.md)
- [Verification checklist](../verification-checklist.md)
- [manage_presentation.md](manage_presentation.md) (path / open / save_as)
- [switch_slide.md](switch_slide.md) (UI navigation; does not edit disk)
- Code: [`macpoint/server.py`](../../macpoint/server.py), [`macpoint/backends/pptx_backend.py`](../../macpoint/backends/pptx_backend.py), [`macpoint/state.py`](../../macpoint/state.py)
