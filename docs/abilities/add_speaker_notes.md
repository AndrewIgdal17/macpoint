# Ability spec: `add_speaker_notes`

## Purpose

Set or replace **speaker notes** for a given slide so agents can script talk tracks, handoff cues, or citations **without** editing on-slide body placeholders. Intended parity with upstream **[powerpoint-mcp](https://github.com/Ayushmaniar/powerpoint-mcp)** for the same tool name.

## Parameters (contract)

Registered on the MCP tool in [`macpoint/server.py`](../../macpoint/server.py) (~131–132). **Mac v0 ignores both parameters** and always returns the not-implemented string.

| Parameter | Type | Required | Notes |
|-----------|------|----------|--------|
| `slide_number` | `str` \| `int` | Yes (signature) | Intended **1-based** slide index after parity audit with reference; not read on Mac v0. |
| `notes_text` | string | Yes | Full notes body to apply (replace vs append semantics **TBD** vs reference). Not read on Mac v0. |

## MacPoint today

**Stub only.** The handler sets `_ = slide_number, notes_text` and returns [`_not_impl("add_speaker_notes")`](../../macpoint/server.py) — the shared message:

`Not implemented on Mac v0: add_speaker_notes. See Tools/MacPoint/docs/mac-deltas.md and Projects/MacPoint/Capability matrix.`

No AppleScript and no `python-pptx` notes write path.

## Reference parity audit

- [ ] Read upstream `add_speaker_notes` in **`repos/powerpoint-mcp/`** when available: **replace** vs **append** to existing notes, newline handling, encoding, max length.
- [ ] Confirm return value (success string, slide echo, error shape).
- [ ] Slide indexing: 0-based vs 1-based; interaction with “current slide” if any.
- [ ] Whether the reference targets only the **active** presentation (expected on Mac too).
- [ ] Update [mac-deltas.md](../mac-deltas.md) when Mac behavior is implemented or explicitly declared unsupported.

## Implementation plan


1. **AppleScript spike** — Inspect Microsoft PowerPoint for Mac’s AppleScript dictionary for notes page / shapes; prototype setting text for slide `n` after optional coordination with [`switch_slide`](switch_slide.md).
2. **python-pptx path** — Evaluate `notes_slide` / notes body APIs for writing without corrupting OOXML; compare **file lock** behavior with [`populate_placeholder`](populate_placeholder.md) (save while PowerPoint has file open).
3. **Fallback** — If neither path is reliable, document **intentional non-support** or partial support in mac-deltas and return a clear MCP error (not silent failure).
4. **Verification** — Add steps to [verification-checklist.md](../verification-checklist.md) when a backend ships (Office version, slide index, read-back confirmation in PowerPoint UI).
5. **Rollup** — Update [reverse engineering checklist](../../reverse_engineering_checklist.md) and this doc’s **MacPoint today** when the stub is replaced.

## Related

- [Reverse engineering checklist](../../reverse_engineering_checklist.md)
- [mac-deltas.md](../mac-deltas.md)
- [Verification checklist](../verification-checklist.md)
- [switch_slide.md](switch_slide.md) (navigate to slide before UI-driven notes automation)
- [manage_presentation.md](manage_presentation.md) (open / active deck)
- [populate_placeholder.md](populate_placeholder.md) (disk write patterns and locks)
- Code: [`macpoint/server.py`](../../macpoint/server.py) (`add_speaker_notes`)
