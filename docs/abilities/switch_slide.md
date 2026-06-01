# Ability spec: `switch_slide`

## Purpose

Move the **active PowerPoint window** to a given **slide index** (1-based) in **slide view**, so operators or agents can drive the UI for demos, manual verification after other tools, or chained steps. This tool does **not** edit the file on disk; it only sends AppleScript to the running app.

Behavioral reference: upstream **[powerpoint-mcp](https://github.com/Ayushmaniar/powerpoint-mcp)** — same tool name for agent portability; Mac implementation is **AppleScript** only.

## Parameters (contract)

| Parameter | Type | Required | Notes |
|-----------|------|----------|--------|
| `slide_number` | `str` \| `int` | Yes | Coerced with `int(slide_number)` in [`server.py`](../../macpoint/server.py). Must be an integer value (e.g. `2` or `"2"`). **1-based** slide index in AppleScript (`go to slide n of active presentation`). |

## MacPoint today

Implementation: [`macpoint/server.py`](../../macpoint/server.py) (`switch_slide`, ~117–127).

1. Parse `slide_number` with `int(...)`; on `TypeError` / `ValueError`, return `Error: slide_number must be integer, got …`.
2. Call `applescript_ppt.switch_slide(n)`.
3. On success, return `Switched to slide {n} (if AppleScript succeeded; verify in PowerPoint).`
4. Any other exception is returned as `Error: {exc}`.

Backend: [`macpoint/backends/applescript_ppt.py`](../../macpoint/backends/applescript_ppt.py) — `switch_slide(slide_number: int)` runs:

```applescript
tell application "Microsoft PowerPoint"
    activate
    tell active window
        set view type of active pane to slide view
        go to slide {n} of active presentation
    end tell
end tell
```

**Scope:** Uses **active window** / **active presentation** only. There is **no** parameter to pick another open deck.

**Range validation:** MacPoint does **not** pre-check that `n` is within `1 … slide count`. Out-of-range or empty-deck behavior depends on **PowerPoint / Office version** (**#verify** per environment).

## Reference parity audit

- [ ] Read upstream `switch_slide` (or equivalent) in **`repos/powerpoint-mcp/`** when available, or on GitHub: parameters, 0-based vs 1-based indexing, error messages.
- [ ] Confirm whether reference can target a **non-active** presentation; if yes, document Mac gap or future `presentation_name`-style extension.
- [ ] Align user-facing error strings with reference where it improves agent UX without breaking existing clients.
- [ ] Document any intentional Mac-only limitation in [mac-deltas.md](../mac-deltas.md) when behavior or guarantees change.

**Operational risks:** Office / macOS variance for AppleScript dictionaries — see [Operational risks (Mac v0)](../../reverse_engineering_checklist.md#operational-risks-mac-v0) in the reverse-engineering checklist (same concern as the **`switch_slide`** row in the parity table).

## Implementation plan

1. **Upstream diff** — Note indexing and error semantics; add a short subsection to [mac-deltas.md](../mac-deltas.md) if Mac behavior is explicitly narrower than reference.
2. **Optional hardening (code)** — If AppleScript can return slide count cheaply: validate `1 <= n <= count` and return a clear error before `go to slide`. Keep string digits working (`"3"`).
3. **Edge cases** — Define behavior for `n < 1` in code (reject) vs rely on PowerPoint (document).
4. **Verification** — Extend [verification-checklist.md](../verification-checklist.md) with Office version + macOS row for `switch_slide`; re-run manual step after substantive AppleScript edits.

**Acceptance criteria (when changing behavior)**

- Tool doc, mac-deltas, and verification checklist stay aligned with `server.py` / `applescript_ppt.py`.

## Related

- [Reverse engineering checklist](../../reverse_engineering_checklist.md)
- [mac-deltas.md](../mac-deltas.md)
- [Verification checklist](../verification-checklist.md)
- [manage_presentation.md](manage_presentation.md) (opens / creates the deck this tool navigates)
- Code: [`macpoint/server.py`](../../macpoint/server.py), [`macpoint/backends/applescript_ppt.py`](../../macpoint/backends/applescript_ppt.py)
