# Ability spec: `analyze_template`

## Purpose

Return **structured information** about a deck or template: slide masters, layouts, placeholder indices/names (and optionally more shape metadata), so agents can call **[`populate_placeholder`](populate_placeholder.md)** and future **`add_slide_with_layout`** (stub today) with valid identifiers without guessing from the UI.

Upstream reference: **[powerpoint-mcp](https://github.com/Ayushmaniar/powerpoint-mcp)** — same tool surface for cross-platform agents.

## Parameters (contract)

Registered in [`macpoint/server.py`](../../macpoint/server.py) (~142–144). **Mac v0 ignores both** and returns `_not_impl`.

| Parameter | Type | Default | Notes |
|-----------|------|---------|--------|
| `source` | string | `"current"` | Intended selector for **which** template/deck to analyze (e.g. active presentation vs path). Exact enum / semantics **#verify** against upstream `repos/powerpoint-mcp/`. |
| `detailed` | bool | `False` | Intended: summary vs expanded listing (extra shapes, XML ids, etc.). Not read on Mac v0. |

## MacPoint today


No `python-pptx` read path and no AppleScript.

## Reference parity audit

- [ ] Read upstream `analyze_template` when **`repos/powerpoint-mcp/`** is available: allowed `source` values; behavior for `"current"` with no open deck.
- [ ] Output format: plain text, JSON, Markdown; size limits and truncation.
- [ ] Difference between `detailed=False` and `True` (which fields appear).
- [ ] Update [mac-deltas.md](../mac-deltas.md) when Mac return contract is fixed.

## Implementation plan


1. **Resolve input** — Map `source="current"` to `state.get_last_active()` / AppleScript active path (same family as [`populate_placeholder`](populate_placeholder.md)); support explicit file path if reference does (add parameter only with parity + mac-deltas).
2. **python-pptx read-only MVP** — Open `.pptx` / `.potx`; enumerate slide masters, layouts, `placeholders` (idx, name, type); optional shape names on title slides. Return a stable, copy-paste-friendly format (decide: JSON vs Markdown table).
3. **`detailed=True`** — Add non-placeholder shapes, dimensions, or slide count sections; guard output size.
4. **Previews** — Optional later integration with [`slide_snapshot`](slide_snapshot.md) for thumbnails once that tool exists.
5. **Verification** — [verification-checklist.md](../verification-checklist.md) steps when shipped (fixture template, golden substring checks).
6. **Rollup** — Update [reverse engineering checklist](../../reverse_engineering_checklist.md) row when stub is replaced.

## Related

- [Reverse engineering checklist](../../reverse_engineering_checklist.md)
- [mac-deltas.md](../mac-deltas.md)
- [Verification checklist](../verification-checklist.md)
- [list_templates.md](list_templates.md) (discover template paths)
- [manage_presentation.md](manage_presentation.md) (open / create deck under analysis)
- [populate_placeholder.md](populate_placeholder.md) (consumer of placeholder names)
- [slide_snapshot.md](slide_snapshot.md) (future visual cross-check)
- Code: [`macpoint/server.py`](../../macpoint/server.py)
