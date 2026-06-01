# Ability spec: `evaluate`

## Purpose

Keep the **`evaluate`** MCP tool name and parameter shape aligned with upstream **[powerpoint-mcp](https://github.com/Ayushmaniar/powerpoint-mcp)** so agents and prompts transfer across platforms. **MacPoint does not run agent-supplied code** against PowerPoint or the host: the handler returns a **fixed guidance string** that steers callers to explicit tools ([`populate_placeholder`](populate_placeholder.md), [`manage_presentation`](manage_presentation.md), and the rest of the surface).


## Parameters (contract)

Registered in [`macpoint/server.py`](../../macpoint/server.py) (~201–212). **All arguments are ignored** for any execution path; they exist for **call-shape parity** with the reference server.

| Parameter | Type | Required | Notes |
|-----------|------|----------|--------|
| `code` | string | Yes | **Never executed** on MacPoint. Upstream may treat this as Python against COM — **#verify** when auditing reference. |
| `slide_number` | `str` \| `int` \| null | No | Ignored today. **#verify** whether reference uses this to scope execution to a slide. |
| `shape_ref` | string \| null | No | Ignored today. **#verify** upstream semantics (shape id / name). |
| `description` | string \| null | No | Ignored today. **#verify** upstream use (logging / UI only). |

## MacPoint today

Implementation: [`macpoint/server.py`](../../macpoint/server.py) (`evaluate`, ~201–212). The handler assigns `_ = code, slide_number, shape_ref, description` and **always** returns the following string (verbatim from code — update this doc if `server.py` changes):

```
MacPoint does not execute arbitrary code against PowerPoint. Use the explicit tools (populate_placeholder, manage_presentation, …). See docs/mac-deltas.md.
```

This is **intentional policy**, not a deferred stub: see [mac-deltas.md](../mac-deltas.md) **Security** and [reverse engineering checklist](../../reverse_engineering_checklist.md) **Per-tool porting workflow** step 6 (**Exception: `evaluate`**).

## Reference parity audit

- [ ] Confirm upstream **`evaluate`** executes user-controlled Python in a COM context (GitHub or `repos/powerpoint-mcp/` when present).
- [ ] Record **intentional delta**: same tool name and parameters for portability; **different execution model** — guidance-only on Mac.
- [ ] If upstream adds or renames parameters, preserve **MCP call shape** on MacPoint and extend this table; still **no** `eval` / `exec` of `code`.
- [ ] Any future “batch” or automation helper must be **allow-listed** operations only; document in [mac-deltas.md](../mac-deltas.md) and the checklist in the **same** change as code.

## Implementation plan

1. **Default — no change** — Keep guidance-only return; safest for vault and operator machines.
2. **Future features** — Prefer **new explicit tools** or a **small allow-listed** command set with fixed verbs; never run arbitrary strings from the model as code.
3. **Optional UX** — Richer static guidance (e.g. keyword hints toward the right tool) is allowed; still **zero** execution of `code`.
4. **Verification** — If behavior changes, update [verification-checklist.md](../verification-checklist.md) and [mac-deltas.md](../mac-deltas.md) in the same commit as `server.py`.

## Related

- [Reverse engineering checklist](../../reverse_engineering_checklist.md)
- [mac-deltas.md](../mac-deltas.md) (Security; tool parity table)
- [Verification checklist](../verification-checklist.md)
- [populate_placeholder.md](populate_placeholder.md)
- [manage_presentation.md](manage_presentation.md)
- Code: [`macpoint/server.py`](../../macpoint/server.py)
