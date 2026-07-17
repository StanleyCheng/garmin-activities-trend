# Plan Review (Round 3): planning/PLAN.md

- Date: 2026-07-16
- Reviewer: Stanley (Round 3 pass)

## Edits since Round 2

1. **§1 "Privacy model" re-framed masking as cosmetic (L3 fix).** Verified — lines 22–26 now read "Masked Garmin username (cosmetic only — first letter + `***`; no domain). Masking is for appearance, not anonymity… Treat the dashboard as fully public." Clear and unambiguous.

2. **Phase 0 step 3 deprecates `DEFAULT_CHART_OUTPUT_FILE` (L8 fix).** Verified — line 314: "**In this phase, set `DEFAULT_CHART_OUTPUT_FILE = None` (or change default to `viz/index.html`) so the legacy inline-HTML path is fully deprecated.**" Two clean options offered.

3. **§9 "First-run expectations" subsection added (M2 fix).** Verified — lines 362–366. Includes expected non-empty `no_pace` count with diagnostic guidance.

4. **§10 wrist-HR filtering added as out-of-scope (M3 fix).** Verified — line 378. Cross-references §3.3 rule bounds (30/230) accurately.

No regressions in the surrounding text of any of the four edits.

## New issues found

### MEDIUM

**M1 — First-run expectations diagnostic is technically inaccurate (regression in the new §9 text).** Lines 364–365 read:

> "If `no_pace` is unexpectedly zero on first run, double-check that `transform.parse_activity_datetime` is correctly extracting `startTimeLocal` from all activities (a date drop reason of `> 0` plus `no_pace` = 0 would suggest a data-extraction bug)."

`parse_activity_datetime` handles `startTimeLocal` (dates). If `parse_activity_datetime` were broken, the symptom would be elevated `date` drop counts, NOT `no_pace = 0`. The two concerns are unrelated. The diagnostic for `no_pace = 0` should check `transform.normalize_activity`'s handling of `averageSpeed` (where `speed_to_pace_seconds` returns `None` for `speed <= 0`). The parenthetical "a date drop reason of `> 0` plus `no_pace` = 0…" implies a connection that does not hold.

**M2 — Dual-handle year-range slider has no library commitment (Plan→Implementation risk).** §4.1 line 178 specifies a "dual-handle slider" but the rest of the plan only commits to Plotly for chart rendering. HTML5 `<input type="range">` does not natively support dual handles. Without a library choice (`noUiSlider`, `ion.rangeSlider`, or an explicit "build from scratch" decision), Phase 3 step 13 will block on this decision. Recommend adding to §4.1: "Implementation: `noUiSlider` (vendored or CDN) or equivalent dual-handle widget; not native HTML."

### LOW

**L1 — §6 row 10 has a half-stale pointer.** Line 289 says: "Kept in `garmin_activities.json` and Excel output; documented as out-of-scope for the chart in §1 Privacy model and §10". The §10 callout landed (lines 380–391), but §1 has no mention of dropped parameters. Drop the §1 reference or add a one-line note to §1.

**L2 — §7 bullet 5 assertion phrasing is loose.** Line 300: "activity count equals sum of `activity_count_dropped` + `activity_count_after_clean`". `activity_count_dropped` is a dict (per the §3.4 example on line 141), not a scalar. The assertion should read: `total_raw_activities == sum(activity_count_dropped.values()) + activity_count_after_clean`. Wording is ambiguous as written.

**L3 — Phase 4 largely duplicates §6 fixes already distributed across Phases 0–3.** Lines 338–340 say "Phase 4 — Code-review fixes list (Section 6)" with step 17 "Apply each fix; tests still pass after each." Inspecting the §6 rows against the phases:
- Rows 1–3 (pace/None drop, TooManyRequests retry, fromisoformat) live in Phase 0's `transform.py` / `get_data.py` work.
- Rows 4, 6, 7, 8 (HTML extraction, root HTML deletion, `publish = "viz"`) live in Phase 2.
- Rows 9, 11 (activity_type allow-list, raw-metres distance check) live in §3.3 (Phase 1).
- Row 10 (parameter regression) is documentation.

Phase 4 will be empty unless the §6 rows are redistributed. Either annotate the §6 table with which Phase each row belongs to, or delete Phase 4 and let §6 remain as a cross-reference checklist.

## Status of unresolved Round 2 items

- **M2: resolved** — §9 lines 362–366 add "First-run expectations" with expected non-empty `no_pace` count. (Caveat: see M1 above — diagnostic text in the new section is inaccurate.)
- **M3: resolved** — §10 line 378 adds "Wrist-HR artefact filtering" as out-of-scope with cross-reference to §3.3 rule bounds.
- **L3: resolved** — §1 lines 22–26 explicitly re-frame masking as cosmetic ("Masking is for appearance, not anonymity… Treat the dashboard as fully public").
- **L8: resolved** — Phase 0 step 3 (line 314) explicitly deprecates `DEFAULT_CHART_OUTPUT_FILE`.

## Verdict

**Approve with minor changes.**

All four targeted edits landed cleanly. No regressions in the surrounding text. The two MEDIUM concerns (inaccurate §9 diagnostic, unspecified dual-handle slider library) should be addressed before Phase 3 implementation begins — M2 in particular because the implementer will block on a library decision otherwise. The three LOW items are cleanups that do not block start of work.