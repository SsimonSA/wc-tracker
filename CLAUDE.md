# World Cup 2026 Pool Tracker

Personal/hobby project (not Sparks Analytics business code). A single-file Python
utility that pulls live 2026 FIFA World Cup results from ESPN's free JSON endpoint
and rebuilds an Excel workbook tracking my fantasy-pool teams plus a full 48-team
standings table.

## The pool
- 12 owners drafted 4 teams each (48 teams total).
- **Winner-take-all** ($240 pot): only the owner of the eventual World Cup champion
  wins. Points / goal difference in the sheet are for tracking and fun, **not** the
  prize.
- My teams (owner `STEVE S`): **Morocco** (the only realistic shot), Canada, Jordan, Iraq.
  - These were originally **John Vinson's** (`JOHN C`) teams. John dropped out and
    Steve Simon took them over (per Doug Vinson's 2026-06-10 email). The draft sheet
    (`World Cup 2026.xlsx`) still labels this row "JOHN C", but the `OWNERS` key and
    the `ME` constant in `update_wc.py` use **`STEVE S`** — the row "led by Canada".
  - Note: the draft sheet says "Monaco" — that's a typo for **Morocco** (carried
    as a name variant in `OWNERS`). Monaco is not in the tournament.
- Full owner→team map lives in the `OWNERS` dict at the top of `update_wc.py`
  (single source of truth — don't duplicate it elsewhere). It was reconstructed
  from `World Cup 2026.xlsx`; each team is a list of name **variants** (first is
  the display name) so ESPN's spellings (Ivory Coast, Cape Verde, United States,
  Türkiye, South Korea, Curaçao, …) all resolve via `norm()`.

## Files
- `update_wc.py` — everything: fetch + parse + build xlsx + emit the live web page.
  Stdlib + openpyxl.
- `run_wc.bat` — Windows double-click launcher (pip-installs openpyxl, runs script, opens xlsx).
- `wc_tracker.xlsx` — generated output, two tabs (`My Teams`, `Pool`). Regenerated every
  run; safe to delete.
- `index.html` — generated (via `--html`) self-contained live standings page for phones,
  published to GitHub Pages. **Fetches ESPN itself in the browser** (ESPN sends
  `Access-Control-Allow-Origin: *`), so it's live for every visitor with no server or
  cron — only republish when `OWNERS` changes. `OWNERS`/`ENDPOINT` are injected from
  the Python (still single source of truth); the JS re-implements `norm()` +
  `pool_table()`. The page is **player-agnostic**: a name dropdown lets each visitor
  pick their owner and highlight their own teams (saved per-device in `localStorage`,
  no default) — so `ME` is NOT used by the web page (only by the xlsx + console).
  In the knockout phase each team also shows a status badge (R32/R16/QF/SF/Final/
  🏆 CHAMP, or OUT) plus a next-match detail line that is **collapsed in the table**
  and revealed on hover (desktop) or tap (mobile) — the badge stays visible for
  at-a-glance status; eliminated teams are greyed/struck. The "my teams" chips at top
  show the detail always (only 4). See `HTML_TEMPLATE` / `build_html()`, `buildInfo()`
  and the `expanded` Set / row-click delegation in the JS. Published at
  https://ssimonsa.github.io/wc-tracker/
- `bracket.html` — generated alongside `index.html` (same `--html` run, written next
  to it). Self-contained live **knockout bracket**. Same self-contained pattern
  (fetches ESPN itself), shares `OWNERS` and the per-device name pick
  (`localStorage "wc-me"`) with the standings page, highlights the chosen owner's
  teams in gold, and has an "only my path" filter. **Two responsive layouts** (CSS
  switches at 900px, re-rendered on resize): a **connector-line tree** on desktop and
  **horizontally-scrolling round columns** on mobile. The tree's feeder topology is
  reconstructed at any seeding stage — `buildModel()` numbers each match by ESPN event
  `id` (sequential in bracket order), then resolves each match's two feeders from the
  placeholder name ("Round of 32 5 Winner") while a slot is unseeded and by
  winner-matching once it's filled (the 2026 bracket pairing is irregular, e.g. R16-1
  is fed by R32 matches 1 & 3 — so index math won't work). Future rounds show ESPN's
  placeholders until seeded. See `HTML_BRACKET` / `build_bracket()` and
  `buildModel()` / `renderTree()` / `renderColumns()`. The two pages cross-link via a
  nav bar.
- `sample.json` — offline test fixture mirroring ESPN's schema.

## Run
```
python update_wc.py                      # live fetch -> xlsx
python update_wc.py --out C:\path\wc.xlsx
python update_wc.py --fixture sample.json # offline, no network
python update_wc.py --html               # also write index.html + bracket.html (live web pages)
python update_wc.py --html out/index.html # ...to a chosen dir (bracket.html lands beside it)
```
Only dependency: `openpyxl`. The HTTP fetch uses stdlib `urllib` (nothing else to install).
The `--html` pages have **no** dependency — they're static files the browser runs.

## Data source
- ESPN (undocumented, free, no key):
  `https://site.api.espn.com/apis/site/v2/sports/soccer/fifa.world/scoreboard?limit=400&dates=20260611-20260719`
- Unofficial — can change or break without notice. If it returns 403, the browser
  `User-Agent` header is the usual fix (already set in `fetch_events`).
- Status `state` values: `pre` / `in` / `post`. **Live (`in`) games count provisionally**
  — a current lead shows as a Win until the match goes final.
- Schema essentials per game: `events[].competitions[0]` →
  `competitors[]` (each has `team.displayName`, `score` string, `homeAway`),
  `status.type.state`, and `altGameNote` (e.g. "FIFA World Cup, Group I").

## Key logic
- `OWNERS`: owner → list of teams; each team is a list of name **variants** (first is
  the display name). Matching is accent/punctuation-insensitive via `norm()`, so
  "Côte d'Ivoire", "Cape Verde", "Korea Republic", "Türkiye", "Bosnia and Herzegovina",
  etc. all resolve.
- `My Teams` tab: per-match detail; Result/Pts are Excel formulas off GF/GA.
- `Pool` tab: all 48 teams, aggregated in Python and written sorted by
  **Points → goal difference → goals for** (standard soccer order). My teams highlighted.
- Tie breaker = goal difference (GD = GF − GA).
- Times shown in ET via a fixed −4h (EDT) offset, valid for the whole tournament.

## Priorities / known gotchas
1. **Verify the 48-team name match against live ESPN data.** This was built and tested
   against fixtures only — the build environment couldn't reach ESPN. First task: run it
   live and confirm every team in `OWNERS` resolves to a real result. Any team showing
   0 games played when it actually played is a name-variant miss → add ESPN's exact
   spelling to that team's variant list in `OWNERS`.
2. **Knockouts:** the script filters the entire tournament date window, so knockout games
   involving my teams appear automatically once ESPN schedules them — no code change
   needed. Just sanity-check that the round label (`altGameNote`) renders cleanly for
   knockout fixtures.

## Possible enhancements (not done yet)
- Windows Task Scheduler entry for periodic auto-refresh (vs. manual double-click).
- Full group tables (all 4 teams per group) for advancement context.
- Put under version control. Hobby convention: personal repo under `SsimonSA`
  (business repos live under the `sparks-analytics` org — this isn't one of those).

## Style
Keep it single-file, dependency-light, and pragmatic. Exact code over abstraction;
performance, robustness, and operational simplicity over cleverness.
