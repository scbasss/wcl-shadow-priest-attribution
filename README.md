# WCL Shadow Priest Debuff Attribution

How much of a raid's magic damage is actually attributable to a Shadow
Priest's Shadow Weaving / Misery uptime — not estimated from average uptime,
but computed hit-by-hit from real aura-stack events pulled from a
[Warcraft Logs](https://www.warcraftlogs.com) report.

Given a report code, this pulls every damage event and every debuff
apply/refresh/remove event in the raid via the WCL v2 GraphQL API, rebuilds
each target's exact debuff timeline, and for every magic-damage hit backs out
what the damage *would have been* without the debuff(s) active on that target
at that exact moment. Sum the difference across the whole raid and you get a
real, event-driven answer instead of a rule-of-thumb multiplier.

It ships wired for TBC Classic's Shadow Weaving (ability 15258, +2%/stack up
to 5 stacks) and Misery (ability 33200, flat +5% while present), but the
approach generalizes to any debuff/damage-school combination — see
"Retargeting at a different debuff" below.

## Why this exists

The common back-of-napkin claim ("a shadow priest's debuffs are worth X% of
raid DPS") glosses over two things: debuff uptime isn't constant across a
raid, and a big chunk of a shadow priest's *real* impact never shows up on
their own damage meter — it shows up as bonus damage on every other caster's
parse instead. This toolkit answers both questions properly:

- exactly how much bonus damage the tracked debuffs added, raid-wide, fight
  by fight
- how much of that bonus was self-buff versus handed to the rest of the raid
- how a priest's *personal damage + the bonus they enabled in everyone else*
  compares to an individual warlock's own damage — the fairer, apples-to-
  apples comparison
- the same analysis run across a batch of reports (e.g. every top-10-speed
  clear of a given raid instance), so one log isn't cherry-picked

## Requirements

- Python 3.9+
- `openpyxl` (`pip install openpyxl`)
- A Warcraft Logs API v2 client ID/secret (free — register a client at
  https://www.warcraftlogs.com/api/clients/)
- For **archived reports** (anything old enough that WCL has archived it —
  which is most reports beyond a few months old), you additionally need a
  personal OAuth **user** access token from an account with an active WCL
  subscription. `client_credentials` alone cannot read event data from an
  archived report, even a public one.

## Setup

```bash
export WCL_CLIENT_ID=...
export WCL_CLIENT_SECRET=...
```

Never commit these, or the user token file below — `.gitignore` already
excludes `wcl_user_token.json` and the `data/`/`output/` directories.

### Getting a user token (only needed for archived reports)

1. In your WCL API client settings, note the client ID/secret and set a
   redirect URI (`http://localhost` works fine for a manual flow).
2. Visit:
   `https://www.warcraftlogs.com/oauth/authorize?client_id=<ID>&redirect_uri=http://localhost&response_type=code&scope=view-private-reports`
3. Log in, approve, and copy the `code` param from the URL you're redirected
   to.
4. Exchange it for a token:
   ```bash
   curl -X POST https://www.warcraftlogs.com/oauth/token \
     -u "$WCL_CLIENT_ID:$WCL_CLIENT_SECRET" \
     -d grant_type=authorization_code \
     -d code=<CODE_FROM_STEP_3> \
     -d redirect_uri=http://localhost
   ```
5. Save the response as `wcl_user_token.json` in the directory you'll run
   the scripts from (or point `WCL_USER_TOKEN_PATH` at wherever you saved
   it). It needs an `access_token` field — that's all the scripts read.

## Usage

**1. Confirm ability IDs for your patch/log** (optional, but recommended
before trusting hardcoded IDs):

```bash
cd scripts
python3 wcl_client.py list-abilities --report <REPORT_CODE>
```

**2. Pull the raid-wide attribution for one report:**

```bash
python3 raidwide_attribution.py <REPORT_CODE> "Optional label" --output-dir ../data
```

This auto-detects who applied the tracked debuffs (no need to know their
name up front), walks every fight in the report, and writes
`../data/raidwide_<REPORT_CODE>.json`.

**3. Set up `raids_config.json`** — points `build_workbook.py` at your
primary report (plus which warlocks in it to show for comparison) and,
optionally, a batch of other reports grouped by instance (e.g. a set of
top-10-speed clears). The included `raids_config.json` is a real example
from a TBC Classic Black Temple / Mount Hyjal speed-run analysis. To reuse
this for a different report or rankings page, edit it directly, or copy it
and pass `--config yourfile.json`.

**4. Run each report in your config through step 2**, then build the
workbook:

```bash
python3 build_workbook.py --data-dir ../data --config raids_config.json --output ../output/attribution.xlsx
```

**5. Recalculate.** openpyxl writes formulas but never evaluates them, so
formula cells read as blank until the file is opened and saved once — either
by hand in Excel/LibreOffice, or headless:

```bash
soffice --headless --convert-to xlsx --outdir output output/attribution.xlsx
```

The resulting workbook has three tabs: the primary report's full analysis
(all on one scrollable tab — summary, a step-by-step attribution
walkthrough, per-fight and per-player breakdowns, methodology), a "Top 10
Speed Priests" tab running the same analysis across every report in your
config, and a "Priests vs Warlocks" tab comparing each raid's priest against
its warlocks directly.

## Retargeting at a different debuff

`raidwide_attribution.py` has two constants at the top (`SW_ID`,
`MISERY_ID`) and the per-hit multiplier logic just below them. To track a
different debuff (or a different class entirely), change the ability ID(s)
and the multiplier math to match; `wcl_client.py`'s `attribute` subcommand
already takes debuff IDs/percentages/stacking behavior as CLI args if you
just want a single-fight, single-source check without touching any code.

## Known limitations

- **Physical damage is always excluded** from the raid-wide totals, since
  the default debuffs here only affect magic schools. If you retarget this
  at a debuff that also boosts physical damage, remove that filter.
- **A handful of fights can fail** on any given API pull with a transient
  connection error; the scripts skip the failed fight and keep going rather
  than aborting the whole report, and the workbook notes when this happened.
- **This computes an event-driven estimate, not ground truth.** It trusts
  WCL's logged `amount` values and school/type classifications; if WCL's own
  data has gaps (a missed event, a misclassified ability), those propagate
  through unchanged.
- **Pet/guardian/totem damage rolls up to its owner** via WCL's `petOwner`
  field, which should cover most summons, but multi-hop summon chains beyond
  a small resolution depth fall back to being counted under the pet itself.

## License

MIT - see `LICENSE`.
