#!/usr/bin/env python3
"""
Warcraft Logs v2 API client + single-fight/single-source debuff damage
attribution tool.

Answers: "how much of caster X's damage on this pull is attributable to a
given debuff (e.g. Shadow Weaving / Misery) being up on the target(s) they
hit?"

Two modes:
  1. list-abilities   - dump every ability name+ID seen in a report, so you
                        can confirm the EXACT ids for whatever debuff you
                        care about on THIS log/patch, rather than trusting
                        hardcoded IDs that may be wrong for your version.
  2. attribute        - do the actual debuff-timeline + damage cross-reference
                        and print the attributed-damage breakdown for one
                        source on one fight.

Setup:
  export WCL_CLIENT_ID=...
  export WCL_CLIENT_SECRET=...
  (keep these as env vars, never commit them or paste them into a script)

For ARCHIVED reports (anything old enough that Warcraft Logs has archived
it), client_credentials alone can't read event data - you need a personal
OAuth user token instead (authorization_code flow, requires a subscribing
WCL account). See README.md for how to obtain one. Once you have it, save
it as JSON ({"access_token": "..."}) and point WCL_USER_TOKEN_PATH at it,
or just drop it at ./wcl_user_token.json (default, and already .gitignored).

Usage:
  python3 wcl_client.py list-abilities --report <REPORT_CODE>
  python3 wcl_client.py attribute --report <REPORT_CODE> \
      --fight <FIGHT_ID> --source "WarlockName" \
      --debuff-ids 15258,33200 \
      --debuff-pcts 2,5 --debuff-stacking true,false

Report code is the string in the WCL URL: warcraftlogs.com/reports/<CODE>
Fight ID is the pull number shown on the report (1, 2, 3, ...).
"""

import argparse
import base64
import json
import os
import sys
import time
import urllib.parse
import urllib.request

TOKEN_URL = "https://www.warcraftlogs.com/oauth/token"
GRAPHQL_URL = "https://www.warcraftlogs.com/api/v2/client"
GRAPHQL_USER_URL = "https://www.warcraftlogs.com/api/v2/user"

DEFAULT_USER_TOKEN_PATH = os.environ.get("WCL_USER_TOKEN_PATH", "wcl_user_token.json")


def load_user_token(path=None):
    """Loads a previously-obtained OAuth user access token (authorization_code
    flow) - needed for archived reports, which client_credentials can't read."""
    with open(path or DEFAULT_USER_TOKEN_PATH) as f:
        return json.load(f)["access_token"]


def get_token(client_id, client_secret):
    """client_credentials flow - works for non-archived, public reports."""
    data = urllib.parse.urlencode({"grant_type": "client_credentials"}).encode()
    req = urllib.request.Request(TOKEN_URL, data=data, method="POST")
    basic = base64.b64encode(f"{client_id}:{client_secret}".encode()).decode()
    req.add_header("Authorization", f"Basic {basic}")
    req.add_header("Content-Type", "application/x-www-form-urlencoded")
    with urllib.request.urlopen(req) as resp:
        return json.load(resp)["access_token"]


def gql(token, query, variables, url=None):
    payload = json.dumps({"query": query, "variables": variables}).encode()
    req = urllib.request.Request(url or GRAPHQL_URL, data=payload, method="POST")
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Content-Type", "application/json")
    with urllib.request.urlopen(req) as resp:
        result = json.load(resp)
    if "errors" in result:
        raise RuntimeError(json.dumps(result["errors"], indent=2))
    return result["data"]


def list_abilities(token, report_code):
    query = """
    query($code: String!) {
      reportData {
        report(code: $code) {
          masterData {
            abilities { gameID name type }
          }
        }
      }
    }
    """
    data = gql(token, query, {"code": report_code})
    abilities = data["reportData"]["report"]["masterData"]["abilities"]
    print(f"Total abilities in report: {len(abilities)}\n")
    for a in sorted(abilities, key=lambda a: (a.get("name") or "")):
        print(f"  id={a['gameID']:<8} type={a.get('type')!s:<8} name={a['name']}")


def get_fight_window(token, report_code, fight_id):
    query = """
    query($code: String!) {
      reportData {
        report(code: $code) {
          fights { id startTime endTime name }
        }
      }
    }
    """
    data = gql(token, query, {"code": report_code})
    for f in data["reportData"]["report"]["fights"]:
        if f["id"] == fight_id:
            return f["startTime"], f["endTime"], f["name"]
    raise ValueError(f"Fight {fight_id} not found in report {report_code}")


def fetch_all_events(token, report_code, start, end, data_type, filter_expression, url=None, hostility_type=None):
    # hostility_type: "Enemies" or "Friendlies" - CRITICAL for dataType=Debuffs.
    # The events query defaults to Friendlies for Debuffs, so a debuff that
    # only ever lands on hostile mobs silently comes back empty unless
    # hostilityType: Enemies is passed explicitly. Confirmed by cross-checking
    # against WCL's own UI (Enemies > Debuffs > Gained By Enemy tab), which
    # showed real counts/uptime that the API returned zero for until this
    # parameter was added.
    query = """
    query($code: String!, $start: Float!, $end: Float!, $dataType: EventDataType, $filter: String, $hostility: HostilityType) {
      reportData {
        report(code: $code) {
          events(startTime: $start, endTime: $end, dataType: $dataType, filterExpression: $filter, hostilityType: $hostility, limit: 10000) {
            data
            nextPageTimestamp
          }
        }
      }
    }
    """
    events = []
    cursor = start
    while cursor is not None and cursor < end:
        data = gql(token, query, {
            "code": report_code, "start": cursor, "end": end,
            "dataType": data_type, "filter": filter_expression,
            "hostility": hostility_type,
        }, url=url)
        page = data["reportData"]["report"]["events"]
        events.extend(page["data"])
        cursor = page["nextPageTimestamp"]
        if cursor:
            time.sleep(0.2)  # be polite to the API
    return events


def build_debuff_timeline(debuff_events, debuff_ids):
    """
    Returns: timeline[targetID] = sorted list of (timestamp, spellID, stacks)
    stacks is the count AFTER this event (0 means removed).
    """
    timeline = {}
    for ev in debuff_events:
        ability_id = ev.get("abilityGameID")
        if ability_id not in debuff_ids:
            continue
        target = ev.get("targetID")
        etype = ev.get("type")
        stacks = ev.get("stack", 1)
        if etype == "removedebuff":
            stacks = 0
        timeline.setdefault(target, []).append((ev["timestamp"], ability_id, stacks))
    for target in timeline:
        timeline[target].sort()
    return timeline


def stacks_at(timeline_for_target, ability_id, timestamp):
    if not timeline_for_target:
        return 0
    current = 0
    for ts, aid, stacks in timeline_for_target:
        if ts > timestamp:
            break
        if aid == ability_id:
            current = stacks
    return current


def attribute_damage(token, report_code, fight_id, source_name, debuff_ids, debuff_pcts, debuff_stacking):
    start, end, fight_name = get_fight_window(token, report_code, fight_id)
    print(f"Analyzing fight {fight_id} ({fight_name}), {start}-{end}ms, source={source_name}\n")

    debuff_events = fetch_all_events(
        token, report_code, start, end, "Debuffs",
        f"ability.id in ({','.join(str(i) for i in debuff_ids)})",
        hostility_type="Enemies",
    )
    timelines = build_debuff_timeline(debuff_events, set(debuff_ids))
    print(f"Collected {len(debuff_events)} debuff events across {len(timelines)} targets.")

    dmg_events = fetch_all_events(
        token, report_code, start, end, "DamageDone",
        f"source.name = \"{source_name}\"",
    )
    print(f"Collected {len(dmg_events)} damage events from {source_name}.\n")

    total_actual = 0
    total_baseline = 0
    per_ability = {}

    for ev in dmg_events:
        amount = ev.get("amount", 0)
        if not amount:
            continue
        target = ev.get("targetID")
        ts = ev["timestamp"]
        multiplier = 1.0
        target_timeline = timelines.get(target, [])
        for did, pct, stacking in zip(debuff_ids, debuff_pcts, debuff_stacking):
            stacks = stacks_at(target_timeline, did, ts)
            if stacks <= 0:
                continue
            if stacking:
                multiplier *= (1 + (pct / 100.0) * stacks)
            else:
                multiplier *= (1 + (pct / 100.0))  # flat, presence-only (e.g. Misery)

        baseline = amount / multiplier if multiplier else amount
        bonus = amount - baseline
        total_actual += amount
        total_baseline += baseline

        ability_name = ev.get("ability", {}).get("name", "Unknown") if isinstance(ev.get("ability"), dict) else str(ev.get("abilityGameID"))
        bucket = per_ability.setdefault(ability_name, {"actual": 0, "baseline": 0})
        bucket["actual"] += amount
        bucket["baseline"] += baseline

    print("=== Per-ability breakdown ===")
    for name, v in sorted(per_ability.items(), key=lambda kv: -kv[1]["actual"]):
        bonus = v["actual"] - v["baseline"]
        pct = (bonus / v["actual"] * 100) if v["actual"] else 0
        print(f"  {name:<28} actual={v['actual']:>10.0f}  baseline={v['baseline']:>10.0f}  debuff-bonus={bonus:>9.0f} ({pct:.1f}%)")

    total_bonus = total_actual - total_baseline
    print(f"\n=== Totals for {source_name} on this pull ===")
    print(f"  Actual (logged) damage:     {total_actual:,.0f}")
    print(f"  Estimated baseline (no debuffs): {total_baseline:,.0f}")
    print(f"  Attributable to the tracked debuff(s): {total_bonus:,.0f} ({(total_bonus/total_actual*100 if total_actual else 0):.1f}% of their total)")


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="cmd", required=True)

    p1 = sub.add_parser("list-abilities")
    p1.add_argument("--report", required=True)

    p2 = sub.add_parser("attribute")
    p2.add_argument("--report", required=True)
    p2.add_argument("--fight", required=True, type=int)
    p2.add_argument("--source", required=True, help="Exact in-game name of the caster to analyze")
    p2.add_argument("--debuff-ids", required=True, help="Comma-separated ability IDs")
    p2.add_argument("--debuff-pcts", required=True, help="Comma-separated percent-per-stack (or flat %) matching debuff-ids order")
    p2.add_argument("--debuff-stacking", required=True, help="Comma-separated true/false matching debuff-ids order (true=stacking per-stack %, false=flat presence-only %)")

    args = parser.parse_args()

    client_id = os.environ.get("WCL_CLIENT_ID")
    client_secret = os.environ.get("WCL_CLIENT_SECRET")
    if not client_id or not client_secret:
        print("Set WCL_CLIENT_ID and WCL_CLIENT_SECRET environment variables first.", file=sys.stderr)
        sys.exit(1)

    token = get_token(client_id, client_secret)

    if args.cmd == "list-abilities":
        list_abilities(token, args.report)
    elif args.cmd == "attribute":
        ids = [int(x) for x in args.debuff_ids.split(",")]
        pcts = [float(x) for x in args.debuff_pcts.split(",")]
        stacking = [x.strip().lower() == "true" for x in args.debuff_stacking.split(",")]
        attribute_damage(token, args.report, args.fight, args.source, ids, pcts, stacking)


if __name__ == "__main__":
    main()
