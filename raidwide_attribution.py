#!/usr/bin/env python3
"""
Raid-wide debuff damage-attribution pass over a single Warcraft Logs report.

Given a report code, this:
  - auto-detects whoever applied the tracked debuffs (by checking who the
    SOURCE of each debuff-apply event was - no need to know their name or
    class up front)
  - builds a per-target debuff timeline (stacks over time, from real
    aura-apply/refresh/remove events, not assumed uptime)
  - pulls every DamageDone event in the raid, across every fight (kills and
    trash/wipes alike)
  - excludes physical-school damage (the default debuffs here - Shadow
    Weaving / Misery - only affect magic schools; adjust school filtering if
    you retarget this at a debuff that also affects physical damage)
  - for every remaining magic-damage hit, backs out what the damage would
    have been WITHOUT the debuff(s) active on that target at that moment, and
    sums the difference
  - rolls pet/guardian/totem damage up under the owning player (via WCL's
    petOwner field), so a priest's Shadowfiend or a warlock's demon counts
    toward their personal total instead of vanishing into an untracked bucket
  - writes the full per-fight / per-source breakdown to a JSON file

By default this is wired for TBC Classic Shadow Priest debuffs (Shadow
Weaving, ability 15258, +2%/stack up to 5 stacks; Misery, ability 33200,
flat +5% while present). To retarget this at a different debuff, change
TRACKED_DEBUFFS below - each entry needs an ability ID, a per-hit multiplier
function, and whether it's the kind of debuff whose SOURCE should be treated
as "the buffer" for auto-detection purposes.

Usage:
  python3 raidwide_attribution.py <REPORT_CODE> [label] [--output-dir DIR]

Requires WCL_USER_TOKEN_PATH (or ./wcl_user_token.json) - see wcl_client.py
and README.md for how to obtain an OAuth user token. Archived reports need
this even for otherwise-public logs; client_credentials alone can't read
their event data.
"""
import argparse
import os
import sys
import time
import json

sys.path.insert(0, os.path.dirname(__file__))
import wcl_client as w

# ---- adjust these two to retarget the analysis at a different debuff pair ----
SW_ID = 15258      # Shadow Weaving: +2% shadow damage taken per stack, up to 5 stacks
MISERY_ID = 33200  # Misery: flat +5% all-magic-school damage taken while present


def build_ability_school_map(token, report_code):
    query = """
    query($code: String!) {
      reportData {
        report(code: $code) {
          masterData { abilities { gameID type } }
        }
      }
    }
    """
    data = w.gql(token, query, {"code": report_code}, url=w.GRAPHQL_USER_URL)
    abilities = data["reportData"]["report"]["masterData"]["abilities"]
    return {a["gameID"]: a["type"] for a in abilities}


def get_fights(token, report_code):
    query = """
    query($code: String!) {
      reportData {
        report(code: $code) {
          fights { id name startTime endTime kill }
        }
      }
    }
    """
    data = w.gql(token, query, {"code": report_code}, url=w.GRAPHQL_USER_URL)
    return data["reportData"]["report"]["fights"]


def get_actors(token, report_code):
    query = """
    query($code: String!) { reportData { report(code: $code) { masterData { actors { id name type subType petOwner } } } } }
    """
    data = w.gql(token, query, {"code": report_code}, url=w.GRAPHQL_USER_URL)
    actors = data["reportData"]["report"]["masterData"]["actors"]
    id_to_name = {a["id"]: a["name"] for a in actors}
    id_to_class = {a["id"]: a.get("subType") for a in actors}

    # Pets/guardians/totems are separate actors (type "Pet"). Map each one to
    # its owning PLAYER's actor id so their damage rolls up under the owner
    # instead of vanishing into an untracked pet-only bucket. petOwner can
    # itself point at another pet in multi-summon chains, so resolve
    # transitively (bounded) up to a real Player actor.
    raw_pet_owner = {a["id"]: a.get("petOwner") for a in actors if a.get("type") == "Pet" and a.get("petOwner")}
    actor_type = {a["id"]: a.get("type") for a in actors}

    def resolve_owner(aid, depth=0):
        if depth > 5:
            return aid
        owner = raw_pet_owner.get(aid)
        if owner is None:
            return aid
        if actor_type.get(owner) == "Pet":
            return resolve_owner(owner, depth + 1)
        return owner

    pet_to_owner = {pid: resolve_owner(pid) for pid in raw_pet_owner}
    return id_to_name, id_to_class, pet_to_owner


def stacks_at(events_sorted, ts):
    current = 0
    for e_ts, _aid, stacks in events_sorted:
        if e_ts > ts:
            break
        current = stacks
    return current


def build_timelines(debuff_events):
    """Returns (timelines, priest_source_ids). timelines[targetID][abilityID]
    is a sorted list of (ts, abilityID, stacks). priest_source_ids is every
    actor ID seen applying/refreshing the tracked debuffs - i.e. auto-detected
    "buffer(s)" for this report, no name/class lookup required up front."""
    timelines = {}
    priest_source_ids = set()
    for ev in debuff_events:
        aid = ev.get("abilityGameID")
        if aid not in (SW_ID, MISERY_ID):
            continue
        target = ev.get("targetID")
        etype = ev.get("type")
        stacks = ev.get("stack", 1)
        if etype == "removedebuff":
            stacks = 0
        else:
            src = ev.get("sourceID")
            if src:
                priest_source_ids.add(src)
        timelines.setdefault(target, {}).setdefault(aid, []).append((ev["timestamp"], aid, stacks))
    for target in timelines:
        for aid in timelines[target]:
            timelines[target][aid].sort()
    return timelines, priest_source_ids


def run(report, label, output_dir):
    token = w.load_user_token()

    print(f"[{report}] Building ability -> school-type map...")
    school_map = build_ability_school_map(token, report)

    print(f"[{report}] Fetching fight list...")
    fights = get_fights(token, report)
    print(f"[{report}] {len(fights)} fights found.")

    id_to_name, id_to_class, pet_to_owner = get_actors(token, report)

    grand_actual = 0.0
    grand_baseline = 0.0
    grand_all_damage = 0.0
    per_fight_results = []
    per_source_bonus = {}
    all_priest_source_ids = set()
    skipped_unclassified = 0
    total_magic_events = 0

    for f in fights:
        fid, name, start, end = f["id"], f["name"], f["startTime"], f["endTime"]
        is_kill = f.get("kill")
        try:
            debuff_events = w.fetch_all_events(
                token, report, start, end, "Debuffs",
                f"ability.id in ({SW_ID},{MISERY_ID})",
                url=w.GRAPHQL_USER_URL,
                hostility_type="Enemies",
            )
            timelines, priest_ids = build_timelines(debuff_events)
            all_priest_source_ids |= priest_ids

            dmg_events = w.fetch_all_events(
                token, report, start, end, "DamageDone", None,
                url=w.GRAPHQL_USER_URL,
            )
        except Exception as e:
            print(f"[{report}] fight {fid} ({name}) FAILED: {e}")
            continue

        fight_actual = 0.0
        fight_baseline = 0.0
        fight_magic_events = 0
        fight_all_damage = 0.0

        for ev in dmg_events:
            amount = ev.get("amount", 0)
            if not amount:
                continue
            fight_all_damage += amount
            aid = ev.get("abilityGameID")
            school_type = school_map.get(aid)
            if school_type is None:
                skipped_unclassified += 1
                continue
            # NOTE: WCL's API returns `type` as a STRING ("1"), not an int -
            # comparing against the Python int 1 is a classic silent bug that
            # lets physical damage leak into the "magic damage" total. Always
            # compare against the string.
            if school_type == "1":
                continue  # physical - the tracked debuffs don't apply here

            fight_magic_events += 1
            target = ev.get("targetID")
            ts = ev["timestamp"]
            target_lines = timelines.get(target, {})

            sw_stacks = stacks_at(target_lines.get(SW_ID, []), ts)
            misery_stacks = stacks_at(target_lines.get(MISERY_ID, []), ts)

            multiplier = 1.0
            if sw_stacks > 0:
                multiplier *= (1 + 0.02 * sw_stacks)
            if misery_stacks > 0:
                multiplier *= 1.05

            baseline = amount / multiplier if multiplier else amount
            bonus = amount - baseline

            fight_actual += amount
            fight_baseline += baseline

            src = ev.get("sourceID")
            src = pet_to_owner.get(src, src)  # roll pets/guardians up under their owner
            b = per_source_bonus.setdefault(src, {"actual": 0.0, "bonus": 0.0})
            b["actual"] += amount
            b["bonus"] += bonus

        fight_bonus = fight_actual - fight_baseline
        per_fight_results.append((fid, name, fight_actual, fight_bonus, fight_all_damage, is_kill))
        grand_actual += fight_actual
        grand_baseline += fight_baseline
        grand_all_damage += fight_all_damage
        total_magic_events += fight_magic_events
        time.sleep(0.2)

    grand_bonus = grand_actual - grand_baseline

    priest_names = sorted({id_to_name.get(pid, f"id={pid}") for pid in all_priest_source_ids})
    priest_rows = [
        {
            "id": pid,
            "name": id_to_name.get(pid, f"id={pid}"),
            "actual": per_source_bonus.get(pid, {"actual": 0.0})["actual"],
            "bonus": per_source_bonus.get(pid, {"bonus": 0.0})["bonus"],
        }
        for pid in all_priest_source_ids
    ]

    all_sources = [
        {
            "id": src,
            "name": id_to_name.get(src, f"id={src}"),
            "class": id_to_class.get(src),
            "actual": v["actual"],
            "bonus": v["bonus"],
        }
        for src, v in per_source_bonus.items()
    ]
    warlock_rows = sorted([s for s in all_sources if s["class"] == "Warlock"], key=lambda s: -s["actual"])

    dump = {
        "report_code": report,
        "label": label,
        "priest_names": priest_names,
        "priests": priest_rows,
        "warlocks": warlock_rows,
        "sources": all_sources,
        "grand_actual_magic": grand_actual,
        "grand_baseline_magic": grand_baseline,
        "grand_bonus": grand_bonus,
        "grand_all_damage": grand_all_damage,
        "total_magic_events": total_magic_events,
        "skipped_unclassified": skipped_unclassified,
        "num_fights": len(fights),
        "fights": [
            {"id": fid, "name": name, "actual": actual, "bonus": bonus, "all_damage": all_dmg, "kill": is_kill}
            for fid, name, actual, bonus, all_dmg, is_kill in per_fight_results
        ],
    }
    os.makedirs(output_dir, exist_ok=True)
    outpath = os.path.join(output_dir, f"raidwide_{report}.json")
    with open(outpath, "w") as fh:
        json.dump(dump, fh, indent=2)
    print(f"[{report}] DONE. priests={priest_names} grand_bonus={grand_bonus:,.0f} grand_actual={grand_actual:,.0f}")
    print(f"[{report}] Wrote {outpath}")
    return dump


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("report", help="Warcraft Logs report code")
    parser.add_argument("label", nargs="?", default=None, help="Human-readable label for this run (defaults to the report code)")
    parser.add_argument("--output-dir", default="data", help="Where to write raidwide_<code>.json (default: ./data)")
    args = parser.parse_args()
    run(args.report, args.label or args.report, args.output_dir)


if __name__ == "__main__":
    main()
