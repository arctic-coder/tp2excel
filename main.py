#!/usr/bin/env python3
"""
tp2excel — export TrainingPeaks training plans to .xlsx files.

Usage:
    python main.py                     # interactive: list plans, pick one
    python main.py --plan-id <id>      # export a specific plan by ID
    python main.py --all               # export all plans
"""

import argparse
import os
import sys
from datetime import date

from dotenv import load_dotenv

from tp_client import TrainingPeaksClient, parse_workout_row
from excel_writer import write_plan


def main():
    load_dotenv()

    args = _parse_args()
    auth_cookie = _get_cookie()
    plan_days_shift = int(os.getenv("TP_PLAN_DAYS_SHIFT", "0"))

    print("Connecting to TrainingPeaks...")
    tp = TrainingPeaksClient(auth_cookie, plan_days_shift)

    try:
        all_plans = tp.get_plans()
    except Exception:
        print("Could not connect. Your cookie may have expired.")
        auth_cookie = _prompt_cookie()
        tp = TrainingPeaksClient(auth_cookie, plan_days_shift)
        all_plans = tp.get_plans()

    if not all_plans:
        print("No plans found.")
        return

    if args.all:
        selected = all_plans
    elif args.plan_id:
        selected = [p for p in all_plans if str(p["planId"]) == str(args.plan_id)]
        if not selected:
            print(f"Plan ID '{args.plan_id}' not found.")
            _print_plans(all_plans)
            sys.exit(1)
    else:
        selected = _interactive_select(all_plans)

    for plan in selected:
        _export_plan(tp, plan)

    input("\nDone. Press Enter to exit.")


def _get_cookie() -> str:
    # .env takes priority when running as a developer
    env_cookie = os.getenv("TP_AUTH_COOKIE")
    if env_cookie:
        return env_cookie
    return _prompt_cookie()


def _prompt_cookie() -> str:
    print()
    print("Your TrainingPeaks auth cookie is needed.")
    print()
    print("How to get it:")
    print("  1. Open https://www.trainingpeaks.com in Chrome and log in")
    print("  2. Press F12 -> Application -> Cookies -> https://www.trainingpeaks.com")
    print("  3. Find 'Production_tpAuth' and copy its value")
    print()
    value = input("Paste cookie value: ").strip()
    if not value.startswith("Production_tpAuth="):
        value = f"Production_tpAuth={value}"
    return value


def _export_plan(tp: TrainingPeaksClient, plan: dict):
    plan_id = str(plan["planId"])
    plan_name = plan.get("title", f"Plan-{plan_id}")
    workout_count = plan.get("workoutCount", "?")

    print(f"\nExporting: {plan_name!r}  ({workout_count} workouts)")
    print("  Fetching workouts from TrainingPeaks...")

    workouts = tp.get_workouts_from_plan(plan_id)
    if not workouts:
        print("  No workouts found — skipping.")
        return

    workouts.sort(key=lambda w: w.get("workoutDay", ""))
    plan_start = date.fromisoformat(workouts[0]["workoutDay"][:10])
    rows = [parse_workout_row(w, plan_start) for w in workouts]

    print("  Writing Excel file...")
    path = write_plan(plan_name, rows)
    print(f"  Saved: {path.resolve()}")


def _interactive_select(plans: list[dict]) -> list[dict]:
    print(f"\nFound {len(plans)} plan(s):\n")
    for i, p in enumerate(plans, 1):
        print(f"  {i:3}.  [{p['planId']}]  {p.get('title', 'Untitled')}  "
              f"({p.get('workoutCount', '?')} workouts)")

    print()
    raw = input("Enter plan number(s) to export (e.g. 1  or  1,3,5  or  all): ").strip()

    if raw.lower() == "all":
        return plans

    indices = []
    for part in raw.replace(",", " ").split():
        try:
            idx = int(part)
            if 1 <= idx <= len(plans):
                indices.append(idx - 1)
            else:
                print(f"  Ignoring out-of-range number: {idx}")
        except ValueError:
            print(f"  Ignoring invalid input: {part!r}")

    if not indices:
        print("Nothing selected.")
        sys.exit(0)

    return [plans[i] for i in indices]


def _print_plans(plans: list[dict]):
    print("\nAvailable plans:")
    for p in plans:
        print(f"  [{p['planId']}]  {p.get('title', 'Untitled')}")


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Export TrainingPeaks plans to Excel")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--plan-id", metavar="ID", help="Export a specific plan by ID")
    group.add_argument("--all", action="store_true", help="Export all plans")
    return parser.parse_args()


if __name__ == "__main__":
    main()
