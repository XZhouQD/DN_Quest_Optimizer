"""CLI: python -m src.main --input-dir templates --out schedule.xlsx"""
from __future__ import annotations

import argparse
from pathlib import Path

from .optimize import solve, _load_inputs
from .schedule import order_battles, write_schedule_with_quests


def main() -> None:
    ap = argparse.ArgumentParser(description="Team weekly quest scheduler")
    ap.add_argument("--input-dir", default="input",
                    help="Directory containing per-member _票.xlsx and _委托.xlsx files.")
    ap.add_argument("--out", default="schedule.xlsx",
                    help="Output schedule xlsx path.")
    ap.add_argument("--time-limit", type=int, default=60,
                    help="Solver time limit (seconds).")
    ap.add_argument("--verbose", action="store_true")
    args = ap.parse_args()

    battles, members = solve(
        args.input_dir,
        time_limit_sec=args.time_limit, verbose=args.verbose,
    )
    battles = order_battles(battles)

    # Reload quest dict to attribute per-member quest credit in the summary.
    _, _, _, _, quests = _load_inputs(Path(args.input_dir))

    write_schedule_with_quests(battles, members, quests, args.out)

    total_q = sum(b.completed for b in battles)
    print(f"Scheduled {len(battles)} battles; {total_q} quests completed.")
    print(f"Output written to: {Path(args.out).resolve()}")


if __name__ == "__main__":
    main()
