"""CLI: python -m src.main --input-dir templates --out schedule.xlsx"""
from __future__ import annotations

import argparse
from pathlib import Path

from .optimize import solve, _load_inputs
from .schedule import finalize_schedule, write_schedule_with_quests


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

    # Reload quest dict to attribute per-member quest credit in the summary.
    _, chars_by_member, _, _, quests = _load_inputs(Path(args.input_dir))

    # Drop redundant battles, order, and reassign free slots (run to fixpoint).
    battles = finalize_schedule(battles, members, quests, chars_by_member)

    write_schedule_with_quests(battles, members, quests, args.out)

    # Count each weekly quest at most once.
    credited: set = set()
    total_q = 0
    for b in battles:
        for m, c in b.participants.items():
            key = (m, c, b.target)
            if quests.get(key, 0) == 1 and key not in credited:
                credited.add(key)
                total_q += 1
    print(f"Scheduled {len(battles)} battles; {total_q} quests completed.")
    print(f"Output written to: {Path(args.out).resolve()}")


if __name__ == "__main__":
    main()
