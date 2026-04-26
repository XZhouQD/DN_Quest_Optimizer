"""Generate a deterministic test case and run the optimizer end-to-end.

Input files are written to tests/input/<member>_票.xlsx and
tests/input/<member>_委托.xlsx. The resulting schedule is written to
tests/output/schedule.xlsx.

Test-case rules (deterministic — seeded):
* 10 characters per member.
* Only the first character of each member has tickets:
  a random integer in [5, 12] total dedicated tickets distributed across
  random targets, plus 2 '选择' wildcard tickets. All other characters have
  zero tickets.
* Each character (including ticket-less ones) has a random number (0-3) of
  weekly quests picked uniformly at random across the 13 targets.

Run:  python -m tests.generate_test_case
"""
from __future__ import annotations

import random
import shutil
import sys
from pathlib import Path

import pandas as pd

from src.config import (
    TARGETS, WILDCARD, CHARACTER_COL,
    TICKET_SUFFIX, QUEST_SUFFIX,
    CHAR_SHEET, QUEST_SHEET,
)
from src.optimize import solve, _load_inputs
from src.schedule import finalize_schedule, write_schedule_with_quests

SEED = 42
CHARS_PER_MEMBER = 10
TICKETS_MIN, TICKETS_MAX = 5, 12
WILDCARD_COUNT = 2
QUESTS_MIN, QUESTS_MAX = 0, 3
TEST_MEMBERS = ["小C", "暗部", "桃核", "蹦蹦"]


def _distribute_tickets(total: int, n_buckets: int, rng: random.Random) -> list[int]:
    """Distribute `total` indivisible tickets into `n_buckets` randomly."""
    out = [0] * n_buckets
    for _ in range(total):
        out[rng.randrange(n_buckets)] += 1
    return out


def build_inputs(input_dir: Path, seed: int = SEED) -> None:
    input_dir.mkdir(parents=True, exist_ok=True)
    rng = random.Random(seed)

    for m in TEST_MEMBERS:
        chars = [f"{m}_C{i+1}" for i in range(CHARS_PER_MEMBER)]

        ticket_rows = []
        quest_rows = []
        for idx, c in enumerate(chars):
            if idx == 0:
                total_tix = rng.randint(TICKETS_MIN, TICKETS_MAX)
                per_target = _distribute_tickets(total_tix, len(TARGETS), rng)
                wild = WILDCARD_COUNT
            else:
                per_target = [0] * len(TARGETS)
                wild = 0

            tr = {CHARACTER_COL: c}
            for t, v in zip(TARGETS, per_target):
                tr[t] = v
            tr[WILDCARD] = wild
            ticket_rows.append(tr)

            # Weekly quests: 0..3 random distinct targets marked 1
            n_quests = rng.randint(QUESTS_MIN, QUESTS_MAX)
            picked = set(rng.sample(TARGETS, n_quests)) if n_quests else set()
            qr = {CHARACTER_COL: c}
            for t in TARGETS:
                qr[t] = 1 if t in picked else 0
            quest_rows.append(qr)

        pd.DataFrame(ticket_rows, columns=[CHARACTER_COL] + list(TARGETS) + [WILDCARD]).to_excel(
            input_dir / f"{m}{TICKET_SUFFIX}.xlsx",
            sheet_name=CHAR_SHEET, index=False,
        )
        pd.DataFrame(quest_rows, columns=[CHARACTER_COL] + list(TARGETS)).to_excel(
            input_dir / f"{m}{QUEST_SUFFIX}.xlsx",
            sheet_name=QUEST_SHEET, index=False,
        )


def run_test(seed: int = SEED, time_limit: int = 60) -> dict:
    root = Path(__file__).resolve().parent
    input_dir = root / "input"
    output_dir = root / "output"

    # Fresh inputs every run so results are reproducible from the seed.
    if input_dir.exists():
        shutil.rmtree(input_dir)
    if output_dir.exists():
        shutil.rmtree(output_dir)
    build_inputs(input_dir, seed=seed)
    output_dir.mkdir(parents=True, exist_ok=True)

    battles, members = solve(input_dir, time_limit_sec=time_limit)
    _, chars_by_member, _, _, quests = _load_inputs(input_dir)
    battles = finalize_schedule(battles, members, quests, chars_by_member)
    out_path = output_dir / "schedule.xlsx"
    write_schedule_with_quests(battles, members, quests, out_path)

    total_tickets_used = len(battles)
    # Count each weekly quest at most once.
    credited: set = set()
    per_member_q = {m: 0 for m in members}
    for b in battles:
        for m, c in b.participants.items():
            key = (m, c, b.target)
            if quests.get(key, 0) == 1 and key not in credited:
                credited.add(key)
                per_member_q[m] += 1
    total_quests = sum(per_member_q.values())

    result = {
        "seed": seed,
        "members": members,
        "input_dir": str(input_dir),
        "output_file": str(out_path),
        "battles": total_tickets_used,
        "total_quests": total_quests,
        "per_member_quests": per_member_q,
    }
    return result


def main() -> None:
    res = run_test()
    print(f"seed:            {res['seed']}")
    print(f"members:         {res['members']}")
    print(f"input_dir:       {res['input_dir']}")
    print(f"output_file:     {res['output_file']}")
    print(f"battles:         {res['battles']}")
    print(f"total_quests:    {res['total_quests']}")
    print(f"per_member_qsts: {res['per_member_quests']}")


if __name__ == "__main__":
    # Allow `python tests/generate_test_case.py` as well as `python -m tests....`
    if __package__ is None:
        sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    main()
