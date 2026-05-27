"""Regression checks for battle-count minimization objective tie-breaker.

Run:
    python -m tests.test_battle_count_tiebreak
"""
from __future__ import annotations

import shutil
from pathlib import Path

import pandas as pd

from src.config import (
    CHAR_SHEET,
    QUEST_SHEET,
    TARGETS,
    WILDCARD,
    CHARACTER_COL,
    TICKET_SUFFIX,
    QUEST_SUFFIX,
)
from src.optimize import _load_inputs, solve
from src.schedule import finalize_schedule


MEMBERS = ["Alpha", "Bravo", "Charlie", "Delta"]


def _zero_row(character: str) -> dict:
    row = {CHARACTER_COL: character}
    for target in TARGETS:
        row[target] = 0
    row[WILDCARD] = 0
    return row


def _quest_row(character: str, targets: set[str]) -> dict:
    row = {CHARACTER_COL: character}
    for target in TARGETS:
        row[target] = 1 if target in targets else 0
    return row


def _build_input(input_dir: Path, extra_wildcards: int) -> None:
    input_dir.mkdir(parents=True, exist_ok=True)
    for member in MEMBERS:
        chars = [f"{member}_A", f"{member}_B"]
        ticket_rows = [_zero_row(chars[0]), _zero_row(chars[1])]
        quest_rows = [
            _quest_row(chars[0], {"狮蝎"}),
            _quest_row(chars[1], {"海龙"}),
        ]

        # Two battles are sufficient: one 狮蝎 and one 海龙, each completing
        # four quests. Extra wildcard tickets should never make the solver use
        # more battles for the same objective score.
        if member == "Alpha":
            ticket_rows[0]["狮蝎"] = 1
            ticket_rows[0][WILDCARD] = extra_wildcards
        if member == "Bravo":
            ticket_rows[1]["海龙"] = 1

        pd.DataFrame(ticket_rows, columns=[CHARACTER_COL] + list(TARGETS) + [WILDCARD]).to_excel(
            input_dir / f"{member}{TICKET_SUFFIX}.xlsx",
            sheet_name=CHAR_SHEET,
            index=False,
        )
        pd.DataFrame(quest_rows, columns=[CHARACTER_COL] + list(TARGETS)).to_excel(
            input_dir / f"{member}{QUEST_SUFFIX}.xlsx",
            sheet_name=QUEST_SHEET,
            index=False,
        )


def _run(input_dir: Path) -> tuple[int, int]:
    battles, members = solve(input_dir, time_limit_sec=30)
    _, chars_by_member, _, _, quests = _load_inputs(input_dir)
    battles = finalize_schedule(battles, members, quests, chars_by_member)
    credited = set()
    for battle in battles:
        for member, character in battle.participants.items():
            key = (member, character, battle.target)
            if quests.get(key, 0) == 1:
                credited.add(key)
    return len(battles), len(credited)


def main() -> None:
    root = Path(__file__).resolve().parent
    base_dir = root / "tmp_tiebreak_base"
    extra_dir = root / "tmp_tiebreak_extra"
    for path in (base_dir, extra_dir):
        if path.exists():
            shutil.rmtree(path)

    _build_input(base_dir, extra_wildcards=0)
    _build_input(extra_dir, extra_wildcards=3)

    base_battles, base_quests = _run(base_dir)
    extra_battles, extra_quests = _run(extra_dir)

    assert base_quests == 8
    assert extra_quests == base_quests
    assert base_battles == 2
    assert extra_battles == 2, (
        f"extra wildcard tickets should not add same-quest battles: got {extra_battles}"
    )

    print("[OK] extra wildcard tickets do not increase battle count for same quest result.")


if __name__ == "__main__":
    main()
