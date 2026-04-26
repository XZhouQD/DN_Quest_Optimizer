"""Validate a schedule against its input: simulate each battle and check invariants.

Usage:
    python -m tests.validate_schedule --input-dir input_260422 --schedule 260422.xlsx
    python -m tests.validate_schedule            # defaults: tests/input  + tests/output/schedule.xlsx

Checks for every battle row, in order:
  1. Target is one of the known TARGETS.
  2. Team size matches the effective TEAM_SIZE for the loaded input members
      (normally 4, but 3 if only 3 members are present; 双生 remains 2).
  3. Each participant is one distinct member's real character (exists in that
     member's ticket file).
  4. For non-双生 battles: every member contributes exactly one character.
  5. ticket_source is among the participants.
  6. ticket_kind is either the target name or the wildcard (选择).
     选择 is valid only for non-双生 targets.
  7. The ticket source has a ticket of that kind left — decrement its stock.

After walking all rows, prints a final summary (quests completed per member,
totals, remaining ticket balances) so you can compare to the solver output.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd

from src.config import (
    TARGETS, TEAM_SIZE, WILDCARD, CHARACTER_COL,
    TICKET_SUFFIX, QUEST_SUFFIX, CHAR_SHEET, QUEST_SHEET,
    SCHEDULE_SHEET,
)

SITOUT_MARK = "—"


def _effective_team_size(target: str, member_count: int) -> int:
    return min(TEAM_SIZE[target], member_count)


def _load_input(input_dir: Path) -> Tuple[
    List[str],
    Dict[str, List[str]],
    Dict[Tuple[str, str, str], int],
    Dict[Tuple[str, str], int],
    Dict[Tuple[str, str, str], int],
]:
    """Return (members, chars_by_member, tickets, wild_tickets, quests)."""
    members: List[str] = []
    chars_by_member: Dict[str, List[str]] = {}
    tickets: Dict[Tuple[str, str, str], int] = {}
    wild_tickets: Dict[Tuple[str, str], int] = {}
    quests: Dict[Tuple[str, str, str], int] = {}

    ticket_members: set[str] = set()
    quest_members: set[str] = set()
    for p in input_dir.glob(f"*{TICKET_SUFFIX}.xlsx"):
        ticket_members.add(p.stem[: -len(TICKET_SUFFIX)])
    for p in input_dir.glob(f"*{QUEST_SUFFIX}.xlsx"):
        quest_members.add(p.stem[: -len(QUEST_SUFFIX)])

    for m in sorted(ticket_members & quest_members):
        tp = input_dir / f"{m}{TICKET_SUFFIX}.xlsx"
        qp = input_dir / f"{m}{QUEST_SUFFIX}.xlsx"
        ch = pd.read_excel(tp, sheet_name=CHAR_SHEET).dropna(subset=[CHARACTER_COL])
        qu = pd.read_excel(qp, sheet_name=QUEST_SHEET).dropna(subset=[CHARACTER_COL])
        ch[CHARACTER_COL] = ch[CHARACTER_COL].astype(str).str.strip()
        qu[CHARACTER_COL] = qu[CHARACTER_COL].astype(str).str.strip()
        ch = ch[ch[CHARACTER_COL] != ""]
        qu = qu[qu[CHARACTER_COL] != ""]
        if ch.empty:
            continue
        for col in list(TARGETS) + [WILDCARD]:
            ch[col] = pd.to_numeric(ch[col], errors="coerce").fillna(0).astype(int).clip(lower=0)
        for col in TARGETS:
            qu[col] = pd.to_numeric(qu[col], errors="coerce").fillna(0).astype(int).clip(lower=0, upper=1)

        members.append(m)
        chars_by_member[m] = ch[CHARACTER_COL].tolist()

        qu_idx = qu.set_index(CHARACTER_COL) if not qu.empty else None
        for _, r in ch.iterrows():
            c = r[CHARACTER_COL]
            wild_tickets[(m, c)] = int(r[WILDCARD])
            for t in TARGETS:
                tickets[(m, c, t)] = int(r[t])
                if qu_idx is not None and c in qu_idx.index:
                    quests[(m, c, t)] = int(qu_idx.loc[c, t])
                else:
                    quests[(m, c, t)] = 0
    return members, chars_by_member, tickets, wild_tickets, quests


def validate(input_dir: Path, schedule_path: Path) -> bool:
    members, chars_by_member, tickets, wild_tickets, quests = _load_input(input_dir)
    if not members:
        print(f"[fail] no member input files found in {input_dir}")
        return False

    known_chars = {m: set(chars_by_member[m]) for m in members}

    sched = pd.read_excel(schedule_path, sheet_name=SCHEDULE_SHEET)
    required_cols = {"order", "target", "ticket_kind", "ticket_source", "quests_completed"} | set(members)
    missing = required_cols - set(sched.columns)
    if missing:
        print(f"[fail] schedule missing columns: {sorted(missing)}")
        return False

    errors: List[str] = []
    credited_quests: set[Tuple[str, str, str]] = set()
    per_member_battles: Dict[str, int] = {m: 0 for m in members}
    per_member_quests: Dict[str, int] = {m: 0 for m in members}
    kind_counts = {"dedicated": 0, "wildcard": 0}

    for _, row in sched.iterrows():
        order = int(row["order"])
        target = str(row["target"]).strip()
        kind = str(row["ticket_kind"]).strip()
        src_raw = str(row["ticket_source"]).strip()

        def err(msg: str) -> None:
            errors.append(f"[battle #{order} {target}] {msg}")

        # 1. Target validity
        if target not in TARGETS:
            err(f"unknown target '{target}'")
            continue

        k = _effective_team_size(target, len(members))

        # 2+3+4. Participants
        participants: Dict[str, str] = {}
        for m in members:
            cell = row.get(m)
            if cell is None or (isinstance(cell, float) and pd.isna(cell)):
                val = SITOUT_MARK
            else:
                val = str(cell).strip()
            if val == SITOUT_MARK or val == "":
                continue
            if val not in known_chars[m]:
                err(f"member '{m}' uses unknown character '{val}'")
                continue
            participants[m] = val

        if len(participants) != k:
            err(f"team size {len(participants)} != expected {k}")
        if target != "双生" and len(participants) < k:
            err(f"non-双生 battle must have one character from every member")

        # 5. ticket_source parsing
        if ":" not in src_raw:
            err(f"malformed ticket_source '{src_raw}'")
            continue
        src_m, src_c = src_raw.split(":", 1)
        src_m, src_c = src_m.strip(), src_c.strip()
        if src_m not in members:
            err(f"ticket_source member '{src_m}' unknown")
            continue
        if src_c not in known_chars[src_m]:
            err(f"ticket_source character '{src_c}' not in {src_m}'s roster")
            continue
        if participants.get(src_m) != src_c:
            err(f"ticket_source {src_m}:{src_c} is not a participant "
                f"(participant was '{participants.get(src_m)}')")

        # 6+7. Ticket kind and stock
        if kind == WILDCARD:
            if target == "双生":
                err("双生 cannot consume a 选择 ticket")
            if wild_tickets.get((src_m, src_c), 0) <= 0:
                err(f"{src_m}:{src_c} has no 选择 tickets left "
                    f"(stock={wild_tickets.get((src_m, src_c), 0)})")
            else:
                wild_tickets[(src_m, src_c)] -= 1
                kind_counts["wildcard"] += 1
        elif kind == target:
            if tickets.get((src_m, src_c, target), 0) <= 0:
                err(f"{src_m}:{src_c} has no '{target}' tickets left "
                    f"(stock={tickets.get((src_m, src_c, target), 0)})")
            else:
                tickets[(src_m, src_c, target)] -= 1
                kind_counts["dedicated"] += 1
        else:
            err(f"ticket_kind '{kind}' must be either '{target}' or '{WILDCARD}'")

        # Tally quest credits (once per (member, character, target))
        for m, c in participants.items():
            per_member_battles[m] += 1
            key = (m, c, target)
            if quests.get(key, 0) == 1 and key not in credited_quests:
                credited_quests.add(key)
                per_member_quests[m] += 1

    ok = not errors
    print(f"=== Validation of {schedule_path.name} against {input_dir} ===")
    print(f"battles read:          {len(sched)}")
    print(f"tickets consumed:      dedicated={kind_counts['dedicated']}, "
          f"wildcard={kind_counts['wildcard']}, total={sum(kind_counts.values())}")
    print(f"quests completed/max per member:")
    for m in members:
        max_q = sum(1 for (mm, _, _), v in quests.items() if mm == m and v == 1)
        print(f"  {m:<6} {per_member_quests[m]:>3} / {max_q:<3}   "
              f"(battles joined: {per_member_battles[m]})")
    print(f"TOTAL quests completed: {sum(per_member_quests.values())}")

    if ok:
        print("\n[OK] schedule is valid: ticket stock never went negative, "
              "team composition correct, and no unknown characters.")
    else:
        print(f"\n[FAIL] {len(errors)} problem(s) found:")
        for e in errors:
            print("  -", e)
    return ok


def main() -> None:
    ap = argparse.ArgumentParser(description="Validate a schedule workbook against its input directory.")
    ap.add_argument("--input-dir", default=str(Path("tests") / "input"),
                    help="Directory with per-member _票 / _委托 xlsx files.")
    ap.add_argument("--schedule", default=str(Path("tests") / "output" / "schedule.xlsx"),
                    help="Path to the schedule xlsx produced by the solver.")
    args = ap.parse_args()

    ok = validate(Path(args.input_dir), Path(args.schedule))
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    if __package__ is None:
        sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    main()
