"""MILP optimizer for weekly team quest scheduling.

Inputs (in a single directory, one pair per member):
    <member>_票.xlsx    — tickets  per (character, target) + 选择 wildcard
    <member>_委托.xlsx  — weekly quest flags per (character, target)

Output (in memory): list of battle instances, each of the form
    { "target": str, "ticket_source": (member, character), "ticket_kind": str,
      "participants": {member: character, ...}, "completed": int }
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
import pulp

from .config import (
    TARGETS, TEAM_SIZE, CHAR_SHEET, QUEST_SHEET, WILDCARD, CHARACTER_COL,
    MEMBERS, TICKET_SUFFIX, QUEST_SUFFIX,
)


@dataclass
class Battle:
    target: str
    ticket_source: Tuple[str, str]            # (member, character)
    ticket_kind: str                          # target name or WILDCARD ("选择")
    participants: Dict[str, str]              # member -> character
    completed: int                            # quests completed by this battle


def _load_inputs(
    input_dir: Path,
) -> Tuple[List[str], Dict[str, List[str]], Dict, Dict, Dict]:
    """Load per-member ticket + quest files from `input_dir`.

    Expected filenames (one pair per member):
        <member>_票.xlsx    sheet=Characters   columns: character, <T>..., 选择
        <member>_委托.xlsx  sheet=Quests       columns: character, <T>...

    A member is included only if BOTH of their files are present and contain
    at least one character. Empty placeholder rows (blank character) are ignored.

    Returns (members, chars_by_member, tickets, wild_tickets, quests).
    """
    input_dir = Path(input_dir)
    if not input_dir.is_dir():
        raise ValueError(f"Input directory not found: {input_dir}")

    members: List[str] = []
    chars_by_member: Dict[str, List[str]] = {}
    tickets: Dict[Tuple[str, str, str], int] = {}
    wild_tickets: Dict[Tuple[str, str], int] = {}
    quests: Dict[Tuple[str, str, str], int] = {}

    expected_ticket_cols = [CHARACTER_COL] + list(TARGETS) + [WILDCARD]
    expected_quest_cols = [CHARACTER_COL] + list(TARGETS)

    for member in MEMBERS:
        ticket_path = input_dir / f"{member}{TICKET_SUFFIX}.xlsx"
        quest_path = input_dir / f"{member}{QUEST_SUFFIX}.xlsx"
        if not ticket_path.exists() or not quest_path.exists():
            print(f"[warn] skipping member '{member}': "
                  f"missing {ticket_path.name if not ticket_path.exists() else quest_path.name}")
            continue

        ch = pd.read_excel(ticket_path, sheet_name=CHAR_SHEET)
        qu = pd.read_excel(quest_path, sheet_name=QUEST_SHEET)

        miss_ch = set(expected_ticket_cols) - set(ch.columns)
        miss_qu = set(expected_quest_cols) - set(qu.columns)
        if miss_ch:
            raise ValueError(f"{ticket_path.name} missing columns: {sorted(miss_ch)}")
        if miss_qu:
            raise ValueError(f"{quest_path.name} missing columns: {sorted(miss_qu)}")

        # Keep only rows with a real character name
        ch = ch.dropna(subset=[CHARACTER_COL]).copy()
        qu = qu.dropna(subset=[CHARACTER_COL]).copy()
        ch[CHARACTER_COL] = ch[CHARACTER_COL].astype(str).str.strip()
        qu[CHARACTER_COL] = qu[CHARACTER_COL].astype(str).str.strip()
        ch = ch[ch[CHARACTER_COL] != ""]
        qu = qu[qu[CHARACTER_COL] != ""]

        # Coerce all numeric columns to non-negative integers (NaN / blank -> 0).
        for col in list(TARGETS) + [WILDCARD]:
            ch[col] = pd.to_numeric(ch[col], errors="coerce").fillna(0).astype(int).clip(lower=0)
        for col in TARGETS:
            qu[col] = pd.to_numeric(qu[col], errors="coerce").fillna(0).astype(int).clip(lower=0, upper=1)

        if ch.empty:
            print(f"[warn] skipping member '{member}': no characters in {ticket_path.name}")
            continue

        if ch[CHARACTER_COL].duplicated().any():
            dups = ch.loc[ch[CHARACTER_COL].duplicated(), CHARACTER_COL].tolist()
            raise ValueError(f"Duplicate character names in {ticket_path.name}: {dups}")

        qu_idx = qu.set_index(CHARACTER_COL) if not qu.empty else None

        char_list: List[str] = []
        for _, r in ch.iterrows():
            c = r[CHARACTER_COL]
            char_list.append(c)
            wild_tickets[(member, c)] = int(r[WILDCARD])
            for t in TARGETS:
                tickets[(member, c, t)] = int(r[t])
                if qu_idx is not None and c in qu_idx.index:
                    quests[(member, c, t)] = int(qu_idx.loc[c, t])
                else:
                    quests[(member, c, t)] = 0

        members.append(member)
        chars_by_member[member] = char_list

    if not members:
        raise ValueError(
            f"No member input files found in {input_dir}. "
            f"Expected files like '<member>{TICKET_SUFFIX}.xlsx' and '<member>{QUEST_SUFFIX}.xlsx'."
        )

    return members, chars_by_member, tickets, wild_tickets, quests


def _upper_bound_slots(
    members: List[str],
    chars_by_member: Dict[str, List[str]],
    tickets: Dict,
    wild_tickets: Dict,
    target: str,
) -> int:
    """Cheap upper bound on number of battles possible at target t.

    Limited by ticket supply and by team composition constraints.
    """
    k = TEAM_SIZE[target]
    # Total tickets available at t (dedicated + wildcard, unless target is 双生)
    total_tickets = sum(tickets[(m, c, target)] for m in members for c in chars_by_member[m])
    if target != "双生":
        total_tickets += sum(wild_tickets[(m, c)] for m in members for c in chars_by_member[m])
    # Per-member capacity: each battle needs 1 char from that member (non-双生) or
    # the member may sit out (双生). Be loose to let MILP decide.
    per_member = {
        m: sum(1 for c in chars_by_member[m] if tickets[(m, c, target)] > 0) or len(chars_by_member[m])
        for m in members
    }
    if target == "双生":
        # Need 2 distinct members per battle; sum of per_member / 2 is an upper bound.
        cap = sum(per_member.values()) // 2
    else:
        # Need 1 char per member; each member may appear in many battles (different chars).
        cap = min(per_member.values()) if per_member else 0
    return max(0, min(total_tickets, cap))


def solve(
    input_dir: str | Path,
    time_limit_sec: int = 60,
    verbose: bool = False,
) -> Tuple[List[Battle], List[str]]:
    """Solve the scheduling MILP. Returns (battles, members)."""
    input_dir = Path(input_dir)

    members, chars_by_member, tickets, wild_tickets, quests = _load_inputs(input_dir)
    if len(members) < 2:
        raise ValueError("Need at least 2 members in characters.xlsx.")

    # Build slot sets per target
    slots: Dict[str, List[int]] = {
        t: list(range(_upper_bound_slots(members, chars_by_member, tickets, wild_tickets, t)))
        for t in TARGETS
    }

    prob = pulp.LpProblem("quest_scheduling", pulp.LpMaximize)

    # p[t,s,m,c]  = 1 if char c of member m participates in battle slot s at target t
    p: Dict[Tuple[str, int, str, str], pulp.LpVariable] = {}
    # q[t,s,m,c]  = 1 if (m,c) spends a DEDICATED ticket_t for slot s at target t
    q: Dict[Tuple[str, int, str, str], pulp.LpVariable] = {}
    # qw[t,s,m,c] = 1 if (m,c) spends a WILDCARD (选择) ticket for slot s at t (t != 双生)
    qw: Dict[Tuple[str, int, str, str], pulp.LpVariable] = {}
    # active[t,s] = 1 if slot s at target t is used
    active: Dict[Tuple[str, int], pulp.LpVariable] = {}

    for t in TARGETS:
        for s in slots[t]:
            active[(t, s)] = pulp.LpVariable(f"a_{t}_{s}", cat="Binary")
            for m in members:
                for c in chars_by_member[m]:
                    p[(t, s, m, c)] = pulp.LpVariable(f"p_{t}_{s}_{m}_{c}", cat="Binary")
                    q[(t, s, m, c)] = pulp.LpVariable(f"q_{t}_{s}_{m}_{c}", cat="Binary")
                    if t != "双生":
                        qw[(t, s, m, c)] = pulp.LpVariable(f"qw_{t}_{s}_{m}_{c}", cat="Binary")

    # Constraints
    for t in TARGETS:
        k = TEAM_SIZE[t]
        for s in slots[t]:
            a = active[(t, s)]
            # Team size: exactly k participants when active, 0 otherwise
            prob += pulp.lpSum(p[(t, s, m, c)] for m in members for c in chars_by_member[m]) == k * a

            # At most 1 character per member per battle slot
            for m in members:
                prob += pulp.lpSum(p[(t, s, m, c)] for c in chars_by_member[m]) <= 1

            if t != "双生":
                # Every member must contribute exactly 1 when active
                for m in members:
                    prob += pulp.lpSum(p[(t, s, m, c)] for c in chars_by_member[m]) == a
            # For 双生 (k=2), the team-size constraint plus "≤1 per member" already
            # forces 2 distinct members to participate when active.

            # Ticket contribution: exactly 1 contributor when active, 0 otherwise.
            # Contributor may spend a dedicated OR wildcard ticket (wildcard only if t != 双生).
            ticket_vars = [q[(t, s, m, c)] for m in members for c in chars_by_member[m]]
            if t != "双生":
                ticket_vars += [qw[(t, s, m, c)] for m in members for c in chars_by_member[m]]
            prob += pulp.lpSum(ticket_vars) == a

            # Ticket contributor must be a participant; at most 1 ticket per (m,c) slot.
            for m in members:
                for c in chars_by_member[m]:
                    if t != "双生":
                        prob += q[(t, s, m, c)] + qw[(t, s, m, c)] <= p[(t, s, m, c)]
                    else:
                        prob += q[(t, s, m, c)] <= p[(t, s, m, c)]

            # Symmetry breaking: force slots to fill in order (s used => s-1 used)
            if s > 0:
                prob += active[(t, s)] <= active[(t, s - 1)]

    # Dedicated-ticket supply per (m, c, t)
    for m in members:
        for c in chars_by_member[m]:
            for t in TARGETS:
                prob += (
                    pulp.lpSum(q[(t, s, m, c)] for s in slots[t])
                    <= tickets[(m, c, t)]
                )

    # Wildcard (选择) ticket supply per (m, c): shared across all non-双生 targets
    for m in members:
        for c in chars_by_member[m]:
            prob += (
                pulp.lpSum(
                    qw[(t, s, m, c)]
                    for t in TARGETS if t != "双生"
                    for s in slots[t]
                )
                <= wild_tickets[(m, c)]
            )

    # Objective: lexicographic (total first, balance as tiebreaker).
    # Total quests completed by member m:
    member_quests = {
        m: pulp.lpSum(
            p[(t, s, m, c)] * quests[(m, c, t)]
            for t in TARGETS for s in slots[t] for c in chars_by_member[m]
        )
        for m in members
    }
    total_quests = pulp.lpSum(member_quests.values())
    min_member = pulp.LpVariable("min_member", lowBound=0, cat="Integer")
    for m in members:
        prob += min_member <= member_quests[m]

    # Scale total_quests high so it dominates; min_member acts only as a
    # tiebreaker among solutions that achieve the same total.
    # Upper bound on min_member: max per-member quest count.
    max_per_member = max(
        (sum(1 for c in chars_by_member[m] for t in TARGETS if quests[(m, c, t)] == 1))
        for m in members
    ) or 1
    prob += (max_per_member + 1) * total_quests + min_member

    solver = pulp.PULP_CBC_CMD(msg=1 if verbose else 0, timeLimit=time_limit_sec)
    prob.solve(solver)

    status = pulp.LpStatus[prob.status]
    if status not in ("Optimal", "Not Solved"):
        # "Not Solved" can occur when time-limited but a feasible solution exists.
        if pulp.value(total_quests) is None:
            raise RuntimeError(f"Solver failed: {status}")

    # Extract solution
    battles: List[Battle] = []
    for t in TARGETS:
        for s in slots[t]:
            if pulp.value(active[(t, s)]) is None or pulp.value(active[(t, s)]) < 0.5:
                continue
            participants: Dict[str, str] = {}
            for m in members:
                for c in chars_by_member[m]:
                    if pulp.value(p[(t, s, m, c)]) >= 0.5:
                        participants[m] = c
            ticket_src: Tuple[str, str] = ("", "")
            ticket_kind: str = t
            found = False
            for m in members:
                for c in chars_by_member[m]:
                    if pulp.value(q[(t, s, m, c)]) >= 0.5:
                        ticket_src = (m, c); ticket_kind = t; found = True; break
                    if t != "双生" and pulp.value(qw[(t, s, m, c)]) >= 0.5:
                        ticket_src = (m, c); ticket_kind = WILDCARD; found = True; break
                if found:
                    break
            completed = sum(
                quests[(m, c, t)] for m, c in participants.items()
            )
            battles.append(Battle(target=t, ticket_source=ticket_src,
                                  ticket_kind=ticket_kind,
                                  participants=participants, completed=completed))

    return battles, members
