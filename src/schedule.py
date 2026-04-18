"""Order battles to minimize per-member character switches, and write output xlsx."""
from __future__ import annotations

from pathlib import Path
from typing import Dict, List

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .config import SCHEDULE_SHEET, SUMMARY_SHEET, TARGETS
from .optimize import Battle


def _switch_cost(prev: Battle | None, b: Battle) -> int:
    """Number of member-level character switches when going prev -> b.

    A switch happens when a member participates in both battles with a different
    character. A member sitting out (only possible for 双生) is free.
    """
    if prev is None:
        return 0
    cost = 0
    for m, c in b.participants.items():
        pc = prev.participants.get(m)
        if pc is not None and pc != c:
            cost += 1
    return cost


def order_battles(battles: List[Battle]) -> List[Battle]:
    """Greedy nearest-neighbor ordering minimizing switches.

    Tries each battle as a starting point and keeps the best sequence.
    Works well for the problem size (usually < 50 battles).
    """
    if not battles:
        return []

    best_seq: List[Battle] = []
    best_cost = None

    for start_idx in range(len(battles)):
        remaining = battles.copy()
        seq = [remaining.pop(start_idx)]
        total = 0
        while remaining:
            prev = seq[-1]
            # Pick the remaining battle with minimum switch cost (ties: same target)
            def key(b: Battle):
                return (_switch_cost(prev, b), 0 if b.target == prev.target else 1)
            nxt = min(remaining, key=key)
            total += _switch_cost(prev, nxt)
            remaining.remove(nxt)
            seq.append(nxt)
        if best_cost is None or total < best_cost:
            best_cost = total
            best_seq = seq

    return best_seq


# ---------- Excel output ----------

HEADER_FILL = PatternFill("solid", fgColor="FFD9E1F2")
HEADER_FONT = Font(bold=True)
SWITCH_FILL = PatternFill("solid", fgColor="FFFCE4D6")  # highlight switches


def write_schedule(
    battles: List[Battle],
    members: List[str],
    out_path: str | Path,
) -> None:
    out_path = Path(out_path)
    wb = Workbook()

    # --- Schedule sheet ---
    ws = wb.active
    ws.title = SCHEDULE_SHEET
    headers = ["order", "target", "ticket_kind", "ticket_source"] + members + ["quests_completed"]
    ws.append(headers)
    for col in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=col)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "B2"

    prev: Battle | None = None
    for i, b in enumerate(battles, start=1):
        ticket_src = f"{b.ticket_source[0]}:{b.ticket_source[1]}" if b.ticket_source[0] else ""
        row = [i, b.target, b.ticket_kind, ticket_src]
        for m in members:
            row.append(b.participants.get(m, "—"))
        row.append(b.completed)
        ws.append(row)
        # Highlight cells where character changed vs previous battle
        if prev is not None:
            for j, m in enumerate(members, start=5):  # col 5+ = member columns
                pc = prev.participants.get(m)
                nc = b.participants.get(m)
                if pc and nc and pc != nc:
                    ws.cell(row=ws.max_row, column=j).fill = SWITCH_FILL
        prev = b

    for col in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.column_dimensions["A"].width = 7
    ws.column_dimensions["B"].width = 10

    # --- Summary sheet ---
    ws2 = wb.create_sheet(SUMMARY_SHEET)
    ws2.append(["member", "quests_completed", "battles_participated",
                "distinct_characters_used", "character_switches"])
    for col in range(1, 6):
        c = ws2.cell(row=1, column=col)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT

    # Compute per-member stats
    per_member: Dict[str, Dict[str, int]] = {
        m: {"quests": 0, "battles": 0, "chars": set(), "switches": 0} for m in members
    }
    prev = None
    for b in battles:
        for m, c in b.participants.items():
            per_member[m]["battles"] += 1
            per_member[m]["chars"].add(c)
            if prev is not None:
                pc = prev.participants.get(m)
                if pc is not None and pc != c:
                    per_member[m]["switches"] += 1
        # Quest credit: only to participants whose quest list included the target.
        # We don't have quest flags here; rely on Battle.completed being already the
        # sum. Attribute proportionally? Instead recount below from battle info is
        # not possible without quests; store it on Battle. For summary we rely on
        # per-battle completed ≤ team size, so split evenly is wrong. We instead
        # recompute quests per member in the caller.
        prev = b

    # Member quest counts must be recomputed with quest flags; the caller passes
    # them via battles (we stashed via per-participant attribute). Simpler:
    # recompute from battle.participants + battle.completed is NOT enough.
    # So: write_schedule also accepts optional per_member_quests.
    for m in members:
        d = per_member[m]
        ws2.append([m, "", d["battles"], len(d["chars"]), d["switches"]])
    for col in range(1, 6):
        ws2.column_dimensions[get_column_letter(col)].width = 22

    total_battles = len(battles)
    total_quests = sum(b.completed for b in battles)
    ws2.append([])
    ws2.append(["TOTAL battles", total_battles])
    ws2.append(["TOTAL quests completed", total_quests])

    # Legend
    ws3 = wb.create_sheet("Legend")
    ws3.append(["Orange cell in Schedule = this member switches character vs previous battle."])
    ws3.append(["ticket_source = member:character who spends 1 ticket for that battle."])
    ws3.append(["— = member sits out (only possible for 双生, team size = 2)."])
    ws3.column_dimensions["A"].width = 90

    wb.save(out_path)


def write_schedule_with_quests(
    battles: List[Battle],
    members: List[str],
    quests: Dict,              # (m,c,t) -> 0/1
    out_path: str | Path,
) -> None:
    """Same as write_schedule but fills per-member quest counts correctly."""
    write_schedule(battles, members, out_path)
    # Reopen to fill quest column B in Summary sheet
    from openpyxl import load_workbook
    wb = load_workbook(out_path)
    ws = wb[SUMMARY_SHEET]
    per_member_q: Dict[str, int] = {m: 0 for m in members}
    for b in battles:
        for m, c in b.participants.items():
            per_member_q[m] += quests.get((m, c, b.target), 0)
    # Rows start at 2 (header at 1), one per member in the order we appended.
    for i, m in enumerate(members, start=2):
        ws.cell(row=i, column=2, value=per_member_q[m])
    wb.save(out_path)
