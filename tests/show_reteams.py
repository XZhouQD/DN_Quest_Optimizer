"""Print each full re-team in a schedule with its neighbor battles."""
import sys
import pandas as pd

path = sys.argv[1] if len(sys.argv) > 1 else "260422_v3.xlsx"
s = pd.read_excel(path, sheet_name='Schedule')
members = [
    c for c in s.columns
    if c not in {"order", "target", "ticket_kind", "ticket_source", "quests_completed"}
]

for i in range(1, len(s)):
    shared = sw = 0
    for m in members:
        p, c = s.iloc[i - 1][m], s.iloc[i][m]
        if pd.notna(p) and pd.notna(c):
            shared += 1
            if p != c:
                sw += 1
    if shared > 0 and sw == shared:
        t0 = s.iloc[i - 1]['target']
        t1 = s.iloc[i]['target']
        ord1 = s.iloc[i]['order']
        print(f"--- re-team at order {ord1} ({t0} -> {t1}) ---")
        print(s.iloc[[i - 1, i]].to_string(index=False))
        print()
