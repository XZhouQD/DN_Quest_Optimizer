"""Quick helper: report full re-teams and switch stats for a schedule xlsx."""
import sys
import pandas as pd

path = sys.argv[1] if len(sys.argv) > 1 else "260422_v2.xlsx"
members = ['小C', '暗部', '桃核', '蹦蹦']
s = pd.read_excel(path, sheet_name='Schedule')
print(f'file: {path}')
print(f'battles: {len(s)}, zero-quest: {(s["quests_completed"]==0).sum()}')

full_reteams = 0
boundaries = 0
for i in range(1, len(s)):
    shared = switches = 0
    for m in members:
        p, c = s.iloc[i - 1][m], s.iloc[i][m]
        if pd.notna(p) and pd.notna(c):
            shared += 1
            if p != c:
                switches += 1
    if shared > 0:
        boundaries += 1
        if switches == shared:
            full_reteams += 1
print(f'full re-teams: {full_reteams} / {boundaries} boundaries')
print()
print(pd.read_excel(path, sheet_name='Summary').to_string(index=False))
