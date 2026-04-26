"""Quick helper: report full re-teams, map changes, and switch stats for a schedule xlsx."""
import sys
import pandas as pd

# Map groups (kept in sync with src/config.py)
MAP_GROUPS = [
    ("狮蝎", "海龙"),
    ("主教", "K博士", "巨人"),
    ("火山", "守卫", "迷雾"),
    ("代达罗斯", "格拉诺"),
    ("卡伊伦",),
    ("台风金",),
    ("双生",),
]
TARGET_TO_MAP = {t: i for i, g in enumerate(MAP_GROUPS) for t in g}

path = sys.argv[1] if len(sys.argv) > 1 else "260422_v2.xlsx"
s = pd.read_excel(path, sheet_name='Schedule')
members = [
    c for c in s.columns
    if c not in {"order", "target", "ticket_kind", "ticket_source", "quests_completed"}
]
print(f'file: {path}')
print(f'battles: {len(s)}, zero-quest: {(s["quests_completed"]==0).sum()}')

full_reteams = 0
boundaries = 0
map_changes = 0
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
    if TARGET_TO_MAP.get(s.iloc[i - 1]['target']) != TARGET_TO_MAP.get(s.iloc[i]['target']):
        map_changes += 1
print(f'full re-teams: {full_reteams} / {boundaries} boundaries')
print(f'map changes:   {map_changes} / {len(s) - 1} transitions')
print()
print(pd.read_excel(path, sheet_name='Summary').to_string(index=False))
