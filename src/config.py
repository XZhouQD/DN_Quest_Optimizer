"""Shared configuration: battle targets and team sizes."""

# All 13 battle targets (order matters for template column layout).
TARGETS = [
    "狮蝎", "海龙", "K博士", "主教", "巨人", "守卫",
    "火山", "迷雾", "卡伊伦", "格拉诺", "代达罗斯", "台风金",
    "双生",
]

# Team size per target. 双生 = 2, all others = 4.
TEAM_SIZE = {t: (2 if t == "双生" else 4) for t in TARGETS}

# Wildcard ticket column name. A 选择 ticket is valid for any target EXCEPT 双生.
WILDCARD = "选择"

# Character-name column label used in the templates.
CHARACTER_COL = "角色"

# Sheet names (kept identical between templates and output for clarity).
CHAR_SHEET = "Characters"
QUEST_SHEET = "Quests"
SCHEDULE_SHEET = "Schedule"
SUMMARY_SHEET = "Summary"

# Team members (fixed). Filenames are derived from these names.
MEMBERS = ["小C", "暗部", "桃核", "蹦蹦"]

# File-name suffixes used to identify which spreadsheet belongs to whom.
#   <member>_票.xlsx   -> tickets input
#   <member>_委托.xlsx -> weekly quests input
TICKET_SUFFIX = "_票"
QUEST_SUFFIX = "_委托"

# Minimum number of blank placeholder rows to emit in each template.
PLACEHOLDER_ROWS = 15
