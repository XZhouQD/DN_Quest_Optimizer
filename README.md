# DN Team Quest Scheduler

[English](README.md) | [中文](README.zh-CN.md)

Optimize weekly battle-quest scheduling from Excel input files.

The tool reads tickets + weekly quests, solves a MILP, and writes an ordered
schedule that:

1. Uses exactly 1 ticket per battle.
2. Maximizes total completed weekly quests first.
3. Uses fairness as a tiebreaker (maximize the minimum per-member completed quests).
4. Reduces character switching and avoids costly full re-team transitions.
5. Considers map changes in ordering.

## Battle targets

`狮蝎, 海龙, K博士, 主教, 巨人, 守卫, 火山, 迷雾, 卡伊伦, 格拉诺, 代达罗斯, 台风金, 双生`

## Dynamic members

Members are discovered from filename pairs in the input directory:

- `<member>_票.xlsx`
- `<member>_委托.xlsx`

No fixed member list is required in code.

If only 3 members are provided, non-`双生` battles are automatically treated as
3-player battles. `双生` remains 2-player.

## Input format

### `<member>_票.xlsx` (sheet: `Characters`)

| Column | Meaning |
| --- | --- |
| `角色` | Character name (unique in this file) |
| `<target>` | Integer ticket count for this target |
| `选择` | Wildcard tickets (valid for all targets except `双生`) |

### `<member>_委托.xlsx` (sheet: `Quests`)

| Column | Meaning |
| --- | --- |
| `角色` | Must match characters in `<member>_票.xlsx` |
| `<target>` | `1` means this target is a weekly quest for this character, else `0` |

Empty numeric cells are accepted and treated as `0`.

## Usage

```powershell
# Install dependencies
pip install -r requirements.txt

# First run: create and seed input/ from templates if needed
python run.py

# Or run directly
python -m src.main --input-dir input --out schedule.xlsx
```

## Output workbook

`schedule.xlsx` contains:

- `Schedule` sheet
  - ordered battles
  - switch-highlighted participant cells (orange)
  - quest-credit participant names in bold
  - target background coloring alternates on each map-group switch (two colors: light grey and light pink)
- `Summary` sheet
  - per-member completed quests, battles, distinct characters, and switches
- `Legend` sheet
  - explanations of symbols and color usage

## Scheduling model highlights

- Ticket constraints are strict (dedicated + wildcard stock never goes negative).
- Quest credit is counted once per `(member, character, target)` per week.
- Objective is lexicographic: total first, balance second.
- Post-processing removes zero-credit battles and optimizes order.
- Ordering cost combines:
  - character switch cost
  - map-change penalty
  - strong penalty for full re-team transitions

## Test and validation

```powershell
# deterministic synthetic test generation
python -m tests.generate_test_case

# validate produced schedule against input
python -m tests.validate_schedule --input-dir tests/input --schedule tests/output/schedule.xlsx

# dynamic feature regression (dynamic members, blanks-as-zero, target colors)
python -m tests.test_dynamic_features
```

## Project layout

```text
DN_Tools/
├── README.md
├── README.zh-CN.md
├── requirements.txt
├── run.py
├── run.bat
├── generate_templates.py
├── src/
│   ├── config.py
│   ├── optimize.py
│   ├── schedule.py
│   ├── templates.py
│   └── main.py
└── tests/
    ├── generate_test_case.py
    ├── validate_schedule.py
    ├── test_dynamic_features.py
    ├── count_reteams.py
    └── show_reteams.py
```
