# DN Team Quest Scheduler

Optimize weekly battle-quest scheduling for a 4-player team. The tool reads
two spreadsheets (tickets + weekly quests), runs a MILP, and writes a third
spreadsheet with an ordered battle plan that:

1. Always uses exactly 1 ticket per battle.
2. Fields a team of 4 characters (2 for `双生`), one from each different member.
3. Maximizes **total** quests completed across the team; ties are broken by
   maximizing the minimum per-member count (soft balance).
4. Orders battles to minimize per-member character switches (logout/login).

## Battle targets

`狮蝎, 海龙, K博士, 主教, 巨人, 守卫, 火山, 迷雾, 卡伊伦, 格拉诺, 代达罗斯, 台风金, 双生`

Team size is 4 for all targets except `双生` (2).

## Inputs

Each of the 4 teammates (`小C`, `暗部`, `桃核`, `蹦蹦`) fills in **two** files
about their own account only. Member identity is derived from the **file name**,
so there is no `member` column inside the sheets.

All 8 files live in the same directory (default: `templates/`).

### `<member>_票.xlsx` — sheet **Characters** (tickets)

| Column      | Meaning                                                                           |
| ----------- | --------------------------------------------------------------------------------- |
| `角色`      | Character name (unique within this file).                                         |
| `<T>`       | One column per battle target; integer weekly tickets (0, 1, 2, ...).              |
| `选择`      | Wildcard tickets — each usable for any target **except `双生`**.                  |

### `<member>_委托.xlsx` — sheet **Quests** (weekly quest flags)

| Column      | Meaning                                                                    |
| ----------- | -------------------------------------------------------------------------- |
| `角色`      | Must match the character names in `<member>_票.xlsx`.                      |
| `<T>`       | `1` if this target is a weekly quest for the character; `0` otherwise.     |

Each generated template ships with at least 15 blank placeholder rows.

## Usage

```powershell
# 1. Install deps (one time)
pip install -r requirements.txt

# 2. Double-click run.bat  — OR —  run: python run.py
#    On the first run an `input/` folder is created and the 8 empty templates
#    are copied into it. Fill them in, then run again.

# 3. Once input/ contains your data, double-click run.bat again to produce
#    schedule.xlsx in the project root.
```

Advanced (use a different folder):

```powershell
python -m src.main --input-dir some/other/dir --out schedule.xlsx
```

## Test suite

A deterministic end-to-end test lives under `tests/`. It builds a small
reproducible input set and runs the solver against it:

```powershell
python -m tests.generate_test_case
```

Rules of the test case (seeded, so identical every run):
* 10 characters per member.
* Only the first character of each member has tickets
  (a random integer in `[5, 12]` dedicated tickets distributed across random
  targets, plus **2** `选择` wildcard tickets). Other characters have none.
* Each character has a random number (0–3) of weekly quests.

Files written:
* `tests/input/<member>_票.xlsx`, `tests/input/<member>_委托.xlsx`
* `tests/output/schedule.xlsx`

## Output

`schedule.xlsx`:

- **Schedule** sheet: one row per battle, in execution order. Columns:
  `order | target | ticket_source | <each member> | quests_completed`.
  Cells where a member changes character vs the previous battle are highlighted
  in orange, so you can eyeball switch points.
- **Summary** sheet: per-member stats (quests completed, battles joined,
  distinct characters used, character switches). Plus overall totals.
- **Legend** sheet: key to the colors and symbols.

## Modeling notes / assumptions

- **Ticket model**: tickets are tracked per `(character, target)` and may be any
  non-negative integer — a character with 3 tickets for `火山` can legitimately
  contribute to 3 separate `火山` battles. In addition, each character has a
  pool of **`选择` wildcard tickets** that may substitute for any target except
  `双生`. The schedule's `ticket_kind` column shows which kind was spent.
- **Quest credit**: a participant "completes a quest" iff the battle's target
  is flagged `1` on their row in `quests.xlsx`.
- **Non-`双生` battles** require exactly one character from **each** of the
  teammates. `双生` battles require exactly 2 distinct members to contribute
  one character each (the other 2 sit out, shown as `—`).
- **Balance**: the objective is lexicographic — first maximize the sum of
  quests completed, then (as a tiebreaker) maximize `min_m quests(m)`. This
  produces the highest total first and a fair distribution only when it costs
  nothing.
- **Ordering**: after the MILP picks the battle set, a nearest-neighbor heuristic
  (with every battle tried as a start point) orders them to minimize total
  character switches across all members.

## Project layout

```
DN_Tools/
├── requirements.txt
├── README.md
├── run.bat                    # double-click launcher (Windows)
├── run.py                     # bootstraps input/ and runs the optimizer
├── generate_templates.py      # (optional) (re)generate empty templates/
└── src/
    ├── config.py              # targets, team sizes, member list, filename suffixes
    ├── templates.py           # builds the 4 × 2 per-member templates
    ├── optimize.py            # MILP model (PuLP + CBC)
    ├── schedule.py            # ordering heuristic + xlsx writer
    └── main.py                # CLI
```
