"""Double-click-friendly entry point.

Reads input/<member>_票.xlsx and input/<member>_委托.xlsx, writes schedule.xlsx.
If the input/ folder is missing, it is created and the templates are copied in.
"""
from __future__ import annotations

import shutil
import sys
import traceback
from pathlib import Path


def _bootstrap_input_dir(root: Path) -> bool:
    """Ensure input/ exists and is populated. Return True if we just seeded it
    (caller should exit so the user can fill the files in)."""
    inp = root / "input"
    if inp.exists() and any(inp.glob("*.xlsx")):
        return False
    inp.mkdir(exist_ok=True)
    tpl = root / "templates"
    if not tpl.exists() or not any(tpl.glob("*.xlsx")):
        from src.templates import main as gen_templates
        gen_templates(tpl)
    for f in tpl.glob("*.xlsx"):
        target = inp / f.name
        if not target.exists():
            shutil.copy(f, target)
    print(f"[info] 'input/' was empty. Seeded templates into: {inp.resolve()}")
    print("[info] Fill them in (one pair per member), then run again.")
    return True


def main() -> None:
    root = Path(__file__).resolve().parent
    # Make imports work regardless of the cwd the user double-clicked from.
    sys.path.insert(0, str(root))
    import os
    os.chdir(root)

    seeded = _bootstrap_input_dir(root)
    if seeded:
        return

    from src.main import main as cli_main
    # If the user passed extra args on the CLI, keep them; otherwise use defaults.
    if len(sys.argv) == 1:
        sys.argv += ["--input-dir", "input", "--out", "schedule.xlsx"]
    cli_main()


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except ValueError as e:
        # Known input-validation error — show clean message only.
        print(f"\n[error] {e}")
        try:
            input("\nPress Enter to close...")
        except EOFError:
            pass
        sys.exit(1)
    except BaseException:
        traceback.print_exc()
        try:
            input("\nPress Enter to close...")
        except EOFError:
            pass
        sys.exit(1)
    else:
        try:
            input("\nDone. Press Enter to close...")
        except EOFError:
            pass
