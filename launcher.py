import subprocess
import sys
import pathlib
import runpy
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def base_dir() -> Path:
    """
    Where bundled files live:
      - Source: directory containing launcher.py
      - PyInstaller onefile: sys._MEIPASS extraction dir
    """
    if is_frozen() and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


ROOT = base_dir()


def dispatch_run_mode() -> None:
    """
    Runs a tool script when invoked as:
      ForgePrograms.exe --run Inventory/scan_in.py
    Optional (for testing):
      python launcher.py --run Inventory/scan_in.py
    """
    if len(sys.argv) >= 3 and sys.argv[1] == "--run":
        rel = sys.argv[2]
        script_path = (ROOT / rel).resolve()

        if not script_path.exists():
            raise FileNotFoundError(f"Cannot run: {script_path}")

        # Make imports like "import shared_functions" work for Inventory tools
        script_dir = script_path.parent
        sys.path.insert(0, str(script_dir))  # e.g., Inventory/
        sys.path.insert(0, str(ROOT))        # repo/bundle root

        # Run as __main__
        sys.argv = [str(script_path)] + sys.argv[3:]
        runpy.run_path(str(script_path), run_name="__main__")
        raise SystemExit(0)


# Only dispatch early if explicitly requested (source testing) OR in frozen EXE mode
dispatch_run_mode()


APPS = {
    "Inventory": ROOT / "Inventory" / "inventory.py",
    "1150 Genner": ROOT / "Genner1150" / "main.py",
    "Scan Out": ROOT / "Inventory" / "scan_out.py",
    "Scan In": ROOT / "Inventory" / "scan_in.py",
    "Import new HR": ROOT / "Inventory" / "import_hr.py",
}


def run_app(path: Path):
    if not path.exists():
        messagebox.showerror("Error", f"Not found: {path}")
        return

    # Source behavior (keep it simple and unchanged): run python.exe <script> with cwd=script folder
    if not is_frozen():
        subprocess.Popen([sys.executable, str(path)], cwd=str(path.parent))
        return

    # Frozen behavior: re-launch the EXE in --run mode to avoid re-opening the launcher GUI
    rel = path.relative_to(ROOT).as_posix()  # e.g. Inventory/scan_in.py
    subprocess.Popen([sys.executable, "--run", rel], cwd=str(path.parent))


def main():
    root = tk.Tk()
    root.title("Toolbox Launcher")
    root.geometry("360x360")

    frm = ttk.Frame(root, padding=16)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text="Choose a program to run:", font=("Segoe UI", 12, "bold")).pack(pady=(0, 12))

    for name, path in APPS.items():
        ttk.Button(frm, text=name, command=lambda p=path: run_app(p)).pack(fill="x", pady=6)

    ttk.Separator(frm).pack(fill="x", pady=12)
    ttk.Button(frm, text="Quit", command=root.destroy).pack()

    root.mainloop()


if __name__ == "__main__":
    main()
