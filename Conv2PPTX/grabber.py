import csv
import os
import sys
import traceback
from pathlib import Path

# ── Dynamic Paths (relative to this script's location) ────────────────────────

BASE_DIR = Path(__file__).parent
DESTINATION_DIR = BASE_DIR / "test_batch" / "test_files"
OUTPUT_CSV_PATH = BASE_DIR / "test_batch" / "test_batch.csv"

# ──────────────────────────────────────────────────────────────────────────────

def normalize_path(raw_path: str) -> str:
    """
    Convert forward-slash UNC paths (//server/share/...) to Windows UNC paths
    (\\\\server\\share\\...) and normalise any mixed separators.
    """
    path = raw_path.strip()
    if path.startswith("//"):
        path = "\\\\" + path[2:]  # // → \\
    path = path.replace("/", "\\")  # remaining / → \
    return os.path.normpath(path)


def generate_csv_from_folder(destination: Path, output_csv: Path) -> bool:
    """
    Scans the destination folder for .pub files and writes a CSV with:
      - File Name      : filename stem (no extension), uppercased e.g. AE1555
      - Full File Path : full absolute path to the file, normalized
    Returns True on success, False on failure.
    """
    if not destination.exists():
        print(f"[ERROR] Destination folder does not exist:\n  {destination}")
        return False

    # Collect all .pub files (case-insensitive)
    dwg_files = sorted(
        f for f in destination.iterdir()
        if f.is_file() and f.suffix.lower() == ".pub"
    )

    if not dwg_files:
        print(f"[WARNING] No .pub files found in:\n  {destination}")
        return False

    output_csv.parent.mkdir(parents=True, exist_ok=True)

    with open(output_csv, "w", newline="", encoding="utf-8") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(["File Name", "Full File Path"])  # Header

        for dwg in dwg_files:
            file_name = dwg.stem.upper()          # e.g. AE1555
            full_path = normalize_path(str(dwg))  # Normalized Windows path
            writer.writerow([file_name, full_path])

    print(f"[SUCCESS] CSV generated with {len(dwg_files)} file(s):")
    print(f"  -> {output_csv}")
    return True


if __name__ == "__main__":
    # IMPORTANT: We deliberately NEVER call sys.exit() with a nonzero code,
    # and we swallow unexpected exceptions at the top level. Reason: when
    # this script is launched by `install.bat` through the Microsoft Store
    # Python alias, an abnormal exit can close the parent console window
    # that the installer is running in. We always print clearly and fall
    # through to a clean exit (0) so install.bat can keep going.
    try:
        success = generate_csv_from_folder(DESTINATION_DIR, OUTPUT_CSV_PATH)
        if not success:
            print("[INFO] grabber.py finished with missing files or errors, "
                  "but will not close the terminal.")
    except Exception as e:
        print(f"[ERROR] grabber.py hit an unexpected error: {e}")
        traceback.print_exc()
        print("[INFO] Continuing without crashing the installer terminal.")
    finally:
        # Make sure everything is on screen before the interpreter exits,
        # since the MS Store alias can be abrupt.
        sys.stdout.flush()
        sys.stderr.flush()