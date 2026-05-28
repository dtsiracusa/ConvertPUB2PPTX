# file: Step2-Convert2PDF.py

import os
import csv
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client
from tqdm import tqdm
import time
import ctypes
from ctypes import wintypes
import subprocess


# ============================================================
# Dialog Dismissal Logic (from pub2pptx converter)
# ============================================================

WM_CLOSE = 0x0010
BM_CLICK = 0x00F5

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)


def get_window_text(hwnd):
    """Get the title text of a window."""
    length = user32.GetWindowTextLengthW(hwnd)
    if length == 0:
        return ""
    buffer = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buffer, length + 1)
    return buffer.value


def get_class_name(hwnd):
    """Get the class name of a window."""
    buffer = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buffer, 256)
    return buffer.value


def is_dialog_window(hwnd):
    """Check if window is likely a dialog box."""
    class_name = get_class_name(hwnd)
    dialog_classes = [
        '#32770', 'Dialog', 'NUIDialog',
        'bosa_sdm_msword', 'bosa_sdm_Microsoft Office Publisher',
    ]
    return class_name in dialog_classes


def find_publisher_dialogs():
    """Find any Publisher dialog windows that might be blocking."""
    dialog_windows = []

    def enum_callback(hwnd, lParam):
        if user32.IsWindowVisible(hwnd):
            title = get_window_text(hwnd)
            class_name = get_class_name(hwnd)

            if is_dialog_window(hwnd):
                if any(keyword in title.lower() for keyword in
                       ['publisher', 'microsoft', 'missing', 'font', 'picture',
                        'error', 'warning', 'recovery', 'unable', 'cannot']):
                    dialog_windows.append((hwnd, title, class_name))

            # Also catch generic Office dialogs
            if class_name == '#32770':
                if title and user32.IsWindowEnabled(hwnd):
                    dialog_windows.append((hwnd, title, class_name))
        return True

    user32.EnumWindows(EnumWindowsProc(enum_callback), 0)
    return dialog_windows


def dismiss_dialog(hwnd, click_button=None):
    """
    Dismiss a dialog window by clicking a specific button or sending WM_CLOSE.
    """
    try:
        if click_button:
            def find_button(parent_hwnd, target_text):
                buttons = []

                def enum_child_callback(child_hwnd, lParam):
                    class_name = get_class_name(child_hwnd)
                    if class_name == 'Button':
                        text = get_window_text(child_hwnd)
                        if target_text.lower() in text.lower():
                            buttons.append(child_hwnd)
                    return True

                user32.EnumChildWindows(parent_hwnd, EnumWindowsProc(enum_child_callback), 0)
                return buttons[0] if buttons else None

            button_hwnd = find_button(hwnd, click_button)
            if button_hwnd:
                user32.SendMessageW(button_hwnd, BM_CLICK, 0, 0)
                return True

        # Fallback: send WM_CLOSE
        user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)
        return True
    except Exception as e:
        print(f"    Warning: Could not dismiss dialog: {e}")
        return False


def dismiss_all_publisher_dialogs(verbose=False):
    """Find and dismiss all Publisher-related dialog windows."""
    dialogs = find_publisher_dialogs()
    dismissed_count = 0

    for hwnd, title, class_name in dialogs:
        if verbose:
            print(f"    Found dialog: '{title}' (class: {class_name})")

        if any(word in title.lower() for word in ['ok', 'continue', 'yes']):
            if (dismiss_dialog(hwnd, "OK") or
                    dismiss_dialog(hwnd, "Yes") or
                    dismiss_dialog(hwnd, "Continue")):
                dismissed_count += 1
        else:
            if dismiss_dialog(hwnd):
                dismissed_count += 1

    return dismissed_count


def kill_publisher_process():
    """Kill any running Publisher processes."""
    try:
        subprocess.run(['taskkill', '/F', '/IM', 'MSPUB.EXE'],
                       capture_output=True, timeout=5)
        time.sleep(1)
        return True
    except Exception as e:
        print(f"    Warning: Could not kill Publisher process: {e}")
        return False


def is_modal_dialog_error(error):
    """Check if an error is due to a modal dialog being active."""
    error_str = str(error).lower()
    return 'modal dialog' in error_str or '-2147221457' in error_str


# ----------------------------
# UI Helpers
# ----------------------------

def select_files():
    """Manual mode: let user pick individual PPT/PPTX/PUB files."""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    file_paths = filedialog.askopenfilenames(
        title="Select PowerPoint or Publisher files",
        filetypes=[
            ("PowerPoint/Publisher Files", "*.ppt *.pptx *.pub"),
            ("All Files", "*.*"),
        ],
    )
    return list(file_paths)


def select_csv_file():
    """Batch mode: pick a CSV that has a 'Full File Path' column."""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    csv_path = filedialog.askopenfilename(
        title="Select CSV file with 'Full File Path' column",
        filetypes=[
            ("CSV Files", "*.csv"),
            ("All Files", "*.*"),
        ],
    )
    return csv_path


def select_folder():
    """Select destination folder for PDFs."""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    folder_path = filedialog.askdirectory(title="Select destination folder for PDFs")
    return folder_path


def ask_processing_mode():
    """
    Ask user in the console which mode to use:
    1 = Manual (file dialog), 2 = Batch from CSV.
    """
    while True:
        print("\nSelect processing mode:")
        print("  1) Manual file selection")
        print("  2) Batch from CSV")
        choice = input("Enter 1 or 2: ").strip()

        if choice == "1":
            return "manual"
        elif choice == "2":
            return "csv"
        else:
            print("Invalid choice. Please enter 1 or 2.")


def ask_max_files(total):
    """
    Let the user declare how many files to process (1..total).
    Empty input means 'all'.
    """
    while True:
        reply = input(
            f"\nHow many files do you want to process? (1–{total}, Enter for all): "
        ).strip()

        if reply == "":
            return None  # All files

        try:
            n = int(reply)
            if 1 <= n <= total:
                return n
            else:
                print(f"Please enter a number between 1 and {total}.")
        except ValueError:
            print("Please enter a valid integer.")


# ----------------------------
# CSV Helpers
# ----------------------------

def load_csv_rows(csv_path, path_column="Full File Path", status_column="Conversion Status"):
    """
    Load rows from CSV and normalize:
    - Ensure status_column exists in fieldnames and rows
    - Clean/strip the path column
    - Build valid_indices for rows that:
        * have a non-empty path, AND
        * are not already marked as 'Converted'
    Returns (rows, fieldnames, valid_indices_to_process)
    """
    with open(csv_path, newline="", encoding="utf-8-sig") as csvfile:
        reader = csv.DictReader(csvfile)
        if reader.fieldnames is None:
            raise ValueError("CSV has no header row.")

        fieldnames = list(reader.fieldnames)
        rows = list(reader)

    # Ensure status column exists
    if status_column not in fieldnames:
        fieldnames.append(status_column)
        for row in rows:
            row[status_column] = ""
    else:
        for row in rows:
            row.setdefault(status_column, "")

    valid_indices = []
    for i, row in enumerate(rows):
        # Normalize path
        raw_path = row.get(path_column, "")
        if raw_path is None:
            raw_path = ""
        cleaned_path = raw_path.strip()
        row[path_column] = cleaned_path

        # Normalize status
        raw_status = row.get(status_column, "")
        if raw_status is None:
            raw_status = ""
        cleaned_status = raw_status.strip()

        # Only consider rows that:
        #  - have a path
        #  - are not already 'Converted'
        if cleaned_path and cleaned_status.lower() != "converted":
            valid_indices.append(i)

    return rows, fieldnames, valid_indices


def write_csv_rows(csv_path, rows, fieldnames):
    """Write the updated rows back to the same CSV."""
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


# ----------------------------
# Conversion Core (Optimized)
# ----------------------------

def ensure_pdf_path(output_folder, file_path):
    """Build the absolute PDF output path from an input file path."""
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    output_path = os.path.join(output_folder, f"{file_name}.pdf")
    return os.path.abspath(output_path).replace("/", "\\")


def _open_publisher_with_dialog_handling(publisher, abs_path, max_retries=3, verbose=False):
    """
    Open a Publisher file with retry logic for modal dialogs.
    Dismisses any blocking dialogs and retries.
    """
    for attempt in range(max_retries):
        try:
            dismiss_all_publisher_dialogs(verbose=verbose)
            doc = publisher.Open(abs_path, False, False)
            return doc
        except Exception as e:
            if is_modal_dialog_error(e):
                if verbose:
                    print(f"    Modal dialog detected on attempt {attempt + 1}, dismissing...")
                dismissed = dismiss_all_publisher_dialogs(verbose=True)
                if dismissed == 0:
                    time.sleep(0.5)
                time.sleep(0.3)
                continue
            raise
    raise RuntimeError(f"Failed to open '{abs_path}' after {max_retries} attempts due to modal dialogs.")


def convert_pptx_to_pdf(pptx_path, output_folder):
    """
    Convert a PowerPoint file (.pptx / .ppt) to PDF using a fresh,
    isolated PowerPoint COM instance (DispatchEx) per file.
    This avoids the 'Python instance cannot be converted to COM object'
    error that occurs when reusing a shared COM instance across files.
    """
    abs_path = os.path.abspath(pptx_path).replace("/", "\\")
    pdf_path = ensure_pdf_path(output_folder, pptx_path)

    pp = None
    presentation = None
    try:
        # DispatchEx creates a brand-new, isolated PowerPoint process per file
        pp = win32com.client.DispatchEx("PowerPoint.Application")
        pp.Visible = True
        pp.DisplayAlerts = False

        # Open: ReadOnly=True (0), Untitled=False (0), WithWindow=False (0)
        presentation = pp.Presentations.Open(abs_path, ReadOnly=True, WithWindow=False)

        # ExportAsFixedFormat: FixedFormatType 2 = PDF
        presentation.SaveAs(pdf_path, 32)
        return True

    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass



def batch_convert_files(file_items, output_folder, rows=None, status_col=None):
    """
    Batch convert .pub (Publisher) and .pptx/.ppt (PowerPoint) files to PDF.

    Publisher ExportAsFixedFormat positional signature (per Microsoft Docs):
        ExportAsFixedFormat(Format, FileName, Intent)
        - Format   = 1  (pbFixedFormatTypePDF)
        - FileName = pdf_path string
        - Intent   = 1  (pbFixedFormatIntentPrint)
    NOTE: win32com does not support keyword arguments for Publisher COM.
    NOTE: time.sleep(1) after Open() is required — known Publisher COM timing issue.

    PowerPoint: SaveAs(FileName, 32) where 32 = ppSaveAsPDF.
    """
    publisher = None
    powerpoint = None

    try:
        with tqdm(total=len(file_items), unit="file", desc="Converting") as pbar:
            for row_idx, file_path in file_items:
                abs_path = os.path.abspath(file_path).replace("/", "\\")
                ext = os.path.splitext(file_path)[1].lower()
                pdf_path = ensure_pdf_path(output_folder, file_path)

                try:
                    if ext == ".pub":
                        # Publisher path — .Visible not supported on Publisher.Application
                        if publisher is None:
                            publisher = win32com.client.Dispatch("Publisher.Application")

                        doc = _open_publisher_with_dialog_handling(publisher, abs_path)
                        time.sleep(1)  # Allow Publisher to fully load before export
                        try:
                            doc.ExportAsFixedFormat(
                                2,          # Format   — pbFixedFormatTypePDF
                                pdf_path,   # FileName — output path
                                1           # Intent   — pbFixedFormatIntentPrint
                            )
                        finally:
                            try:
                                doc.Close()
                            except Exception:
                                pass

                    elif ext in (".pptx", ".ppt"):
                        # PowerPoint path
                        if powerpoint is None:
                            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
                            powerpoint.Visible = 1  # msoTrue — required for PowerPoint COM

                        prs = powerpoint.Presentations.Open(abs_path, WithWindow=False)
                        try:
                            prs.SaveAs(pdf_path, 32)  # ppSaveAsPDF = 32
                        finally:
                            try:
                                prs.Close()
                            except Exception:
                                pass

                    else:
                        raise ValueError(f"Unsupported file type: {ext}")

                    # Mark row as successfully converted
                    if rows is not None and status_col:
                        rows[row_idx][status_col] = "Converted"

                except Exception as e:
                    tqdm.write(f"\n❌ Failed to convert {file_path}: {e}")

                pbar.update(1)

    finally:
        # Clean up COM objects
        if publisher is not None:
            try:
                publisher.Quit()
            except Exception:
                pass
        if powerpoint is not None:
            try:
                powerpoint.Quit()
            except Exception:
                pass


# ----------------------------
# High-Level Workflows
# ----------------------------

def process_manual_mode():
    """Manual selection of files; no CSV modification (just faster conversion)."""
    files = select_files()
    if not files:
        print("No files selected.")
        return

    output_folder = select_folder()
    if not output_folder:
        print("No output folder selected.")
        return

    # Sort smallest to largest
    file_items = sorted(
        enumerate(files),
        key=lambda x: os.path.getsize(x[1]) if os.path.exists(x[1]) else 0
    )

    n = ask_max_files(len(file_items))
    if n is not None:
        file_items = file_items[:n]

    print(f"\nConverting {len(file_items)} file(s) to PDF (manual mode, smallest to largest)...")
    batch_convert_files(file_items, output_folder)
    print("\n✅ Manual mode conversion complete!")


def process_csv_mode():
    """Batch processing from CSV, with status written back into the CSV, skipping 'Converted' rows,
    and processing pending files in order of smallest to largest file size.
    """
    csv_file = select_csv_file()
    if not csv_file:
        print("No CSV file selected.")
        return

    output_folder = select_folder()
    if not output_folder:
        print("No output folder selected.")
        return

    rows, fieldnames, valid_indices = load_csv_rows(csv_file)
    status_col = "Conversion Status"

    if not valid_indices:
        print("No pending files found in the CSV (all are converted or paths are empty).")
        return

    print(f"\nFound {len(valid_indices)} pending file(s) in the CSV (excluding 'Converted').")

    # Sort pending files by file size (smallest first)
    def get_size(idx):
        path = rows[idx].get("Full File Path", "")
        return os.path.getsize(path) if path and os.path.exists(path) else 0

    sorted_indices = sorted(valid_indices, key=get_size)

    n = ask_max_files(len(sorted_indices))
    if n is not None:
        sorted_indices = sorted_indices[:n]

    file_items = [(idx, rows[idx]["Full File Path"]) for idx in sorted_indices]

    print(f"\nConverting {len(file_items)} file(s) to PDF (from CSV, smallest to largest)...")
    batch_convert_files(file_items, output_folder, rows=rows, status_col=status_col)

    write_csv_rows(csv_file, rows, fieldnames)
    print("\n✅ CSV updated with conversion status in column 'Conversion Status'.")
    print("✅ CSV mode conversion complete!")


# ----------------------------
# Main
# ----------------------------

if __name__ == "__main__":
    print("=== Convert PPTX/PUB to PDF (Optimized, Skip Converted) ===")
    mode = ask_processing_mode()

    if mode == "manual":
        process_manual_mode()
    elif mode == "csv":
        process_csv_mode()