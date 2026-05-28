import os
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime

def select_directory():
    """Open a dialog for the user to select a directory."""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    folder_selected = filedialog.askdirectory(title="Select root folder to scan")
    root.destroy()
    return folder_selected if folder_selected else None

def count_files(root_folder):
    """Quickly count total files to set the progress bar maximum."""
    total = 0
    for _, _, filenames in os.walk(root_folder):
        total += len(filenames)
    return total

def scan_directory(root_folder, progress_callback=None):
    """
    Walk the root_folder and gather file information while updating progress.
    """
    rows = []
    max_depth = 0
    files_processed = 0

    for dirpath, dirnames, filenames in os.walk(root_folder):
        for filename in filenames:
            full_path = os.path.join(dirpath, filename)
            file_name, file_ext = os.path.splitext(filename)

            try:
                file_size = os.path.getsize(full_path)
            except OSError:
                file_size = 0

            rel_dir = os.path.relpath(dirpath, root_folder)
            folder_list = [] if rel_dir == '.' else rel_dir.split(os.sep)

            if len(folder_list) > max_depth:
                max_depth = len(folder_list)

            rows.append({
                "file_name": file_name,
                "file_ext": file_ext,
                "full_path": full_path,
                "file_size": file_size,
                "folders": folder_list
            })

            # Update Progress
            files_processed += 1
            if progress_callback:
                progress_callback(files_processed, filename)

    return rows, max_depth

def write_csv(rows, max_depth, output_path):
    """Write the collected file data to a CSV file."""
    header = ["File Name", "Full File Path", "File Extension", "File Size"]
    for i in range(1, max_depth + 1):
        header.append(f"Folder {i}")

    with open(output_path, mode="w", newline="", encoding="utf-8") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(header)
        for row in rows:
            csv_row = [row["file_name"], row["full_path"], row["file_ext"], row["file_size"]]
            folders = row["folders"]
            for i in range(max_depth):
                csv_row.append(folders[i] if i < len(folders) else "")
            writer.writerow(csv_row)

def main():
    # 1. Select Directory
    root_folder = select_directory()
    if not root_folder:
        return

    # 2. Setup Progress Window
    progress_root = tk.Tk()
    progress_root.title("Scanning Files...")
    progress_root.geometry("400x150")
    progress_root.attributes('-topmost', True)
    
    label_var = tk.StringVar(value="Counting files...")
    label = tk.Label(progress_root, textvariable=label_var, wraplength=350)
    label.pack(pady=20)

    progress_bar = ttk.Progressbar(progress_root, orient="horizontal", length=300, mode="determinate")
    progress_bar.pack(pady=10)
    
    # Pre-count files for progress bar scale
    progress_root.update()
    total_files = count_files(root_folder)
    progress_bar["maximum"] = total_files
    
    def update_progress(count, current_file):
        progress_bar["value"] = count
        label_var.set(f"Scanning: {current_file}")
        # Crucial: update the UI so it doesn't freeze
        progress_root.update()

    # 3. Perform the Scan
    rows, max_depth = scan_directory(root_folder, update_progress)
    
    # Close progress window
    progress_root.destroy()

    # 4. Save CSV
    if rows:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        output_csv = os.path.join(script_dir, f"directory_report_{timestamp}.csv")
        write_csv(rows, max_depth, output_csv)
        
        # Show Completion
        final_root = tk.Tk()
        final_root.withdraw()
        messagebox.showinfo("Complete", f"Done! Scanned {len(rows)} files.\nCSV saved to: {output_csv}")
        final_root.destroy()

if __name__ == "__main__":
    main()