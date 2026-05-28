import os
import sys
import argparse
from pathlib import Path
import win32com.client
from win32com.client import constants
import pythoncom
import tempfile
import time
from tkinter import Tk, filedialog
import tkinter.messagebox as messagebox
import csv
import json
from datetime import datetime
import traceback
import ctypes
from ctypes import wintypes
import subprocess

# Windows API constants for dismissing dialogs
WM_CLOSE = 0x0010
HWND_TOPMOST = -1
GW_OWNER = 4
GA_ROOT = 2

BASE_DIR = Path(__file__).parent

# CHANGE ME
# Maximum number of files to attempt per run (Task Scheduler-friendly)
BATCH_LIMIT = 2

# CHANGE ME
# Replace "test_batch.csv" with the .csv made from the detect_files.py; STORE THE CSV IN THE SAME FOLDER AS THIS SCRIPT.
CSV_PATH = BASE_DIR / "test_of_sops.csv"

# Load Windows DLLs
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

# Define callback type for EnumWindows
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
    # Common dialog class names
    dialog_classes = ['#32770', 'Dialog', 'NUIDialog', 'bosa_sdm_msword', 'bosa_sdm_Microsoft Office Publisher']
    return class_name in dialog_classes


def find_publisher_dialogs():
    """Find any Publisher dialog windows that might be blocking."""
    dialog_windows = []
    
    def enum_callback(hwnd, lParam):
        if user32.IsWindowVisible(hwnd):
            title = get_window_text(hwnd)
            class_name = get_class_name(hwnd)
            
            # Look for Publisher-related dialogs
            if is_dialog_window(hwnd):
                # Check if it's related to Publisher or Microsoft Office
                if any(keyword in title.lower() for keyword in 
                       ['publisher', 'microsoft', 'missing', 'font', 'picture', 
                        'error', 'warning', 'recovery', 'unable', 'cannot']):
                    dialog_windows.append((hwnd, title, class_name))
            
            # Also look for generic Office dialogs
            if class_name == '#32770':  # Standard Windows dialog
                if title and user32.IsWindowEnabled(hwnd):
                    dialog_windows.append((hwnd, title, class_name))
        return True
    
    user32.EnumWindows(EnumWindowsProc(enum_callback), 0)
    return dialog_windows


def dismiss_dialog(hwnd, click_button=None):
    """
    Dismiss a dialog window.
    
    Args:
        hwnd: Window handle
        click_button: Optional button text to click (e.g., "OK", "Yes", "Cancel")
    """
    try:
        if click_button:
            # Try to find and click a specific button
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
                # Send button click message
                BM_CLICK = 0x00F5
                user32.SendMessageW(button_hwnd, BM_CLICK, 0, 0)
                return True
        
        # Fall back to sending WM_CLOSE
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
        
        # Try different approaches based on dialog type
        if any(word in title.lower() for word in ['ok', 'continue', 'yes']):
            # Dialog with OK/Continue - click it
            if dismiss_dialog(hwnd, "OK") or dismiss_dialog(hwnd, "Yes") or dismiss_dialog(hwnd, "Continue"):
                dismissed_count += 1
        elif 'cancel' in title.lower():
            dismiss_dialog(hwnd, "Cancel")
            dismissed_count += 1
        else:
            # Try common buttons in order of preference
            for button in ["OK", "Yes", "Continue", "Close", "Cancel", "No"]:
                if dismiss_dialog(hwnd, button):
                    dismissed_count += 1
                    break
            else:
                # Last resort - send close message
                dismiss_dialog(hwnd)
                dismissed_count += 1
        
        time.sleep(0.2)  # Brief pause between dismissals
    
    return dismissed_count


def kill_publisher_process():
    """Kill any running Publisher processes."""
    try:
        subprocess.run(['taskkill', '/F', '/IM', 'MSPUB.EXE'], 
                      capture_output=True, timeout=5)
        time.sleep(1)  # Give it time to fully close
        return True
    except Exception as e:
        print(f"    Warning: Could not kill Publisher process: {e}")
        return False


def kill_powerpoint_process():
    """Kill any running PowerPoint processes."""
    try:
        subprocess.run(['taskkill', '/F', '/IM', 'POWERPNT.EXE'], 
                      capture_output=True, timeout=5)
        time.sleep(1)  # Give it time to fully close
        return True
    except Exception as e:
        print(f"    Warning: Could not kill PowerPoint process: {e}")
        return False


def is_modal_dialog_error(error):
    """Check if an error is due to a modal dialog being active."""
    error_str = str(error).lower()
    return 'modal dialog' in error_str or '-2147221457' in error_str

class ConversionLogger:
    """Handles logging for Publisher to PowerPoint conversions."""
    
    def __init__(self, log_directory=None):
        """
        Initialize the logger.
        
        Args:
            log_directory: Directory to save log files. 
                          Default: "C:/ConversionLogs" (REPLACE THIS PATH)
        """
        if log_directory is None:
            # REPLACE THIS PATH WITH YOUR DESIRED LOG DIRECTORY
            log_directory = BASE_DIR = Path(__file__).parent / "ConversionLogs"
        
        self.log_dir = Path(log_directory)
        self.log_dir.mkdir(parents=True, exist_ok=True)
        
        # Generate run-specific log filename
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        self.run_log_path = self.log_dir / f"conversion_log_{timestamp}.txt"
        
        # Master log file path
        self.master_log_path = self.log_dir / "master_conversion_log.txt"
        
        # Maximum master log size before rotation (10 MB)
        self.max_master_size = 10 * 1024 * 1024
        
        # Buffer for collecting all console output
        self.console_buffer = []
        
        # Initialize run log
        self.run_log_data = {
            "session_start": datetime.now().isoformat(),
            "files": [],
            "summary": {
                "total": 0,
                "success": 0,
                "failed": 0,
                "skipped": 0
            }
        }
        
        self._write_header()
    
    def _write_header(self):
        """Write header to run log file."""
        header = "="*80 + "\n"
        header += "PUBLISHER TO POWERPOINT CONVERSION LOG\n"
        header += "="*80 + "\n"
        header += f"Session Start: {self.run_log_data['session_start']}\n"
        header += "="*80 + "\n\n"
        
        with open(self.run_log_path, 'w', encoding='utf-8') as f:
            f.write(header)
        
        # Also add to console buffer
        self.console_buffer.append(header)
    
    def print_and_log(self, message):
        """
        Print message to console AND write to log file.
        This replaces all print() calls to ensure logs capture everything.
        
        Args:
            message: Message to print and log
        """
        # Print to console
        print(message)
        
        # Add to buffer
        if not message.endswith('\n'):
            message = message + '\n'
        self.console_buffer.append(message)
        
        # Write to run log immediately
        self._append_to_run_log(message)
    
    def log_file_start(self, file_path):
        """Log the start of a file conversion."""
        msg = f"\n{'─'*80}\nProcessing: {file_path}\nStart Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        self._append_to_run_log(msg)
        print(msg)
    
    def log_file_complete(self, file_path, status, output_path=None, error_msg=None, 
                         stats=None):
        """
        Log completion of a file conversion.
        
        Args:
            file_path: Source Publisher file path
            status: 'SUCCESS', 'FAILED', or 'SKIPPED'
            output_path: Path to output PPTX file (if successful)
            error_msg: Error message (if failed/skipped)
            stats: Dictionary with conversion statistics
        """
        end_time = datetime.now()
        
        # Update summary counts
        self.run_log_data["summary"]["total"] += 1
        if status == "SUCCESS":
            self.run_log_data["summary"]["success"] += 1
        elif status == "FAILED":
            self.run_log_data["summary"]["failed"] += 1
        elif status == "SKIPPED":
            self.run_log_data["summary"]["skipped"] += 1
        
        # Build file entry
        file_entry = {
            "file": str(file_path),
            "status": status,
            "end_time": end_time.isoformat(),
            "output": str(output_path) if output_path else None,
            "error": error_msg if error_msg else None,
            "stats": stats if stats else {}
        }
        
        self.run_log_data["files"].append(file_entry)
        
        # Format log message
        msg = f"Status: [{status}]\n"
        msg += f"End Time: {end_time.strftime('%Y-%m-%d %H:%M:%S')}\n"
        
        if output_path:
            msg += f"Output: {output_path}\n"
        
        if error_msg:
            msg += f"Error Details:\n{error_msg}\n"
        
        if stats:
            msg += f"Statistics:\n"
            for key, value in stats.items():
                msg += f"  {key}: {value}\n"
        
        self._append_to_run_log(msg)
        print(msg)
    
    def log_message(self, message, level="INFO"):
        """Log a general message."""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        msg = f"[{timestamp}] [{level}] {message}\n"
        self._append_to_run_log(msg)
        print(message)
    
    def _append_to_run_log(self, message):
        """Append message to run log file."""
        with open(self.run_log_path, 'a', encoding='utf-8') as f:
            f.write(message)
    
    def finalize(self):
        """Finalize the log session and update master log."""
        end_time = datetime.now()
        self.run_log_data["session_end"] = end_time.isoformat()
        
        # Write summary to run log
        summary_msg = "\n" + "="*80 + "\n"
        summary_msg += "CONVERSION SUMMARY\n"
        summary_msg += "="*80 + "\n"
        summary_msg += f"Session Start: {self.run_log_data['session_start']}\n"
        summary_msg += f"Session End: {end_time.isoformat()}\n"
        summary_msg += f"Total Files: {self.run_log_data['summary']['total']}\n"
        summary_msg += f"Successful: {self.run_log_data['summary']['success']}\n"
        summary_msg += f"Failed: {self.run_log_data['summary']['failed']}\n"
        summary_msg += f"Skipped: {self.run_log_data['summary']['skipped']}\n"
        summary_msg += "="*80 + "\n"
        
        self._append_to_run_log(summary_msg)
        print(summary_msg)
        
        # Save JSON version for programmatic access
        json_log_path = self.run_log_path.with_suffix('.json')
        with open(json_log_path, 'w', encoding='utf-8') as f:
            json.dump(self.run_log_data, f, indent=2)
        
        # Append to master log
        self._append_to_master_log(summary_msg)
        
        print(f"\nLogs saved to:")
        print(f"  Run Log: {self.run_log_path}")
        print(f"  JSON Log: {json_log_path}")
        print(f"  Master Log: {self.master_log_path}")
    
    def _append_to_master_log(self, summary):
        """Append summary to master log with rotation if needed."""
        # Check if master log needs rotation
        if self.master_log_path.exists():
            if self.master_log_path.stat().st_size >= self.max_master_size:
                self._rotate_master_log()
        
        # Append to master log
        with open(self.master_log_path, 'a', encoding='utf-8') as f:
            f.write(summary)
            f.write("\n")
    
    def _rotate_master_log(self):
        """Rotate master log when it gets too large."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        archive_name = f"master_conversion_log_archive_{timestamp}.txt"
        archive_path = self.log_dir / archive_name
        
        self.master_log_path.rename(archive_path)
        print(f"\nMaster log rotated to: {archive_path}")


class PublisherToPowerPointConverter:
    """Convert Microsoft Publisher files to PowerPoint preserving editable content."""
    
    def __init__(self, preserve_text=True, preserve_images=True, preserve_shapes=True, 
                 template_path=None, verbose=False, dry_run=False, logger=None):
        self.publisher = None
        self.powerpoint = None
        self.preserve_text = preserve_text
        self.preserve_images = preserve_images
        self.preserve_shapes = preserve_shapes
        self.temp_dir = None
        self.template_path = template_path
        self.verbose = verbose
        self.dry_run = dry_run
        self.logger = logger
        self.current_file_stats = {}
        self._publisher_needs_restart = False
        self._com_initialized = False
    
    def _print(self, message):
        """
        Helper method to print and log simultaneously.
        Routes all output through logger if available.
        """
        if self.logger:
            self.logger.print_and_log(message)
        else:
            print(message)
        
    def __enter__(self):
            """Initialize COM objects."""
            if self.dry_run:
                print("DRY RUN MODE - No files will be created or modified")
                return self
                
            try:
                pythoncom.CoInitialize()
                self._com_initialized = True
                
                # --- UPDATED DEEP-CLEANING CACHE LOGIC ---
                try:
                    self._init_publisher()
                    self._init_powerpoint()
                except AttributeError as e:
                    # Catch the specific gen_py cache corruption errors
                    if "win32com.gen_py" in str(e) or "CLSIDToClassMap" in str(e) or "CLSIDToPackageMap" in str(e):
                        print("\n[!] Detected corrupted Microsoft Office COM cache.")
                        print("    Attempting a deep clean of the cache and retrying...")
                        
                        import win32com
                        import shutil
                        import os
                        import sys
                        
                        try:
                            # 1. Remove the folder from the hard drive
                            cache_path = win32com.__gen_path__
                            if cache_path and os.path.exists(cache_path):
                                shutil.rmtree(cache_path, ignore_errors=True)
                            
                            # 2. Clear the corrupted modules from Python's active memory
                            modules_to_remove = [mod for mod in sys.modules if mod.startswith("win32com.gen_py")]
                            for mod in modules_to_remove:
                                del sys.modules[mod]
                                
                            print("    Cache wiped from disk and memory. Retrying initialization...\n")
                            
                            # Brief pause to let the filesystem catch up, then retry
                            time.sleep(1)
                            self._init_publisher()
                            self._init_powerpoint()
                            
                        except Exception as clear_err:
                            print(f"    Warning: Failed to clear cache: {clear_err}")
                            raise e  # Re-raise the original error if cleanup fails
                    else:
                        raise e  # Re-raise if it's a different AttributeError
                # --- END NEW LOGIC ---
                
                # Create a temporary directory for image exports
                self.temp_dir = Path(tempfile.mkdtemp(prefix="pub2ppt_"))
                print(f"Using temporary directory: {self.temp_dir}")
                
                return self
                
            except Exception as e:
                error_msg = f"Error initializing Microsoft Office applications: {e}\n"
                error_msg += "Make sure Microsoft Publisher and PowerPoint are installed."
                print(error_msg)
                sys.exit(1)
    
    def _init_publisher(self):
        """Initialize or reinitialize Publisher COM object."""
        if self.publisher:
            try:
                # Try to close existing Publisher gracefully
                self.publisher.Quit()
            except:
                pass
            self.publisher = None
            time.sleep(0.5)
        
        # Dismiss any lingering dialogs before starting
        dismiss_all_publisher_dialogs(verbose=self.verbose)
        
        self.publisher = win32com.client.Dispatch("Publisher.Application")
        self._publisher_needs_restart = False
    
    def _init_powerpoint(self):
        """Initialize PowerPoint COM object."""
        if self.powerpoint:
            try:
                self.powerpoint.Quit()
            except:
                pass
            self.powerpoint = None
            time.sleep(0.5)
        
        self.powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        self.powerpoint.Visible = True
    
    def _restart_publisher(self):
        """Restart Publisher to recover from modal dialog state."""
        print("    Restarting Publisher to recover from error...")
        if self.logger:
            self.logger.log_message("Restarting Publisher to recover from modal dialog")
        
        # First try to dismiss any dialogs
        dismissed = dismiss_all_publisher_dialogs(verbose=self.verbose)
        if dismissed > 0:
            print(f"    Dismissed {dismissed} dialog(s)")
        
        # Try to close Publisher gracefully
        try:
            if self.publisher:
                self.publisher.Quit()
        except:
            pass
        
        self.publisher = None
        time.sleep(0.5)
        
        # If that didn't work, force kill
        try:
            # Check if dialogs are still present
            remaining_dialogs = find_publisher_dialogs()
            if remaining_dialogs:
                print("    Force-killing Publisher process...")
                kill_publisher_process()
                time.sleep(1)
        except:
            pass
        
        # Reinitialize
        try:
            self._init_publisher()
            print("    Publisher restarted successfully")
            return True
        except Exception as e:
            print(f"    Warning: Failed to restart Publisher: {e}")
            return False
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Clean up COM objects."""
        if self.dry_run:
            return
        
        # First, try to dismiss any modal dialogs
        try:
            dismissed = dismiss_all_publisher_dialogs(verbose=self.verbose)
            if dismissed > 0 and self.verbose:
                print(f"Dismissed {dismissed} dialog(s) during cleanup")
        except:
            pass
        
        # Try to close Publisher gracefully
        if self.publisher:
            try:
                self.publisher.Quit()
            except Exception as e:
                if is_modal_dialog_error(e):
                    # Modal dialog blocking - try to dismiss and retry
                    print("Modal dialog detected during cleanup, attempting to dismiss...")
                    dismiss_all_publisher_dialogs()
                    time.sleep(0.5)
                    try:
                        self.publisher.Quit()
                    except:
                        # Force kill as last resort
                        print("Force-killing Publisher process...")
                        kill_publisher_process()
                else:
                    # Other error - try force kill
                    try:
                        kill_publisher_process()
                    except:
                        pass
        
        # Close PowerPoint
        if self.powerpoint:
            try:
                # First close any open presentations without saving
                try:
                    for i in range(self.powerpoint.Presentations.Count, 0, -1):
                        try:
                            self.powerpoint.Presentations(i).Close()
                        except:
                            pass
                except:
                    pass
                
                # Now quit PowerPoint
                self.powerpoint.Quit()
                print("PowerPoint closed successfully.")
            except Exception as e:
                print(f"Warning: Could not close PowerPoint gracefully: {e}")
                # Force kill as last resort
                try:
                    kill_powerpoint_process()
                    print("PowerPoint force-closed.")
                except:
                    pass
        
        # Force-kill both processes to ensure they are fully dead.
        # Without this, the COM proxy Release() calls below can hang
        # waiting on an RPC channel to a half-exited process.
        kill_publisher_process()
        kill_powerpoint_process()

        # Clean up temporary directory
        if self.temp_dir and self.temp_dir.exists():
            import shutil
            try:
                shutil.rmtree(self.temp_dir)
            except:
                pass
        
        # Uninitialize COM
        if self._com_initialized:
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    
    def get_shape_type_name(self, shape_type):
        """Get human-readable shape type name."""
        shape_types = {
            1: "TextBox",
            2: "Picture",
            5: "Rectangle",
            9: "Line",
            12: "Shape",
            13: "Picture",
            17: "Table",
            18: "Table",
            19: "Table"
        }
        return shape_types.get(shape_type, f"Unknown({shape_type})")
    
    def copy_table_to_powerpoint(self, pub_table, ppt_slide, left, top, width, height):
        """Copy a Publisher table to PowerPoint, preserving structure, content, and merged cells."""
        try:
            # Get table dimensions
            num_rows = pub_table.Rows.Count
            num_cols = pub_table.Columns.Count
            
            print(f"      Table: {num_rows} rows x {num_cols} columns")
            
            # Calculate actual table dimensions from Publisher
            actual_width = 0
            actual_height = 0
            
            try:
                # Sum up actual column widths
                for col_idx in range(1, num_cols + 1):
                    actual_width += pub_table.Columns(col_idx).Width
                
                # Sum up actual row heights
                for row_idx in range(1, num_rows + 1):
                    actual_height += pub_table.Rows(row_idx).Height
                
                print(f"      Original table size: {actual_width:.1f} x {actual_height:.1f} points")
            except:
                # Fallback to scaled dimensions if we can't get actual sizes
                actual_width = width
                actual_height = height
            
            # ================================================================
            # PHASE 1: Detect merged cells in Publisher table
            # ================================================================
            # In Publisher, cell.Width and cell.Height return the SPAN values
            # (number of columns/rows the cell spans), NOT dimensions in points!
            
            # Build a map of merged cell regions
            # Key: (start_row, start_col), Value: (row_span, col_span, cell_object)
            merged_regions = {}
            # Track which cells are part of a merge (but not the top-left origin)
            cells_in_merge = set()
            
            try:
                # Iterate through all cells to detect merges using span values
                for row_idx in range(1, num_rows + 1):
                    for col_idx in range(1, num_cols + 1):
                        if (row_idx, col_idx) in cells_in_merge:
                            continue  # Skip cells already identified as part of a merge
                        
                        try:
                            # Access the cell via Rows().Cells() - correct for Publisher
                            pub_cell = None
                            try:
                                pub_row = pub_table.Rows(row_idx)
                                pub_cell = pub_row.Cells(col_idx)
                            except:
                                # Fallback: Try flat index
                                try:
                                    flat_index = (row_idx - 1) * num_cols + col_idx
                                    pub_cell = pub_table.Cells.Item(flat_index)
                                except:
                                    continue
                            
                            if pub_cell is None:
                                continue
                            
                            # Get span values from Width/Height (these ARE the spans in Publisher!)
                            col_span = 1
                            row_span = 1
                            
                            try:
                                cell_width_span = pub_cell.Width
                                if cell_width_span is not None and cell_width_span > 1:
                                    col_span = int(cell_width_span)
                            except:
                                pass
                            
                            try:
                                cell_height_span = pub_cell.Height
                                if cell_height_span is not None and cell_height_span > 1:
                                    row_span = int(cell_height_span)
                            except:
                                pass
                            
                            # If we detected a merge, record it
                            if row_span > 1 or col_span > 1:
                                merged_regions[(row_idx, col_idx)] = (row_span, col_span, pub_cell)
                                # Mark all cells in this merge region
                                for r in range(row_idx, row_idx + row_span):
                                    for c in range(col_idx, col_idx + col_span):
                                        if (r, c) != (row_idx, col_idx):
                                            cells_in_merge.add((r, c))
                                
                                if self.verbose:
                                    print(f"        Detected merged cell at ({row_idx},{col_idx}): {row_span} rows x {col_span} cols")
                        
                        except Exception as e:
                            if self.verbose:
                                print(f"        Could not check cell ({row_idx}, {col_idx}) for merge: {e}")
                            continue
                
                if merged_regions:
                    print(f"      Detected {len(merged_regions)} merged cell region(s)")
                    
            except Exception as e:
                if self.verbose:
                    print(f"      Warning: Could not fully detect merged cells: {e}")
            
            # ================================================================
            # PHASE 2: Create PowerPoint table
            # ================================================================
            # Read Publisher column widths BEFORE creating the table so we can
            # pass the exact total width to AddTable (prevents equal-distribution)
            pub_col_widths = {}
            try:
                for col_idx in range(1, num_cols + 1):
                    pub_col_widths[col_idx] = pub_table.Columns(col_idx).Width
                print(f"      Publisher column widths: {pub_col_widths}")
            except:
                pass
            # Use sum of Publisher column widths if available, otherwise fall back
            total_pub_width = sum(pub_col_widths.values()) if pub_col_widths else actual_width
            ppt_table_shape = ppt_slide.Shapes.AddTable(
                NumRows=num_rows,
                NumColumns=num_cols,
                Left=left,
                Top=top,
                Width=total_pub_width,
                Height=actual_height
            )
            
            ppt_table = ppt_table_shape.Table
            
            # Remove all default table formatting for speed
            try:
                # Apply "No Style, No Grid" - completely blank table
                ppt_table_shape.Table.ApplyStyle("{5940675A-B579-460E-94D1-54222C63F5DA}")
            except:
                # If style application fails, manually remove borders
                try:
                    for row_idx in range(1, num_rows + 1):
                        for col_idx in range(1, num_cols + 1):
                            cell = ppt_table.Cell(row_idx, col_idx)
                            # Make all borders invisible by default
                            for border_idx in range(1, 5):  # 1=top, 2=left, 3=bottom, 4=right
                                try:
                                    cell.Borders(border_idx).Visible = False
                                except:
                                    pass
                except:
                    pass
            
            # pub_col_widths was already read before AddTable.
            # Let ApplyStyle's async relayout settle before touching widths
            time.sleep(0.2)
            if pub_col_widths:
                # Log widths for debugging
                if self.logger:
                    self.logger.log_message(f"Publisher column widths: {pub_col_widths}")
                # Lock table shape width to exact sum
                try:
                    ppt_table_shape.Width = total_pub_width
                except:
                    pass
                # Set columns left-to-right (3 passes -- PowerPoint
                # redistributes from neighbours on each set, so repeated
                # passes converge toward the target widths)
                for _pass in range(3):
                    for col_idx in range(1, num_cols + 1):
                        if col_idx in pub_col_widths:
                            try:
                                ppt_table.Columns(col_idx).Width = pub_col_widths[col_idx]
                            except:
                                pass  

            # ================================================================
            # PHASE 3: Copy cell contents and formatting FIRST
            # ================================================================
            # We copy content BEFORE merging because PowerPoint's Merge() 
            # keeps only the origin cell's content
            #
            # IMPORTANT: Publisher's Cells collection may index differently when
            # there are merged cells. We need to track the actual grid position.
            
            # Build a map of Publisher cell content by grid position
            # This handles the case where Publisher's cell indices don't match grid positions
            pub_cell_content = {}  # (row, col) -> cell data
            
            for row_idx in range(1, num_rows + 1):
                try:
                    pub_row = pub_table.Rows(row_idx)
                    cells_in_row = pub_row.Cells
                    num_cells_in_row = cells_in_row.Count
                    
                    # Track the current grid column as we iterate through cells
                    grid_col = 1
                    
                    for cell_idx in range(1, num_cells_in_row + 1):
                        try:
                            pub_cell = cells_in_row(cell_idx)
                            
                            # Get the cell's span
                            col_span = 1
                            row_span = 1
                            try:
                                col_span = int(pub_cell.Width) if pub_cell.Width > 1 else 1
                            except:
                                pass
                            try:
                                row_span = int(pub_cell.Height) if pub_cell.Height > 1 else 1
                            except:
                                pass
                            
                            # Store the cell content at the current grid position
                            pub_cell_content[(row_idx, grid_col)] = {
                                'cell': pub_cell,
                                'col_span': col_span,
                                'row_span': row_span
                            }
                            
                            if self.verbose:
                                try:
                                    text_preview = pub_cell.TextRange.Text[:20].strip()
                                    print(f"        Publisher cell at grid ({row_idx},{grid_col}): '{text_preview}' span={row_span}x{col_span}")
                                except:
                                    pass
                            
                            # Advance grid column by the cell's span
                            grid_col += col_span
                            
                        except Exception as e:
                            if self.verbose:
                                print(f"        Could not read cell {cell_idx} in row {row_idx}: {e}")
                            grid_col += 1  # Assume single column if we can't read
                            
                except Exception as e:
                    if self.verbose:
                        print(f"        Could not process row {row_idx}: {e}")
            
            # Now copy content from Publisher to PowerPoint using grid positions
            for row_idx in range(1, num_rows + 1):
                for col_idx in range(1, num_cols + 1):
                    try:
                        # Get Publisher cell content for this grid position
                        if (row_idx, col_idx) not in pub_cell_content:
                            # This cell is part of a merged region (not the origin)
                            continue
                        
                        cell_data = pub_cell_content[(row_idx, col_idx)]
                        pub_cell = cell_data['cell']
                        
                        # Get the PowerPoint cell at the same grid position
                        try:
                            ppt_cell = ppt_table.Cell(row_idx, col_idx)
                        except:
                            continue

                        # --- FAST PATH: Skip empty cells entirely ---
                        # Check if cell has text or visible fill; if neither, skip
                        # all formatting work to avoid unnecessary COM round-trips
                        cell_is_empty = True
                        try:
                            cell_text_check = pub_cell.TextRange.Text
                            if cell_text_check and cell_text_check.strip():
                                cell_is_empty = False
                        except:
                            pass
                        
                        if cell_is_empty:
                            # Check if the cell has a visible non-white fill we need to copy
                            try:
                                if pub_cell.Fill.Visible and pub_cell.Fill.Type == 1:
                                    fill_rgb = pub_cell.Fill.ForeColor.RGB
                                    if fill_rgb != 16777215:  # Skip white (default)
                                        cell_is_empty = False
                            except:
                                pass
                        
                        if cell_is_empty:
                            continue
                        # --- END FAST PATH ---
                        
                        # Get references once to minimize COM calls
                        try:
                            ppt_shape = ppt_cell.Shape
                            ppt_text_frame = ppt_shape.TextFrame
                            ppt_text_range = ppt_text_frame.TextRange
                            pub_text_range = pub_cell.TextRange
                        except:
                            continue
                        
                        # Copy text content (stripped of trailing whitespace)
                        try:
                            cell_text = pub_text_range.Text
                            cell_text = cell_text.rstrip('\r\n\x0b\x0c')
                            ppt_text_range.Text = cell_text
                        except:
                            pass
                        
                        # Copy font properties: name, size, bold, italic (NO text color)
                        try:
                            pub_font = pub_text_range.Font
                            ppt_font = ppt_text_range.Font
                            ppt_font.Name = pub_font.Name
                            ppt_font.Size = pub_font.Size
                            ppt_font.Bold = pub_font.Bold
                            ppt_font.Italic = pub_font.Italic
                        except:
                            pass
                        
                        # Copy paragraph alignment
                        try:
                            pub_para = pub_text_range.ParagraphFormat
                            ppt_para = ppt_text_range.ParagraphFormat
                            pub_align = pub_para.Alignment
                            align_map = {0: 1, 1: 2, 2: 3, 3: 4}
                            ppt_para.Alignment = align_map.get(pub_align, pub_align + 1)
                        except:
                            pass

                        
                        # Copy text frame margins
                        try:
                            pub_tf = pub_cell.TextFrame
                            ppt_text_frame.MarginLeft = pub_tf.MarginLeft
                            ppt_text_frame.MarginRight = pub_tf.MarginRight
                            ppt_text_frame.MarginTop = pub_tf.MarginTop
                            ppt_text_frame.MarginBottom = pub_tf.MarginBottom
                        except:
                            pass
                        
                        # Copy cell fill/background color (skip white -- it's the default)
                        try:
                            if pub_cell.Fill.Visible and pub_cell.Fill.Type == 1:
                                fill_rgb = pub_cell.Fill.ForeColor.RGB
                                if fill_rgb != 16777215:  # Not white
                                    ppt_shape.Fill.Visible = True
                                    ppt_shape.Fill.ForeColor.RGB = fill_rgb
                        except:
                            pass
                        
                    except Exception as e:
                        if self.verbose:
                            print(f"      Warning: Could not copy cell ({row_idx}, {col_idx}): {e}")
                        continue
            
            # Re-enforce column widths after content (auto-fit may have shifted them)
            for col_idx, col_w in pub_col_widths.items():
                try:
                    ppt_table.Columns(col_idx).Width = col_w
                except:
                    pass
            
            # Set exact row heights from Publisher
            try:
                for row_idx in range(1, num_rows + 1):
                    try:
                        pub_row_height = pub_table.Rows(row_idx).Height
                        ppt_table.Rows(row_idx).Height = pub_row_height
                    except Exception as e:
                        if self.verbose:
                            print(f"        Warning: Could not set row {row_idx} height: {e}")
            except:
                pass
            
            # ================================================================
            # PHASE 4: Apply merged cells to PowerPoint table AFTER content
            # ================================================================
            # We merge cells AFTER copying content because PowerPoint's Merge() 
            # keeps only the origin cell's content - all other content is preserved
            # because we already copied it to the origin cell position
            
            if merged_regions:
                print(f"      Applying {len(merged_regions)} merge(s) to PowerPoint table...")
            
            for (start_row, start_col), (row_span, col_span, pub_cell) in merged_regions.items():
                try:
                    # Get the starting cell
                    ppt_start_cell = ppt_table.Cell(start_row, start_col)
                    
                    # Get the ending cell (bottom-right of merge region)
                    end_row = start_row + row_span - 1
                    end_col = start_col + col_span - 1
                    ppt_end_cell = ppt_table.Cell(end_row, end_col)
                    
                    print(f"        Merging cells ({start_row},{start_col}) to ({end_row},{end_col}) - {row_span} rows x {col_span} cols")
                    
                    # Merge the cells in PowerPoint
                    ppt_start_cell.Merge(ppt_end_cell)
                    
                    print(f"        ✓ Merge successful")
                        
                except Exception as e:
                    # ALWAYS show merge errors - they're important!
                    print(f"        ✗ ERROR merging cells ({start_row},{start_col}): {e}")
            

            # ================================================================
            # PHASE 5: Re-enforce column widths and row heights AFTER merges
            # ================================================================
            # PowerPoint's Merge() and content auto-fit can redistribute columns
            # and expand rows. Lock table shape width then re-apply exact
            # Publisher dimensions.

            # Re-enforce column widths after merges (same technique as Phase 2)
            if pub_col_widths:
                try:
                    ppt_table_shape.Width = total_pub_width
                except:
                    pass
                for _pass in range(3):
                    for col_idx in range(1, num_cols + 1):
                        if col_idx in pub_col_widths:
                            try:
                                ppt_table.Columns(col_idx).Width = pub_col_widths[col_idx]
                            except:
                                pass            

            # Re-enforce row heights
            try:
                for row_idx in range(1, num_rows + 1):
                    try:
                        pub_row_height = pub_table.Rows(row_idx).Height
                        ppt_table.Rows(row_idx).Height = pub_row_height
                    except Exception as e:
                        if self.verbose:
                            print(f"        Warning: Could not re-enforce row {row_idx} height: {e}")
            except:
                pass

            # Report success with merge info
            if merged_regions:
                print(f"      Table conversion successful! ({len(merged_regions)} merged region(s) preserved)")
            else:
                print(f"      Table conversion successful!")
            return ppt_table_shape
            
        except Exception as e:
            print(f"      Error copying table: {e}")
            return None
    
    def copy_text_properties(self, pub_text_range, ppt_text_range):
        """Copy text formatting from Publisher to PowerPoint."""
        try:
            # Font properties - copy each individually to handle failures gracefully
            try:
                ppt_text_range.Font.Name = pub_text_range.Font.Name
            except:
                pass
            
            try:
                ppt_text_range.Font.Size = pub_text_range.Font.Size
            except:
                pass
            
            try:
                ppt_text_range.Font.Bold = pub_text_range.Font.Bold
            except:
                pass
            
            try:
                ppt_text_range.Font.Italic = pub_text_range.Font.Italic
            except:
                pass
            
            # Underline - Publisher and PowerPoint use different constants
            # Publisher: 0=None, 1=Single, etc.
            # PowerPoint: msoNoUnderline=-4142, msoUnderlineSingleLine=1, etc.
            # Direct assignment can cause "value out of range" errors
            try:
                pub_underline = pub_text_range.Font.Underline
                # Map common underline values - use boolean for maximum compatibility
                if pub_underline and pub_underline != 0:
                    ppt_text_range.Font.Underline = True  # Use boolean for safety
                else:
                    ppt_text_range.Font.Underline = False
            except:
                pass
            
            # Color
            try:
                ppt_text_range.Font.Color.RGB = pub_text_range.Font.Color.RGB
            except:
                pass
                
        except Exception as e:
            if self.verbose:
                print(f"    Warning: Could not copy all text properties: {e}")
    
    def copy_alignment_properties(self, pub_text_frame, ppt_text_frame):
        """
        Copy text alignment properties from Publisher to PowerPoint.
        Handles both horizontal paragraph alignment and vertical text frame alignment.
        
        CRITICAL: Publisher uses 0-based indexing, PowerPoint uses 1-based!
        """
        try:
            # Copy vertical alignment (text frame level)
            try:
                pub_vert_align = pub_text_frame.VerticalTextAlignment
                
                # Publisher vertical alignment (0-based):
                # 0 = Top (or default)
                # 1 = Middle
                # 2 = Bottom
                
                # PowerPoint vertical alignment (different constants):
                # ppAnchorTop = 1
                # ppAnchorMiddle = 3
                # ppAnchorBottom = 4
                
                vertical_alignment_map = {
                    0: 1,  # Top/Default -> ppAnchorTop
                    1: 3,  # Middle -> ppAnchorMiddle
                    2: 4   # Bottom -> ppAnchorBottom
                }
                
                if pub_vert_align in vertical_alignment_map:
                    ppt_text_frame.VerticalAnchor = vertical_alignment_map[pub_vert_align]
                    if self.verbose:
                        vert_names = {0: "Top", 1: "Middle", 2: "Bottom"}
                        print(f"      Set vertical alignment: {vert_names.get(pub_vert_align, pub_vert_align)}")
                        
            except Exception as e:
                if self.verbose:
                    print(f"      Could not copy vertical alignment: {e}")
            
            # Copy horizontal alignment (paragraph level)
            try:
                # Get the Publisher text range to access paragraph properties
                pub_text_range = pub_text_frame.TextRange
                
                # Publisher horizontal alignment (0-based):
                # 0 = Left
                # 1 = Center
                # 2 = Right
                # 3 = Justify
                
                # PowerPoint horizontal alignment (1-based):
                # ppAlignLeft = 1
                # ppAlignCenter = 2
                # ppAlignRight = 3
                # ppAlignJustify = 4
                
                # Try to get paragraph alignment from Publisher
                try:
                    pub_horiz_align = pub_text_range.ParagraphFormat.Alignment
                    
                    # Map Publisher (0-based) to PowerPoint (1-based) by adding 1
                    horizontal_alignment_map = {
                        0: 1,  # Left -> ppAlignLeft
                        1: 2,  # Center -> ppAlignCenter
                        2: 3,  # Right -> ppAlignRight
                        3: 4   # Justify -> ppAlignJustify
                    }
                    
                    if pub_horiz_align in horizontal_alignment_map:
                        ppt_text_frame.TextRange.ParagraphFormat.Alignment = horizontal_alignment_map[pub_horiz_align]
                        
                        if self.verbose:
                            horiz_names = {0: "Left", 1: "Center", 2: "Right", 3: "Justify"}
                            print(f"      Set horizontal alignment: {horiz_names.get(pub_horiz_align, pub_horiz_align)}")
                    else:
                        if self.verbose:
                            print(f"      Unknown alignment value: {pub_horiz_align}, defaulting to left")
                        
                except Exception as e:
                    if self.verbose:
                        print(f"      Could not copy horizontal alignment: {e}")
                        
            except Exception as e:
                if self.verbose:
                    print(f"      Could not access paragraph alignment: {e}")
                    
        except Exception as e:
            if self.verbose:
                print(f"    Warning: Could not copy alignment properties: {e}")
    
    def convert_shape_to_powerpoint(self, pub_shape, ppt_slide, page_width, page_height, slide_width, slide_height):
        """Convert a Publisher shape to PowerPoint, preserving editability."""
        try:
            # Calculate scaling factors
            scale_x = slide_width / page_width
            scale_y = slide_height / page_height
            
            # Get position and size (in points)
            # For text boxes, use exact Publisher values to preserve text layout
            # For other shapes, scale to fit the slide
            if pub_shape.HasTextFrame:
                # Text boxes: NO SCALING - use exact Publisher values
                left = pub_shape.Left
                top = pub_shape.Top
                width = pub_shape.Width
                height = pub_shape.Height
            else:
                # Other shapes: SCALE to fit slide dimensions
                left = pub_shape.Left * scale_x
                top = pub_shape.Top * scale_y
                width = pub_shape.Width * scale_x
                height = pub_shape.Height * scale_y

            shape_type = pub_shape.Type
            shape_name = self.get_shape_type_name(shape_type)
            
            # Debug: Print shape type for all shapes
            if self.verbose:
                print(f"      Shape type: {shape_type} ({shape_name})")
                if self.logger:
                    self.logger.log_message(f"Shape type: {shape_type} ({shape_name})")
            
            shape_converted = False
            
            # Handle tables - Check if shape has a Table property
            try:
                if hasattr(pub_shape, 'Table') and pub_shape.Table is not None:
                    print(f"      Found table object, attempting to convert...")
                    if self.logger:
                        self.logger.log_message("Found table object, attempting to convert...")
                    
                    ppt_table_shape = self.copy_table_to_powerpoint(
                        pub_shape.Table,
                        ppt_slide,
                        left, top, width, height
                    )

                    if ppt_table_shape:
                        try:
                            ppt_table_shape.Rotation = pub_shape.Rotation
                        except:
                            pass
                        shape_converted = True
                        self.current_file_stats["tables"] = self.current_file_stats.get("tables", 0) + 1

            except Exception as e:
                if self.verbose:
                    print(f"      Table check failed: {e}")
                    if self.logger:
                        self.logger.log_message(f"Table check failed: {e}", "WARNING")
            
            # Alternative table check using Type
            if not shape_converted and shape_type in [17, 18, 19]:  # pbTable types
                try:
                    if self.verbose:
                        print(f"      Shape type indicates table, accessing Table property...")
                        if self.logger:
                            self.logger.log_message("Shape type indicates table, accessing Table property...")
                    
                    table_obj = pub_shape.Table
                    ppt_table_shape = self.copy_table_to_powerpoint(
                        table_obj,
                        ppt_slide,
                        left, top, width, height
                    )
                    
                    if ppt_table_shape:
                        try:
                            ppt_table_shape.Rotation = pub_shape.Rotation
                        except:
                            pass
                        shape_converted = True
                        self.current_file_stats["tables"] = self.current_file_stats.get("tables", 0) + 1

                except Exception as e:
                    if self.verbose:
                        print(f"      Warning: Could not access table from shape type {shape_type}: {e}")
                        if self.logger:
                            self.logger.log_message(f"Could not access table from shape type {shape_type}: {e}", "WARNING")
            
            # If we already converted a table, return success
            if shape_converted:
                return True
            
            # Handle group shapes - recursively convert each child
            if shape_type == 6:  # msoGroup
                try:
                    group_items = pub_shape.GroupItems
                    group_count = group_items.Count
                    
                    if self.verbose:
                        print(f"      Group shape with {group_count} child shape(s)")
                    if self.logger:
                        self.logger.log_message(f"Group shape with {group_count} child shape(s)")
                    
                    children_converted = 0
                    children_failed = 0
                    
                    for child_idx in range(1, group_count + 1):
                        try:
                            child_shape = group_items(child_idx)
                            if self.convert_shape_to_powerpoint(
                                child_shape, ppt_slide,
                                page_width, page_height,
                                slide_width, slide_height
                            ):
                                children_converted += 1
                            else:
                                children_failed += 1
                        except Exception as child_err:
                            children_failed += 1
                            if self.verbose:
                                print(f"        Warning: Failed to convert group child {child_idx}: {child_err}")
                            if self.logger:
                                self.logger.log_message(f"Failed to convert group child {child_idx}: {child_err}", "WARNING")
                    
                    if self.verbose or children_failed > 0:
                        print(f"      Group result: {children_converted} converted, {children_failed} failed")
                    
                    return children_converted > 0
                    
                except Exception as e:
                    print(f"    Warning: Could not process group shape: {e}")
                    if self.logger:
                        self.logger.log_message(f"Could not process group shape: {e}", "WARNING")
                    return False
            
            # Handle text boxes and text frames
            if (pub_shape.HasTextFrame and shape_type not in [1, 2, 5, 6, 9, 13]) and self.preserve_text:
                try:
                    # Safely read text content — may fail for type 1 shapes without a true TextFrame
                    text_content = ""
                    has_text = False
                    try:
                        raw_text = pub_shape.TextFrame.TextRange.Text
                        raw_text = raw_text.rstrip('\r\n\x0b\x0c')

                        if raw_text.strip():
                            text_content = raw_text
                            has_text = True
                    except:
                        pass
                    
                    # Always create text box in PowerPoint — even for empty type 1 shapes
                    ppt_shape = ppt_slide.Shapes.AddTextbox(
                        Orientation=1,  # msoTextOrientationHorizontal
                        Left=left,
                        Top=top,
                        Width=width,
                        Height=height
                    )
                    
                    ppt_text_frame = ppt_shape.TextFrame
                    
                    # IMPORTANT: Disable auto-sizing BEFORE setting text to prevent PowerPoint
                    # from auto-adjusting the dimensions
                    try:
                        ppt_text_frame.AutoSize = 0  # No auto-size
                    except:
                        pass
                    
                    # Temporarily disable WordWrap BEFORE writing text to prevent
                    # PowerPoint from wrapping+clipping in tight text boxes
                    pub_word_wrap = -1  # default to wrap ON
                    try:
                        pub_word_wrap = pub_shape.TextFrame.WordWrap
                    except:
                        pass
                    try:
                        ppt_text_frame.WordWrap = 0  # OFF during text write
                    except:
                        pass
                    
                    # Copy text content and formatting only if there is actual text
                    if has_text:
                        pub_text_range = pub_shape.TextFrame.TextRange
                        ppt_text_range = ppt_text_frame.TextRange
                        ppt_text_range.Text = text_content
                        

                        # Per-paragraph formatting: font, alignment, bullets/numbers
                        try:
                            # Count paragraphs from the text we set (avoids COM .Count failures)
                            para_count = text_content.count('\r') + 1
                            
                            align_map = {0: 1, 1: 2, 2: 3, 3: 4}
                            
                            for para_idx in range(1, para_count + 1):
                                try:
                                    pub_para = pub_text_range.Paragraphs(para_idx)
                                    ppt_para = ppt_text_range.Paragraphs(para_idx)
                                    
                                    # --- Per-run font copy ---
                                    # Each paragraph can have multiple "runs" with
                                    # different formatting (bold word in normal sentence, etc.)
                                    try:
                                        pub_run_count = pub_para.Runs().Count
                                        ppt_run_count = ppt_para.Runs().Count
                                        run_count = min(pub_run_count, ppt_run_count)
                                        
                                        if run_count > 1:
                                            # Multiple runs: copy each individually
                                            for run_idx in range(1, run_count + 1):
                                                try:
                                                    pub_run = pub_para.Runs(run_idx)
                                                    ppt_run = ppt_para.Runs(run_idx)
                                                    ppt_run.Font.Name = pub_run.Font.Name
                                                    ppt_run.Font.Size = pub_run.Font.Size
                                                    ppt_run.Font.Bold = pub_run.Font.Bold
                                                    ppt_run.Font.Italic = pub_run.Font.Italic
                                                    try:
                                                        ppt_run.Font.Color.RGB = pub_run.Font.Color.RGB
                                                    except:
                                                        pass
                                                    try:
                                                        pub_ul = pub_run.Font.Underline
                                                        ppt_run.Font.Underline = True if (pub_ul and pub_ul != 0) else False
                                                    except:
                                                        pass
                                                except:
                                                    continue
                                        else:
                                            # Single run or mismatched count: copy from first Publisher run
                                            try:
                                                src = pub_para.Runs(1)
                                            except:
                                                try:
                                                    src = pub_para.Characters(1, 1)
                                                except:
                                                    src = pub_para
                                            ppt_para.Font.Name = src.Font.Name
                                            ppt_para.Font.Size = src.Font.Size
                                            ppt_para.Font.Bold = src.Font.Bold
                                            ppt_para.Font.Italic = src.Font.Italic
                                            try:
                                                ppt_para.Font.Color.RGB = src.Font.Color.RGB
                                            except:
                                                pass
                                            try:
                                                pub_ul = src.Font.Underline
                                                ppt_para.Font.Underline = True if (pub_ul and pub_ul != 0) else False
                                            except:
                                                pass
                                    except:
                                        #  Fallback: Runs().Count failed, try Runs(1) or Characters(1,1)
                                        try:
                                            src = pub_para.Runs(1)
                                        except:
                                            try:
                                                src = pub_para.Characters(1, 1)
                                            except:
                                                src = pub_para
                                        try:
                                            ppt_para.Font.Name = src.Font.Name
                                            ppt_para.Font.Size = src.Font.Size
                                            ppt_para.Font.Bold = src.Font.Bold
                                            ppt_para.Font.Italic = src.Font.Italic
                                        except:
                                            pass
                                    
                                    # --- Per-paragraph alignment (mapped) ---
                                    try:
                                        pub_align = pub_para.ParagraphFormat.Alignment
                                        ppt_para.ParagraphFormat.Alignment = align_map.get(pub_align, pub_align + 1)
                                    except:
                                        pass                                    
                                    # --- Per-paragraph bullet/number formatting ---
                                    try:
                                        pub_pf = pub_para.ParagraphFormat
                                        ppt_pf = ppt_para.ParagraphFormat
                                        list_type = pub_pf.ListType

                                        if list_type == 255:
                                            # pbListTypeNone -- no list
                                            try:
                                                ppt_pf.Bullet.Type = 0
                                            except:
                                                pass

                                        elif list_type == 23:
                                            # Bullet styles (23 = standard bullet, etc.)
                                            ppt_pf.Bullet.Type = 1  # ppBulletUnnumbered

                                            # Copy bullet character from Publisher
                                            try:
                                                bt = pub_pf.ListBulletText
                                                if bt:
                                                    ppt_pf.Bullet.Character = ord(bt[0])
                                                else:
                                                    ppt_pf.Bullet.Character = 8226  # default bullet
                                            except:
                                                ppt_pf.Bullet.Character = 8226
                                            
                                            ###$ Force standard bullet if Publisher returned middle dot
                                            #try:
                                                #if ppt_pf.Bullet.Character == 183:
                                                   # ppt_pf.Bullet.Character = 8226
                                            #except:
                                                #pass
                                            # Copy bullet font -- must be set AFTER Type and Character
                                            try:
                                                bfn = pub_pf.ListBulletFontName
                                                if bfn:
                                                    ppt_pf.Bullet.Font.Name = str(bfn)
                                            except:
                                                pass

                                            # Copy bullet font size
                                            try:
                                                bfs = pub_pf.ListBulletFontSize
                                                if bfs and bfs > 0:
                                                    ppt_pf.Bullet.Font.Size = bfs
                                                else:
                                                    ppt_pf.Bullet.Font.Size = pub_para.Font.Size
                                            except:
                                                try:
                                                    ppt_pf.Bullet.Font.Size = pub_para.Font.Size
                                                except:
                                                    pass

                                        else:
                                            # All remaining list_type values are numbered variants
                                            # PbListType enumeration (MS docs):
                                            #   0  = pbListTypeArabic          (1, 2, 3)
                                            #   1  = pbListTypeUpperCaseRoman  (I, II, III)
                                            #   2  = pbListTypeLowerCaseRoman  (i, ii, iii)
                                            #   3  = pbListTypeUpperCaseLetter (A, B, C)
                                            #   4  = pbListTypeLowerCaseLetter (a, b, c)
                                            #   5  = pbListTypeOrdinal
                                            #   6  = pbListTypeCardinalText
                                            #   7  = pbListTypeOrdinalText
                                            #  22  = pbListTypeArabicLeadingZero (01, 02, 03)
                                            # PpNumberedBulletStyle:
                                            #   1 = ppBulletStyleArabicPeriod
                                            #   3 = ppBulletStyleAlphaUCPeriod
                                            #   5 = ppBulletStyleAlphaLCPeriod
                                            #   7 = ppBulletStyleRomanUCPeriod
                                            #   9 = ppBulletStyleRomanLCPeriod
                                            ppt_pf.Bullet.Type = 2  # ppBulletNumbered
                                            num_style_map = {
                                                0: 3,   # Arabic (1,2,3)       -> ppBulletStyleArabicPeriod
                                                1: 7,   # UpperCaseRoman (I)   -> ppBulletStyleRomanUCPeriod
                                                2: 9,   # LowerCaseRoman (i)   -> ppBulletStyleRomanLCPeriod
                                                3: 3,   # UpperCaseLetter (A)  -> ppBulletStyleAlphaUCPeriod
                                                4: 5,   # LowerCaseLetter (a)  -> ppBulletStyleAlphaLCPeriod
                                                22: 1,  # ArabicLeadingZero    -> ppBulletStyleArabicPeriod
                                            }
                                            ppt_pf.Bullet.Style = num_style_map.get(list_type, 1)

                                            # Copy start value
                                            try:
                                                sv = pub_pf.ListNumberStart
                                                if sv and sv > 0:
                                                    ppt_pf.Bullet.StartValue = sv
                                            except:
                                                pass
                                        
                                            # Copy paragraph indents (hanging indent for bullets)
                                            try:
                                                ppt_pf.LeftMargin = pub_pf.LeftIndent
                                            except:
                                                pass
                                            try:
                                                ppt_pf.FirstLineIndent = pub_pf.FirstLineIndent
                                            except:
                                                pass
                                    except:
                                        pass
                                
                                except:
                                    continue
                        except Exception as e:
                            if self.verbose:
                                print(f"      Per-paragraph formatting error: {e}")
                            # Fallback: copy whole-range font as before
                            try:
                                ppt_text_range.Font.Name = pub_text_range.Font.Name
                                ppt_text_range.Font.Size = pub_text_range.Font.Size
                                ppt_text_range.Font.Bold = pub_text_range.Font.Bold
                                ppt_text_range.Font.Italic = pub_text_range.Font.Italic
                            except:
                                pass
                    
                    # --- Everything below applies to ALL shapes (empty or not) ---
                    
                    # Copy text frame margins
                    try:
                        pub_text_frame = pub_shape.TextFrame
                        ppt_text_frame.MarginLeft = pub_text_frame.MarginLeft
                        ppt_text_frame.MarginRight = pub_text_frame.MarginRight
                        ppt_text_frame.MarginTop = pub_text_frame.MarginTop
                        ppt_text_frame.MarginBottom = pub_text_frame.MarginBottom
                    except:
                        # Fallback to Publisher's typical margin
                        try:
                            ppt_text_frame.MarginLeft = 2.88
                            ppt_text_frame.MarginRight = 2.88
                            ppt_text_frame.MarginTop = 2.88
                            ppt_text_frame.MarginBottom = 2.88
                        except:
                            pass

                    # Copy vertical alignment
                    try:
                        pub_va = pub_shape.TextFrame.VerticalTextAlignment
                        va_map = {0: 1, 1: 3, 2: 4}
                        if pub_va in va_map:
                            ppt_text_frame.VerticalAnchor = va_map[pub_va]
                    except:
                        pass
                    
                    # Copy shape fill/background color
                    try:
                        if pub_shape.Fill.Visible:
                            ppt_shape.Fill.Visible = True
                            if pub_shape.Fill.Type == 1:  # Solid fill
                                ppt_shape.Fill.ForeColor.RGB = pub_shape.Fill.ForeColor.RGB
                    except:
                        pass

                    # Restore WordWrap from Publisher AFTER text is written
                    try:
                        ppt_text_frame.WordWrap = pub_word_wrap
                        if self.verbose:
                            wrap_status = "ON" if pub_word_wrap else "OFF"
                            print(f"      WordWrap: {wrap_status}")
                    except:
                        pass
                    
                    # FINAL STEP: Force exact dimensions from Publisher
                    # This must be done AFTER all content/formatting to override any auto-adjustments
                    try:
                        ppt_shape.Left = pub_shape.Left
                        ppt_shape.Top = pub_shape.Top
                        ppt_shape.Width = pub_shape.Width
                        ppt_shape.Height = pub_shape.Height
                    except:
                        pass
                    
                    try:
                        if pub_shape.Line.Visible:
                            ppt_shape.Line.Visible = True
                            ppt_shape.Line.Weight = pub_shape.Line.Weight
                            ppt_shape.Line.ForeColor.RGB = pub_shape.Line.ForeColor.RGB
                    except:
                        pass
                    
                    # Apply rotation
                    try:
                        ppt_shape.Rotation = pub_shape.Rotation
                    except:
                        pass
                    
                    self.current_file_stats["text_boxes"] = self.current_file_stats.get("text_boxes", 0) + 1
                    return True
                except Exception as e:
                    print(f"    Warning: Could not fully convert text shape: {e}")
                    if self.logger:
                        self.logger.log_message(f"Could not fully convert text shape: {e}", "WARNING")
            # Handle picture shapes
            elif shape_type in [2, 13] and self.preserve_images:
                try:
                    # Save original rotation and dimensions
                    original_rotation = 0
                    try:
                        original_rotation = pub_shape.Rotation
                    except:
                        pass
                    
                    # Temporarily remove rotation so SaveAsPicture exports
                    # the raw unrotated image (rotation is baked into the export otherwise)
                    if original_rotation != 0:
                        try:
                            pub_shape.Rotation = 0
                        except:
                            pass
                    
                    # Create unique filename
                    timestamp = int(time.time() * 1000000)
                    temp_img = self.temp_dir / f"shape_{timestamp}_{id(pub_shape)}.png"
                    temp_img_str = str(temp_img.resolve())
                    
                    # Export picture
                    pub_shape.SaveAsPicture(temp_img_str)
                    
                    # Restore Publisher shape rotation immediately
                    if original_rotation != 0:
                        try:
                            pub_shape.Rotation = original_rotation
                        except:
                            pass
                    
                    time.sleep(0.1)
                    
                    if not temp_img.exists():
                        print(f"    Warning: Image file was not created: {temp_img.name}")
                        if self.logger:
                            self.logger.log_message(f"Image file was not created: {temp_img.name}", "WARNING")
                        return False
                    
                    # Use exact Publisher position/size (no scaling -- slide matches page)
                    img_left = pub_shape.Left
                    img_top = pub_shape.Top
                    img_width = pub_shape.Width
                    img_height = pub_shape.Height
                    
                    ppt_shape = ppt_slide.Shapes.AddPicture(
                        FileName=temp_img_str,
                        LinkToFile=False,
                        SaveWithDocument=True,
                        Left=img_left,
                        Top=img_top,
                        Width=img_width,
                        Height=img_height
                    )
                    
                    # Unlock aspect ratio so dimensions are exact
                    try:
                        ppt_shape.LockAspectRatio = False
                    except:
                        pass
                    
                    # Force exact dimensions after insertion
                    try:
                        ppt_shape.Left = img_left
                        ppt_shape.Top = img_top
                        ppt_shape.Width = img_width
                        ppt_shape.Height = img_height
                    except:
                        pass
                    
                    # Apply rotation in PowerPoint (from the original Publisher value)
                    try:
                        ppt_shape.Rotation = original_rotation
                    except:
                        pass

                    # Copy flip state
                    try:
                        if pub_shape.HorizontalFlip:
                            ppt_shape.Flip(0)  # msoFlipHorizontal
                    except:
                        pass
                    try:
                        if pub_shape.VerticalFlip:
                            ppt_shape.Flip(1)  # msoFlipVertical
                    except:
                        pass
                            
                    # Clean up temp file
                    try:
                        time.sleep(0.1)
                        temp_img.unlink()
                    except:
                        pass
                    
                    self.current_file_stats["images"] = self.current_file_stats.get("images", 0) + 1
                    return True
                except Exception as e:
                    # Restore Publisher rotation on failure too
                    try:
                        if original_rotation != 0:
                            pub_shape.Rotation = original_rotation
                    except:
                        pass
                    print(f"    Warning: Could not convert picture shape: {e}")
                    if self.logger:
                        self.logger.log_message(f"Could not convert picture shape: {e}", "WARNING")
            
            # Handle AutoShapes (type 1): ovals, rounded rects, arrows, callouts, etc.
            elif shape_type == 1 and self.preserve_shapes:
                try:
                    # Read the specific AutoShape sub-type from Publisher
                    auto_shape_type = 1  # default to msoShapeRectangle
                    try:
                        auto_shape_type = pub_shape.AutoShapeType
                    except:
                        pass
                    
                    if auto_shape_type < 0:
                        # Negative AutoShapeType = connector/line that AddShape can't handle
                        # Fall back to AddLine using the shape's bounding box
                        # Read flip state BEFORE creating the line so we can
                        # bake the direction into the begin/end coordinates
                        h_flip = False
                        v_flip = False
                        try:
                            h_flip = bool(pub_shape.HorizontalFlip)
                        except:
                            pass
                        try:
                            v_flip = bool(pub_shape.VerticalFlip)
                        except:
                            pass
                        # Compute correct begin/end points based on flip state
                        # Default (no flip): top-left -> bottom-right
                        begin_x = left
                        begin_y = top
                        end_x = left + width
                        end_y = top + height
                        if h_flip:
                            begin_x, end_x = end_x, begin_x
                        if v_flip:
                            begin_y, end_y = end_y, begin_y
                        ppt_shape = ppt_slide.Shapes.AddLine(
                            BeginX=begin_x,
                            BeginY=begin_y,
                            EndX=end_x,
                            EndY=end_y
                        )
                        # Copy arrow properties
                        try:
                            ppt_shape.Line.EndArrowheadStyle = pub_shape.Line.EndArrowheadStyle
                        except:
                            pass
                        try:
                            ppt_shape.Line.BeginArrowheadStyle = pub_shape.Line.BeginArrowheadStyle
                        except:
                            pass
                        try:
                            ppt_shape.Line.EndArrowheadLength = pub_shape.Line.EndArrowheadLength
                        except:
                            pass
                        try:
                            ppt_shape.Line.EndArrowheadWidth = pub_shape.Line.EndArrowheadWidth
                        except:
                            pass
                        try:
                            ppt_shape.Line.BeginArrowheadLength = pub_shape.Line.BeginArrowheadLength
                        except:
                            pass
                        try:
                            ppt_shape.Line.BeginArrowheadWidth = pub_shape.Line.BeginArrowheadWidth
                        except:
                            pass
                        try:
                            ppt_shape.Line.DashStyle = pub_shape.Line.DashStyle
                        except:
                            pass

                    else:
                        ppt_shape = ppt_slide.Shapes.AddShape(
                            Type=auto_shape_type,
                            Left=left,
                            Top=top,
                            Width=width,
                            Height=height
                        )
                    
                    # Copy fill
                    try:
                        if pub_shape.Fill.Visible:
                            ppt_shape.Fill.Visible = True
                            if pub_shape.Fill.Type == 1:
                                ppt_shape.Fill.ForeColor.RGB = pub_shape.Fill.ForeColor.RGB
                        else:
                            ppt_shape.Fill.Visible = False
                    except:
                        pass
                    
                    # Copy line/border
                    try:
                        if pub_shape.Line.Visible:
                            ppt_shape.Line.Visible = True
                            ppt_shape.Line.Weight = pub_shape.Line.Weight
                            ppt_shape.Line.ForeColor.RGB = pub_shape.Line.ForeColor.RGB
                            try:
                                ppt_shape.Line.DashStyle = pub_shape.Line.DashStyle
                            except:
                                pass
                        else:
                            ppt_shape.Line.Visible = False
                    except:
                        pass
                    
                    # Copy text if the AutoShape has a text frame
                    if pub_shape.HasTextFrame:
                        try:
                            pub_tf = pub_shape.TextFrame
                            ppt_tf = ppt_shape.TextFrame
                            try:
                                ppt_tf.AutoSize = 0
                            except:
                                pass

                            # Copy WordWrap
                            try:
                                ppt_tf.WordWrap = pub_tf.WordWrap
                            except:
                                try:
                                    ppt_tf.WordWrap = -1
                                except:
                                    pass

                            original_text = pub_tf.TextRange.Text
                            raw_text = original_text.rstrip('\r\n\x0b\x0c')

                            # Write text content -- if stripping remove everything,
                            # fall back to original to \r-only shapes keep paragraph structure.
                            text_to_write = raw_text if raw_text else original_text
                            if text_to_write:
                                ppt_tf.TextRange.Text = text_to_write

                            # Re-obtain source TextRange fresh for formatting
                            pub_text_range = pub_shape.TextFrame.TextRange
                            ppt_text_range = ppt_tf.TextRange

                            # Per-paragraph formatting -- ALWAYS runs
                            align_map = {0: 1, 1: 2, 2: 3, 3: 4}
                            try:
                                pub_pc = pub_text_range.Paragraphs().Count
                                ppt_pc = ppt_text_range.Paragraphs().Count
                                for pi in range(1, min(pub_pc, ppt_pc) + 1):
                                    try:
                                        pp = pub_text_range.Paragraphs(pi)
                                        tp = ppt_text_range.Paragraphs(pi)

                                        # Try per-run font copy first (matches TextBox handler)
                                        try:
                                            pub_run_count = pp.Runs().Count
                                            ppt_run_count = tp.Runs().Count
                                            if pub_run_count > 0 and pub_run_count == ppt_run_count:
                                                for ri in range(1, pub_run_count + 1):
                                                    try:
                                                        pub_run = pp.Runs(ri)
                                                        ppt_run = tp.Runs(ri)
                                                        ppt_run.Font.Name = pub_run.Font.Name
                                                        ppt_run.Font.Size = pub_run.Font.Size
                                                        ppt_run.Font.Bold = pub_run.Font.Bold
                                                        ppt_run.Font.Italic = pub_run.Font.Italic
                                                        try:
                                                            ppt_run.Font.Color.RGB = pub_run.Font.Color.RGB
                                                        except:
                                                            ppt_run.Font.Color.RGB = 0
                                                        try:
                                                            pub_ul = pub_run.Font.Underline
                                                            ppt_run.Font.Underline = True if (pub_ul and pub_ul != 0) else False
                                                        except:
                                                            pass
                                                    except:
                                                        continue
                                            else:
                                                # Fallback: copy from first run or paragraph level
                                                try:
                                                    src = pp.Runs(1)
                                                except:
                                                    try:
                                                        src = pp.Characters(1, 1)
                                                    except:
                                                        src = pp
                                                tp.Font.Name = src.Font.Name
                                                tp.Font.Size = src.Font.Size
                                                tp.Font.Bold = src.Font.Bold
                                                tp.Font.Italic = src.Font.Italic
                                                try:
                                                    tp.Font.Color.RGB = src.Font.Color.RGB
                                                except:
                                                    pass
                                                try:
                                                    pub_ul = src.Font.Underline
                                                    tp.Font.Underline = True if (pub_ul and pub_ul != 0) else False
                                                except:
                                                    pass
                                        except:
                                            # Runs().Count failed -- try paragraph-level font
                                            try:
                                                tp.Font.Name = pp.Font.Name
                                                tp.Font.Size = pp.Font.Size
                                                tp.Font.Bold = pp.Font.Bold
                                                tp.Font.Italic = pp.Font.Italic
                                            except:
                                                pass

                                        # Per-paragraph alignment
                                        try:
                                            pa = pp.ParagraphFormat.Alignment
                                            tp.ParagraphFormat.Alignment = align_map.get(pa, pa + 1)
                                        except:
                                            pass
                                    except:
                                        continue
                            except:
                                # Paragraph loop failed entirely -- whole-range fallback
                                try:
                                    ppt_text_range.Font.Name = pub_text_range.Font.Name
                                    ppt_text_range.Font.Size = pub_text_range.Font.Size
                                    ppt_text_range.Font.Bold = pub_text_range.Font.Bold
                                    ppt_text_range.Font.Italic = pub_text_range.Font.Italic
                                    try:
                                        ppt_text_range.Font.Color.RGB = pub_text_range.Font.Color.RGB
                                    except:
                                        ppt_text_range.Font.Color.RGB = 0
                                except:
                                    pass

                            # Copy margins -- ALWAYS runs
                            try:
                                ppt_tf.MarginLeft = pub_tf.MarginLeft
                                ppt_tf.MarginRight = pub_tf.MarginRight
                                ppt_tf.MarginTop = pub_tf.MarginTop
                                ppt_tf.MarginBottom = pub_tf.MarginBottom
                            except:
                                pass

                            # Copy vertical alignment -- ALWAYS runs
                            try:
                                pub_va = pub_tf.VerticalTextAlignment
                                va_map = {0: 1, 1: 3, 2: 4}
                                if pub_va in va_map:
                                    ppt_tf.VerticalAnchor = va_map[pub_va]
                            except:
                                pass
                        except:
                            pass
                    
                    # Rotation
                    try:
                        ppt_shape.Rotation = pub_shape.Rotation
                    except:
                        pass
                    # Copy flip state (skip for negative AutoShapeType -- flip is
                    # already baked into the AddLine begin/end coordinates above)
                    if auto_shape_type >= 0:
                        try:
                            if pub_shape.HorizontalFlip:
                                ppt_shape.Flip(0)  # msoFlipHorizontal
                        except:
                            pass
                        try:
                            if pub_shape.VerticalFlip:
                                ppt_shape.Flip(1)  # msoFlipVertical
                        except:
                            pass
                    self.current_file_stats["other_shapes"] = self.current_file_stats.get("other_shapes", 0) + 1

                    return True
                except Exception as e:
                    print(f"    Warning: Could not convert AutoShape: {e}")
                    if self.logger:
                        self.logger.log_message(f"Could not convert AutoShape: {e}", "WARNING")
            
            # Handle other shapes (rectangles, lines, etc.)
            elif self.preserve_shapes:
                try:
                    if shape_type == 5:  # Rectangle
                        ppt_shape = ppt_slide.Shapes.AddShape(
                            Type=1,  # msoShapeRectangle
                            Left=left,
                            Top=top,
                            Width=width,
                            Height=height
                        )
                    elif shape_type == 9:  # Line
                        # Read flip state BEFORE creating the line so we can
                        # bake the direction into the begin/end coordinates
                        h_flip = False
                        v_flip = False
                        try:
                            h_flip = bool(pub_shape.HorizontalFlip)
                        except:
                            pass
                        try:
                            v_flip = bool(pub_shape.VerticalFlip)
                        except:
                            pass
                        # Compute correct begin/end points based on flip state
                        # Default (no flip): top-left -> bottom-right
                        begin_x = left
                        begin_y = top
                        end_x = left + width
                        end_y = top + height
                        if h_flip:
                            begin_x, end_x = end_x, begin_x
                        if v_flip:
                            begin_y, end_y = end_y, begin_y
                        ppt_shape = ppt_slide.Shapes.AddLine(
                            BeginX=begin_x,
                            BeginY=begin_y,
                            EndX=end_x,
                            EndY=end_y
                        )
                        # Copy arrow properties
                        try:
                            ppt_shape.Line.BeginArrowheadStyle = pub_shape.Line.BeginArrowheadStyle
                        except:
                            pass
                        try:
                            ppt_shape.Line.EndArrowheadStyle = pub_shape.Line.EndArrowheadStyle
                        except:
                            pass
                        try:
                            ppt_shape.Line.BeginArrowheadLength = pub_shape.Line.BeginArrowheadLength
                        except:
                            pass
                        try:
                            ppt_shape.Line.BeginArrowheadWidth = pub_shape.Line.BeginArrowheadWidth
                        except:
                            pass
                        try:
                            ppt_shape.Line.EndArrowheadLength = pub_shape.Line.EndArrowheadLength
                        except:
                            pass
                        try:
                            ppt_shape.Line.EndArrowheadWidth = pub_shape.Line.EndArrowheadWidth
                        except:
                            pass
                        try:
                            ppt_shape.Line.DashStyle = pub_shape.Line.DashStyle
                        except:
                            pass

                    else:
                        # Generic shape -- try to read AutoShapeType, fall back to rectangle
                        fallback_type = 1  # msoShapeRectangle
                        try:
                            fallback_type = pub_shape.AutoShapeType
                        except:
                            pass
                        ppt_shape = ppt_slide.Shapes.AddShape(
                            Type=fallback_type,
                            Left=left,
                            Top=top,
                            Width=width,
                            Height=height
                        )
                    
                    # Copy fill properties
                    try:
                        if pub_shape.Fill.Visible:
                            ppt_shape.Fill.Visible = True
                            if pub_shape.Fill.Type == 1:  # Solid fill
                                ppt_shape.Fill.ForeColor.RGB = pub_shape.Fill.ForeColor.RGB
                    except:
                        pass
                    
                    # Copy line properties
                    try:
                        if pub_shape.Line.Visible:
                            ppt_shape.Line.Visible = True
                            ppt_shape.Line.Weight = pub_shape.Line.Weight
                            ppt_shape.Line.ForeColor.RGB = pub_shape.Line.ForeColor.RGB
                    except:
                        pass

                    try:
                        ppt_shape.Rotation = pub_shape.Rotation
                    except:
                        pass
                    # Copy flip state (skip for type 9 lines -- flip is
                    # already baked into the AddLine begin/end coordinates above)
                    if shape_type != 9:
                        try:
                            if pub_shape.HorizontalFlip:
                                ppt_shape.Flip(0)  # msoFlipHorizontal
                        except:
                            pass
                        try:
                            if pub_shape.VerticalFlip:
                                ppt_shape.Flip(1)  # msoFlipVertical
                        except:
                            pass
                    self.current_file_stats["other_shapes"] = self.current_file_stats.get("other_shapes", 0) + 1

                    return True
                except Exception as e:
                    print(f"    Warning: Could not convert shape: {e}")
                    if self.logger:
                        self.logger.log_message(f"Could not convert shape: {e}", "WARNING")
            
            return False
            
        except Exception as e:
            print(f"    Error processing shape: {e}")
            if self.logger:
                self.logger.log_message(f"Error processing shape: {e}", "ERROR")
            return False
    
    def convert_file(self, input_path, output_path=None, fallback_to_image=True):
        """
        Convert a single Publisher file to PowerPoint.
        
        Args:
            input_path: Path to Publisher file
            output_path: Path for output PPTX file (optional)
            fallback_to_image: If True, fall back to image export on failure
            
        Returns:
            True if successful, False otherwise
        """
        input_path = Path(input_path).resolve()
        
        if not input_path.exists():
            error_msg = f"Input file not found: {input_path}"
            if self.logger:
                self.logger.log_message(error_msg, "ERROR")
            else:
                print(f"Error: {error_msg}")
            return False
        
        if input_path.suffix.lower() != '.pub':
            warning_msg = f"Warning: File may not be a Publisher file: {input_path}"
            if self.logger:
                self.logger.log_message(warning_msg, "WARNING")
            else:
                print(warning_msg)
        
        # Determine output path
        if output_path is None:
            output_path = input_path.with_suffix('.pptx')
        else:
            output_path = Path(output_path).resolve()
        
        # Print conversion header
        print(f"Converting: {input_path.name}")
        print(f"Output: {output_path}")
        print(f"Mode: {'DRY RUN - ' if self.dry_run else ''}Preserving editable content")
        
        if self.dry_run:
            msg = f"[DRY RUN] Would convert {input_path.name} to {output_path.name}"
            if self.logger:
                self.logger.log_message(msg)
            print(msg)
            return True
        
        # Reset stats for this file
        self.current_file_stats = {
            "pages": 0,
            "shapes_processed": 0,
            "shapes_converted": 0,
            "shapes_failed": 0,
            "text_boxes": 0,
            "images": 0,
            "tables": 0,
            "other_shapes": 0
        }
        
        # Log file start
        if self.logger:
            self.logger.log_file_start(input_path)
        
        try:
            # Check if Publisher needs restart from previous error
            if self._publisher_needs_restart:
                self._restart_publisher()
            
            # Dismiss any lingering dialogs before opening
            dismiss_all_publisher_dialogs(verbose=self.verbose)
            
            # Open Publisher document with retry logic for modal dialogs
            pub_doc = None
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    pub_doc = self.publisher.Open(str(input_path))
                    
                    # Brief pause to let any dialogs appear
                    time.sleep(0.3)
                    
                    # Check for and dismiss any dialogs that appeared during open
                    dismissed = dismiss_all_publisher_dialogs(verbose=self.verbose)
                    if dismissed > 0:
                        print(f"    Dismissed {dismissed} dialog(s) after opening file")
                    
                    break  # Success, exit retry loop
                    
                except Exception as open_error:
                    if is_modal_dialog_error(open_error):
                        print(f"    Modal dialog detected (attempt {attempt + 1}/{max_retries}), attempting to dismiss...")
                        
                        # Try to dismiss dialogs
                        dismissed = dismiss_all_publisher_dialogs(verbose=self.verbose)
                        if dismissed > 0:
                            print(f"    Dismissed {dismissed} dialog(s)")
                            time.sleep(0.5)
                            
                            # If this is the last attempt, restart Publisher
                            if attempt == max_retries - 1:
                                self._restart_publisher()
                        else:
                            # No dialogs found but error occurred - restart Publisher
                            self._restart_publisher()
                    else:
                        # Not a modal dialog error, re-raise
                        raise
            
            if pub_doc is None:
                raise Exception(f"Failed to open {input_path.name} after {max_retries} attempts")
            
            # Create PowerPoint presentation from template or blank
            if self.template_path and Path(self.template_path).exists():
                print(f"Using template: {Path(self.template_path).name}")
                if self.logger:
                    self.logger.log_message(f"Using template: {Path(self.template_path).name}")
                
                # Open template and create new presentation from it
                ppt_pres = self.powerpoint.Presentations.Open(str(Path(self.template_path).resolve()))
                
                # Remove any existing slides from the template
                slide_count = ppt_pres.Slides.Count
                if slide_count > 0:
                    if self.logger:
                        self.logger.log_message(f"Deleting {slide_count} existing template slides...")
                    while ppt_pres.Slides.Count > 0:
                        ppt_pres.Slides(1).Delete()
            else:
                if self.template_path:
                    warning_msg = f"Warning: Template not found at {self.template_path}, using blank presentation"
                    print(warning_msg)
                    if self.logger:
                        self.logger.log_message(warning_msg, "WARNING")
                # Create blank presentation
                ppt_pres = self.powerpoint.Presentations.Add()
            
            # Get Publisher page dimensions (in points)
            num_pages = pub_doc.Pages.Count
            
            # Get page dimensions from the document's page setup
            page_width = pub_doc.PageSetup.PageWidth
            page_height = pub_doc.PageSetup.PageHeight
            
            # Get slide dimensions
            slide_width = ppt_pres.PageSetup.SlideWidth
            slide_height = ppt_pres.PageSetup.SlideHeight
            
            # Set PowerPoint slide size to match Publisher page size
            # This prevents scaling issues and ensures text boxes render correctly
            ppt_pres.PageSetup.SlideWidth = page_width
            ppt_pres.PageSetup.SlideHeight = page_height
            
            # Update slide dimensions after setting them
            slide_width = page_width
            slide_height = page_height

            print(f"Found {num_pages} page(s)")
            print(f"Publisher page size: {page_width:.1f} x {page_height:.1f} points")
            print(f"PowerPoint slide size: {slide_width:.1f} x {slide_height:.1f} points")
            
            if self.logger:
                self.logger.log_message(f"Processing {num_pages} pages...")
                self.logger.log_message(f"Publisher page size: {page_width:.1f} x {page_height:.1f} points")
                self.logger.log_message(f"PowerPoint slide size: {slide_width:.1f} x {slide_height:.1f} points")
            
            self.current_file_stats["pages"] = num_pages
            
            # Process each page
            for page_num in range(1, num_pages + 1):
                print(f"\n  Processing page {page_num} of {num_pages}...")
                if self.logger:
                    self.logger.log_message(f"Processing page {page_num} of {num_pages}")
                
                page = pub_doc.Pages(page_num)
                
                # Add slide to PowerPoint (always use Add method)
                slide = ppt_pres.Slides.Add(page_num, 12)  # 12 = ppLayoutBlank
                
                # Get all shapes on the page
                shapes_converted = 0
                shapes_failed = 0
                
                try:
                    shape_count = page.Shapes.Count
                    print(f"    Found {shape_count} shape(s) on page")
                    if self.logger:
                        self.logger.log_message(f"Page {page_num}: Found {shape_count} shapes")
                    
                    # Process shapes in z-order (index 1 = back, higher = front)
                    # In PowerPoint, shapes added first go to back, added later go to front
                    # So we iterate from 1 to shape_count to preserve layering
                    for i in range(1, shape_count + 1):
                        try:
                            pub_shape = page.Shapes(i)
                            if self.convert_shape_to_powerpoint(
                                pub_shape, slide, page_width, page_height,
                                slide_width, slide_height
                            ):
                                shapes_converted += 1
                            else:
                                shapes_failed += 1
                        except Exception as e:
                            shapes_failed += 1
                            warning_msg = f"Failed to convert shape {i}: {e}"
                            print(f"    Warning: {warning_msg}")
                            if self.logger:
                                self.logger.log_message(warning_msg, "WARNING")
                    
                    print(f"    Converted: {shapes_converted} shapes")
                    if shapes_failed > 0:
                        print(f"    Failed: {shapes_failed} shapes")
                    
                    if self.logger:
                        self.logger.log_message(f"Page {page_num}: Converted {shapes_converted}, Failed {shapes_failed}")
                    
                    # Update cumulative stats
                    self.current_file_stats["shapes_converted"] += shapes_converted
                    self.current_file_stats["shapes_failed"] += shapes_failed
                    
                    # If most shapes failed and fallback is enabled, use image
                    if fallback_to_image and shapes_converted == 0 and shape_count > 0:
                        print(f"    Using image fallback for page {page_num}")
                        if self.logger:
                            self.logger.log_message(f"Page {page_num}: Using image fallback")
                        
                        timestamp = int(time.time() * 1000000)
                        temp_img = self.temp_dir / f"page_{timestamp}_{page_num}.png"
                        page.SaveAsPicture(str(temp_img.resolve()), constants.pbPictureTypePNG)
                        
                        time.sleep(0.1)
                        
                        if temp_img.exists():
                            slide.Shapes.AddPicture(
                                FileName=str(temp_img.resolve()),
                                LinkToFile=False,
                                SaveWithDocument=True,
                                Left=0,
                                Top=0,
                                Width=slide_width,
                                Height=slide_height
                            )
                            
                            try:
                                time.sleep(0.1)
                                temp_img.unlink()
                            except:
                                pass
                
                except Exception as e:
                    error_msg = f"Error processing page {page_num}: {e}"
                    print(f"    {error_msg}")
                    if self.logger:
                        self.logger.log_message(error_msg, "ERROR")
                    
                    if fallback_to_image:
                        print(f"    Using image fallback")
                        if self.logger:
                            self.logger.log_message(f"Page {page_num}: Using image fallback due to error")
                        
                        timestamp = int(time.time() * 1000000)
                        temp_img = self.temp_dir / f"page_{timestamp}_{page_num}.png"
                        try:
                            page.SaveAsPicture(str(temp_img.resolve()), constants.pbPictureTypePNG)
                            time.sleep(0.1)
                            
                            if temp_img.exists():
                                slide.Shapes.AddPicture(
                                    FileName=str(temp_img.resolve()),
                                    LinkToFile=False,
                                    SaveWithDocument=True,
                                    Left=0,
                                    Top=0,
                                    Width=slide_width,
                                    Height=slide_height
                                )
                                time.sleep(0.1)
                                temp_img.unlink()
                        except Exception as fallback_error:
                            if self.logger:
                                self.logger.log_message(f"Image fallback also failed: {fallback_error}", "ERROR")
            
            # Save PowerPoint file
            #ppt_pres.SaveAs(str(output_path), 24)  # ppSaveAsOpenXMLPresentation
            #ppt_pres.Close()

            # Save PowerPoint file
            try:
                ppt_pres.SaveAs(str(output_path), 24)  # ppSaveAsOpenXMLPresentation
            except Exception as save_err:
                print(f"    Warning: Issue Saving PowerPoint file: {save_err}")

            # Add a small delay to ensure PowerPoint finishes saving to disk
            time.sleep(1.0)
            
            # Try closing the presentation safely
            try:
                    if ppt_pres is not None:
                        ppt_pres.Close()
            except Exception as close_error:
                warning_msg = f"Could not cleanly close PowerPoint presentation: {close_error}"
                print(f"    Warning: {warning_msg}")
                if self.logger:
                    self.logger.log_message(warning_msg, "WARNING")
            
            # Clean up Publisher safely
            try:
                if pub_doc is not None:
                    pub_doc.Close()
            except Exception as pub_close_error:
                warning_msg = f"Publisher object lost (it may have closed early: {pub_close_error})"
                print(f"    Warning:{warning_msg}")
                if self.logger:
                    self.logger.log_message(warning_msg, "WARNING")
                # Flag Publisher for a fresh restart on the next file in the batch!
                self._publisher_needs_restart = True

            print(f"\n✓ Successfully converted to: {output_path}\n") 

            # Log Success
            if self.logger:
                self.logger.log_file_complete(
                    input_path,
                    "SUCCESS",
                    output_path,
                    stats=self.current_file_stats
                )
            
            return True
                # Silently pass or log if Publisher also resists closing
            
            # Clean up
            pub_doc.Close()
            
            print(f"\n✓ Successfully converted to: {output_path}\n")
            
            # Log success
            if self.logger:
                self.logger.log_file_complete(
                    input_path, 
                    "SUCCESS", 
                    output_path, 
                    stats=self.current_file_stats
                )
            
            return True
            
        except Exception as e:
            error_msg = f"Error converting {input_path.name}: {str(e)}\n{traceback.format_exc()}"
            print(f"\n✗ {error_msg}")
            
            # Check if this was a modal dialog error
            if is_modal_dialog_error(e):
                self._publisher_needs_restart = True
                print("    Publisher will be restarted before next file")
                
                # Try to dismiss any dialogs
                dismiss_all_publisher_dialogs(verbose=self.verbose)
            
            # Close any open documents
            try:
                if 'pub_doc' in locals() and pub_doc is not None:
                    try:
                        pub_doc.Close()
                    except:
                        pass  # May fail if modal dialog is blocking
                if 'ppt_pres' in locals() and ppt_pres is not None:
                    try:
                        ppt_pres.Close()
                    except:
                        pass
            except:
                pass
            
            # Log failure
            if self.logger:
                self.logger.log_file_complete(input_path, "FAILED", error_msg=error_msg)
            
            return False
    
    def convert_to_images(self, input_path, output_path):
        """Fallback method: convert Publisher pages to images in PowerPoint."""
        try:
            pub_doc = self.publisher.Open(str(Path(input_path).absolute()))
            ppt_pres = self.powerpoint.Presentations.Add()
            
            # Set slide size
            first_page = pub_doc.Pages(1)
            ppt_pres.PageSetup.SlideWidth = first_page.Width
            ppt_pres.PageSetup.SlideHeight = first_page.Height
            
            if self.logger:
                self.logger.log_message(f"Exporting {pub_doc.Pages.Count} pages as images...")
            else:
                print(f"Exporting {pub_doc.Pages.Count} pages as images...")
            
            # Export each page as image
            for page_num in range(1, pub_doc.Pages.Count + 1):
                pub_page = pub_doc.Pages(page_num)
                
                # Export as image
                temp_image = self.temp_dir / f"page_{page_num}.png"
                pub_page.SaveAsPicture(str(temp_image))
                
                # Add slide
                if page_num == 1:
                    ppt_slide = ppt_pres.Slides(1)
                else:
                    ppt_slide = ppt_pres.Slides.Add(page_num, 12)
                
                # Add image to slide
                ppt_slide.Shapes.AddPicture(
                    FileName=str(temp_image),
                    LinkToFile=False,
                    SaveWithDocument=True,
                    Left=0,
                    Top=0,
                    Width=first_page.Width,
                    Height=first_page.Height
                )
                
                # Clean up
                if temp_image.exists():
                    temp_image.unlink()
            
            # Save
            ppt_pres.SaveAs(str(Path(output_path).absolute()))
            ppt_pres.Close()
            pub_doc.Close()
            
            if self.logger:
                self.logger.log_message("Image fallback successful")
            else:
                print("✓ Image fallback successful")
            
            return True
            
        except Exception as e:
            error_msg = f"Image fallback failed: {str(e)}"
            if self.logger:
                self.logger.log_message(error_msg, "ERROR")
            else:
                print(f"✗ {error_msg}")
            return False
    
    def convert_batch(self, input_paths, output_dir=None, fallback_to_image=True, 
                     checkpoint_interval=5):
        """
        Convert multiple Publisher files to PowerPoint with checkpoint verification.
        
        Args:
            input_paths: List of Publisher file paths or directories
            output_dir: Output directory for converted files
            fallback_to_image: If True, fall back to image export on failure
            checkpoint_interval: Pause after this many files for user verification
        """
        # Collect all .pub files
        pub_files = []
        for path in input_paths:
            path = Path(path)
            if path.is_file() and path.suffix.lower() == '.pub':
                pub_files.append(path)
            elif path.is_dir():
                pub_files.extend(path.rglob('*.pub'))
        
        if not pub_files:
            msg = "No Publisher files found to convert."
            if self.logger:
                self.logger.log_message(msg, "WARNING")
            else:
                print(msg)
            return
        
        total_files = len(pub_files)
        msg = f"Found {total_files} Publisher file(s) to convert"
        if self.logger:
            self.logger.log_message(msg)
        else:
            print(f"\n{msg}")
        
        # Determine output directory
        if output_dir:
            output_dir = Path(output_dir)
            output_dir.mkdir(parents=True, exist_ok=True)
        
        converted_count = 0
        consecutive_failures = 0
        max_consecutive_failures = 3
        
        # Convert each file with checkpoint verification
        for idx, pub_file in enumerate(pub_files, 1):
            # Determine output path
            if output_dir:
                output_path = output_dir / pub_file.with_suffix('.pptx').name
            else:
                output_path = pub_file.with_suffix('.pptx')
            
            # Check if Publisher needs restart (from previous failure)
            if self._publisher_needs_restart:
                print(f"\n    Recovering Publisher before processing {pub_file.name}...")
                self._restart_publisher()
            
            # Convert file
            try:
                success = self.convert_file(pub_file, output_path, fallback_to_image)
                if success:
                    converted_count += 1
                    consecutive_failures = 0  # Reset failure counter on success
                else:
                    consecutive_failures += 1
            except Exception as e:
                consecutive_failures += 1
                error_msg = f"Unexpected error converting {pub_file}: {str(e)}\n{traceback.format_exc()}"
                
                # Check if it's a modal dialog error
                if is_modal_dialog_error(e):
                    self._publisher_needs_restart = True
                    # Try to dismiss dialogs immediately
                    dismiss_all_publisher_dialogs(verbose=self.verbose)
                
                if self.logger:
                    self.logger.log_file_complete(pub_file, "FAILED", error_msg=error_msg)
                else:
                    print(f"✗ {error_msg}")
            
            # If too many consecutive failures, force restart Publisher
            if consecutive_failures >= max_consecutive_failures:
                print(f"\n    {consecutive_failures} consecutive failures - restarting Publisher...")
                self._restart_publisher()
                consecutive_failures = 0
            
            # Small delay between files to let things stabilize
            time.sleep(0.3)
            
            # (checkpoint logic removed)
        
        # Final summary
        msg = f"\n{'='*80}\nBatch conversion complete: {converted_count}/{total_files} files converted successfully"
        if self.logger:
            self.logger.log_message(msg)
        else:
            print(msg)
    






def load_batch_csv(csv_path):
    """
    Loads test_batch.csv and ensures a 'Conversion Status' column exists.
    Returns: (rows, fieldnames)
    rows: list of dicts (DictReader rows)
    fieldnames: list of column headers
    """
    csv_path = Path(csv_path)

    if not csv_path.exists():
        raise FileNotFoundError(f"CSV not found: {csv_path}")

    with csv_path.open("r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        fieldnames = list(reader.fieldnames) if reader.fieldnames else []
        rows = list(reader)

    if "Conversion Status" not in fieldnames:
        fieldnames.append("Conversion Status")
        for r in rows:
            r["Conversion Status"] = ""

    return rows, fieldnames


def write_batch_csv(csv_path, rows, fieldnames):
    """
    Writes the CSV back to disk (atomic-ish replace) to persist status updates.
    """
    csv_path = Path(csv_path)
    temp_path = csv_path.with_suffix(".tmp")

    with temp_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    temp_path.replace(csv_path)


def should_process_row(row):
    """
    Skip rows marked Success; process blanks and Failures.
    """
    status = (row.get("Conversion Status") or "").strip().lower()
    return status != "success"


def get_conversionlogs_dir(args_log_dir):
    """
    Resolve the ConversionLogs directory.
    If --log-dir is provided, use it; otherwise default to C:/ConversionLogs.
    """
    if args_log_dir:
        return Path(args_log_dir)
    return Path("C:/ConversionLogs")


def write_summary_log(summary_dir, attempted, successes, failures, left_after, failed_entries):
    """
    Write a run summary log file into ConversionLogs with prefix 'summary_log_'.
    """
    summary_dir = Path(summary_dir)
    summary_dir.mkdir(parents=True, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    summary_path = summary_dir / f"summary_log_{ts}.txt"

    lines = []
    lines.append("Publisher -> PowerPoint Batch Run Summary")
    lines.append(f"Timestamp: {datetime.now().isoformat(timespec='seconds')}")
    lines.append(f"Batch limit (attempted per run): {BATCH_LIMIT}")
    lines.append("")
    lines.append(f"Attempted this run: {attempted}")
    lines.append(f"Successes this run: {successes}")
    lines.append(f"Failures this run: {failures}")
    lines.append(f"Left remaining (not Success in CSV after run): {left_after}")
    lines.append("")
    lines.append("Failed items (this run):")
    if failed_entries:
        for entry in failed_entries:
            lines.append(f" - {entry}")
    else:
        lines.append(" - (none)")
    lines.append("")

    summary_path.write_text("\n".join(lines), encoding="utf-8")
    return summary_path



def main():
    parser = argparse.ArgumentParser(
        description="Convert Microsoft Publisher files to PowerPoint, preserving editable content."
    )

    # Keep only non-interactive options that still matter
    parser.add_argument(
        "-t", "--template",
        ## TEMPLATE FILE PATH
        default=BASE_DIR / "Pub-Pow Blank Template.potx",
        help="PowerPoint template file (.potx) to use"
    )
    parser.add_argument(
        "--log-dir",
        help="Directory for log files (default: C:/ConversionLogs)"
    )
    parser.add_argument(
        "--no-fallback",
        action="store_true",
        help="Disable image fallback for failed conversions"
    )
    parser.add_argument(
        "--no-text",
        action="store_true",
        help="Do not preserve text content"
    )
    parser.add_argument(
        "--no-images",
        action="store_true",
        help="Do not preserve images"
    )
    parser.add_argument(
        "--no-shapes",
        action="store_true",
        help="Do not preserve shapes"
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose output for debugging"
    )

    args = parser.parse_args()

    script_dir = Path(__file__).resolve().parent
    csv_path = CSV_PATH
    output_dir = script_dir / "converted_pptxs"
    output_dir.mkdir(parents=True, exist_ok=True)

    # Initialize logger
    logger = ConversionLogger(log_directory=args.log_dir)
    logger.log_message("Publisher to PowerPoint Converter Started (CSV batch mode)")

    # Load CSV
    rows, fieldnames = load_batch_csv(csv_path)

    try:
        with PublisherToPowerPointConverter(
            preserve_text=not args.no_text,
            preserve_images=not args.no_images,
            preserve_shapes=not args.no_shapes,
            template_path=args.template,
            verbose=args.verbose,
            dry_run=False, # Always run full logic (no dry-run prompt)
            logger=logger
        ) as converter:

            attempted = 0
            successes = 0
            failures = 0
            failed_entries = []

            for row in rows:
                if not should_process_row(row):
                    continue

                # Enforce batch limit (count every attempted file/row)
                if attempted >= BATCH_LIMIT:
                    break

                attempted += 1

                display_label = (row.get("File Name") or "").strip()
                full_path = (row.get("Full File Path") or "").strip()

                if not full_path:
                    row["Conversion Status"] = "Failure"
                    failures += 1
                    failed_entries.append(f"{display_label or '(no File Name)'} | Missing Full File Path")
                    write_batch_csv(csv_path, rows, fieldnames)
                    continue

                input_path = Path(full_path)
                output_path = output_dir / input_path.with_suffix(".pptx").name

                try:
                    success = converter.convert_file(
                        input_path,
                        output_path,
                        fallback_to_image=not args.no_fallback
                    )
                    row["Conversion Status"] = "Success" if success else "Failure"
                    if success:
                        successes += 1
                    else:
                        failures += 1
                        failed_entries.append(f"{display_label or input_path.name} | {str(input_path)}")
                except Exception:
                    row["Conversion Status"] = "Failure"
                    failures += 1
                    failed_entries.append(f"{display_label or input_path.name} | {str(input_path)} | Exception during convert_file")

                # Persist after each file
                write_batch_csv(csv_path, rows, fieldnames)

            # After the run, compute how many are left (anything not marked Success)
            left_after = sum(1 for r in rows if should_process_row(r))

            # Write summary log into ConversionLogs (prefix summary_log_)
            summary_dir = get_conversionlogs_dir(args.log_dir)
            summary_path = write_summary_log(
                summary_dir=summary_dir,
                attempted=attempted,
                successes=successes,
                failures=failures,
                left_after=left_after,
                failed_entries=failed_entries
            )

            # Also note it in the main conversion logger (optional, but non-interactive)
            try:
                logger.log_message(f"Summary log written: {summary_path}")
            except Exception:
                pass

    finally:
        logger.finalize()

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nConversion interrupted by user.")
    except Exception as e:
        print(f"\n\nUnexpected error: {e}")
        print(traceback.format_exc())
    finally:
        # Keep console window open
        print("\n" + "="*60)
        input("Press Enter to close this window...")
        print("="*60)