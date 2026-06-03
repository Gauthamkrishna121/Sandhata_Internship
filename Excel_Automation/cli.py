import os
import sys
import json
from datetime import datetime, date, timedelta
from excel_manager import (
    DEFAULT_FILENAME,
    get_workbook,
    find_next_empty_slot,
    save_log,
    get_existing_log
)

CONFIG_FILENAME = "tracker_config.json"

def load_start_date():
    if os.path.exists(CONFIG_FILENAME):
        try:
            with open(CONFIG_FILENAME, 'r') as f:
                data = json.load(f)
                return datetime.strptime(data["start_date"], "%Y-%m-%d").date()
        except:
            pass
    return None

def save_start_date(start_date):
    try:
        with open(CONFIG_FILENAME, 'w') as f:
            json.dump({"start_date": start_date.strftime("%Y-%m-%d")}, f)
    except Exception as e:
        print(f"Warning: Could not save config: {e}")

def get_default_start_date():
    # Monday of the current week
    today = date.today()
    return today - timedelta(days=today.weekday())

def get_current_slot(start_date):
    today = date.today()
    days_diff = (today - start_date).days
    
    # Calculate week number (7 days per week)
    week_num = (days_diff // 7) + 1
    # Day number (Monday=1, Tuesday=2, ..., Sunday=7)
    day_num = today.weekday() + 1
    
    # If weekend, default to Day 5 (Friday) of the same week
    if day_num > 5:
        day_num = 5
        
    return week_num, day_num

def get_multi_line_input():
    print("Enter your updates for today:")
    print(" (Type your notes. Press Enter on a blank line when you are finished)")
    print("-" * 60)
    lines = []
    while True:
        try:
            line = input()
            if line.strip() == "":
                break
            lines.append(line)
        except (KeyboardInterrupt, EOFError):
            print("\nInput cancelled.")
            sys.exit(0)
    return "\n".join(lines)

def main():
    print("=" * 60)
    print("        SANDHATA INTERNSHIP LOG AUTOMATION")
    print("=" * 60)
    
    filepath = DEFAULT_FILENAME
    
    # Check start date config
    start_date = load_start_date()
    if not start_date:
        start_date = get_default_start_date()
        save_start_date(start_date)
        print(f"[Info] Internship start date initialized to Monday, {start_date.strftime('%B %d, %Y')} (based on system clock).")
        
    # Calculate today's slot
    today_w, today_d = get_current_slot(start_date)
    weekday_name = date.today().strftime("%A")
    today_str = date.today().strftime("%B %d, %Y")
    
    print(f"\n[Status] Today is {weekday_name}, {today_str}")
    print(f"Calculated slot: WEEK {today_w}, DAY {today_d}")
    
    # 2. Confirm or select week/day
    choice = input(f"Use this slot? (Y/n): ").strip().lower()
    if choice in ("", "y", "yes"):
        w = today_w
        d = today_d
    else:
        # User wants to manually specify slot
        while True:
            try:
                w_in = input("Enter Week Number (e.g. 1): ").strip()
                w = int(w_in)
                if w <= 0:
                    raise ValueError
                break
            except ValueError:
                print("Invalid week number. Please enter a positive integer.")
                
        while True:
            try:
                d_in = input("Enter Day Number (1-5): ").strip()
                d = int(d_in)
                if d < 1 or d > 5:
                    raise ValueError
                break
            except ValueError:
                print("Invalid day number. Please enter a number between 1 and 5.")
                
    # Check if there is an existing log for this day
    try:
        existing_log = get_existing_log(filepath, w, d)
    except Exception:
        existing_log = None
        
    append_choice = True
    if existing_log:
        print(f"\n[Notice] Existing log found for WEEK {w}, DAY {d}:")
        print(f"  \"{existing_log}\"")
        while True:
            action = input("Choose action: (a) Append new notes, (o) Overwrite today's notes, (c) Cancel [Default: a]: ").strip().lower()
            if not action or action == 'a':
                append_choice = True
                break
            elif action == 'o':
                append_choice = False
                break
            elif action == 'c':
                print("Cancelled. Exiting without saving.")
                return
            else:
                print("Invalid choice. Please select 'a', 'o', or 'c'.")

    # 3. Get log content
    raw_text = get_multi_line_input()
    if not raw_text.strip():
        print("No content entered. Exiting without saving.")
        return
        
    # 4. Save and format with retry loop for open files
    while True:
        try:
            formatted_text = save_log(filepath, w, d, raw_text, append=append_choice)
            print("-" * 60)
            print(f"[SUCCESS] Log saved to WEEK {w}, DAY {d}.")
            print(f"Formatted log text:\n  {formatted_text}")
            print("-" * 60)
            break
        except PermissionError:
            print(f"\n[ERROR] Permission denied: Cannot write to '{filepath}'.")
            print("This usually happens because the Excel file is open in Microsoft Excel.")
            input("--> Please CLOSE the Excel file and press Enter to try saving again...")
        except Exception as e:
            print(f"Error saving log to Excel: {e}")
            return
        
    # 5. Open Excel sheet option
    open_choice = input("Would you like to open the Excel sheet? (y/N): ").strip().lower()
    if open_choice in ("y", "yes"):
        try:
            print(f"Opening {filepath}...")
            # On Windows, os.startfile opens the file with the default associated application
            if hasattr(os, 'startfile'):
                os.startfile(filepath)
            else:
                # Fallback for non-Windows (just in case, though the user is on Windows)
                import subprocess
                subprocess.call(['open', filepath])
        except Exception as e:
            print(f"Could not open spreadsheet: {e}")
            
    print("\nThank you for updating your log! Have a great day.")

if __name__ == "__main__":
    main()
