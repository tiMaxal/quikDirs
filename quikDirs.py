import os, sys, shutil, json, ctypes, subprocess
from datetime import datetime

# Version tracking
VERSION = "0.8.0"  # Fixed get_junction_target, added broken junction repair, updated help

# =================== HUB SETTINGS ===================
HUB = r"C:\$quikDirs"
APP_NAME = "$quikDirs"  # Using $ prefix moves it to top of Send To menu
INSTALLED_SCRIPT = os.path.join(HUB, "quikdirs.py")
SENDTO = os.path.join(os.environ["APPDATA"], r"Microsoft\Windows\SendTo")
SENDTO_LNK = os.path.join(SENDTO, f"{APP_NAME}.lnk")
LOG_FILE = os.path.join(HUB, "quikdirs.log")
QUIKDIRS_CSV = os.path.join(HUB, "quikdirs.csv")

# Readme content for Help button
README_TEXT = """
# QuikDirs – Help

Windows 11 does not support alphabetical sorting of pinned Quick Access items.
QuikDirs works around this by pinning a single hub folder and letting Explorer
alphabetically sort the contents of that folder.

## Key Features:
• Hub folder: C:\\$quikDirs
• Uses NTFS junctions (not shortcuts) for instant navigation
• Automatic duplicate detection - won't create multiple junctions to same folder
• Target folder shortcuts - each linked folder gets a shortcut back to $quikDirs
• CSV tracking of all junctions with creation dates
• Broken junction detection and repair
• Deleting a junction is safe - only removes the link, not the target

## Collision Naming:
• Same folder name, different paths: Name[1], Name[2], etc.
• Same target path: reuses existing junction (no duplicates)

## Usage:
• Right-click a folder → Send To → $quikDirs
• Junction appears instantly in the hub folder
• Each target folder gets a $quikDirs__[name].lnk shortcut
• Optional import of pinned Quick Access folders via GUI
• Undo last operations as a single step

## GUI Features:
• Import: Green button, creates junctions for selected folders
• Undo Last: Yellow button, reverses last import operation
• Refresh: Updates list from Quick Access
• Help: Shift+? opens this help, Esc closes
• Close on completion: Checkbox, optional auto-close after import
• Orphaned Quick Access items: Can be replaced or removed
• Broken junctions: Detected and repairable

## Keyboard Shortcuts:
• Shift+? : Open help window
• Esc (in help): Close help window

## Safety:
• Only pinned Quick Access items are affected
• Frequent / recent folders are ignored
• No registry modifications
• Junctions are just directory links - safe to delete
• CSV file tracks all junctions: C:\\$quikDirs\\quikdirs.csv

## Technical Details:
• NTFS junction creation via mklink
• Automatic cleanup of orphaned shortcuts
• Path normalization for reliable duplicate detection
• Filesystem-based verification (not reliant on cached data)
"""


# ---------------- Privilege / dependency ----------------
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


def ensure_pywin32():
    try:
        import win32com.client  # noqa

        return True
    except ImportError:
        return False


def install_pywin32():
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])


# ---------------- Shell helpers ----------------
def pin_to_quick_access(path):
    import win32com.client

    try:
        shell = win32com.client.Dispatch("Shell.Application")
        folder = shell.Namespace(path)
        if folder and folder.Self:
            folder.Self.InvokeVerb("pintohome")
            return True
        return False
    except Exception:
        return False


def clear_quick_access_file():
    """Clear the Quick Access automatic destinations file (nuclear option)"""
    import glob
    qa_path = os.path.join(os.environ["APPDATA"], 
                           r"Microsoft\Windows\Recent\AutomaticDestinations")
    try:
        # Backup first
        backup_path = os.path.join(HUB, "quickaccess_backup")
        os.makedirs(backup_path, exist_ok=True)
        
        qa_files = glob.glob(os.path.join(qa_path, "*.automaticDestinations-ms"))
        for qa_file in qa_files:
            if os.path.exists(qa_file):
                backup_file = os.path.join(backup_path, 
                                          f"{os.path.basename(qa_file)}.{datetime.now().strftime('%Y%m%d%H%M%S')}")
                shutil.copy2(qa_file, backup_file)
                os.remove(qa_file)
        return True
    except Exception as e:
        return False


def unpin_from_quick_access(path):
    import win32com.client
    import time

    shell = win32com.client.Dispatch("Shell.Application")
    qa = shell.Namespace("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}")
    found = False
    
    # Normalize target path
    target_path_normalized = os.path.normcase(os.path.normpath(path))
    
    # Try up to 3 times with delays
    for attempt in range(3):
        for item in qa.Items():
            try:
                # Normalize item path
                item_path_normalized = os.path.normcase(os.path.normpath(item.Path))
                
                if item_path_normalized == target_path_normalized:
                    item.InvokeVerb("unpinfromhome")
                    found = True
                    time.sleep(0.5)  # Give Windows time to process
                    break
            except Exception:
                pass
        
        if found:
            # Verify it was actually removed
            time.sleep(0.5)
            qa = shell.Namespace("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}")
            still_there = False
            for item in qa.Items():
                try:
                    item_path_normalized = os.path.normcase(os.path.normpath(item.Path))
                    if item_path_normalized == target_path_normalized:
                        still_there = True
                        break
                except Exception:
                    pass
            
            if not still_there:
                return True  # Successfully removed
            else:
                found = False  # Try again
        
        if attempt < 2:
            time.sleep(0.5)
    
    return found


def add_sendto_shortcut(target_script):
    import win32com.client

    shell = win32com.client.Dispatch("WScript.Shell")
    s = shell.CreateShortcut(SENDTO_LNK)
    s.TargetPath = sys.executable
    s.Arguments = f'"{target_script}"'
    s.WorkingDirectory = HUB
    s.IconLocation = sys.executable
    s.Save()


def exclude_from_google_drive():
    """Create .gdignore file to exclude hub from Google Drive monitoring"""
    try:
        parent_dir = os.path.dirname(HUB)
        gdignore_file = os.path.join(parent_dir, ".gdignore")
        hub_name = os.path.basename(HUB)
        
        # Read existing content if file exists
        existing_rules = set()
        if os.path.exists(gdignore_file):
            with open(gdignore_file, "r") as f:
                existing_rules = set(line.strip() for line in f if line.strip())
        
        # Add our exclusion rule if not present
        if hub_name not in existing_rules:
            existing_rules.add(hub_name)
            with open(gdignore_file, "w") as f:
                f.write("\n".join(sorted(existing_rules)) + "\n")
    except Exception:
        pass  # Silently fail if Google Drive not installed or permission issues


# ---------------- Junction / collision ----------------
def parent_name(path):
    return os.path.basename(os.path.dirname(path.rstrip("\\/")))


def get_junction_target(junction_path):
    """Get the target path of a junction"""
    try:
        import win32file
        # Open WITHOUT FILE_FLAG_OPEN_REPARSE_POINT so we follow the junction
        handle = win32file.CreateFile(
            junction_path,
            0,  # No access needed, just query
            win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE,
            None,
            win32file.OPEN_EXISTING,
            win32file.FILE_FLAG_BACKUP_SEMANTICS,  # Required for directories, but DON'T use OPEN_REPARSE_POINT
            None
        )
        try:
            target = win32file.GetFinalPathNameByHandle(handle, 0)
            # Remove \\?\ prefix and normalize
            target = target.replace("\\\\?\\", "")
            target = target.replace("\\\\?\\UNC\\", "\\\\")
            return os.path.normpath(target)
        finally:
            handle.Close()
    except Exception:
        return None


def find_existing_junctions():
    """Find all existing junctions in HUB and their targets"""
    junctions = {}
    if not os.path.exists(HUB):
        return junctions
    for item in os.listdir(HUB):
        item_path = os.path.join(HUB, item)
        if os.path.isdir(item_path):
            target = get_junction_target(item_path)
            if target:
                target_normalized = os.path.normcase(os.path.normpath(target))
                junctions[item_path] = target_normalized
    return junctions


def find_duplicate_junctions(target):
    """Find junctions pointing to the same target, return shortest name"""
    target_normalized = os.path.normcase(os.path.normpath(target))
    existing = find_existing_junctions()
    duplicates = [j for j, t in existing.items() if t == target_normalized]
    return duplicates


def remove_duplicate_junctions(keep_junction):
    """Remove all junctions with same target, keeping the shortest named one"""
    if not os.path.exists(keep_junction):
        return []
    
    target = get_junction_target(keep_junction)
    if not target:
        return []
    
    duplicates = find_duplicate_junctions(target)
    if keep_junction in duplicates:
        duplicates.remove(keep_junction)
    
    # Keep shortest name
    if duplicates:
        all_junctions = [keep_junction] + duplicates
        shortest = min(all_junctions, key=lambda x: len(os.path.basename(x)))
        if shortest != keep_junction:
            duplicates.remove(shortest)
            duplicates.append(keep_junction)
            keep_junction = shortest
    
    removed = []
    junctions = read_junctions_csv()
    for dup in duplicates:
        try:
            # Remove from CSV
            jname = os.path.basename(dup)
            if jname in junctions:
                del junctions[jname]
            # Remove junction
            os.rmdir(dup)
            removed.append(dup)
        except Exception:
            pass
    
    # Write updated CSV
    if removed:
        write_junctions_csv(junctions)
    
    return removed


def junction_dest(target):
    base = os.path.basename(target.rstrip("\\/"))
    dest = os.path.join(HUB, base)
    
    # Normalize target for comparison
    target_normalized = os.path.normcase(os.path.normpath(target))
    
    if not os.path.exists(dest):
        return dest
    
    # Check if existing item is a junction pointing to the same target
    existing_target = get_junction_target(dest)
    if existing_target:
        existing_normalized = os.path.normcase(os.path.normpath(existing_target))
        if existing_normalized == target_normalized:
            return dest  # Same target, use existing
    
    # Real collision - different target, same name - use numeric suffix
    counter = 1
    while True:
        numbered_dest = os.path.join(HUB, f"{base}[{counter}]")
        if not os.path.exists(numbered_dest):
            return numbered_dest
        # Check if this one points to our target
        existing_target = get_junction_target(numbered_dest)
        if existing_target:
            existing_normalized = os.path.normcase(os.path.normpath(existing_target))
            if existing_normalized == target_normalized:
                return numbered_dest  # Found existing junction to same target
        counter += 1
        # Safety check to prevent infinite loop
        if counter > 100:
            break
    return numbered_dest


def create_junction(target):
    # Normalize target path for consistent comparison
    target_normalized = os.path.normpath(target)
    target_check = os.path.normcase(target_normalized)
    
    # Don't create junction to HUB itself
    hub_check = os.path.normcase(os.path.normpath(HUB))
    if target_check == hub_check:
        print(f"Skipping: Cannot create junction to $quikDirs itself")
        return None
    
    # FIRST: Scan all existing junctions in HUB to find if target already exists
    if os.path.exists(HUB):
        for item in os.listdir(HUB):
            item_path = os.path.join(HUB, item)
            if os.path.isdir(item_path):
                # Check if it's a junction and get its target
                existing_target = get_junction_target(item_path)
                if existing_target:
                    existing_check = os.path.normcase(os.path.normpath(existing_target))
                    if existing_check == target_check:
                        # Found existing junction to same target!
                        update_junctions_csv(item_path, target_normalized)
                        return item_path
    
    # No existing junction found, determine best name
    base = os.path.basename(target_normalized.rstrip("\\/"))
    dest = os.path.join(HUB, base)
    
    # Check if base name is available
    if not os.path.exists(dest):
        # Base name is free, use it
        subprocess.run(
            ["cmd", "/c", "mklink", "/J", dest, target_normalized],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            shell=True,
        )
        update_junctions_csv(dest, target_normalized)
        return dest
    
    # Base name exists, need to find a numbered variant
    counter = 1
    while counter <= 100:
        dest = os.path.join(HUB, f"{base}[{counter}]")
        if not os.path.exists(dest):
            # Found free numbered slot
            subprocess.run(
                ["cmd", "/c", "mklink", "/J", dest, target_normalized],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                shell=True,
            )
            update_junctions_csv(dest, target_normalized)
            return dest
        counter += 1
    
    # Fallback if somehow we hit 100 variants
    return dest


def read_junctions_csv():
    """Read the junctions CSV file, return dict of {junction_name: (target, date_created)}"""
    junctions = {}
    if os.path.exists(QUIKDIRS_CSV):
        try:
            import csv
            with open(QUIKDIRS_CSV, "r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    junctions[row["junction_name"]] = (row["target_path"], row["date_created"])
        except Exception:
            pass
    return junctions


def write_junctions_csv(junctions_dict):
    """Write junctions dict to CSV, sorted alphabetically by junction name"""
    try:
        import csv
        with open(QUIKDIRS_CSV, "w", encoding="utf-8", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=["junction_name", "target_path", "date_created"])
            writer.writeheader()
            for jname in sorted(junctions_dict.keys(), key=str.lower):
                target, date_created = junctions_dict[jname]
                writer.writerow({
                    "junction_name": jname,
                    "target_path": target,
                    "date_created": date_created
                })
    except Exception:
        pass


def update_junctions_csv(junction_path, target):
    """Update CSV file with new junction, cleanup orphaned shortcuts, and add missing junctions"""
    junction_name = os.path.basename(junction_path)
    junctions = read_junctions_csv()
    
    # Cleanup: check all existing entries, remove if junction no longer exists
    to_remove = []
    for jname, (jtarget, _) in junctions.items():
        jpath = os.path.join(HUB, jname)
        if not os.path.exists(jpath):
            # Junction removed, cleanup shortcut in target folder
            to_remove.append(jname)
            remove_target_shortcut(jtarget, jname)
    
    for jname in to_remove:
        del junctions[jname]
    
    # Scan $quikDirs directory for junctions not in CSV and add them
    if os.path.exists(HUB):
        for item in os.listdir(HUB):
            item_path = os.path.join(HUB, item)
            if os.path.isdir(item_path) and item not in junctions:
                # Check if it's a junction
                item_target = get_junction_target(item_path)
                if item_target:
                    # Found a junction not in CSV, add it
                    junctions[item] = (item_target, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    
    # Add or update current junction
    junctions[junction_name] = (target, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    
    # Ensure all target folders have shortcuts (using actual junction targets from filesystem)
    for jname, (jtarget, _) in junctions.items():
        jpath = os.path.join(HUB, jname)
        if os.path.exists(jpath):
            # Get the ACTUAL target from the junction, not from CSV
            actual_target = get_junction_target(jpath)
            if actual_target and os.path.exists(actual_target):
                create_target_shortcut(jpath, actual_target)
    
    # Write updated CSV
    write_junctions_csv(junctions)


def create_target_shortcut(junction_path, target):
    """Create a shortcut IN THE TARGET FOLDER that points to the target folder itself"""
    try:
        import win32com.client
        junction_name = os.path.basename(junction_path)
        target_dir_name = os.path.basename(target.rstrip("\\/"))
        shortcut_name = f"$quikDirs__{target_dir_name}.lnk"
        shortcut_path = os.path.join(target, shortcut_name)
        
        shell = win32com.client.Dispatch("WScript.Shell")
        s = shell.CreateShortcut(shortcut_path)
        s.TargetPath = target  # Point to the target folder itself
        s.WorkingDirectory = target
        s.Description = f"Shortcut to {target_dir_name} (linked from $quikDirs)"
        s.IconLocation = "shell32.dll,3"  # Folder icon
        s.Save()
    except Exception:
        pass  # Silently fail if we can't create the shortcut


def remove_target_shortcut(target, junction_name):
    """Remove shortcut from target folder when junction is removed"""
    try:
        target_dir_name = os.path.basename(target.rstrip("\\/"))
        shortcut_name = f"$quikDirs__{target_dir_name}.lnk"
        shortcut_path = os.path.join(target, shortcut_name)
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)
    except Exception:
        pass


# ---------------- Logging / Undo ----------------
def log_operation(actions):
    entry = {
        "timestamp": datetime.now().isoformat(),
        "action": "operation",
        "actions": actions,
    }
    data = []
    if os.path.exists(LOG_FILE):
        try:
            with open(LOG_FILE, "r") as f:
                data = json.load(f)
        except Exception:
            data = []
    data.append(entry)
    with open(LOG_FILE, "w") as f:
        json.dump(data, f, indent=2)


def undo_last_operation():
    if not os.path.exists(LOG_FILE):
        return
    with open(LOG_FILE, "r") as f:
        data = json.load(f)
    if not data:
        return
    entry = data.pop()
    if entry.get("action") == "operation":
        for act in reversed(entry["actions"]):
            if act["type"] == "create_junction" and os.path.exists(act["junction"]):
                # Remove shortcut from target folder first
                target = act.get("target")
                if target:
                    remove_target_shortcut(target, os.path.basename(act["junction"]))
                # Remove junction
                os.rmdir(act["junction"])
                # Update CSV to remove this entry
                junctions = read_junctions_csv()
                jname = os.path.basename(act["junction"])
                if jname in junctions:
                    del junctions[jname]
                    write_junctions_csv(junctions)
            elif act["type"] == "remove_pin":
                try:
                    pin_to_quick_access(act["target"])
                except Exception:
                    pass
    with open(LOG_FILE, "w") as f:
        json.dump(data, f, indent=2)


# ---------------- Quick Access enumeration ----------------
def get_pinned_quick_access():
    import win32com.client

    shell = win32com.client.Dispatch("Shell.Application")
    qa = shell.Namespace("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}")
    paths = []
    orphaned = []
    for item in qa.Items():
        try:
            if item.IsFolder:
                if os.path.isdir(item.Path):
                    paths.append(item.Path)
                else:
                    orphaned.append(item.Path)
        except Exception:
            pass
    return sorted(set(paths), key=str.lower), orphaned


def find_broken_junctions():
    """Find all junctions that point to non-existent directories"""
    broken = []
    if os.path.exists(HUB):
        for item in os.listdir(HUB):
            item_path = os.path.join(HUB, item)
            if os.path.isdir(item_path):
                target = get_junction_target(item_path)
                if target and not os.path.exists(target):
                    broken.append((item, item_path, target))
    return broken


def show_broken_junctions_dialog(parent, broken_junctions):
    """Show dialog to repair or remove broken junctions"""
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog
    
    dialog = tk.Toplevel(parent)
    dialog.title("Repair Broken Junctions")
    dialog.resizable(True, True)
    dialog.geometry("800x400")
    
    frame = ttk.Frame(dialog, padding=10)
    frame.pack(fill="both", expand=True)
    
    ttk.Label(frame, text="The following junctions point to folders that no longer exist:").pack(anchor="w", pady=(0, 10))
    
    # Scrollable list
    canvas = tk.Canvas(frame)
    scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
    scroll_frame = ttk.Frame(canvas)
    scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    for jname, jpath, old_target in broken_junctions:
        item_frame = ttk.Frame(scroll_frame)
        item_frame.pack(anchor="w", fill="x", pady=5, padx=5)
        
        ttk.Label(item_frame, text=f"❌ {jname}", font=("TkDefaultFont", 10, "bold")).pack(anchor="w")
        ttk.Label(item_frame, text=f"   Broken target: {old_target}", foreground="red").pack(anchor="w")
        
        btn_frame = ttk.Frame(item_frame)
        btn_frame.pack(anchor="w", pady=(2, 0))
        
        def make_repair(jp=jpath, jn=jname, ot=old_target, if_=item_frame):
            def repair():
                new_target = filedialog.askdirectory(title=f"Select new location for: {jn}")
                if new_target and os.path.isdir(new_target):
                    # Remove old junction
                    try:
                        os.rmdir(jp)
                        # Create new junction
                        subprocess.run(
                            ["cmd", "/c", "mklink", "/J", jp, new_target],
                            stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL,
                            shell=True,
                        )
                        # Update CSV
                        update_junctions_csv(jp, new_target)
                        # Update UI
                        for widget in if_.winfo_children():
                            widget.destroy()
                        ttk.Label(if_, text=f"✓ {jn} → {new_target}", foreground="green").pack(anchor="w")
                        messagebox.showinfo(APP_NAME, f"Repaired: {jn}")
                    except Exception as e:
                        messagebox.showerror(APP_NAME, f"Failed to repair: {e}")
            return repair
        
        def make_remove(jp=jpath, jn=jname, if_=item_frame):
            def remove():
                try:
                    os.rmdir(jp)
                    # Update CSV
                    junctions = read_junctions_csv()
                    if jn in junctions:
                        del junctions[jn]
                        write_junctions_csv(junctions)
                    # Update UI
                    if_.destroy()
                    messagebox.showinfo(APP_NAME, f"Removed: {jn}")
                except Exception as e:
                    messagebox.showerror(APP_NAME, f"Failed to remove: {e}")
            return remove
        
        ttk.Button(btn_frame, text="Repair (Select New Location)", command=make_repair()).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="Remove Junction", command=make_remove()).pack(side="left", padx=2)
    
    # Close button
    close_btn = tk.Button(frame, text="Close", bg="red", fg="white", width=15, command=dialog.destroy)
    close_btn.pack(pady=(10, 0))
    
    dialog.focus_force()


# ---------------- GUI ----------------
def run_import_gui():
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog
    import tkinter.font as tkfont

    pinned, orphaned = get_pinned_quick_access()
    if not pinned and not orphaned:
        messagebox.showinfo(APP_NAME, "No pinned Quick Access folders found.")
        return

    root = tk.Tk()
    root.title("QuikDirs")
    root.resizable(True, True)

    rows = []
    orphan_replacements = {}  # Maps orphaned path to replacement path

    frm_outer = ttk.Frame(root, padding=10)
    frm_outer.pack(fill="both", expand=True)

    # Scrollable canvas
    canvas = tk.Canvas(frm_outer)
    scrollbar = ttk.Scrollbar(frm_outer, orient="vertical", command=canvas.yview)
    scroll_frame = ttk.Frame(canvas)
    scroll_frame.bind(
        "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="top", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Header with refresh button
    header_frame = ttk.Frame(scroll_frame)
    header_frame.pack(anchor="w", fill="x", pady=(0, 5))
    ttk.Label(header_frame, text="Select pinned Quick Access folders to import:").pack(
        side="left"
    )
    
    def refresh_list():
        nonlocal rows, orphaned, orphan_replacements
        # Clear existing rows
        for widget in scroll_frame.winfo_children():
            if widget != header_frame:
                widget.destroy()
        rows.clear()
        orphan_replacements.clear()
        
        # Reload pinned items
        pinned, orphaned = get_pinned_quick_access()
        
        # Show orphaned items first if any
        if orphaned:
            orphan_section = ttk.LabelFrame(scroll_frame, text="⚠ Orphaned Quick Access Items", padding=5)
            orphan_section.pack(anchor="w", fill="both", expand=True, pady=(0, 10))
            
            for orphan_path in orphaned:
                orphan_line = ttk.Frame(orphan_section)
                orphan_line.pack(anchor="w", fill="x", pady=2)
                
                ttk.Label(orphan_line, text="❌", foreground="red").pack(side="left")
                ttk.Label(orphan_line, text=orphan_path, anchor="w").pack(side="left", fill="x", expand=True, padx=5)
                
                def make_picker(op=orphan_path, ol=orphan_line):
                    def pick_replacement():
                        replacement = filedialog.askdirectory(title=f"Replace: {op}")
                        if replacement and os.path.isdir(replacement):
                            # First unpin the orphan
                            unpin_from_quick_access(op)
                            
                            # Try to pin the replacement
                            success = pin_to_quick_access(replacement)
                            
                            if success:
                                orphan_replacements[op] = replacement
                                
                                # Update label
                                for w in ol.winfo_children():
                                    w.destroy()
                                ttk.Label(ol, text="✓", foreground="green").pack(side="left")
                                ttk.Label(ol, text=f"{op} → {replacement}", anchor="w").pack(side="left", fill="x", expand=True, padx=5)
                                
                                messagebox.showinfo(APP_NAME, 
                                    f"Orphan removed from Quick Access.\n"
                                    f"Replacement pinned: {replacement}\n\n"
                                    f"Click 'Refresh' to see it in the list.")
                            else:
                                messagebox.showerror(APP_NAME, 
                                    f"Failed to pin replacement to Quick Access.\n"
                                    f"The orphan has been removed, but you may need to\n"
                                    f"manually pin: {replacement}")
                    return pick_replacement
                
                def make_remover(op=orphan_path):
                    def remove_orphan():
                        unpin_from_quick_access(op)
                        messagebox.showinfo(APP_NAME, f"Removed: {op}")
                        refresh_list()
                    return remove_orphan
                
                ttk.Button(orphan_line, text="Replace...", command=make_picker()).pack(side="right", padx=2)
                ttk.Button(orphan_line, text="Remove", command=make_remover()).pack(side="right", padx=2)
        
        if not pinned:
            messagebox.showinfo(APP_NAME, "No valid pinned Quick Access folders found.")
            return
        
        # Recreate rows for valid items
        for p in pinned:
            # Check if already exists as junction
            existing_junctions = find_duplicate_junctions(p)
            already_imported = bool(existing_junctions)
            
            v_import = tk.BooleanVar(value=not already_imported)
            v_remove = tk.BooleanVar(value=(os.path.normcase(p) != os.path.normcase(HUB)))
            rows.append((p, v_import, v_remove))
            line = ttk.Frame(scroll_frame)
            line.pack(anchor="w", fill="x", pady=2)
            
            if already_imported:
                ttk.Label(line, text="✓", foreground="green").pack(side="left")
            
            ttk.Checkbutton(line, variable=v_import).pack(side="left")
            status = " (already imported)" if already_imported else ""
            ttk.Label(line, text=p + status, width=80, anchor="w").pack(side="left", padx=5)
            ttk.Checkbutton(line, text="Remove", variable=v_remove).pack(
                side="left", padx=5
            )
        canvas.configure(scrollregion=canvas.bbox("all"))
    
    ttk.Button(header_frame, text="Refresh", command=refresh_list).pack(side="left", padx=10)
    
    def clear_qa():
        if messagebox.askyesno(APP_NAME, 
            "⚠ Nuclear Option: Clear All Quick Access?\n\n"
            "This will:\n"
            "• Backup current Quick Access to $quikDirs\\quickaccess_backup\n"
            "• Clear ALL Quick Access items\n"
            "• Restart Explorer\n\n"
            "You'll need to re-pin folders you want.\n\n"
            "Continue?"):
            if clear_quick_access_file():
                messagebox.showinfo(APP_NAME, 
                    "Quick Access cleared!\n"
                    "Restarting Explorer...\n"
                    "Please re-pin $quikDirs after restart.")
                subprocess.run(["taskkill", "/F", "/IM", "explorer.exe"], 
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                subprocess.Popen("explorer.exe")
            else:
                messagebox.showerror(APP_NAME, "Failed to clear Quick Access.")
    
    ttk.Button(header_frame, text="⚠ Clear All QA", command=clear_qa).pack(side="left", padx=10)

    # Initial population
    refresh_list()

    # Check for broken junctions
    broken_junctions = find_broken_junctions()
    if broken_junctions:
        if messagebox.askyesno(APP_NAME, 
            f"Found {len(broken_junctions)} broken junction(s) pointing to non-existent folders.\n\n"
            "Would you like to review and repair them?"):
            show_broken_junctions_dialog(root, broken_junctions)

    # Options frame
    options_frame = ttk.Frame(frm_outer)
    options_frame.pack(anchor="w", pady=5)
    
    close_var = tk.BooleanVar(value=False)
    ttk.Checkbutton(options_frame, text="Close on completion", variable=close_var).pack(
        anchor="w"
    )

    # Buttons frame at bottom
    btns = ttk.Frame(frm_outer)
    btns.pack(fill="x", pady=(10, 0))

    # Select All/None buttons frame
    select_frame = ttk.Frame(btns)
    select_frame.pack(side="top", fill="x", pady=(0, 5))
    ttk.Button(
        select_frame,
        text="All Import",
        command=lambda: [v_i.set(True) for _, v_i, _ in rows],
    ).pack(side="left", padx=2)
    ttk.Button(
        select_frame,
        text="None Import",
        command=lambda: [v_i.set(False) for _, v_i, _ in rows],
    ).pack(side="left", padx=2)
    ttk.Button(
        select_frame,
        text="All Remove",
        command=lambda: [
            v_r.set(True)
            for p, _, v_r in rows
            if os.path.normcase(p) != os.path.normcase(HUB)
        ],
    ).pack(side="left", padx=2)
    ttk.Button(
        select_frame,
        text="None Remove",
        command=lambda: [v_r.set(False) for _, _, v_r in rows],
    ).pack(side="left", padx=2)

    # Action functions
    def do_import():
        actions = []
        failed_removals = []
        
        # Handle orphan replacements first
        for orphan_path, replacement in orphan_replacements.items():
            success = unpin_from_quick_access(orphan_path)
            if success:
                actions.append({"type": "remove_pin", "target": orphan_path})
            else:
                failed_removals.append(orphan_path)
            
            junction = create_junction(replacement)
            if junction:
                actions.append({
                    "type": "create_junction",
                    "target": replacement,
                    "junction": junction,
                })
        
        # Process regular imports
        for path, v_i, v_r in rows:
            if v_i.get():
                junction = create_junction(path)
                if junction:
                    actions.append(
                        {
                            "type": "create_junction",
                            "target": path,
                            "junction": junction,
                        }
                    )
            if v_r.get() and os.path.normcase(path) != os.path.normcase(HUB):
                success = unpin_from_quick_access(path)
                if success:
                    actions.append({"type": "remove_pin", "target": path})
                else:
                    failed_removals.append(path)
        
        if actions:
            log_operation(actions)
        
        msg = "Import completed."
        if orphan_replacements:
            msg += f"\n{len(orphan_replacements)} orphaned item(s) replaced."
        if failed_removals:
            msg += f"\n\n⚠ Failed to remove {len(failed_removals)} item(s) from Quick Access:"
            for fr in failed_removals[:5]:  # Show max 5
                msg += f"\n  • {fr}"
            if len(failed_removals) > 5:
                msg += f"\n  ... and {len(failed_removals) - 5} more"
        
        messagebox.showinfo(APP_NAME, msg)
        if close_var.get():
            root.destroy()

    def do_undo():
        undo_last_operation()
        messagebox.showinfo(APP_NAME, "Last import operation undone (if any).")

    def exit_app():
        root.destroy()

    def show_help():
        help_win = tk.Toplevel(root)
        help_win.title("QuikDirs Help")
        help_win.resizable(True, True)
        help_win.geometry("900x700")
        
        # Create frame for text and button
        help_frame = ttk.Frame(help_win)
        help_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        txt = tk.Text(help_frame, wrap="word", width=100, height=35)
        txt.insert("1.0", README_TEXT)
        txt.config(state="disabled")
        txt.pack(fill="both", expand=True)
        
        # Exit button at bottom
        btn_frame = ttk.Frame(help_frame)
        btn_frame.pack(fill="x", pady=(10, 0))
        exit_btn = tk.Button(btn_frame, text="Close (Esc)", bg="red", fg="white", width=15, command=help_win.destroy)
        exit_btn.pack(side="right")
        
        # Bind Esc to close
        help_win.bind("<Escape>", lambda e: help_win.destroy())
        
        # Focus the help window
        help_win.focus_force()
        txt.focus_set()

    # Buttons row at bottom
    b_import = tk.Button(
        btns, text="Import", bg="green", fg="white", width=12, command=do_import
    )
    b_import.pack(side="left", padx=5)
    b_undo = tk.Button(
        btns, text="Undo Last", bg="yellow", fg="black", width=12, command=do_undo
    )
    b_undo.pack(side="left", padx=5)
    b_help = tk.Button(btns, text="Help", width=12, command=show_help)
    b_help.pack(side="left", padx=5)
    b_exit = tk.Button(
        btns, text="Exit", bg="red", fg="white", width=12, command=exit_app
    )
    b_exit.pack(side="left", padx=5)

    # ----------- Auto width based on folder text -----------
    default_font = tkfont.nametofont("TkDefaultFont")
    max_text_width = max([default_font.measure(p) for p, _, _ in rows] + [600])
    extra_width = 150  # for checkboxes + padding
    root.geometry(f"{max_text_width + extra_width}x600")  # 600px initial height

    # Bind Shift+? to open help
    root.bind("<question>", lambda e: show_help() if e.state & 0x1 else None)  # Check for Shift modifier
    root.bind("<?>", lambda e: show_help())  # Alternative binding

    root.mainloop()


# ---------------- Main ----------------
def main():
    os.makedirs(HUB, exist_ok=True)
    exclude_from_google_drive()

    # Dependency check
    if not ensure_pywin32():
        if not is_admin():
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, __file__, None, 1
            )
            sys.exit(0)
        install_pywin32()

    # Self-install - always update if this is not the installed version
    current_file = os.path.abspath(__file__)
    installed_file = os.path.abspath(INSTALLED_SCRIPT)
    
    if current_file != installed_file:
        print(f"Installing from: {current_file}")
        print(f"Installing to: {installed_file}")
        shutil.copy2(current_file, installed_file)
        print(f"Installation complete. Restarting from installed location...")
        add_sendto_shortcut(INSTALLED_SCRIPT)
        pin_to_quick_access(HUB)
        subprocess.Popen([sys.executable, INSTALLED_SCRIPT] + sys.argv[1:])
        sys.exit(0)
    
    print(f"Running QuikDirs v{VERSION} from: {current_file}")

    # Send To invocation
    if len(sys.argv) > 1:
        for arg in sys.argv[1:]:
            if os.path.isdir(arg):
                create_junction(arg)
        return

    # Interactive run
    add_sendto_shortcut(INSTALLED_SCRIPT)
    pin_to_quick_access(HUB)
    
    # Rebuild CSV from existing junctions on startup to ensure it's correct
    rebuild_csv_from_filesystem()
    
    run_import_gui()


def rebuild_csv_from_filesystem():
    """Scan all junctions in HUB and rebuild CSV with correct targets"""
    junctions = {}
    if os.path.exists(HUB):
        for item in os.listdir(HUB):
            item_path = os.path.join(HUB, item)
            if os.path.isdir(item_path):
                target = get_junction_target(item_path)
                if target and os.path.exists(target):
                    junctions[item] = (target, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    if junctions:
        write_junctions_csv(junctions)
        print(f"Rebuilt CSV with {len(junctions)} junctions")


if __name__ == "__main__":
    main()
