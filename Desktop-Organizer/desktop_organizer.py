import os
import shutil
import json
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path

# Try to import win32com for resolving .lnk files
try:
    import win32com.client
    def get_lnk_target(lnk_path):
        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(str(lnk_path))
            return shortcut.Targetpath
        except Exception:
            return None
except ImportError:
    def get_lnk_target(lnk_path):
        return None

def get_url_content(url_path):
    try:
        with open(url_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line.upper().startswith('URL'):
                    # Split on first '='
                    parts = line.split('=', 1)
                    if len(parts) == 2:
                        return parts[1].strip()
    except Exception:
        return None
    return None

# ------------------- Configuration -------------------
CONFIG_FILE = "organizer_config.json"
DEFAULT_CATEGORIES = {
    "Images": [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".webp"],
    "Documents": [".pdf", ".doc", ".docx", ".txt", ".xls", ".xlsx", ".ppt", ".pptx", ".csv", ".rtf"],
    "Videos": [".mp4", ".avi", ".mkv", ".mov", ".wmv", ".flv", ".webm", ".mpeg"],
    "Games": [".iso", ".csd", ".bin", ".cue", ".mdf", ".mds", ".nrg", ".img", ".ddf"],
    "Applications": [".exe", ".lnk", ".bat", ".cmd", ".msi", ".wsf", ".vbs", ".js", ".jar", ".py"],
}

KEYWORD_CATEGORIES = {
    "steam": "Steam Games",
    "riot": "Riot Games",
}

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
                # Ensure default categories exist
                for cat, exts in DEFAULT_CATEGORIES.items():
                    if cat not in config:
                        config[cat] = exts
                return config
        except Exception:
            pass
    return DEFAULT_CATEGORIES.copy()

def save_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4, ensure_ascii=False)

# ------------------- Core Logic -------------------
def get_desktop_path():
    return Path.home() / "Desktop"

def organize_desktop(categories):
    desktop = get_desktop_path()
    if not desktop.exists():
        messagebox.showerror("Error", "Desktop folder not found.")
        return

    moved_count = 0
    skipped_count = 0

    for item in desktop.iterdir():
        if item.is_file():
            # Skip the script itself and config file
            if item.name in ("desktop_organizer.py", CONFIG_FILE):
                skipped_count += 1
                continue

            placed = False
            lower_name = item.name.lower()
            ext = item.suffix.lower()

            # First, check keyword categories (including .lnk target and .url content)
            for keyword, category in KEYWORD_CATEGORIES.items():
                if keyword in lower_name:
                    target_dir = desktop / category
                    target_dir.mkdir(exist_ok=True)
                    try:
                        shutil.move(str(item), str(target_dir / item.name))
                        moved_count += 1
                    except Exception as e:
                        print(f"Move error {item.name}: {e}")
                    placed = True
                    break
            if not placed and ext == ".lnk":
                # Check .lnk target for keywords
                target_path = get_lnk_target(item)
                if target_path:
                    target_lower = target_path.lower()
                    for keyword, category in KEYWORD_CATEGORIES.items():
                        if keyword in target_lower:
                            target_dir = desktop / category
                            target_dir.mkdir(exist_ok=True)
                            try:
                                shutil.move(str(item), str(target_dir / item.name))
                                moved_count += 1
                            except Exception as e:
                                print(f"Move error {item.name}: {e}")
                            placed = True
                            break
            if not placed and ext == ".url":
                # Check .url content for keywords
                url_content = get_url_content(item)
                if url_content:
                    url_lower = url_content.lower()
                    for keyword, category in KEYWORD_CATEGORIES.items():
                        if keyword in url_lower:
                            target_dir = desktop / category
                            target_dir.mkdir(exist_ok=True)
                            try:
                                shutil.move(str(item), str(target_dir / item.name))
                                moved_count += 1
                            except Exception as e:
                                print(f"Move error {item.name}: {e}")
                            placed = True
                            break

            # If not placed by keyword, check extension categories
            if not placed:
                for category, extensions in categories.items():
                    if ext in extensions:
                        target_dir = desktop / category
                        target_dir.mkdir(exist_ok=True)
                        try:
                            shutil.move(str(item), str(target_dir / item.name))
                            moved_count += 1
                        except Exception as e:
                            print(f"Move error {item.name}: {e}")
                        placed = True
                        break
            if not placed:
                # Optional: move to "Other" or leave as is
                pass

    messagebox.showinfo(
        "Completed",
        f"{moved_count} files moved.\n{skipped_count} files skipped (own script/config).",
    )

# ------------------- GUI -------------------
class OrganizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Desktop Organizer")
        self.root.geometry("500x400")
        self.root.resizable(False, False)

        # Dark neon theme colors
        self.bg_color = "#0d0d0d"
        self.fg_color = "#00ffef"  # neon cyan
        self.accent_color = "#ff00ff"  # neon magenta
        self.button_bg = "#1a1a1a"
        self.button_fg = "#00ffef"
        self.entry_bg = "#1a1a1a"
        self.entry_fg = "#00ffef"

        self.root.configure(bg=self.bg_color)

        self.categories = load_config()

        self.create_widgets()

    def create_widgets(self):
        # Title
        title_label = tk.Label(
            self.root,
            text="Desktop Organizer",
            font=("Helvetica", 18, "bold"),
            bg=self.bg_color,
            fg=self.fg_color,
        )
        title_label.pack(pady=20)

        # Info
        info_label = tk.Label(
            self.root,
            text="Organizes your files by type into folders.\nYou can add custom categories.",
            font=("Helvetica", 10),
            bg=self.bg_color,
            fg=self.fg_color,
            justify="center",
        )
        info_label.pack(pady=10)

        # Organize Button
        organize_btn = tk.Button(
            self.root,
            text="Organize Desktop",
            command=self.start_organize,
            font=("Helvetica", 12, "bold"),
            bg=self.button_bg,
            fg=self.button_fg,
            activebackground=self.accent_color,
            activeforeground="#ffffff",
            relief="raised",
            bd=3,
            padx=20,
            pady=10,
        )
        organize_btn.pack(pady=15)

        # Add Category Button
        add_cat_btn = tk.Button(
            self.root,
            text="Add Custom Category",
            command=self.add_category_dialog,
            font=("Helvetica", 12),
            bg=self.button_bg,
            fg=self.button_fg,
            activebackground=self.accent_color,
            activeforeground="#ffffff",
            relief="raised",
            bd=3,
            padx=20,
            pady=10,
        )
        add_cat_btn.pack(pady=10)

        # Status
        self.status_var = tk.StringVar(value=f"Ready. {len(self.categories)} categories loaded.")
        status_label = tk.Label(
            self.root,
            textvariable=self.status_var,
            font=("Helvetica", 9),
            bg=self.bg_color,
            fg=self.accent_color,
        )
        status_label.pack(pady=20, side="bottom")

    def start_organize(self):
        if not messagebox.askyesno(
            "Confirmation",
            "This will move all files on your desktop to folders based on their categories.\nAre you sure you want to continue?"
        ):
            return
        self.status_var.set("Organizing...")
        self.root.update()
        organize_desktop(self.categories)
        self.status_var.set("Organization completed.")

    def add_category_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Custom Category")
        dialog.geometry("350x200")
        dialog.resizable(False, False)
        dialog.configure(bg=self.bg_color)
        dialog.transient(self.root)
        dialog.grab_set()

        tk.Label(
            dialog,
            text="Category Name:",
            bg=self.bg_color,
            fg=self.fg_color,
            font=("Helvetica", 10),
        ).pack(pady=(20,5))
        cat_entry = tk.Entry(
            dialog,
            bg=self.entry_bg,
            fg=self.entry_fg,
            insertbackground=self.fg_color,
            font=("Helvetica", 10),
            width=30,
        )
        cat_entry.pack(pady=5)
        cat_entry.focus()

        tk.Label(
            dialog,
            text="Extensions (comma-separated, e.g. .xyz,.abc):",
            bg=self.bg_color,
            fg=self.fg_color,
            font=("Helvetica", 10),
        ).pack(pady=(15,5))
        ext_entry = tk.Entry(
            dialog,
            bg=self.entry_bg,
            fg=self.entry_fg,
            insertbackground=self.fg_color,
            font=("Helvetica", 10),
            width=30,
        )
        ext_entry.pack(pady=5)

        def save_category():
            cat_name = cat_entry.get().strip()
            ext_text = ext_entry.get().strip()
            if not cat_name or not ext_text:
                messagebox.showwarning("Warning", "Please fill in all fields.", parent=dialog)
                return
            # Process extensions
            ext_list = []
            for part in ext_text.split(","):
                part = part.strip()
                if not part.startswith("."):
                    part = "." + part
                ext_list.append(part.lower())
            # Add to categories
            self.categories[cat_name] = ext_list
            save_config(self.categories)
            messagebox.showinfo("Successful", f"Category '{cat_name}' added.", parent=dialog)
            self.status_var.set(f"Category added. Total {len(self.categories)} categories.")
            dialog.destroy()

        save_btn = tk.Button(
            dialog,
            text="Add",
            command=save_category,
            bg=self.button_bg,
            fg=self.button_fg,
            activebackground=self.accent_color,
            activeforeground="#ffffff",
            font=("Helvetica", 10, "bold"),
            width=10,
        )
        save_btn.pack(pady=20)

if __name__ == "__main__":
    root = tk.Tk()
    app = OrganizerApp(root)
    root.mainloop()