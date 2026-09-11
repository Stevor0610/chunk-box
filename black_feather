import sys
import os
import threading

if __name__ == "__main__":
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
            
import json
import time
import string
import csv
import tkinter as tk
import customtkinter as ctk
from tksheet import Sheet
from tkinter import filedialog, messagebox


try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False

# Global Theme Configuration
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

PREFS_FILE = os.path.join(os.path.expanduser("~"), ".black_feather_prefs.json")


class ModernEditableEditor(ctk.CTk, TkinterDnD.DnDWrapper if HAS_DND else object):

    # ==================== 🚀 Initialization & UI Construction ====================

    def __init__(self):
    
        self._selection_stats_timer = None
        self._last_search_keyword = None
        self._loading_active = False
        self._converting_active = False
        self._saving_active = False
       
        if HAS_DND:
            ctk.CTk.__init__(self)
            TkinterDnD.DnDWrapper.__init__(self)
            self.TkdndVersion = TkinterDnD._require(self)
        else:
            super().__init__()

        self.title("Black Feather")
        self.set_window_icon()     
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        self.geometry(f"{screen_w}x{screen_h}+0+0")
        
        # ====== Tab Ledger Data Registry ======
        self.tabs_data = {}
        self.tab_counter = 0

        self.headers = list(string.ascii_uppercase)
        self.current_font_size = 11
        self.history = []

        # ==================== 🛠️ 1. Sidebar Panel ====================
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0, fg_color=("#f1f3f5", "#111115"))
        self.sidebar.pack(side="left", fill="y")
        
        # Branding Header
        self.logo_label = ctk.CTkLabel(
            self.sidebar, 
            text="🪶 Black Feather", 
            font=ctk.CTkFont(family="Segoe UI", size=22, weight="bold")
        )
        self.logo_label.pack(padx=15, pady=(30, 8))
        
        self.sub_logo = ctk.CTkLabel(
            self.sidebar,
            text="PARQUET & FEATHER EDITOR",
            font=ctk.CTkFont(family="Segoe UI", size=9, weight="bold"),
            text_color=("#6c757d", "#4b5563")
        )
        self.sub_logo.pack(padx=15, pady=(0, 20))

        self.add_separator()

        # Navigation Action Button Frame
        self.btn_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.btn_frame.pack(padx=14, pady=12, fill="x")

        self.btn_new = ctk.CTkButton(
            self.btn_frame, 
            text="📝  New Blank Tab", 
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color=("#e8d5f5", "#2d1f3d"),
            text_color=("#7b2d8e", "#d8b4fe"),
            hover_color=("#d4b5e9", "#3d2b54"),
            height=42,
            corner_radius=8,
            border_width=1,
            border_color=("#d4b5e9", "#6b21a8"),
            command=self.new_blank_tab
        )
        self.btn_new.pack(fill="x", pady=6)

        # File Ingestion Buttons Layout
        self.open_row = ctk.CTkFrame(self.btn_frame, fg_color="transparent")
        self.open_row.pack(fill="x", pady=6)

        self.btn_open = ctk.CTkButton(
            self.open_row, 
            text="📂  Open File", 
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color=("#d0ebff", "#1f2937"),
            text_color=("#1864ab", "#f3f4f6"),
            hover_color=("#a5d8ff", "#374151"),
            height=42,
            corner_radius=8,
            border_width=1,
            border_color=("#a5d8ff", "#4b5563"),
            command=self.load_file
        )
        self.btn_open.pack(side="left", fill="x", expand=True)

        self.btn_history = ctk.CTkButton(
            self.open_row,
            text="⏱️",
            width=42,
            height=42,
            font=ctk.CTkFont(size=15),
            fg_color=("#d0ebff", "#1f2937"),
            text_color=("#1864ab", "#f3f4f6"),
            hover_color=("#a5d8ff", "#374151"),
            corner_radius=8,
            border_width=1,
            border_color=("#a5d8ff", "#4b5563"),
            command=self.show_history_menu
        )
        self.btn_history.pack(side="right", padx=(4, 0))

        # Recent History Dropdown Menu
        self.history_menu = None
                
        # Layout Autofit Execution Button
        self.btn_autofit = ctk.CTkButton(
            self.btn_frame, 
            text="📐  Autofit Column Widths", 
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color=("#d3f9d8", "#1e3a1e"),
            text_color=("#2b8a3e", "#a7f3d0"),
            hover_color=("#b2f2bb", "#2d5a2d"),
            height=42,
            corner_radius=8,
            border_width=1,
            border_color=("#b2f2bb", "#064e3b"),
            command=self.autofit_columns,
            state="normal"
        )
        self.btn_autofit.pack(fill="x", pady=6)
        
        # Save & Compile Actions
        self.btn_save = ctk.CTkButton(
            self.btn_frame, 
            text="💾  Save / Export File", 
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color=("#1c7ed6", "#1e3a8a"),
            text_color=("#ffffff", "#dbeafe"),
            hover_color=("#1971c2", "#2563eb"),
            height=42,
            corner_radius=8,
            command=self.save_file_as
        )
        self.btn_save.pack(fill="x", pady=6)

        # Inspect Active Grid Schema structures
        self.btn_schema = ctk.CTkButton(
            self.btn_frame, 
            text="📊  Inspect Column Types", 
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color=("#fff3bf", "#78350f"),
            text_color=("#d9480f", "#fef3c7"),
            hover_color=("#ffe066", "#92400e"),
            height=42,
            corner_radius=8,
            border_width=1,
            border_color=("#ffe066", "#d9480f"),
            command=self.show_schema_viewer
        )
        self.btn_schema.pack(fill="x", pady=6)

        self.add_separator()

        # ==================== 🔄 Quick Convert Panel ====================
        self.convert_card = ctk.CTkFrame(
            self.sidebar,
            fg_color=("#e9ecef", "#1e1e24"),
            corner_radius=8,
            border_width=2,
            border_color=("#dee2e6", "#2d2d34")
        )
        self.convert_card.pack(padx=14, pady=12, fill="x")

        self.lbl_convert_title = ctk.CTkLabel(
            self.convert_card,
            text="🔄 Quick Converter",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            text_color=("#495057", "#9ca3af")
        )
        self.lbl_convert_title.pack(padx=10, pady=(8, 4), anchor="w")

        self.drop_zone = ctk.CTkLabel(
            self.convert_card,
            text="📥\nDrag and drop files\nhere to convert",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=("#868e96", "#6b7280"),
            fg_color=("#f8f9fa", "#111115"),
            corner_radius=6,
            height=120
        )
        self.drop_zone.pack(padx=10, pady=(0, 6), fill="x")

        self.convert_fmt_row = ctk.CTkFrame(self.convert_card, fg_color="transparent")
        self.convert_fmt_row.pack(padx=10, pady=(0, 8), fill="x")

        self.lbl_convert_to = ctk.CTkLabel(
            self.convert_fmt_row,
            text="Format:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=("#495057", "#9ca3af")
        )
        self.lbl_convert_to.pack(side="left")

        self.convert_fmt = ctk.CTkOptionMenu(
            self.convert_fmt_row,
            values=["Parquet", "Feather", "CSV", "Excel"],
            fg_color=("#ffffff", "#111115"),
            button_color=("#ced4da", "#2d2d34"),
            button_hover_color=("#adb5bd", "#3f3f46"),
            text_color=("#212529", "#e5e7eb"),
            height=26,
            width=100,
            corner_radius=6,
            font=ctk.CTkFont(size=12)
        )
        self.convert_fmt.set("Parquet")
        self.convert_fmt.pack(side="right")

        # ==================== 🛠️ 2. Sidebar Control panel ====================
        self.bottom_ctrl = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.bottom_ctrl.pack(side="bottom", padx=14, pady=20, fill="x")
        
        # 🌙 Palette Theme Switch
        self.switch_theme = ctk.CTkSwitch(
            self.bottom_ctrl, 
            text="🌙 Dark Mode",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=("#495057", "#9ca3af"),
            command=self.toggle_appearance_mode
        )
        self.switch_theme.select()
        self.switch_theme.pack(fill="x", pady=(0, 15))

        # ==================== 🛠️ 3. Main Dashboard Workspace ====================
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(side="right", fill="both", expand=True, padx=20, pady=20)
        
        # 💡 Status & Progress Indicator Strip
        self.status_card = ctk.CTkFrame(
            self.main_frame, 
            height=50, 
            corner_radius=10, 
            fg_color=("#e9ecef", "#18181c"),
            border_width=1,
            border_color=("#dee2e6", "#2d2d34")
        )
        self.status_card.pack(fill="x", pady=(0, 15))
        self.status_card.pack_propagate(False)
        
        self.led_indicator = ctk.CTkLabel(
            self.status_card,
            text="●",
            font=ctk.CTkFont(size=18),
            text_color=("#007bff", "#60a5fa")
        )
        self.led_indicator.pack(side="left", padx=(15, 5))
        
        self.lbl_status = ctk.CTkLabel(
            self.status_card, 
            text="Edit below to build a new file, or drop/open a Parquet/Feather database to inspect...", 
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=("#495057", "#9ca3af")
        )
        self.lbl_status.pack(side="left", pady=10)

        # 🔍 Table Zoom / Font Size controls
        self.lbl_font_val = ctk.CTkLabel(
            self.status_card,
            text=f"{self.current_font_size}px",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=("#868e96", "#6b7280")
        )
        self.lbl_font_val.pack(side="right", padx=(0, 12))

        self.font_slider = ctk.CTkSlider(
            self.status_card,
            from_=8,
            to=24,
            number_of_steps=16,
            height=14,
            width=100,
            command=self.change_table_font_size
        )
        self.font_slider.set(self.current_font_size)
        self.font_slider.pack(side="right", padx=2)

        self.lbl_font_icon = ctk.CTkLabel(
            self.status_card,
            text="🔍",
            font=ctk.CTkFont(size=12),
            text_color=("#495057", "#9ca3af")
        )
        self.lbl_font_icon.pack(side="right", padx=(8, 0))

        # ==================== 🔍 Inline Quick Search Panel ====================
        self.search_bar = ctk.CTkFrame(
            self.main_frame,
            height=45,
            corner_radius=10,
            fg_color=("#dbe4ff", "#1a1a2e"),
            border_width=1,
            border_color=("#748ffc", "#4c6ef5")
        )
        self.search_bar_visible = False

        self.search_entry = ctk.CTkEntry(
            self.search_bar,
            placeholder_text="Search content...",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            height=32,
            width=280,
            corner_radius=6
        )
        self.search_entry.pack(side="left", padx=(12, 6), pady=6)
        self.search_entry.bind("<Return>", lambda e: self.find_next())
        self.search_entry.bind("<Shift-Return>", lambda e: self.find_prev())
        self.search_entry.bind("<Escape>", lambda e: self.hide_search_bar())

        self.btn_find_prev = ctk.CTkButton(
            self.search_bar, text="▲", width=36, height=32,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=("#748ffc", "#364fc7"),
            hover_color=("#5c7cfa", "#4c6ef5"),
            corner_radius=6,
            command=self.find_prev
        )
        self.btn_find_prev.pack(side="left", padx=2, pady=6)

        self.btn_find_next = ctk.CTkButton(
            self.search_bar, text="▼", width=36, height=32,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=("#748ffc", "#364fc7"),
            hover_color=("#5c7cfa", "#4c6ef5"),
            corner_radius=6,
            command=self.find_next
        )
        self.btn_find_next.pack(side="left", padx=2, pady=6)

        self.lbl_search_result = ctk.CTkLabel(
            self.search_bar,
            text="",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=("#495057", "#9ca3af")
        )
        self.lbl_search_result.pack(side="left", padx=10, pady=6)

        self.btn_search_close = ctk.CTkButton(
            self.search_bar, text="✕", width=36, height=32,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="transparent",
            hover_color=("#ff6b6b", "#c92a2a"),
            text_color=("#495057", "#9ca3af"),
            corner_radius=6,
            command=self.hide_search_bar
        )
        self.btn_search_close.pack(side="right", padx=(0, 8), pady=6)

        self.search_matches = []
        self.search_current_idx = -1
        
        # 💡 Sheet Tabview Frame Wrapper
        self.table_container = ctk.CTkFrame(
            self.main_frame, 
            corner_radius=10,
            border_width=1,
            border_color=("#dee2e6", "#2d2d34")
        )
        self.table_container.pack(fill="both", expand=True)

        # ====== 💡 Dynamic Multi-Tab Module Layout ======
        self.tab_view = ctk.CTkTabview(self.table_container, command=self.on_tab_changed)
        self.tab_view.pack(expand=True, fill="both", padx=2, pady=2)

        # Context Menu structure for sheet selection buttons
        self.tab_context_menu = None

        try:
            self.tab_view._segmented_button.bind("<Button-3>", self.show_tab_context_menu)
        except Exception:
            try:
                self.tab_view._canvas.bind("<Button-3>", self.show_tab_context_menu)
            except Exception:
                pass

        # Load initial startup sheet
        self.create_new_tab("untitled", [["" for _ in range(26)] for _ in range(256)], self.headers, None)

        # Global Hotkey Integrations
        self.bind_all("<Control-y>", self.trigger_redo)
        self.bind_all("<Control-Y>", self.trigger_redo)
        self.bind_all("<Control-s>", self.trigger_save)
        self.bind_all("<Control-S>", self.trigger_save)
        self.bind_all("<Control-Shift-s>", self.trigger_save_as)
        self.bind_all("<Control-Shift-S>", self.trigger_save_as)
        self.bind_all("<Control-f>", self.toggle_search_bar)
        self.bind_all("<Control-F>", self.toggle_search_bar)
        self.bind_all("<F1>", self.show_shortcuts)

        if HAS_DND:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self.handle_file_drop)
            self.drop_zone.drop_target_register(DND_FILES)
            self.drop_zone.dnd_bind("<<Drop>>", self.on_convert_drop)
            self.drop_zone.dnd_bind("<<DragEnter>>", self.on_convert_drag_enter)
            self.drop_zone.dnd_bind("<<DragLeave>>", self.on_convert_drag_leave)

        self.after(50, self.load_preferences)
        self.after(150, self.check_startup_file) 
        
        self.protocol("WM_DELETE_WINDOW", self.on_app_close)
        
        self.after(50, lambda: self.state('zoomed'))

    # ==================== 🎨 Window Chrome & Sidebar Helpers ====================

    def add_separator(self):
        sep = ctk.CTkFrame(self.sidebar, height=1, fg_color=("#dee2e6", "#2d2d34"))
        sep.pack(fill="x", padx=15, pady=5)

    def set_window_icon(self):
        icon_filename = r"D:\lhwangq_pyfile\app\BlackFeather\feather.ico"
        base_path = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(base_path, icon_filename)
        if os.path.exists(icon_path):
            try:
                self.after(200, lambda: self.iconbitmap(icon_path))
            except Exception as e:
                print(f"Failed to load application ico file: {e}")

    # ==================== 🌙 Theme & Appearance ====================

    def toggle_appearance_mode(self):
        if ctk.get_appearance_mode() == "Dark":
            ctk.set_appearance_mode("light")
            self.switch_theme.configure(text="☀️ Light Mode")
            for tid, info in self.tabs_data.items():
                info["sheet"].change_theme(theme="light green")
                info["sheet"].redraw()
        else:
            ctk.set_appearance_mode("dark")
            self.switch_theme.configure(text="🌙 Dark Mode")
            for tid, info in self.tabs_data.items():
                info["sheet"].change_theme(theme="dark blue")
                info["sheet"].redraw()

    # ==================== 💾 User Preferences & History ====================

    def load_preferences(self):
        if os.path.exists(PREFS_FILE):
            try:
                with open(PREFS_FILE, "r", encoding="utf-8") as f:
                    prefs = json.load(f)
                    self.history = prefs.get("history", [])
                    self.history = [p for p in self.history if os.path.exists(p)]
                    self.update_history_menu()
            except Exception:
                self.history = []

    def save_preference(self, new_path):
        if not new_path:
            return
            
        new_path = os.path.normpath(new_path)
        
        if new_path in self.history:
            self.history.remove(new_path)
        self.history.insert(0, new_path)
        self.history = self.history[:5]
        try:
            with open(PREFS_FILE, "w", encoding="utf-8") as f:
                json.dump({"history": self.history}, f, ensure_ascii=False, indent=4)
        except Exception:
            pass
        self.update_history_menu()

    def update_history_menu(self):
        self.history = [p for p in self.history if os.path.exists(p)]

    def show_history_menu(self):
        """Displays historical documents dropdown relative to history button position"""
        if self.history_menu is None:
            self.history_menu = tk.Menu(
                self, tearoff=0,
                bg="#111115", fg="white",
                activebackground="#1f538d",
                font=("Segoe UI", 10),
                borderwidth=0
            )

        self.history_menu.delete(0, "end")
        
        if not self.history:
            self.history_menu.add_command(label="(No history recorded)", state="disabled")
        else:
            for path in self.history:
                name = os.path.basename(path)
                if len(name) > 45:
                    name = name[:42] + "..."
                self.history_menu.add_command(
                    label=f"📄 {name}",
                    command=lambda p=path: self.load_file(target_file=p)
                )
        
        x = self.btn_history.winfo_rootx()
        y = self.btn_history.winfo_rooty() + self.btn_history.winfo_height()
        self.history_menu.post(x, y)

    # ==================== 🔍 Font Size & Zoom Control ====================

    def change_table_font_size(self, value):
        self.current_font_size = int(float(value))
        self.lbl_font_val.configure(text=f"{self.current_font_size}px")
        
        for tid, info in self.tabs_data.items():
            sheet = info.get("sheet")
            if sheet:
                try:
                    sheet.set_options(
                        table_font=("Segoe UI", self.current_font_size, "normal"),
                        header_font=("Segoe UI", self.current_font_size, "bold"),
                        index_font=("Segoe UI", self.current_font_size, "normal")
                    )
                    sheet.refresh()
                except Exception as e:
                    print(f"Failed to synchronize table size scaling (Tab ID: {tid}): {e}")

    def on_sheet_zoom(self, event=None):
        tid = self.get_active_tab_id()
        if not tid or tid not in self.tabs_data:
            return
            
        sheet = self.tabs_data[tid]["sheet"]
        try:
            font_opt = sheet.options.table_font
            actual_size = font_opt[1]
            
            self.current_font_size = int(actual_size)
            
            self.lbl_font_val.configure(text=f"{self.current_font_size}px")
            
            slider_min = self.font_slider.cget("from")
            slider_max = self.font_slider.cget("to")
            safe_size = max(slider_min, min(self.current_font_size, slider_max))
            
            self.font_slider.configure(command=None)
            self.font_slider.set(safe_size)
            self.font_slider.configure(command=self.change_table_font_size)
            
        except Exception as e:
            print(f"Failed to sync slider on zoom: {e}")

    def _on_ctrl_mousewheel(self, event):
        if event.delta > 0:
            new_size = min(24, self.current_font_size + 1)
        else:
            new_size = max(8, self.current_font_size - 1)

        if new_size != self.current_font_size:
            self.current_font_size = new_size

            self.font_slider.configure(command=None)
            self.font_slider.set(new_size)
            self.font_slider.configure(command=self.change_table_font_size)

            self.lbl_font_val.configure(text=f"{new_size}px")

            for tid, info in self.tabs_data.items():
                s = info.get("sheet")
                if s:
                    try:
                        s.set_options(
                            table_font=("Segoe UI", new_size, "normal"),
                            header_font=("Segoe UI", new_size, "bold"),
                            index_font=("Segoe UI", new_size, "normal")
                        )
                        s.refresh()
                    except Exception:
                        pass

        return "break"

    # ==================== ⌨️ Hotkeys & Shortcut Help ====================

    def show_shortcuts(self, event=None):
        shortcuts = [
            ("Ctrl + S", "Quick Save File"),
            ("Ctrl + Shift + S", "Export / Save As File"),
            ("Ctrl + F", "Toggle Inline Search Panel"),
            ("Ctrl + Z", "Undo Action"),
            ("Ctrl + Y", "Redo Action"),
            ("Ctrl + C / X / V", "Copy / Cut / Paste Elements"),
            ("Ctrl + Tab", "Switch to Next Tab View"),
            ("Ctrl + Scroll", "Zoom Table Font Size"),
            ("Double-Click Header", "Inline Rename Selected Column"),
            ("Double-Click Col Divider", "Autofit Column Layout Width"),
            ("F1", "Open Shortcuts Quick Guide"),
        ]
        
        win = ctk.CTkToplevel(self)
        win.title("⌨️ Keyboard Shortcuts Reference")
        win.geometry("400x420")
        win.attributes("-topmost", True)
        win.resizable(False, False)

        header = ctk.CTkFrame(win, height=55, corner_radius=0, fg_color=("#e9ecef", "#18181c"))
        header.pack(fill="x")
        header.pack_propagate(False)
        
        ctk.CTkLabel(
            header, text="⌨️ Shortcuts Reference",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(padx=20, pady=12, anchor="w")

        scroll = ctk.CTkScrollableFrame(win, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=15, pady=15)
        
        for key, desc in shortcuts:
            row = ctk.CTkFrame(scroll, fg_color=("#f1f3f5", "#1e1e24"), corner_radius=6,
                               border_width=1, border_color=("#dee2e6", "#2d2d34"))
            row.pack(fill="x", pady=3)
            
            ctk.CTkLabel(
                row, text=key,
                font=ctk.CTkFont(family="Consolas", size=12, weight="bold"),
                text_color=("#1c7ed6", "#60a5fa"),
                width=180
            ).pack(side="left", padx=12, pady=8)
            
            ctk.CTkLabel(
                row, text=desc,
                font=ctk.CTkFont(size=12),
                text_color=("#495057", "#9ca3af")
            ).pack(side="right", padx=12, pady=8)

        ctk.CTkButton(
            win, text="Close", width=100, height=34,
            fg_color=("#1c7ed6", "#1e3a8a"),
            hover_color=("#1971c2", "#2563eb"),
            command=win.destroy
        ).pack(pady=(0, 15))

    def trigger_redo(self, event):
        tid = self.get_active_tab_id()
        if tid and tid in self.tabs_data:
            self.tabs_data[tid]["sheet"].redo(event)
            
    def trigger_save(self, event):
        self.save_file_direct()
        return "break"

    def trigger_save_as(self, event):
        self.save_file_as()
        return "break"

    # ==================== 📂 Tab & Multi-Document Management ====================

    def get_active_tab_id(self):
        try:
            active_tab_name = self.tab_view.get()
            for tid, info in self.tabs_data.items():
                if info["tab_name"] == active_tab_name:
                    return tid
        except Exception:
            return None
        return None

    def on_tab_changed(self):
        tid = self.get_active_tab_id()
        if tid and tid in self.tabs_data:
            info = self.tabs_data[tid]
            sheet = info["sheet"]

            try:
                sheet.set_options(
                    table_font=("Segoe UI", self.current_font_size, "normal"),
                    header_font=("Segoe UI", self.current_font_size, "bold"),
                    index_font=("Segoe UI", self.current_font_size, "normal")
                )
                sheet.refresh()
            except Exception as e:
                print(f"Failed to synchronize font size scales on tab transition: {e}")

            rows_count = len(sheet.get_sheet_data())
            
            if info["file_path"]:
                base_name = os.path.basename(info["file_path"])
                status_text = f"📄 Active Tab: {base_name} | Row Count: {rows_count}"
            else:
                status_text = f"📄 Active Tab: {info['tab_name']} | Row Count: {rows_count}"

            if info["modified"]:
                status_text += " | ⚠️ Document modified, saving pending..."

            self.led_indicator.configure(text_color=("#1b5e20", "#2ecc71") if not info["modified"] else ("#b45309", "#ffcc00"))
            self.lbl_status.configure(text=status_text, text_color=("#495057", "#9ca3af") if not info["modified"] else ("#b45309", "#ffcc00"))

    def mark_as_modified(self, tab_id):
        if tab_id in self.tabs_data and not self.tabs_data[tab_id]["modified"]:
            self.tabs_data[tab_id]["modified"] = True
            old_name = self.tabs_data[tab_id]["tab_name"]
            new_name = f"📄 {old_name.replace('📄 ', '')} *"
            self.tab_view.rename(old_name, new_name)
            self.tabs_data[tab_id]["tab_name"] = new_name
            self.on_tab_changed()

    def show_tab_context_menu(self, event):
        if self.tab_context_menu is None:
            self.tab_context_menu = tk.Menu(
                self, tearoff=0,
                bg="#111115", fg="white",
                activebackground="#1f538d",
                borderwidth=0
            )
            self.tab_context_menu.add_command(label="❌ Close Active Tab", command=self.close_active_tab)
            self.tab_context_menu.add_command(label="💾 Save Active Tab", command=self.save_file_direct)
            self.tab_context_menu.add_separator()
            self.tab_context_menu.add_command(label="🧹 Close All Other Tabs", command=self.close_other_tabs)

        self.tab_context_menu.post(event.x_root, event.y_root)
        
    def _on_tab_double_click(self, event):
        tid = self.get_active_tab_id()
        if not tid or tid not in self.tabs_data:
            return

        info = self.tabs_data[tid]
        old_tab_name = info["tab_name"]
        file_path = info.get("file_path", "")
            
        if file_path:
            current_name = os.path.splitext(os.path.basename(file_path))[0]
        else:
            current_name = old_tab_name.replace("📄 ", "").replace("● ", "").replace(" *", "")

        dialog = ctk.CTkInputDialog(
            text="Enter new file name:",
            title="Rename"
        )
        new_name = dialog.get_input()

        if not new_name or not new_name.strip():
            return
        new_name = new_name.strip()

        if file_path:
            directory = os.path.dirname(file_path)
            ext = os.path.splitext(file_path)[1]
            new_path = os.path.join(directory, new_name + ext)

            if os.path.exists(new_path) and os.path.normpath(new_path) != os.path.normpath(file_path):
                confirm = messagebox.askyesno(
                    "File Exists",
                    f"'{new_name + ext}' already exists. Overwrite?"
                )
                if not confirm:
                    return
            try:
                os.rename(file_path, new_path)
                info["file_path"] = new_path
            except Exception as e:
                messagebox.showerror("Rename Error", str(e))
                return

            prefix = "● " if info.get("modified") else "📄 "
            new_tab_name = prefix + new_name + ext
        else:
            prefix = "● " if info.get("modified") else "📄 "
            new_tab_name = prefix + new_name

        self.tab_view.rename(old_tab_name, new_tab_name)
        info["tab_name"] = new_tab_name

        try:
            tab_button = self.tab_view._segmented_button._buttons_dict[new_tab_name]
            tab_button.bind("<Button-3>", self.show_tab_context_menu)
            tab_button.bind("<Double-Button-1>", self._on_tab_double_click)
        except Exception:
            pass

        self.on_tab_changed()
        
    def create_new_tab(self, display_name, rows_data, headers, arrow_table, file_path=None):
        self.tab_counter += 1
        tab_id = f"tab_{self.tab_counter}"

        tab_title = f"📄 {display_name}"
        
        existing_names = [info["tab_name"] for info in self.tabs_data.values()]
        counter = 1
        while tab_title in existing_names:
            tab_title = f"📄 {display_name} ({counter})"
            counter += 1

        self.tab_view.add(tab_title)
        
        try:
            tab_button = self.tab_view._segmented_button._buttons_dict[tab_title]
            tab_button.bind("<Button-3>", self.show_tab_context_menu)
            tab_button.bind("<Double-Button-1>", self._on_tab_double_click)
        except Exception:
            pass

        tab_frame = self.tab_view.tab(tab_title)
        tab_frame.grid_rowconfigure(0, weight=1)
        tab_frame.grid_columnconfigure(0, weight=1)

        sheet_theme = "light green" if ctk.get_appearance_mode() == "Light" else "dark blue"
        
        # ✅ 建空表
        sheet = Sheet(
            tab_frame,
            data=[],
            headers=headers,
            theme=sheet_theme,
            header_font=("Segoe UI", 11, "bold"),
            font=("Segoe UI", self.current_font_size, "normal"),
            table_wrap="w"
        )
        sheet.enable_bindings(
            "single_select",
            "drag_select",
            "column_select",
            "row_select",
            "column_width_resize",
            "row_height_resize",
            "double_click_column_resize",
            "edit_cell",
            "copy",
            "cut",
            "paste",
            "undo",
            "arrowkeys",
            "rc_select",
            "right_click_popup_menu",
            "rc_insert_column",
            "rc_delete_column",
            "edit_header",
            "rc_insert_row",
            "rc_delete_row"
        )
        
        sheet.extra_bindings("cell_select", self.update_selection_stats)      
        sheet.extra_bindings("select_cells", self.update_selection_stats)
        sheet.extra_bindings("drag_select_cells", self.update_selection_stats)
        sheet.extra_bindings("deselect_cells", self.clear_selection_stats)
        sheet.extra_bindings([("double_click_column_resize", self.on_double_click_resize)])
        sheet.extra_bindings("sheet_modified", lambda event, tid=tab_id: self.mark_as_modified(tid))

        sheet.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)

        sheet.set_options(zoom=100)
        for child in sheet.winfo_children():
            child.bind("<Control-MouseWheel>", self._on_ctrl_mousewheel, add=False)
        sheet.bind("<Control-MouseWheel>", self._on_ctrl_mousewheel, add=False)

        if HAS_DND:
            sheet.drop_target_register(DND_FILES)
            sheet.dnd_bind("<<Drop>>", self.handle_file_drop)

        self.tabs_data[tab_id] = {
            "file_path": file_path,
            "table_data": arrow_table,
            "sheet": sheet,
            "modified": False,
            "tab_name": tab_title,
            "headers": headers
        }

        if file_path is not None:
            default_tid = None
            for tid, info in self.tabs_data.items():
                if info["tab_name"] == "📄 untitled" and info["file_path"] is None and tid != tab_id:
                    default_tid = tid
                    break
            if default_tid:
                try:
                    self.tabs_data[default_tid]["sheet"].destroy()
                except Exception:
                    pass
                self.tab_view.delete("📄 untitled")
                del self.tabs_data[default_tid]

        self.tab_view.set(tab_title)

        # ✅ 延遲填入資料，讓 UI 先顯示空表
        if rows_data:
            self.after(50, lambda: self._deferred_set_data(tab_id, rows_data))
        else:
            self.on_tab_changed()

        return tab_id

    def _deferred_set_data(self, tab_id, rows_data):
        if tab_id not in self.tabs_data:
            return

        sheet = self.tabs_data[tab_id]["sheet"]
        sheet.set_sheet_data(rows_data, redraw=False)
        sheet.redraw()
        self.on_tab_changed()

    def new_blank_tab(self):
        self.create_new_tab(
            "untitled",
            [["" for _ in range(26)] for _ in range(256)],
            list(string.ascii_uppercase),
            None
        )

    def close_active_tab(self):
        tid = self.get_active_tab_id()
        if not tid:
            return

        info = self.tabs_data[tid]
        tab_name = info["tab_name"]

        if info["modified"]:
            ans = messagebox.askyesnocancel(
                "Unsaved File Changes Detected", 
                f"Tab '{tab_name.replace('📄 ', '')}' has modified values. Do you want to save file changes before closing?"
            )
            if ans is True:
                self.save_file_direct()
            elif ans is None:
                return

        try:
            info["sheet"].destroy()
        except Exception:
            pass

        self.tab_view.delete(tab_name)
        del self.tabs_data[tid]

        if not self.tabs_data:
            self.create_new_tab("untitled", [["" for _ in range(26)] for _ in range(256)], list(string.ascii_uppercase), None)
            self.led_indicator.configure(text_color=("#007bff", "#60a5fa"))
            self.lbl_status.configure(text="Edit below to build a new file, or drop/open a Parquet/Feather database to inspect...", text_color=("#495057", "#9ca3af"))
        else:
            self.on_tab_changed()

    def close_other_tabs(self):
        active_tid = self.get_active_tab_id()
        if not active_tid:
            return

        tids_to_close = [tid for tid in self.tabs_data if tid != active_tid]
        
        for tid in tids_to_close:
            info = self.tabs_data[tid]
            if info["modified"]:
                clean_name = info["tab_name"].replace(" *", "").replace("📄 ", "").strip()
                confirm = messagebox.askyesno(
                    "⚠️ Close Unsaved File View",
                    f"Document '{clean_name}' has changes that were not saved.\n\n"
                    "Are you sure you want to discard changes and close this tab view?"
                )
                if not confirm:
                    continue
            self._execute_tab_closure(tid)

        if active_tid in self.tabs_data:
            self.tab_view.set(self.tabs_data[active_tid]["tab_name"])

    def _execute_tab_closure(self, tid):
        info = self.tabs_data[tid]
        tab_name = info["tab_name"]
        
        try:
            self.tab_view.delete(tab_name)
        except Exception:
            pass
            
        if tid in self.tabs_data:
            del self.tabs_data[tid]
            
        self.on_tab_changed()

    # ==================== 📂 File Ingestion Pipeline ====================

    def check_startup_file(self):
        if len(sys.argv) > 1:
            input_file = sys.argv[1]
            if os.path.exists(input_file):
                self.load_file(target_file=input_file)

    def handle_file_drop(self, event):
        raw_data = event.data
        if not raw_data:
            return
        paths = self.tk.splitlist(raw_data) if hasattr(self, "tk") else [raw_data.strip('{}')]
        
        supported_ext = ('.parquet', '.feather', '.ftr', '.csv', '.xlsx')
        valid_files = [p for p in paths if p.lower().endswith(supported_ext)]
        invalid_files = [p for p in paths if not p.lower().endswith(supported_ext)]
        
        if invalid_files and not valid_files:
            self.led_indicator.configure(text_color=("#b71c1c", "#e74c3c"))
            self.lbl_status.configure(
                text="⚠️ File type not supported! Please import .parquet, .feather, .csv, or .xlsx database structures", 
                text_color=("#b71c1c", "#e74c3c")
            )
            return

        for file_path in valid_files:
            norm_path = os.path.normpath(file_path)
            already_open = False
            for tid, info in self.tabs_data.items():
                if info["file_path"] and os.path.normpath(info["file_path"]) == norm_path:
                    if len(valid_files) == 1:
                        self.led_indicator.configure(text_color=("#1b5e20", "#2ecc71"))
                        self.lbl_status.configure(
                            text="ℹ️ Target file is already loaded. Automatically focused target file view.", 
                            text_color=("#1b5e20", "#2ecc71")
                        )
                        self.tab_view.set(info["tab_name"])
                    already_open = True
                    break
            
            if not already_open:
                self.load_file(target_file=file_path)
        
        if invalid_files:
            skipped = ", ".join(os.path.basename(f) for f in invalid_files)
            self.led_indicator.configure(text_color=("#b45309", "#ffcc00"))
            self.lbl_status.configure(
                text=f"⚠️ Skipped unsupported file paths: {skipped}", 
                text_color=("#b45309", "#ffcc00")
            )

    def load_file(self, target_file=None):
        file_types = [
            ("Supported Grid Files", "*.parquet *.feather *.ftr *.csv *.xlsx"),
            ("Arrow Datasets (Parquet/Feather)", "*.parquet *.feather *.ftr"),
            ("Delimited Datasets (CSV)", "*.csv")
        ]
        if HAS_OPENPYXL:
            file_types.append(("Excel Workbooks", "*.xlsx"))
        else:
            file_types.append(("Excel Workbooks (Requires: pip install openpyxl)", "*.xlsx_disabled"))

        file_path = target_file if target_file else filedialog.askopenfilename(filetypes=file_types)
        if not file_path:
            return

        norm_path = os.path.normpath(file_path)
        for tid, info in self.tabs_data.items():
            if info["file_path"] and os.path.normpath(info["file_path"]) == norm_path:
                self.led_indicator.configure(text_color=("#1b5e20", "#2ecc71"))
                self.lbl_status.configure(
                    text="ℹ️ File is already open in another tab. Automatically shifted view focus.",
                    text_color=("#1b5e20", "#2ecc71")
                )
                self.tab_view.set(info["tab_name"])
                return

        if file_path.endswith(".xlsx_disabled"):
            messagebox.showwarning("Incomplete Setup dependencies", "Please run 'pip install openpyxl' to enable full .xlsx parsing logic.")
            return

        if getattr(self, "_loading_active", False):
            messagebox.showwarning("Please Wait", "A file is already being loaded.")
            return

        self._loading_active = True
        self.led_indicator.configure(text_color=("#b45309", "#ffcc00"))
        self.lbl_status.configure(
            text=f"⏳ Recompiling & Parsing schema models: {os.path.basename(file_path)}",
            text_color=("#b45309", "#ffcc00")
        )
        self.btn_open.configure(state="disabled")
        self.update()

        threading.Thread(
            target=self._load_file_worker,
            args=(file_path,),
            daemon=True
        ).start()

    def _load_file_worker(self, file_path):
        try:
            import pyarrow as pa
            import pyarrow.parquet as pq
            import pyarrow.feather as ft
            import pyarrow.csv as pacsv

            start_time = time.time()
            arrow_table = None
            rows_data = []
            file_headers = []
            lower_path = file_path.lower()

            if lower_path.endswith('.csv'):
                parse_options = pacsv.ParseOptions(newlines_in_values=True)
                arrow_table = pacsv.read_csv(file_path, parse_options=parse_options)

            elif lower_path.endswith('.xlsx'):
                if not HAS_OPENPYXL:
                    raise ImportError("Missing required module 'openpyxl' dependency.")

                wb = openpyxl.load_workbook(file_path, data_only=True)
                ws = wb.active
                excel_rows = list(ws.iter_rows(values_only=True))
                wb.close()

                if excel_rows:
                    file_headers = [str(h) if h is not None else f"Column_{idx}" for idx, h in enumerate(excel_rows[0])]
                    for r in excel_rows[1:]:
                        rows_data.append(["" if val is None else str(val) for val in r])

                arrays = [
                    pa.array([row[i] if i < len(row) else None for row in rows_data], type=pa.string())
                    for i in range(len(file_headers))
                ]
                arrow_table = pa.Table.from_arrays(arrays, names=file_headers)

            elif lower_path.endswith(('.feather', '.ftr')):
                arrow_table = ft.read_table(file_path)

            else:
                arrow_table = pq.read_table(file_path)

            if not rows_data and arrow_table is not None:
                file_headers = arrow_table.column_names
                num_rows = arrow_table.num_rows
                num_cols = len(file_headers)

                col_lists = []
                for col_name in file_headers:
                    str_list = []
                    for val in arrow_table.column(col_name).to_pylist():
                        if val is None:
                            str_list.append("")
                        else:
                            s = str(val)
                            str_list.append("" if s.lower() in ("nan", "nat", "none") else s)
                    col_lists.append(str_list)

                # 欄 → 行
                rows_data = [
                    [col_lists[c][i] for c in range(num_cols)]
                    for i in range(num_rows)
                ]

            self.after(0, lambda: self._on_file_loaded(
                file_path, rows_data, file_headers, arrow_table, start_time
            ))

        except Exception as e:
            try:
                self.after(0, lambda: self._on_file_load_error(str(e)))
            except RuntimeError:
                pass

    def _on_file_loaded(self, file_path, rows_data, file_headers, arrow_table, start_time):
        try:
            base_name = os.path.basename(file_path)
            self.create_new_tab(base_name, rows_data, file_headers, arrow_table, file_path)

            elapsed_time = time.time() - start_time
            file_size_mb = os.path.getsize(file_path) / (1024 * 1024)

            self.led_indicator.configure(text_color=("#1b5e20", "#2ecc71"))
            self.lbl_status.configure(
                text=f"✅ Import successful! | Count: {len(rows_data)} | Size: {file_size_mb:.2f} MB | Time elapsed: {elapsed_time:.3f} s",
                text_color=("#1b5e20", "#2ecc71")
            )
            self.save_preference(file_path)

        except Exception as e:
            self._on_file_load_error(str(e))
        finally:
            self._loading_active = False
            self.btn_open.configure(state="normal")

    def _on_file_load_error(self, error_msg):
        self._loading_active = False
        self.btn_open.configure(state="normal")
        self.led_indicator.configure(text_color=("#b71c1c", "#e74c3c"))
        self.lbl_status.configure(
            text="❌ Processing Stream failed to correctly serialize.",
            text_color=("#b71c1c", "#e74c3c")
        )
        messagebox.showerror("IO Processing Exception", f"Failed to correctly load source stream structures: {error_msg}")

    # ==================== 💾 File Compilation & Save Engine ====================

    def save_file_direct(self):
        tid = self.get_active_tab_id()
        if not tid:
            return

        info = self.tabs_data[tid]

        if not info["modified"]:
            self.led_indicator.configure(text_color=("#1b5e20", "#2ecc71"))
            self.lbl_status.configure(
                text="ℹ️ Active document modifications not found. Duplicate action bypassed.", 
                text_color=("#1b5e20", "#2ecc71")
            )
            return

        if not info["file_path"] or "untitled" in info["tab_name"]:
            self.save_file_as(force_save_as=True)
            return

        self._execute_save(tid, info["file_path"])

    def save_file_as(self, force_save_as=False):
        tid = self.get_active_tab_id()
        if not tid:
            return

        info = self.tabs_data[tid]
        file_path = info["file_path"]

        if not info["modified"] and not force_save_as:
            if not file_path or "untitled" in info["tab_name"]:
                self.led_indicator.configure(text_color=("#1b5e20", "#2ecc71"))
                self.lbl_status.configure(
                    text="ℹ️ Initial blank templates do not require file generation until modified.", 
                    text_color=("#1b5e20", "#2ecc71")
                )
                return

        try:
            file_types = [
                ("Parquet Dataset", "*.parquet"),
                ("Feather Dataset", "*.feather *.ftr"),
                ("CSV Spreadsheet Data", "*.csv")
            ]
            if HAS_OPENPYXL:
                file_types.append(("Excel Spreadsheet Workbook", "*.xlsx"))
            else:
                file_types.append(("Excel Workbooks (Requires: pip install openpyxl)", "*.xlsx_disabled"))

            if file_path:
                ext = "." + file_path.split('.')[-1]
                initial_file = os.path.basename(file_path)
            else:
                ext = ".parquet"
                initial_file = "untitled.parquet"
            
            save_path = filedialog.asksaveasfilename(
                initialfile=initial_file,
                defaultextension=ext,
                filetypes=file_types
            )
            if not save_path:
                return

            if save_path.endswith(".xlsx_disabled"):
                messagebox.showwarning("Dependency Error", "Please execute 'pip install openpyxl' to compile target as .xlsx formats")
                return

            self._execute_save(tid, save_path)

        except Exception as e:
            messagebox.showerror("File Writing Error", f"Could not export document stream correctly: {e}")

    def _execute_save(self, tid, save_path):
        if self._saving_active:
            messagebox.showwarning("Please Wait", "A save is already in progress.")
            return
        self._saving_active = True
        info = self.tabs_data[tid]
        sheet = info["sheet"]
        tab_headers = sheet.headers()
        rows = sheet.get_sheet_data()
        table_data = info["table_data"]

        self.led_indicator.configure(text_color=("#b45309", "#ffcc00"))
        self.lbl_status.configure(
            text=f"⏳ Saving: {os.path.basename(save_path)}...",
            text_color=("#b45309", "#ffcc00")
        )
        self.btn_save.configure(state="disabled")
        self.update()

        threading.Thread(
            target=self._save_worker,
            args=(tid, save_path, tab_headers, rows, table_data),
            daemon=True
        ).start()

    def _save_worker(self, tid, save_path, tab_headers, rows, table_data):
        try:
            import pyarrow as pa
            import pyarrow.parquet as pq
            import pyarrow.feather as ft

            save_ext = save_path.lower()

            if save_ext.endswith(".csv"):
                with open(save_path, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    writer.writerow(tab_headers)
                    writer.writerows(rows)

            elif save_ext.endswith(".xlsx"):
                if not HAS_OPENPYXL:
                    raise ImportError("openpyxl not installed.")
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Sheet1"
                ws.append(tab_headers)
                for r in rows:
                    ws.append(r)
                wb.save(save_path)
                wb.close()

            else:
                col_data = {col: [] for col in tab_headers}
                for row in rows:
                    for col_idx, col_name in enumerate(tab_headers):
                        if col_idx < len(row):
                            val = row[col_idx]
                            col_data[col_name].append(None if val == "" or val is None else val)
                        else:
                            col_data[col_name].append(None)

                arrays = []
                new_table = None
                if table_data is not None:
                    original_schema = table_data.schema
                    new_fields = []
                    for col_name in tab_headers:
                        try:
                            field_type = original_schema.field(col_name).type
                            arrays.append(pa.array(col_data[col_name], type=field_type))
                            new_fields.append(pa.field(col_name, field_type))
                        except Exception:
                            arr = pa.array(col_data[col_name])
                            arrays.append(arr)
                            new_fields.append(pa.field(col_name, arr.type))
                    new_schema = pa.schema(new_fields)
                    new_table = pa.Table.from_arrays(arrays, schema=new_schema)
                else:
                    for col_name in tab_headers:
                        arrays.append(pa.array(col_data[col_name], type=pa.string()))
                    new_table = pa.Table.from_arrays(arrays, names=tab_headers)

                if save_ext.endswith(('.feather', '.ftr')):
                    ft.write_feather(new_table, save_path)
                else:
                    pq.write_table(new_table, save_path, compression="snappy")

            self.after(0, lambda: self._on_save_done(tid, save_path, 
                new_table if not save_ext.endswith((".csv", ".xlsx")) else None))

        except Exception as e:
            self.after(0, lambda: self._on_save_error(str(e)))

    def _on_save_done(self, tid, save_path, new_table):
        self._saving_active = False
        self.btn_save.configure(state="normal")

        if tid not in self.tabs_data:
            return

        info = self.tabs_data[tid]

        if new_table is not None:
            info["table_data"] = new_table

        old_title = info["tab_name"]
        clean_name = f"📄 {os.path.basename(save_path)}"
        self.tab_view.rename(old_title, clean_name)

        try:
            tab_button = self.tab_view._segmented_button._buttons_dict[clean_name]
            tab_button.bind("<Button-3>", self.show_tab_context_menu)
            tab_button.bind("<Double-Button-1>", self._on_tab_double_click)
        except Exception:
            pass

        info["file_path"] = save_path
        info["tab_name"] = clean_name
        info["modified"] = False

        self.on_tab_changed()
        self.save_preference(save_path)

        self.led_indicator.configure(text_color=("#1b5e20", "#2ecc71"))
        self.lbl_status.configure(
            text=f"💾 Saved: {os.path.basename(save_path)}",
            text_color=("#1b5e20", "#2ecc71")
        )

    def _on_save_error(self, error_msg):
        self._saving_active = False
        self.btn_save.configure(state="normal")
        self.led_indicator.configure(text_color=("#b71c1c", "#e74c3c"))
        self.lbl_status.configure(
            text="❌ Save failed.",
            text_color=("#b71c1c", "#e74c3c")
        )
        messagebox.showerror("Save Error", error_msg)

    def on_app_close(self):
        unsaved_tabs = []
        
        for tid, info in self.tabs_data.items():
            if info["modified"]:
                clean_name = info["tab_name"].replace(" *", "").strip()
                unsaved_tabs.append(clean_name)
        
        if unsaved_tabs:
            file_list_str = "\n".join([f"• {name}" for name in unsaved_tabs])
            
            confirm = messagebox.askyesno(
                "⚠️ Unsaved Document Variations Detected",
                f"The following open files have unsaved changes:\n\n{file_list_str}\n\n"
                "👉 Press [Yes]: Force termination (⚠️ All unsaved modifications will be permanently lost!)\n"
                "👉 Press [No]: Abort close sequence and return to save workflow manually",
                icon='warning'
            )
            
            if confirm:
                self.destroy()
            else:
                return
        else:
            self.destroy()

    # ==================== 📐 Grid Operations & Column Logic ====================

    def autofit_columns(self):
        tid = self.get_active_tab_id()
        if not tid:
            return

        info = self.tabs_data[tid]
        sheet = info["sheet"]
        tab_headers = sheet.headers()

        self.led_indicator.configure(text_color=("#b45309", "#ffcc00"))
        self.lbl_status.configure(
            text="⚡ Executing fast layout autofit calculation models...", 
            text_color=("#b45309", "#ffcc00")
        )
        self.update()
        
        try:
            rows = sheet.get_sheet_data()
            if not rows:
                return
                
            num_cols = len(tab_headers)
            new_widths = []
            char_multiplier = self.current_font_size * 0.75  
            
            for col_idx in range(num_cols):
                max_width = 50  
                col_name = tab_headers[col_idx]
                
                header_len = sum(2 if ord(char) > 127 else 1 for char in col_name)
                max_width = max(max_width, header_len * char_multiplier + 30)
                
                for row in rows[:1000]:
                    if col_idx < len(row):
                        cell_val = str(row[col_idx])
                        if cell_val:
                            val_len = sum(2 if ord(char) > 127 else 1 for char in cell_val)
                            max_width = max(max_width, val_len * char_multiplier + 20)
                
                new_widths.append(max(50, min(int(max_width), 800)))
            
            sheet.set_column_widths(new_widths)
            sheet.redraw()
            
            total_rows = len(rows)
            self.led_indicator.configure(text_color=("#1b5e20", "#2ecc71"))
            self.lbl_status.configure(
                text=f"✅ Realignment compilation completed! (Calculated via sampled metadata) | Total rows: {total_rows}", 
                text_color=("#1b5e20", "#2ecc71")
            )
        except Exception as e:
            self.led_indicator.configure(text_color=("#b71c1c", "#e74c3c"))
            self.lbl_status.configure(
                text=f"❌ Failed to calculate appropriate widths: {e}", 
                text_color=("#b71c1c", "#e74c3c")
            )

    def on_double_click_resize(self, event):
        tid = self.get_active_tab_id()
        if not tid:
            return

        info = self.tabs_data[tid]
        sheet = info["sheet"]
        tab_headers = sheet.headers()

        try:
            col_idx = event[1]
        except (IndexError, TypeError):
            return
            
        rows = sheet.get_sheet_data()
        if not rows or col_idx < 0 or col_idx >= len(tab_headers):
            return
        
        char_multiplier = self.current_font_size * 0.75
        col_name = tab_headers[col_idx]
        header_len = sum(2 if ord(char) > 127 else 1 for char in col_name)
        max_width = header_len * char_multiplier + 30
        
        for row in rows[:1000]:
            if col_idx < len(row):
                cell_val = str(row[col_idx])
                if cell_val:
                    val_len = sum(2 if ord(char) > 127 else 1 for char in cell_val)
                    max_width = max(max_width, val_len * char_multiplier + 20)
                    
        final_width = max(50, min(int(max_width), 800))
        sheet.set_column_width(col_idx, final_width)
        sheet.redraw()

    # ==================== 📊 Schema Viewer ====================

    def show_schema_viewer(self):
        tid = self.get_active_tab_id()
        if not tid:
            return

        info = self.tabs_data[tid]
        if info["table_data"] is None:
            messagebox.showwarning("Incompatible Structure", "This active tab does not inherit typed binary database properties!")
            return

        schema_win = ctk.CTkToplevel(self)
        schema_win.title("📊 Schema Definition Diagnostics")
        schema_win.geometry("520x450")
        schema_win.resizable(False, False)
        schema_win.attributes("-topmost", True)
        
        header_frame = ctk.CTkFrame(schema_win, height=60, corner_radius=0, fg_color=("#e9ecef", "#18181c"))
        header_frame.pack(fill="x", side="top")
        
        lbl_title = ctk.CTkLabel(
            header_frame, 
            text="📊 Arrow Column Schema & Physical Types", 
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold")
        )
        lbl_title.pack(padx=20, pady=15, anchor="w")

        scroll_frame = ctk.CTkScrollableFrame(schema_win, fg_color="transparent")
        scroll_frame.pack(fill="both", expand=True, padx=15, pady=15)

        legend_frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
        legend_frame.pack(fill="x", pady=(0, 8))
        
        ctk.CTkLabel(legend_frame, text="Column Key", font=ctk.CTkFont(size=11, weight="bold"), text_color="#868e96").pack(side="left", padx=10)
        ctk.CTkLabel(legend_frame, text="Arrow Physical Type", font=ctk.CTkFont(size=11, weight="bold"), text_color="#868e96").pack(side="right", padx=10)

        schema = info["table_data"].schema
        tab_headers = info["sheet"].headers()
        for idx, col_name in enumerate(tab_headers):
            try:
                arrow_type = schema.field(col_name).type
            except Exception:
                arrow_type = "string (default)"

            item_frame = ctk.CTkFrame(
                scroll_frame, 
                height=38, 
                fg_color=("#f1f3f5", "#1e1e24"), 
                border_width=1, 
                border_color=("#dee2e6", "#2d2d34")
            )
            item_frame.pack(fill="x", pady=3)
            item_frame.pack_propagate(False)

            lbl_col = ctk.CTkLabel(
                item_frame, 
                text=f"{idx+1}.  {col_name}", 
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold")
            )
            lbl_col.pack(side="left", padx=15)

            lbl_type = ctk.CTkLabel(
                item_frame, 
                text=str(arrow_type).upper(), 
                font=ctk.CTkFont(family="Consolas", size=11, weight="bold"),
                text_color=("#d9480f", "#f59e0b")
            )
            lbl_type.pack(side="right", padx=15)

    # ==================== 📊 Selection Statistics ====================

    def update_selection_stats(self, event=None):
        if hasattr(self, "_selection_stats_timer") and self._selection_stats_timer is not None:
            self.after_cancel(self._selection_stats_timer)
            self._selection_stats_timer = None

        self._selection_stats_timer = self.after(30, self._execute_selection_stats_update)

    def _execute_selection_stats_update(self):
        self._selection_stats_timer = None
        
        tid = self.get_active_tab_id()
        if not tid or tid not in self.tabs_data:
            return

        info = self.tabs_data[tid]
        sheet = info["sheet"]
        
        rows_count = len(sheet.get_sheet_data())
        if info["file_path"]:
            base_name = os.path.basename(info["file_path"])
            status_text = f"📄 Active Tab: {base_name} | Rows: {rows_count}"
        else:
            status_text = f"📄 Active Tab: {info['tab_name']} | Rows: {rows_count}"

        if info["modified"]:
            status_text += " | ⚠️ Unsaved Changes"

        try:
            selections = sheet.get_all_selection_boxes()
            if selections:
                r1, c1, r2, c2 = selections[-1]
                
                selected_rows = abs(r2 - r1)
                selected_cols = abs(c2 - c1)
                total_cells = selected_rows * selected_cols
                
                if total_cells > 1:
                    stats_str = f"  |  🔲 Selected: {selected_rows} Row(s) × {selected_cols} Col(s) ({total_cells} cells)"
                    status_text += stats_str
        except Exception:
            pass

        if info["modified"]:
            self.led_indicator.configure(text_color=("#b45309", "#ffcc00"))
            self.lbl_status.configure(text=status_text, text_color=("#b45309", "#ffcc00"))
        else:
            self.led_indicator.configure(text_color=("#1b5e20", "#2ecc71"))
            self.lbl_status.configure(text=status_text, text_color=("#495057", "#9ca3af"))

    def clear_selection_stats(self, event=None):
        self.on_tab_changed()

    # ==================== 🔍 Inline Search Panel ====================

    def toggle_search_bar(self, event=None):
        if self.search_bar_visible:
            self.hide_search_bar()
        else:
            self.show_search_bar()
        return "break"

    def show_search_bar(self):
        if not self.search_bar_visible:
            self.search_bar.pack(fill="x", pady=(0, 8), before=self.table_container)
            self.search_bar_visible = True
        self.search_entry.focus_set()
        tid = self.get_active_tab_id()
        if tid and tid in self.tabs_data:
            sheet = self.tabs_data[tid]["sheet"]
            try:
                selected = sheet.get_currently_selected()
                if selected:
                    row, col = selected.row, selected.column
                    val = sheet.get_cell_data(row, col)
                    if val:
                        self.search_entry.delete(0, "end")
                        self.search_entry.insert(0, str(val))
                        self.search_entry.select_range(0, "end")
            except Exception:
                pass

    def hide_search_bar(self):
        if self.search_bar_visible:
            self.search_bar.pack_forget()
            self.search_bar_visible = False
        self.search_matches = []
        self.search_current_idx = -1
        self._last_search_keyword = None
        self.lbl_search_result.configure(text="")
        self._clear_search_highlight()

    def _clear_search_highlight(self):
        tid = self.get_active_tab_id()
        if tid and tid in self.tabs_data:
            sheet = self.tabs_data[tid]["sheet"]
            try:
                sheet.dehighlight_cells(all_=True)
                sheet.redraw()
            except Exception:
                pass

    def _execute_search(self):
        keyword = self.search_entry.get().strip().lower()

        if not keyword:
            self.search_matches = []
            self.search_current_idx = -1
            self.lbl_search_result.configure(text="")
            self._clear_search_highlight()
            return

        if hasattr(self, "_last_search_keyword") and self._last_search_keyword == keyword:
            return
        self._last_search_keyword = keyword

        self.search_matches = []
        self.search_current_idx = -1

        tid = self.get_active_tab_id()
        if not tid or tid not in self.tabs_data:
            return

        sheet = self.tabs_data[tid]["sheet"]
        rows = sheet.get_sheet_data()

        self._clear_search_highlight()

        for r_idx, row in enumerate(rows):
            for c_idx, cell in enumerate(row):
                if keyword in str(cell).lower():
                    self.search_matches.append((r_idx, c_idx))
                    sheet.highlight_cells(row=r_idx, column=c_idx, bg="#fff3bf", fg="#000000")

        if self.search_matches:
            self.lbl_search_result.configure(
                text=f"Found {len(self.search_matches)} occurrences",
                text_color=("#1b5e20", "#2ecc71")
            )
        else:
            self.lbl_search_result.configure(
                text="No matching search entries found",
                text_color=("#b71c1c", "#e74c3c")
            )
        sheet.redraw()

    def find_next(self):
        self._execute_search()
        if not self.search_matches:
            return
        self.search_current_idx = (self.search_current_idx + 1) % len(self.search_matches)
        self._jump_to_match()

    def find_prev(self):
        self._execute_search()
        if not self.search_matches:
            return
        self.search_current_idx = (self.search_current_idx - 1) % len(self.search_matches)
        self._jump_to_match()

    def _jump_to_match(self):
        if not self.search_matches or self.search_current_idx < 0:
            return

        tid = self.get_active_tab_id()
        if not tid:
            return

        sheet = self.tabs_data[tid]["sheet"]
        r, c = self.search_matches[self.search_current_idx]

        self._clear_search_highlight()
        for mr, mc in self.search_matches:
            sheet.highlight_cells(row=mr, column=mc, bg="#fff3bf", fg="#000000")
        sheet.highlight_cells(row=r, column=c, bg="#ff922b", fg="#ffffff")

        sheet.see(row=r, column=c)
        sheet.select_cell(r, c)
        sheet.redraw()

        self.lbl_search_result.configure(
            text=f"{self.search_current_idx + 1} / {len(self.search_matches)} matches",
            text_color=("#1b5e20", "#2ecc71")
        )

    # ==================== 🔄 Quick Converter Drag & Drop ====================

    def on_convert_drag_enter(self, event):
        self.drop_zone.configure(
            fg_color=("#d0ebff", "#1e3a5f"),
            text="📥\nRelease mouse button\nto convert!"
        )

    def on_convert_drag_leave(self, event):
        self.drop_zone.configure(
            fg_color=("#f8f9fa", "#111115"),
            text="📥\nDrag and drop files\nhere to convert"
        )

    def on_convert_drop(self, event):
        self.drop_zone.configure(
            fg_color=("#f8f9fa", "#111115"),
            text="📥\nDrag and drop files\nhere to convert"
        )

        raw = event.data
        if not raw:
            return
        paths = self.tk.splitlist(raw) if hasattr(self, "tk") else [raw.strip('{}')]

        target_fmt = self.convert_fmt.get()

        if self._converting_active:
            messagebox.showwarning("Please Wait", "A conversion is already in progress.")
            return

        self._converting_active = True
        self.led_indicator.configure(text_color=("#b45309", "#ffcc00"))
        self.lbl_status.configure(
            text=f"⏳ Converting {len(paths)} file(s) → {target_fmt}...",
            text_color=("#b45309", "#ffcc00")
        )
        self.drop_zone.configure(text="⏳ Converting...")
        self.update()

        threading.Thread(
            target=self._convert_worker,
            args=(list(paths), target_fmt),
            daemon=True
        ).start()

    def _convert_worker(self, paths, target_fmt):
        try:
            import pyarrow as pa
            import pyarrow.parquet as pq
            import pyarrow.feather as ft
            import pyarrow.csv as pacsv
            if HAS_OPENPYXL:
                import openpyxl

            ext_map = {
                "Parquet": ".parquet",
                "Feather": ".feather",
                "CSV": ".csv",
                "Excel": ".xlsx"
            }
            target_ext = ext_map[target_fmt]

            converted = 0
            errors = []

            for path in paths:
                path = path.strip()
                if not os.path.isfile(path):
                    continue
                src_ext = os.path.splitext(path)[1].lower()
                if src_ext == target_ext:
                    continue
                try:
                    # ── Read ──
                    if src_ext == ".parquet":
                        table = pq.read_table(path)
                    elif src_ext in (".feather", ".ftr"):
                        table = ft.read_table(path)
                    elif src_ext == ".csv":
                        parse_options = pacsv.ParseOptions(newlines_in_values=True)
                        table = pacsv.read_csv(path, parse_options=parse_options)
                    elif src_ext in (".xlsx", ".xls"):
                        if not HAS_OPENPYXL:
                            continue
                        wb = openpyxl.load_workbook(path, read_only=True)
                        ws = wb.active
                        data = list(ws.values)
                        wb.close()
                        if not data:
                            continue
                        headers = [str(h) if h else f"col_{i}" for i, h in enumerate(data[0])]
                        rows = data[1:]
                        arrays = []
                        for ci in range(len(headers)):
                            col_vals = [row[ci] if ci < len(row) else None for row in rows]
                            arrays.append(pa.array(col_vals))
                        table = pa.table({h: a for h, a in zip(headers, arrays)})
                    else:
                        continue

                    # ── Write ──
                    out_path = os.path.splitext(path)[0] + target_ext
                    if target_fmt == "Parquet":
                        pq.write_table(table, out_path)
                    elif target_fmt == "Feather":
                        ft.write_feather(table, out_path)
                    elif target_fmt == "CSV":
                        pacsv.write_csv(table, out_path)
                    elif target_fmt == "Excel":
                        if not HAS_OPENPYXL:
                            continue
                        wb = openpyxl.Workbook()
                        ws = wb.active
                        ws.append(table.column_names)
                        for row in zip(*[col.to_pylist() for col in table.columns]):
                            ws.append(list(row))
                        wb.save(out_path)

                    converted += 1

                except Exception as e:
                    errors.append(f"{os.path.basename(path)}: {e}")
                    continue

            self.after(0, lambda: self._on_convert_done(converted, target_fmt, errors))

        except Exception as e:
            self.after(0, lambda: self._on_convert_error(str(e)))

    def _on_convert_done(self, converted, target_fmt, errors):
        self._converting_active = False

        if converted > 0:
            self.led_indicator.configure(text_color=("#1b5e20", "#2ecc71"))
            status = f"✅ Conversion complete on {converted} structure(s) → {target_fmt}"
            if errors:
                status += f"  |  ⚠️ {len(errors)} failed"
            self.lbl_status.configure(
                text=status,
                text_color=("#1b5e20", "#2ecc71")
            )
            self.drop_zone.configure(text=f"✅ Complete ({converted})")
            self.after(2000, lambda: self.drop_zone.configure(
                text="📥\nDrag and drop files\nhere to convert"
            ))
        else:
            self.led_indicator.configure(text_color=("#b45309", "#ffcc00"))
            if errors:
                self.lbl_status.configure(
                    text=f"❌ All conversions failed ({len(errors)} error(s))",
                    text_color=("#b71c1c", "#e74c3c")
                )
                messagebox.showerror(
                    "Conversion Errors",
                    "\n".join(errors[:10])
                )
            else:
                self.lbl_status.configure(
                    text="⚠️ No compatible structures compiled during operation.",
                    text_color=("#b45309", "#ffcc00")
                )
            self.drop_zone.configure(text="📥\nDrag and drop files\nhere to convert")

    def _on_convert_error(self, error_msg):
        self._converting_active = False
        self.led_indicator.configure(text_color=("#b71c1c", "#e74c3c"))
        self.lbl_status.configure(
            text="❌ Conversion failed.",
            text_color=("#b71c1c", "#e74c3c")
        )
        self.drop_zone.configure(text="📥\nDrag and drop files\nhere to convert")
        messagebox.showerror("Conversion Error", error_msg)


# ==================== 🚀 Application Entry Point ====================
if __name__ == "__main__":
    app = ModernEditableEditor()
    app.mainloop()
