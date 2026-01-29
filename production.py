""" Modern Production Tool - HIGHLIGHTER STYLE - Complete Integration
Full conversion from box selection to highlighter annotations
AUTO-OPENS PRODUCTION MODE when cabinet is loaded from queue
UPDATED: Proper highlighter display, box annotations removed
"""
import tkinter as tk
from tkinter import messagebox, simpledialog, Menu
from PIL import Image, ImageTk, ImageDraw, ImageFont
import fitz  # PyMuPDF
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from datetime import datetime
import os
import sys
import json
import getpass
import re
import sqlite3
import numpy as np
from handover_database import HandoverDB
from database_manager import DatabaseManager
import sys

User = sys.argv[1] if len(sys.argv) > 1 else None
Name = sys.argv[2] if len(sys.argv) > 2 else None

print(f"✓ Production Tool started by: {Name} (username: {User})")


def getbase():
    """Returns the directory where the app is running from"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


class ManagerDB:
    """Manager database integration for status tracking"""
    
    def __init__(self, db_path):
        self.db_path = db_path
        self.initializedatabase()
    
    def initializedatabase(self):
        """Initialize tables if they don't exist"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''CREATE TABLE IF NOT EXISTS cabinets (
            cabinet_id TEXT PRIMARY KEY,
            project_name TEXT,
            sales_order_no TEXT,
            total_pages INTEGER DEFAULT 0,
            annotated_pages INTEGER DEFAULT 0,
            total_punches INTEGER DEFAULT 0,
            open_punches INTEGER DEFAULT 0,
            implemented_punches INTEGER DEFAULT 0,
            closed_punches INTEGER DEFAULT 0,
            status TEXT DEFAULT 'quality_inspection',
            created_date TEXT,
            last_updated TEXT,
            storage_location TEXT,
            excel_path TEXT
        )''')
        
        cursor.execute('''CREATE TABLE IF NOT EXISTS category_occurrences (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cabinet_id TEXT,
            project_name TEXT,
            category TEXT,
            subcategory TEXT,
            occurrence_date TEXT
        )''')
        
        # Add columns if they don't exist
        try:
            cursor.execute('ALTER TABLE cabinets ADD COLUMN storage_location TEXT')
        except sqlite3.OperationalError:
            pass
        
        try:
            cursor.execute('ALTER TABLE cabinets ADD COLUMN excel_path TEXT')
        except sqlite3.OperationalError:
            pass
        
        conn.commit()
        conn.close()
    
    def updcab(self, cabinet_id, project_name, sales_order_no, total_pages, annotated_pages,
                      total_punches, open_punches, implemented_punches, closed_punches, status,
                      storage_location=None, excel_path=None):
        """Update cabinet statistics WITH excel_path and storage_location"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                INSERT OR REPLACE INTO cabinets
                (cabinet_id, project_name, sales_order_no, total_pages, annotated_pages,
                 total_punches, open_punches, implemented_punches, closed_punches, status,
                 storage_location, excel_path, created_date, last_updated)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                        COALESCE((SELECT created_date FROM cabinets WHERE cabinet_id = ?), ?),
                        ?)
            ''', (cabinet_id, project_name, sales_order_no, total_pages, annotated_pages,
                  total_punches, open_punches, implemented_punches, closed_punches, status,
                  storage_location, excel_path,
                  cabinet_id, datetime.now().isoformat(),
                  datetime.now().isoformat()))
            
            conn.commit()
            conn.close()
            print(f"✓ Manager DB: Updated {cabinet_id} - Status: {status}")
            return True
        except Exception as e:
            print(f"Manager DB update error: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def updstats(self, cabinet_id, status):
        """Update cabinet status only"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                UPDATE cabinets
                SET status = ?, last_updated = ?
                WHERE cabinet_id = ?
            ''', (status, datetime.now().isoformat(), cabinet_id))
            
            conn.commit()
            conn.close()
            print(f"✓ Manager DB: Status updated for {cabinet_id} → {status}")
            return True
        except Exception as e:
            print(f"Status update error: {e}")
            return False


class ProductionTool:
    def __init__(self, root):
        self.root = root
        self.logged_in_username = User
        self.logged_in_fullname = Name
        self.root.title("Production Tool - Highlighter Mode")
        self.root.geometry("1400x900")
        # Bind window close event to auto-save
        self.root.protocol("WM_DELETE_WINDOW", self.closing)
        
        # Data / files
        self.pdf_document = None
        self.current_pdf_path = None
        self.current_page = 0
        self.project_name = ""
        self.sales_order_no = ""
        self.cabinet_id = ""
        self.storage_location = ""
        self.annotations = []
        
        base = getbase()
        self.handover_db = HandoverDB(os.path.join(base, "handover_db.json"))
        self.db = DatabaseManager(os.path.join(base, "inspection_tool.db"))
        self.manager_db = ManagerDB(os.path.join(base, "manager.db"))
        
        self.excel_file = None
        self.working_excel_path = None
        self.zoom_level = 1.0
        self.current_sr_no = 1
        self.current_page_image = None
        self.session_refs = set()
        
        # Visual navigation for production mode
        self.production_highlight_tags = []
        self.production_dialog_open = False
        
        # Highlighter colors with RGBA for semi-transparency
        self.highlighter_colors = {
            'yellow': {'rgb': (255, 255, 0), 'rgba': (255, 255, 0, 100)},
            'green': {'rgb': (0, 255, 0), 'rgba': (0, 255, 0, 100)},
            'blue': {'rgb': (0, 191, 255), 'rgba': (0, 191, 255, 100)},
            'pink': {'rgb': (255, 105, 180), 'rgba': (255, 105, 180, 100)},
            'orange': {'rgb': (255, 165, 0), 'rgba': (255, 165, 0, 100)}
        }
        
        # Column mapping
        self.punch_sheet_name = 'Punch Sheet'
        self.punch_cols = {
            'sr_no': 'A',
            'ref_no': 'B',
            'desc': 'C',
            'category': 'D',
            'implemented_name': 'G',
            'implemented_date': 'H',
            'closed_name': 'I',
            'closed_date': 'J'
        }
        
        self.interphase_sheet_name = 'Interphase'
        self.interphase_cols = {
            'ref_no': 'B',
            'description': 'C',
            'status': 'D',
        }
        
        self.header_cells = {
            "Interphase": {
                "project_name": "C4",
                "sales_order": "C6",
                "cabinet_id": "F6"
            },
            "Punch Sheet": {
                "project_name": "C2",
                "sales_order": "C4",
                "cabinet_id": "H4"
            }
        }
        
        # Highlighter drawing state - NO BOX SELECTION
        self.drawing = False
        self.highlighter_start_x = None
        self.highlighter_start_y = None
        self.temp_highlight_id = None
        self.selected_annotation = None
        
        # Tool modes (pen, text)
        self.current_tool = None  # None, 'pen', 'text'
        self.tool_mode = None  # Alias for current_tool
        self.pen_points = []
        self.temp_pen_line = None
        self.temp_line_ids = []  # Store temporary drawing line IDs
        self.drawing_type = None  # 'pen', 'text'
        self.text_pos_x = None
        self.text_pos_y = None
        
        # Highlighter state
        self.active_highlighter = False
        
        # Undo stack
        self.undo_stack = []
        self.max_undo = 50
        
        self.uisetup()
        self.current_sr_no = self.getnextsr()

    # ================================================================
    # MANAGER SYNC - PRODUCTION SPECIFIC
    # ================================================================
    
    def syncmgrstats(self):
        """Sync current cabinet statistics to manager database"""
        if not self.cabinet_id:
            return
        
        try:
            # Count from Excel - start from row 9
            implemented_punches = 0
            closed_punches = 0
            total_punches = 0
            
            if self.excel_file and os.path.exists(self.excel_file):
                try:
                    wb = load_workbook(self.excel_file, data_only=True)
                    ws = wb[self.punch_sheet_name] if self.punch_sheet_name in wb.sheetnames else wb.active
                    
                    row = 9  # Start from row 9
                    while row <= ws.max_row + 5:
                        checked = self.read_cell(ws, row, 'E')
                        if not checked:
                            row += 1
                            if row > 2000:
                                break
                            continue
                        
                        total_punches += 1
                        implemented = self.read_cell(ws, row, self.punch_cols['implemented_name'])
                        closed = self.read_cell(ws, row, self.punch_cols['closed_name'])
                        
                        if closed:
                            closed_punches += 1
                        elif implemented:
                            implemented_punches += 1
                        
                        row += 1
                        if row > 2000:
                            break
                    
                    wb.close()
                except Exception as e:
                    print(f"Excel read error: {e}")
            
            open_punches = total_punches - implemented_punches - closed_punches
            
            self.manager_db.updcab(
                self.cabinet_id,
                self.project_name,
                self.sales_order_no,
                0,
                0,
                total_punches,
                open_punches,
                implemented_punches,
                closed_punches,
                'in_progress',
                storage_location=getattr(self, 'storage_location', None),
                excel_path=self.excel_file
            )
        
        except Exception as e:
            print(f"Manager sync error: {e}")
            import traceback
            traceback.print_exc()
    
    def syncmgrstatsonly(self):
        """Lightweight sync without full recount - for display updates"""
        # Only sync if we have the necessary data loaded
        if self.cabinet_id and self.excel_file:
            self.syncmgrstats()

    # ================================================================
    # CELL HELPERS
    # ================================================================
    
    def split_cell(self, cell_ref):
        m = re.match(r"([A-Z]+)(\d+)", cell_ref)
        if not m:
            raise ValueError(f"Invalid cell reference: {cell_ref}")
        col, row = m.groups()
        return int(row), col
    
    def _resolve_merged_target(self, ws, row, col_idx):
        for merged in ws.merged_cells.ranges:
            if merged.min_row <= row <= merged.max_row and merged.min_col <= col_idx <= merged.max_col:
                return merged.min_row, merged.min_col
        return row, col_idx
    
    def write_cell(self, ws, row, col, value):
        if isinstance(col, str):
            col_idx = column_index_from_string(col)
        else:
            col_idx = int(col)
        target_row, target_col = self._resolve_merged_target(ws, int(row), col_idx)
        ws.cell(row=target_row, column=target_col).value = value
    
    def read_cell(self, ws, row, col):
        if isinstance(col, str):
            col_idx = column_index_from_string(col)
        else:
            col_idx = int(col)
        target_row, target_col = self._resolve_merged_target(ws, int(row), col_idx)
        return ws.cell(row=target_row, column=target_col).value

    # ================================================================
    # MODERN UI SETUP
    # ================================================================
    
    def uisetup(self):
        """Setup modern professional UI with highlighter mode"""
        # Main toolbar
        toolbar = tk.Frame(self.root, bg='#1e293b', height=70)
        toolbar.pack(side=tk.TOP, fill=tk.X)
        
        # Enhanced Menu Bar
        menubar = Menu(self.root, bg='#1e293b', fg='white', activebackground='#3b82f6')
        self.root.config(menu=menubar)
        
        # File Menu
        file_menu = Menu(menubar, tearoff=0, bg='#1e293b', fg='white', activebackground='#3b82f6')
        menubar.add_cascade(label="📁 File", menu=file_menu)
        file_menu.add_command(label="Load from Production Queue", command=self.loadfrmhandover, accelerator="Ctrl+O")
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Tools Menu
        tools_menu = Menu(menubar, tearoff=0, bg='#1e293b', fg='white', activebackground='#3b82f6')
        menubar.add_cascade(label="🛠️ Tools", menu=tools_menu)
        tools_menu.add_command(label="🏭 Production Mode", command=self.prodmode, accelerator="Ctrl+P")
        tools_menu.add_separator()
        tools_menu.add_command(label="✅ Complete & Handback", command=self.compreworkhndbck, accelerator="Ctrl+H")
        
        # View Menu
        view_menu = Menu(menubar, tearoff=0, bg='#1e293b', fg='white', activebackground='#3b82f6')
        menubar.add_cascade(label="👁️ View", menu=view_menu)
        view_menu.add_command(label="Zoom In", command=self.zoom, accelerator="Ctrl++")
        view_menu.add_command(label="Zoom Out", command=self.zoomout, accelerator="Ctrl+-")
        view_menu.add_command(label="Reset Zoom", command=lambda: setattr(self, 'zoom_level', 1.0) or self.display())
        
        # Keyboard shortcuts
        self.root.bind_all("<Control-o>", lambda e: self.loadfrmhandover())
        self.root.bind_all("<Control-p>", lambda e: self.prodmode())
        self.root.bind_all("<Control-h>", lambda e: self.compreworkhndbck())
        self.root.bind_all("<Control-plus>", lambda e: self.zoom())
        self.root.bind_all("<Control-minus>", lambda e: self.zoomout())
        self.root.bind_all("<Control-z>", lambda e: self.undolast())
        self.root.bind_all("<Escape>", lambda e: self.deactivate_all())
        
        # Left section - Load operations
        left_frame = tk.Frame(toolbar, bg='#1e293b')
        left_frame.pack(side=tk.LEFT, padx=10, pady=10)
        
        tk.Button(left_frame, text="📦 Load from Queue", command=self.loadfrmhandover,
                 bg='#8b5cf6', fg='white', padx=15, pady=10,
                 font=('Segoe UI', 10, 'bold'), relief=tk.FLAT, borderwidth=0,
                 cursor='hand2').pack(side=tk.LEFT, padx=3)
        
        # Center section - Navigation
        center_frame = tk.Frame(toolbar, bg='#1e293b')
        center_frame.pack(side=tk.LEFT, padx=20)
        
        self.page_label = tk.Label(center_frame, text="Page: 0/0", bg='#1e293b', fg='white',
                                   font=('Segoe UI', 10, 'bold'))
        self.page_label.pack(side=tk.LEFT, padx=10)
        
        nav_btn_style = {
            'bg': '#64748b',
            'fg': 'white',
            'font': ('Segoe UI', 9, 'bold'),
            'relief': tk.FLAT,
            'cursor': 'hand2'
        }
        
        tk.Button(center_frame, text="◀", command=self.prev, width=3,
                 **nav_btn_style).pack(side=tk.LEFT, padx=2)
        tk.Button(center_frame, text="▶", command=self.next, width=3,
                 **nav_btn_style).pack(side=tk.LEFT, padx=2)
        
        # Zoom controls
        zoom_frame = tk.Frame(center_frame, bg='#1e293b')
        zoom_frame.pack(side=tk.LEFT, padx=15)
        
        zoom_btn_style = nav_btn_style.copy()
        zoom_btn_style['bg'] = '#10b981'
        
        tk.Button(zoom_frame, text="🔍+", command=self.zoom, width=4,
                 **zoom_btn_style).pack(side=tk.LEFT, padx=2)
        tk.Button(zoom_frame, text="🔍−", command=self.zoomout, width=4,
                 **zoom_btn_style).pack(side=tk.LEFT, padx=2)
        
        # Tool section - Pen, Text, Undo
        tool_frame = tk.Frame(toolbar, bg='#1e293b')
        tool_frame.pack(side=tk.LEFT, padx=10)

        tk.Label(tool_frame, text="Tools:", bg='#1e293b', fg='#94a3b8', 
                 font=('Segoe UI', 9)).pack(side=tk.LEFT, padx=(0, 8))

        # Load icons or use fallback
        self.pen_btn = None
        self.text_btn = None
        
        try:
            assets_dir = os.path.join(os.path.dirname(getbase()), "assets")
            icon_size = (44, 44)
            
            pen_icon_path = os.path.join(assets_dir, "pen_icon.png")
            pen_img = Image.open(pen_icon_path).resize(icon_size, Image.Resampling.LANCZOS)
            self.pen_icon = ImageTk.PhotoImage(pen_img)
            
            text_icon_path = os.path.join(assets_dir, "text_icon.png")
            text_img = Image.open(text_icon_path).resize(icon_size, Image.Resampling.LANCZOS)
            self.text_icon = ImageTk.PhotoImage(text_img)
            
            undo_icon_path = os.path.join(assets_dir, "undo_icon.png")
            undo_img = Image.open(undo_icon_path).resize(icon_size, Image.Resampling.LANCZOS)
            self.undo_icon = ImageTk.PhotoImage(undo_img)
            
            self.pen_btn = tk.Button(tool_frame, image=self.pen_icon, 
                                     command=lambda: self.settlmd("pen"),
                                     bg='#334155', width=48, height=48, 
                                     relief=tk.FLAT, cursor='hand2')
            self.pen_btn.pack(side=tk.LEFT, padx=2)
            
            self.text_btn = tk.Button(tool_frame, image=self.text_icon, 
                                      command=lambda: self.settlmd("text"),
                                      bg='#334155', width=48, height=48, 
                                      relief=tk.FLAT, cursor='hand2')
            self.text_btn.pack(side=tk.LEFT, padx=2)
            
            self.undo_btn = tk.Button(tool_frame, image=self.undo_icon,
                                      command=self.undolast,
                                      bg='#334155', width=48, height=48, 
                                      relief=tk.FLAT, cursor='hand2')
            self.undo_btn.pack(side=tk.LEFT, padx=2)
            
        except Exception as e:
            print(f"Could not load tool icons: {e}")
            # Fallback to text buttons
            self.pen_btn = tk.Button(tool_frame, text="✏️ Pen", 
                     command=lambda: self.settlmd("pen"),
                     bg='#334155', fg='white', padx=10, pady=8,
                     font=('Segoe UI', 9, 'bold'), relief=tk.FLAT,
                     cursor='hand2')
            self.pen_btn.pack(side=tk.LEFT, padx=2)
            
            self.text_btn = tk.Button(tool_frame, text="🅰️ Text", 
                     command=lambda: self.settlmd("text"),
                     bg='#334155', fg='white', padx=10, pady=8,
                     font=('Segoe UI', 9, 'bold'), relief=tk.FLAT,
                     cursor='hand2')
            self.text_btn.pack(side=tk.LEFT, padx=2)
            
            tk.Button(tool_frame, text="↶ Undo",
                     command=self.undolast,
                     bg='#334155', fg='white', padx=10, pady=8,
                     font=('Segoe UI', 9, 'bold'), relief=tk.FLAT,
                     cursor='hand2').pack(side=tk.LEFT, padx=2)
        
        # Right section - Action buttons
        right_frame = tk.Frame(toolbar, bg='#1e293b')
        right_frame.pack(side=tk.RIGHT, padx=10, pady=10)
        
        tk.Button(right_frame, text="🏭 Production Mode", command=self.prodmode,
                 bg='#f59e0b', fg='white', padx=15, pady=10,
                 font=('Segoe UI', 9, 'bold'), relief=tk.FLAT, borderwidth=0,
                 cursor='hand2').pack(side=tk.LEFT, padx=3)
        
        tk.Button(right_frame, text="✅ Handback to Quality", command=self.compreworkhndbck,
                 bg='#10b981', fg='white', padx=15, pady=10,
                 font=('Segoe UI', 9, 'bold'), relief=tk.FLAT, borderwidth=0,
                 cursor='hand2').pack(side=tk.LEFT, padx=3)
        
        # Canvas with scrollbars
        canvas_frame = tk.Frame(self.root, bg='#f1f5f9')
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        v_scrollbar = tk.Scrollbar(canvas_frame, orient=tk.VERTICAL)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        h_scrollbar = tk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.canvas = tk.Canvas(canvas_frame, bg='#f8fafc',
                               yscrollcommand=v_scrollbar.set,
                               xscrollcommand=h_scrollbar.set,
                               highlightthickness=0)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        v_scrollbar.config(command=self.canvas.yview)
        h_scrollbar.config(command=self.canvas.xview)
        
        # Bind mouse events - CRITICAL FOR PEN AND TEXT TOOLS
        self.canvas.bind("<ButtonPress-1>", self.leftclick)
        self.canvas.bind("<B1-Motion>", self.leftdrag)
        self.canvas.bind("<ButtonRelease-1>", self.leftrls)
        self.canvas.bind("<Double-Button-1>", self.doubleclick)
        self.canvas.bind("<Double-Button-3>", self.doubleright)
        
        # Modern status bar
        status_bar = tk.Frame(self.root, bg='#334155', height=40)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        instructions_text = "Pen: Freehand | Text: Click to add | Esc: Deactivate | Ctrl+Z: Undo"
        tk.Label(status_bar, text=instructions_text, bg='#334155', fg='#e2e8f0',
                font=('Segoe UI', 9), pady=10).pack()
    
    def updtoolpane(self):
        """Placeholder for tool pane update - not needed in production mode"""
        pass

    # ================================================================
    # LOAD FROM HANDOVER QUEUE - WITH AUTO-OPEN PRODUCTION MODE
    # ================================================================
    
    def loadfrmhandover(self):
        """Load item from production handover queue"""
        pending_items = self.handover_db.get_pending_production_items()
        
        if not pending_items:
            messagebox.showinfo("No Items", 
                              "✓ No items in production queue.\n"
                              "All items have been processed!", 
                              icon='info')
            return
        
        # Create selection dialog
        dlg = tk.Toplevel(self.root)
        dlg.title("Production Queue")
        dlg.geometry("1000x600")
        dlg.configure(bg='#f8fafc')
        dlg.transient(self.root)
        dlg.grab_set()
        
        # Header
        header_frame = tk.Frame(dlg, bg='#8b5cf6', height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="Production Queue - Select Item",
                bg='#8b5cf6', fg='white',
                font=('Segoe UI', 14, 'bold')).pack(pady=15)
        
        # Info bar
        info_frame = tk.Frame(dlg, bg='#eff6ff')
        info_frame.pack(fill=tk.X, padx=20, pady=(15, 5))
        
        tk.Label(info_frame, text=f"Total items in queue: {len(pending_items)}",
                bg='#eff6ff', fg='#1e40af',
                font=('Segoe UI', 10, 'bold')).pack(pady=8)
        
        # Listbox frame
        list_frame = tk.Frame(dlg, bg='white')
        list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        tk.Label(list_frame, text="Select item to load:",
                font=('Segoe UI', 10, 'bold'), bg='white', fg='#1e293b').pack(anchor='w', pady=(0, 10))
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(list_frame, font=('Consolas', 9),
                            yscrollcommand=scrollbar.set,
                            bg='#f8fafc', relief=tk.FLAT,
                            selectmode=tk.SINGLE, height=18)
        listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        # Populate listbox
        for item in pending_items:
            status_icon = " " if item['status'] == 'in_progress' else "📦"
            display = (
                f"{status_icon} {item['cabinet_id']:20} | {item['project_name']:30} | "
                f"Punches: {item['open_punches']:3} | By: {item['handed_over_by']:15} | "
                f"{item['handed_over_date'][:10]}"
            )
            listbox.insert(tk.END, display)
        
        def load_selected():
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("No Selection", "Please select an item first.")
                return
            
            item = pending_items[selection[0]]
            dlg.destroy()
            self.loadhndovritm(item)
        
        # Buttons
        btn_frame = tk.Frame(dlg, bg='#f8fafc')
        btn_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
        
        btn_style = {
            'font': ('Segoe UI', 10, 'bold'),
            'relief': tk.FLAT,
            'cursor': 'hand2',
            'padx': 20,
            'pady': 12
        }
        
        tk.Button(btn_frame, text="Load Selected", command=load_selected,
                 bg='#3b82f6', fg='white', **btn_style).pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_frame, text="Cancel", command=dlg.destroy,
                 bg='#64748b', fg='white', **btn_style).pack(side=tk.RIGHT, padx=5)
        
        listbox.bind('<Double-Button-1>', lambda e: load_selected())
    
    def loadhndovritm(self, item):
        """Load a handover item - WITH AUTO-OPEN PRODUCTION MODE"""
        try:
            # Verify files exist
            if not os.path.exists(item['pdf_path']):
                messagebox.showerror("File Not Found", 
                                   f"PDF not found:\n{item['pdf_path']}")
                return
            
            if not os.path.exists(item['excel_path']):
                messagebox.showerror("File Not Found", 
                                   f"Excel not found:\n{item['excel_path']}")
                return
            
            # Get project from database
            project_data = self.db.get_project(item['cabinet_id'])
            if not project_data:
                messagebox.showerror("Error", "Project not found in database")
                return
            
            # Load PDF
            self.pdf_document = fitz.open(item['pdf_path'])
            self.current_pdf_path = item['pdf_path']
            self.current_page = 0
            self.zoom_level = 1.0
            
            # Set project details
            self.cabinet_id = item['cabinet_id']
            self.project_name = item['project_name']
            self.sales_order_no = item['sales_order_no']
            self.storage_location = project_data['storage_location']
            
            # Set Excel
            self.excel_file = item['excel_path']
            self.working_excel_path = item['excel_path']
            
            # Load session if available
            print(f"\n{'='*60}")
            print(f"Loading handover item: {self.cabinet_id}")
            print(f"PDF: {item['pdf_path']}")
            print(f"Excel: {item['excel_path']}")
            print(f"Session path from item: {item.get('session_path')}")
            
            if item.get('session_path') and os.path.exists(item['session_path']):
                print(f"✓ Session file exists, loading...")
                self.loadsessfrompath(item['session_path'])
                print(f"After loading: {len(self.annotations)} annotations loaded")
                
                # Debug: Show what's in annotations
                highlight_count = sum(1 for a in self.annotations if a.get('type') == 'highlight')
                error_count = sum(1 for a in self.annotations if a.get('type') == 'error')
                print(f"  Highlights: {highlight_count}, Errors: {error_count}")
                
                for i, ann in enumerate(self.annotations[:3]):  # First 3 only
                    print(f"  Annotation {i}: type={ann.get('type')}, "
                          f"page={ann.get('page')}, "
                          f"has_points_page={'points_page' in ann}, "
                          f"has_bbox_page={'bbox_page' in ann}, "
                          f"sr_no={ann.get('sr_no')}")
            else:
                print(f"⚠️ No session file found")
                self.annotations = []
                self.session_refs.clear()
            
            print(f"{'='*60}\n")
            
            # Mark as in progress
            username = self.logged_in_fullname or "Unknown User"
            
            self.handover_db.update_production_status(
                item['cabinet_id'],
                status='in_progress',
                user=username
            )
            
            # Update manager status
            self.manager_db.updstats(self.cabinet_id, 'in_progress')
            self.syncmgrstats()
            
            self.display()
            
            
            # AUTO-OPEN PRODUCTION MODE
            self.root.after(500, self.prodmode)
        
        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load item:\n{e}")
            import traceback
            traceback.print_exc()
    def closing(self):
        """Handle application closing with auto-save"""
        if self.pdf_document and hasattr(self, 'project_dirs'):
            try:
                print("\n🔄 Auto-saving before closing...")
                self.savesess()
                print("✓ Session auto-saved successfully")
                
                # Sync stats one last time
                self.syncmgrstatsonly()
                print("✓ Statistics synced")
                
            except Exception as e:
                print(f"⚠️ Auto-save on close failed: {e}")
                # Ask user if they want to close anyway
                proceed = messagebox.askyesno(
                    "Save Failed",
                    f"Failed to auto-save:\n{e}\n\nClose anyway?",
                    icon='warning'
                )
                if not proceed:
                    return  # Don't close the application
        
        # Close the application
        self.root.destroy()
    # ================================================================
    # COMPLETE REWORK & HANDBACK - CHECK IMPLEMENTED COLUMN
    # ================================================================
    
    def compreworkhndbck(self):
        """Complete rework and handback to Quality"""
        if not self.pdf_document or not self.excel_file:
            messagebox.showwarning("No Item Loaded", 
                                 "Please load an item from the production queue first.")
            return
        
        item = self.handover_db.get_item_by_cabinet_id(self.cabinet_id, "quality_to_production")
        if not item:
            messagebox.showwarning("Not from Queue", 
                                 "This item was not loaded from the handover queue.")
            return
        
        # Check for punches without implementation
        not_implemented = self.findnotimplemented()
        if not_implemented:
            self.shownotimplemented(not_implemented)
            return
        
        # AUTO-SAVE SESSION BEFORE HANDBACK
        print("Auto-saving session before handback...")
        try:
            self.savesess()
            print("✓ Session auto-saved successfully")
        except Exception as e:
            print(f"⚠️ Session auto-save failed: {e}")
            # Continue anyway - not critical
        remarks=None
        
        username = self.logged_in_fullname or "Unknown User"
        
        handback_data = {
            "cabinet_id": self.cabinet_id,
            "project_name": self.project_name,
            "sales_order_no": self.sales_order_no,
            "pdf_path": self.current_pdf_path,
            "excel_path": self.excel_file,
            "session_path": self.getsesspathforpdf(),
            "rework_completed_by": username,
            "rework_completed_date": datetime.now().isoformat(),
            "production_remarks": remarks or "No remarks"
        }
        
        success = self.handover_db.add_production_handback(handback_data)
        
        if success:
            self.syncmgrstats()
            self.manager_db.updstats(self.cabinet_id, 'being_closed_by_quality')
            
            messagebox.showinfo(
                "Handback Complete",
                f"✓ Successfully handed back to Quality:\n\n"
                f"Cabinet: {self.cabinet_id}\n"
                f"Project: {self.project_name}\n\n"
                f"Session auto-saved\n"
                f"Quality team will verify and close this item.",
                icon='info'
            )
            
            # Clear current work
            self.pdf_document = None
            self.current_pdf_path = None
            self.excel_file = None
            self.annotations = []
            self.canvas.delete("all")
            self.page_label.config(text="Page: 0/0")
            self.root.title("Production Tool - Highlighter Mode")
        else:
            messagebox.showerror("Error", "Failed to handback item to Quality.")
    
    def findnotimplemented(self):
        """Get list of punches without 'Implemented By'"""
        not_implemented = []
        
        try:
            if not self.excel_file or not os.path.exists(self.excel_file):
                return not_implemented
            
            wb = load_workbook(self.excel_file, data_only=True)
            ws = wb[self.punch_sheet_name] if self.punch_sheet_name in wb.sheetnames else wb.active
            
            row = 9
            while row <= ws.max_row + 5:
                checked = self.read_cell(ws, row, 'E')
                if not checked:
                    row += 1
                    if row > 2000:
                        break
                    continue
                
                closed = self.read_cell(ws, row, self.punch_cols['closed_name'])
                if closed:
                    row += 1
                    continue
                
                implemented = self.read_cell(ws, row, self.punch_cols['implemented_name'])
                if not implemented:
                    sr_no = self.read_cell(ws, row, self.punch_cols['sr_no'])
                    ref_no = self.read_cell(ws, row, self.punch_cols['ref_no'])
                    desc = self.read_cell(ws, row, self.punch_cols['desc'])
                    category = self.read_cell(ws, row, self.punch_cols['category'])
                    
                    not_implemented.append({
                        'row': row,
                        'sr_no': sr_no,
                        'ref_no': ref_no,
                        'description': desc,
                        'category': category
                    })
                
                row += 1
                if row > 2000:
                    break
            
            wb.close()
            return not_implemented
        
        except Exception as e:
            print(f"Error checking implementation: {e}")
            return []
    
    def shownotimplemented(self, not_implemented):
        """Show dialog listing punches without implementation"""
        dlg = tk.Toplevel(self.root)
        dlg.title("⚠️ Implementation Required")
        dlg.geometry("800x600")
        dlg.configure(bg='#fef3c7')
        dlg.transient(self.root)
        dlg.grab_set()
        
        header_frame = tk.Frame(dlg, bg='#f59e0b', height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="⚠️ IMPLEMENTATION REQUIRED",
                bg='#f59e0b', fg='white',
                font=('Segoe UI', 14, 'bold')).pack(pady=15)
        
        info_frame = tk.Frame(dlg, bg='#fef3c7')
        info_frame.pack(fill=tk.X, padx=20, pady=15)
        
        tk.Label(info_frame, 
                text=f"The following {len(not_implemented)} punch(es) have not been marked as 'Implemented'.\n"
                     "Please complete implementation before handing back to Quality.",
                font=('Segoe UI', 11), bg='#fef3c7', fg='#78350f',
                justify='left').pack(anchor='w')
        
        list_frame = tk.Frame(dlg, bg='white')
        list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        tk.Label(list_frame, text="Punches requiring implementation:",
                font=('Segoe UI', 10, 'bold'), bg='white', fg='#1e293b').pack(anchor='w', padx=10, pady=(10, 5))
        
        scroll_frame = tk.Frame(list_frame, bg='white')
        scroll_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        scrollbar = tk.Scrollbar(scroll_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        text_widget = tk.Text(scroll_frame, wrap=tk.WORD, font=('Courier New', 9),
                            yscrollcommand=scrollbar.set, bg='#f8fafc', relief=tk.FLAT,
                            padx=10, pady=10)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=text_widget.yview)
        
        for idx, punch in enumerate(not_implemented, 1):
            text_widget.insert(tk.END, f"\n{'='*70}\n")
            text_widget.insert(tk.END, f"#{idx} - SR No: {punch['sr_no']} | Ref: {punch['ref_no']}\n")
            text_widget.insert(tk.END, f"Category: {punch['category']}\n")
            text_widget.insert(tk.END, f"\nDescription:\n{punch['description']}\n")
        
        text_widget.config(state=tk.DISABLED)
        
        tk.Button(dlg, text="OK - I'll Complete Implementation First",
                 command=dlg.destroy, bg='#f59e0b', fg='white',
                 font=('Segoe UI', 10, 'bold'), padx=20, pady=12,
                 relief=tk.FLAT, cursor='hand2').pack(pady=20)

    # ================================================================
    # ENHANCED PRODUCTION MODE WITH HIGHLIGHTER NAVIGATION
    # ================================================================
    
    def prodmode(self):
        """Production mode with highlighter navigation"""
        if not self.pdf_document or not self.excel_file:
            messagebox.showwarning("No Item", 
                                 "Please load an item from the production queue first.")
            return
        
        punches = self.openpunches()
        
        if not punches:
            messagebox.showinfo("No Punches", 
                              "✓ All punches are closed!\n"
                              "You can now handback to Quality.", 
                              icon='info')
            return
        
        punches.sort(key=lambda p: (p['implemented'], p['sr_no']))
        
        dlg = tk.Toplevel(self.root)
        dlg.title("Production Mode - Highlighter")
        dlg.geometry("900x550")
        dlg.configure(bg='#f8fafc')
        dlg.transient(self.root)
        dlg.grab_set()
        
        self.production_dialog_open = True
        
        header_frame = tk.Frame(dlg, bg='#f59e0b', height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="🖍️ PRODUCTION MODE - HIGHLIGHTER",
                bg='#f59e0b', fg='white',
                font=('Segoe UI', 14, 'bold')).pack(pady=15)
        
        progress_frame = tk.Frame(dlg, bg='#f8fafc')
        progress_frame.pack(fill=tk.X, padx=20, pady=(15, 5))
        
        idx_label = tk.Label(progress_frame, text="",
                           font=('Segoe UI', 11, 'bold'),
                           bg='#f8fafc', fg='#1e293b')
        idx_label.pack()
        
        info_frame = tk.Frame(dlg, bg='#f8fafc')
        info_frame.pack(fill=tk.X, padx=20, pady=10)
        
        sr_card = tk.Frame(info_frame, bg='#dbeafe', relief=tk.FLAT)
        sr_card.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        tk.Label(sr_card, text="SR No.", font=('Segoe UI', 8),
                bg='#dbeafe', fg='#1e40af').pack(anchor='w', padx=10, pady=(8, 2))
        
        sr_label = tk.Label(sr_card, text="", font=('Segoe UI', 12, 'bold'),
                          bg='#dbeafe', fg='#1e293b')
        sr_label.pack(anchor='w', padx=10, pady=(0, 8))
        
        ref_card = tk.Frame(info_frame, bg='#e0e7ff', relief=tk.FLAT)
        ref_card.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        tk.Label(ref_card, text="Reference", font=('Segoe UI', 8),
                bg='#e0e7ff', fg='#4338ca').pack(anchor='w', padx=10, pady=(8, 2))
        
        ref_label = tk.Label(ref_card, text="", font=('Segoe UI', 12, 'bold'),
                           bg='#e0e7ff', fg='#1e293b')
        ref_label.pack(anchor='w', padx=10, pady=(0, 8))
        
        status_card = tk.Frame(info_frame, bg='#fef3c7', relief=tk.FLAT)
        status_card.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        
        tk.Label(status_card, text="Status", font=('Segoe UI', 8),
                bg='#fef3c7', fg='#92400e').pack(anchor='w', padx=10, pady=(8, 2))
        
        impl_label = tk.Label(status_card, text="", font=('Segoe UI', 12, 'bold'),
                            bg='#fef3c7', fg='#1e293b')
        impl_label.pack(anchor='w', padx=10, pady=(0, 8))
        
        content_frame = tk.Frame(dlg, bg='white', relief=tk.FLAT)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        tk.Label(content_frame, text="Punch Description:",
                font=('Segoe UI', 9, 'bold'), bg='white', fg='#64748b',
                anchor='w').pack(fill=tk.X, padx=15, pady=(10, 5))
        
        text_widget = tk.Text(content_frame, wrap=tk.WORD, height=12,
                            font=('Segoe UI', 10), bg='#f8fafc', relief=tk.FLAT,
                            padx=10, pady=10)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
        text_widget.config(state=tk.DISABLED)
        
        pos = [0]
        
        def show_item():
            p = punches[pos[0]]
            
            progress_text = f"Item {pos[0]+1} of {len(punches)}"
            progress_pct = f"({int((pos[0]+1)/len(punches)*100)}% complete)"
            idx_label.config(text=f"{progress_text} {progress_pct}")
            
            sr_label.config(text=str(p['sr_no']))
            ref_label.config(text=str(p['ref_no']))
            
            impl_status = "✓ Implemented" if p['implemented'] else "⚠ Not Implemented"
            impl_color = '#10b981' if p['implemented'] else '#f59e0b'
            impl_label.config(text=impl_status, fg=impl_color)
            
            text_widget.config(state=tk.NORMAL)
            text_widget.delete("1.0", tk.END)
            text_widget.insert(tk.END, p['punch_text'])
            text_widget.insert(tk.END, f"\n\n───────────────────────\n")
            text_widget.insert(tk.END, f"Category: {p['category']}\n")
            text_widget.insert(tk.END, f"Implementation: {'YES' if p['implemented'] else 'NO'}\n")
            
            # Find annotation - checks for both SR number and excel row
            ann = next((a for a in self.annotations 
                       if a.get('sr_no') == p['sr_no'] 
                       or a.get('excel_row') == p['row']), None)
            
            # Display quality remarks from quality team
            if ann and ann.get('quality_remark'):
                text_widget.insert(tk.END, f"\n───────────────────────\n")
                text_widget.insert(tk.END, "📋 Quality Remarks:\n")
                text_widget.insert(tk.END, ann['quality_remark'])
            
            # Display previous implementation remarks
            if ann and ann.get('implementation_remark'):
                text_widget.insert(tk.END, f"\n───────────────────────\n")
                text_widget.insert(tk.END, "Previous Implementation Remarks:\n")
                text_widget.insert(tk.END, ann['implementation_remark'])
            
            text_widget.config(state=tk.DISABLED)
            
            self.navtopunch(p['sr_no'], p['punch_text'])
        
        show_item()
        
        def mark_implemented():
            p = punches[pos[0]]
            
            default_user = self.logged_in_fullname or "Unknown User"
            
            name = default_user
            if not name:
                return
            
            remark = simpledialog.askstring("Remarks (optional)",
                                          "Add remarks about the implementation (optional):",
                                          parent=dlg)
            
            try:
                wb = load_workbook(self.excel_file)
                ws = wb[self.punch_sheet_name]
                
                self.write_cell(ws, p['row'], self.punch_cols['implemented_name'], name)
                # Updated to include timestamp + date
                self.write_cell(ws, p['row'], self.punch_cols['implemented_date'],
                              datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                
                wb.save(self.excel_file)
                wb.close()
                
                self.syncmgrstats()
            
            except PermissionError:
                messagebox.showerror("File Locked",
                                   "⚠️ Please close the Excel file and try again.",
                                   parent=dlg)
                return
            except Exception as e:
                messagebox.showerror("Excel Error", str(e), parent=dlg)
                return
            
            # Find annotation and update implementation status
            ann = next((a for a in self.annotations 
                       if a.get('sr_no') == p['sr_no'] 
                       or a.get('excel_row') == p['row']), None)
            
            if ann:
                ann['implemented'] = True
                ann['implemented_name'] = name
                ann['implemented_date'] = datetime.now().isoformat()
                if remark:
                    ann['implementation_remark'] = remark
            
            if pos[0] < len(punches) - 1:
                pos[0] += 1
                show_item()
            else:
                messagebox.showinfo("Complete",
                                  "✓ All punches reviewed!\n"
                                  "You can now handback to Quality.",
                                  icon='info', parent=dlg)
                self.clrborderhighlight()
                self.production_dialog_open = False
                dlg.destroy()
        
        def next_item():
            if pos[0] < len(punches) - 1:
                pos[0] += 1
                show_item()
        
        def prev_item():
            if pos[0] > 0:
                pos[0] -= 1
                show_item()
        
        def on_close():
            self.clrborderhighlight()
            self.production_dialog_open = False
            dlg.destroy()
        
        dlg.protocol("WM_DELETE_WINDOW", on_close)
        
        btn_frame = tk.Frame(dlg, bg='#f8fafc')
        btn_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
        
        btn_style = {
            'font': ('Segoe UI', 10, 'bold'),
            'relief': tk.FLAT,
            'borderwidth': 0,
            'cursor': 'hand2',
            'padx': 20,
            'pady': 12
        }
        
        tk.Button(btn_frame, text="◀ Previous", command=prev_item,
                 bg='#94a3b8', fg='white', width=12, **btn_style).pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_frame, text="✓ MARK DONE", command=mark_implemented,
                 bg='#10b981', fg='white', width=16, **btn_style).pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_frame, text="Next ▶", command=next_item,
                 bg='#94a3b8', fg='white', width=12, **btn_style).pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_frame, text="Close", command=on_close,
                 bg='#64748b', fg='white', width=10, **btn_style).pack(side=tk.RIGHT, padx=5)
    
    def navtopunch(self, sr_no, punch_text):
        """Navigate to highlighter annotation and highlight it - UPDATED FOR HIGHLIGHTER"""
        target_ann = None
        
        # Try SR No match - looking for 'error' type annotations (which are highlighter marks)
        for ann in self.annotations:
            if ann.get('sr_no') == sr_no and ann.get('type') in ('error', 'highlight'):
                target_ann = ann
                print(f"✓ Found annotation by SR No: {sr_no}, type: {ann.get('type')}")
                break
        
        # Fuzzy text match if no direct SR match
        if not target_ann:
            best_match = None
            best_score = 0
            
            for ann in self.annotations:
                if ann.get('type') in ('error', 'highlight') and ann.get('punch_text'):
                    ann_text = str(ann['punch_text']).lower()
                    search_text = str(punch_text).lower()
                    
                    if search_text in ann_text or ann_text in search_text:
                        score = len(set(search_text.split()) & set(ann_text.split()))
                        if score > best_score:
                            best_score = score
                            best_match = ann
            
            if best_match:
                target_ann = best_match
                print(f"✓ Found annotation by text match, SR: {best_match.get('sr_no')}")
        
        self.clrborderhighlight()
        
        if target_ann:
            print(f"Navigating to annotation:")
            print(f"  Type: {target_ann.get('type')}")
            print(f"  SR No: {target_ann.get('sr_no')}")
            print(f"  Has points_page: {'points_page' in target_ann}")
            print(f"  Has bbox_page: {'bbox_page' in target_ann}")
            
            if target_ann.get('page') is not None:
                self.current_page = target_ann['page']
                self.display()
            
            # Highlight the annotation visually
            if 'points_page' in target_ann or 'bbox_page' in target_ann:
                self.highlightannonvisual(target_ann)
                self._last_highlighted_ann = target_ann
        else:
            print(f"⚠️ No annotation found for SR {sr_no}")
            print(f"Available annotation types: {set(a.get('type') for a in self.annotations)}")
            print(f"Available SR numbers: {set(a.get('sr_no') for a in self.annotations if a.get('sr_no'))}")
    
    def highlightannonvisual(self, annotation):
        """Draw visual indicators for highlighter annotation - UPDATED"""
        # Calculate bounding box from points_page or use bbox_page
        if 'points_page' in annotation and annotation['points_page']:
            # Calculate bbox from highlighter points
            points = annotation['points_page']
            xs = [p[0] for p in points]
            ys = [p[1] for p in points]
            bbox_page = (min(xs), min(ys), max(xs), max(ys))
            bbox_display = self.bbox_page_to_display(bbox_page)
            print(f"  Using points_page to calculate bbox: {bbox_page}")
        elif 'bbox_page' in annotation:
            bbox_display = self.bbox_page_to_display(annotation['bbox_page'])
            print(f"  Using bbox_page: {annotation['bbox_page']}")
        else:
            print("⚠️ Annotation has no points_page or bbox_page - cannot highlight")
            return
        
        x1, y1, x2, y2 = bbox_display
        
        # Calculate center
        cx = (x1 + x2) / 2
        cy = (y1 + y2) / 2
        
        padding = 15
        
        # Glow layers - pulsing effect
        for i in range(3):
            glow_padding = padding + (i * 5)
            
            glow_id = self.canvas.create_rectangle(
                x1 - glow_padding, y1 - glow_padding,
                x2 + glow_padding, y2 + glow_padding,
                outline='#fbbf24', width=2, dash=(8, 4),
                tags='production_highlight'
            )
            self.production_highlight_tags.append(glow_id)
        
        # Main highlight border - bright orange
        main_id = self.canvas.create_rectangle(
            x1 - padding, y1 - padding,
            x2 + padding, y2 + padding,
            outline='#f59e0b', width=4, dash=(10, 5),
            tags='production_highlight'
        )
        self.production_highlight_tags.append(main_id)
        
        # Arrow pointing to the annotation
        arrow_start_x = cx - 120
        arrow_start_y = cy - 120
        
        # Arrow background (shadow)
        arrow_bg = self.canvas.create_line(
            arrow_start_x, arrow_start_y,
            cx - 20, cy - 20,
            arrow=tk.LAST, fill='#fbbf24', width=6,
            tags='production_highlight'
        )
        self.production_highlight_tags.append(arrow_bg)
        
        # Arrow foreground
        arrow_fg = self.canvas.create_line(
            arrow_start_x, arrow_start_y,
            cx - 20, cy - 20,
            arrow=tk.LAST, fill='#f59e0b', width=3,
            tags='production_highlight'
        )
        self.production_highlight_tags.append(arrow_fg)
        
        # Label background
        label_bg = self.canvas.create_rectangle(
            arrow_start_x - 60, arrow_start_y - 35,
            arrow_start_x + 10, arrow_start_y - 5,
            fill='#fef3c7', outline='#f59e0b', width=2,
            tags='production_highlight'
        )
        self.production_highlight_tags.append(label_bg)
        
        # Label text
        label_text = f"SR {annotation.get('sr_no', '?')}"
        label_txt = self.canvas.create_text(
            arrow_start_x - 25, arrow_start_y - 20,
            text=label_text,
            fill='#92400e',
            font=('Segoe UI', 12, 'bold'),
            tags='production_highlight'
        )
        self.production_highlight_tags.append(label_txt)
        
        # Scroll to make visible
        bbox_all = self.canvas.bbox("all")
        if bbox_all:
            self.canvas.yview_moveto(max(0, (y1 - 150) / max(1, bbox_all[3])))
            self.canvas.xview_moveto(max(0, (x1 - 150) / max(1, bbox_all[2])))
        
        print(f"✓ Visual highlight added at display coords: {bbox_display}")
    
    def clrborderhighlight(self):
        """Clear production mode visual indicators"""
        self.canvas.delete('production_highlight')
        self.production_highlight_tags.clear()
    
    def openpunches(self):
        """Read open punches from Excel - row 9 onwards"""
        punches = []
        
        if not self.excel_file or not os.path.exists(self.excel_file):
            return punches
        
        wb = load_workbook(self.excel_file, data_only=True)
        ws = wb[self.punch_sheet_name] if self.punch_sheet_name in wb.sheetnames else wb.active
        
        row = 9
        while True:
            checked = self.read_cell(ws, row, 'E')
            if not checked:
                row += 1
                if row > 2000:
                    break
                continue
            
            closed = self.read_cell(ws, row, self.punch_cols['closed_name'])
            if closed:
                row += 1
                continue
            
            implemented = bool(self.read_cell(ws, row, self.punch_cols['implemented_name']))
            sr = self.read_cell(ws, row, self.punch_cols['sr_no'])
            
            punches.append({
                'sr_no': sr,
                'row': row,
                'ref_no': self.read_cell(ws, row, self.punch_cols['ref_no']),
                'punch_text': self.read_cell(ws, row, self.punch_cols['desc']),
                'category': self.read_cell(ws, row, self.punch_cols['category']),
                'implemented': implemented
            })
            
            row += 1
            if row > 2000:
                break
        
        wb.close()
        return punches

    # ================================================================
    # TOOL MODES - PEN, TEXT, UNDO
    # ================================================================
    
    def settlmd(self, mode):
        """Set tool mode (pen or text)"""
        # Deactivate highlighter if active (not applicable in production tool, but kept for consistency)
        if hasattr(self, 'active_highlighter') and self.active_highlighter:
            self.active_highlighter = None
        
        # Toggle tool mode
        if self.tool_mode == mode:
            self.tool_mode = None
            if mode == "pen":
                self.pen_btn.config(bg='#334155', relief=tk.FLAT)
            else:
                self.text_btn.config(bg='#334155', relief=tk.FLAT)
        else:
            self.tool_mode = mode
            if mode == "pen":
                self.pen_btn.config(bg='#3b82f6', relief=tk.SUNKEN)
                self.text_btn.config(bg='#334155', relief=tk.FLAT)
            else:
                self.text_btn.config(bg='#3b82f6', relief=tk.SUNKEN)
                self.pen_btn.config(bg='#334155', relief=tk.FLAT)
        
        print(f"Tool mode: {self.tool_mode}")
    
    def deactivate_all(self):
        """Deactivate all tools"""
        if self.tool_mode:
            self.settlmd(self.tool_mode)
        
        self.drawing = False
        self.drawing_type = None
        self.pen_points = []
        self.temp_line_ids = []
        self.display()
    
    def updtoolpane(self):
        """Update annotation statistics - placeholder"""
        pass
    
    def _flash_status(self, message, bg='#10b981'):
        """Show a temporary status message"""
        status_label = tk.Label(
            self.root, 
            text=message, 
            bg=bg, 
            fg='white', 
            font=('Segoe UI', 10, 'bold'),
            padx=25, 
            pady=12,
            relief=tk.FLAT
        )
        status_label.place(relx=0.5, rely=0.08, anchor='center')
        self.root.after(1500, lambda: status_label.destroy())
    
    def clear_temp_drawings(self):
        """Clear temporary drawing elements from canvas"""
        for line_id in self.temp_line_ids:
            try:
                self.canvas.delete(line_id)
            except:
                pass
        self.temp_line_ids.clear()
    
    # ================================================================
    # UNDO FUNCTIONALITY
    # ================================================================
    
    def addtoundostck(self, action_type, annotation):
        """Add an action to the undo stack"""
        self.undo_stack.append({
            'type': action_type,
            'annotation': annotation.copy()
        })
        
        if len(self.undo_stack) > self.max_undo:
            self.undo_stack.pop(0)
    
    def undolast(self):
        """Undo the last annotation action"""
        if not self.undo_stack:
            messagebox.showinfo("Nothing to Undo", "No actions to undo.", icon='info')
            return
        
        last_action = self.undo_stack.pop()
        
        if last_action['type'] == 'add_annotation':
            annotation = last_action['annotation']
            if annotation in self.annotations:
                self.annotations.remove(annotation)
                self.display()
                self._flash_status("✓ Annotation removed", bg='#10b981')
        
        self.updtoolpane()
    
    # ================================================================
    # MOUSE EVENT HANDLERS - PEN AND TEXT
    # ================================================================
    
    def leftclick(self, event):
        """Handle left mouse button press"""
        if not self.pdf_document:
            messagebox.showwarning("Warning", "Please load a PDF first")
            return

        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)

        # -------- PEN TOOL --------
        if self.tool_mode == "pen":
            self.drawing = True
            self.drawing_type = "pen"
            self.pen_points = [(x, y)]
            self.clear_temp_drawings()
            return

        # -------- TEXT TOOL --------
        if self.tool_mode == "text":
            self.drawing = True
            self.drawing_type = "text"
            self.text_pos_x = x
            self.text_pos_y = y
            return
    
    def leftdrag(self, event):
        """Handle left mouse button drag"""
        if not self.drawing:
            return

        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)

        # -------- PEN TOOL DRAWING --------
        if self.drawing_type == "pen":
            if len(self.pen_points) > 0:
                last_x, last_y = self.pen_points[-1]
                line_id = self.canvas.create_line(
                    last_x, last_y, x, y,
                    fill="red", width=3,
                    capstyle=tk.ROUND, smooth=True
                )
                self.temp_line_ids.append(line_id)
            self.pen_points.append((x, y))
            return
    
    def leftrls(self, event):
        """Handle left mouse button release"""
        if not self.pdf_document or not self.drawing:
            return

        # -------- PEN TOOL FINISH --------
        if self.drawing_type == "pen":
            if len(self.pen_points) >= 2:
                points_page = self.display_to_page_coords(self.pen_points)
                annotation = {
                    'type': 'pen',
                    'page': self.current_page,
                    'points': points_page,
                    'timestamp': datetime.now().isoformat()
                }
                self.annotations.append(annotation)
                self.addtoundostck('add_annotation', annotation)
            self.pen_points = []
            self.clear_temp_drawings()
            self.drawing = False
            self.drawing_type = None
            self.display()
            self.updtoolpane()
            self._flash_status("✓ Pen stroke added", bg='#10b981')
            return

        # -------- TEXT TOOL FINISH --------
        if self.drawing_type == "text":
            txt = simpledialog.askstring("Text", "Enter text:", parent=self.root)
            if txt and txt.strip():
                pos_page = self.display_to_page_coords((self.text_pos_x, self.text_pos_y))
                annotation = {
                    'type': 'text',
                    'page': self.current_page,
                    'pos_page': pos_page,
                    'text': txt.strip(),
                    'timestamp': datetime.now().isoformat()
                }
                self.annotations.append(annotation)
                self.addtoundostck('add_annotation', annotation)
                self.display()
                self._flash_status("✓ Text added", bg='#10b981')
            self.drawing = False
            self.drawing_type = None
            self.updtoolpane()
            return

    # ================================================================
    # DISPLAY PAGE - HIGHLIGHTER RENDERING ONLY (NO BOXES)
    # ================================================================
    
    def display(self):
        """Render the current PDF page with HIGHLIGHTER annotations ONLY - NO BOXES"""
        if not self.pdf_document:
            self.canvas.delete("all")
            self.page_label.config(text="Page: 0/0")
            return

        try:
            page = self.pdf_document[self.current_page]
            mat = fitz.Matrix(self.page_to_display_scale(), self.page_to_display_scale())
            pix = page.get_pixmap(matrix=mat)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            self.current_page_image = np.array(img)
            draw = ImageDraw.Draw(img, 'RGBA')

            # Try to load a font for text
            try:
                font_size = max(12, int(14 * self.zoom_level))
                font = ImageFont.truetype("arial.ttf", font_size)
            except:
                font = ImageFont.load_default()

            # Count annotations by type for debugging
            page_annotations = [ann for ann in self.annotations if ann.get('page') == self.current_page]
            print(f"\n=== Rendering Page {self.current_page + 1} ===")
            print(f"Total annotations on this page: {len(page_annotations)}")
            
            highlight_count = 0
            error_count = 0
            pen_count = 0
            text_count = 0
            box_count = 0

            for ann in self.annotations:
                if ann.get('page') != self.current_page:
                    continue

                ann_type = ann.get('type')

                # -------- HIGHLIGHTER STROKES (type='highlight' or type='error') --------
                if ann_type in ('highlight', 'error') and 'points_page' in ann:
                    points_page = ann['points_page']
                    if len(points_page) >= 2:
                        points_display = self.page_to_display_coords(points_page)
                        color_key = ann.get('color', 'yellow')
                        rgba = self.highlighter_colors.get(color_key, self.highlighter_colors['yellow'])['rgba']
                        
                        # Draw thick semi-transparent strokes
                        stroke_width = max(15, int(15 * self.zoom_level))
                        for i in range(len(points_display) - 1):
                            x1, y1 = points_display[i]
                            x2, y2 = points_display[i + 1]
                            draw.line([x1, y1, x2, y2], fill=rgba, width=stroke_width)
                        
                        # Add closed indicator if applicable
                        if ann.get('closed_by'):
                            # Calculate bbox if not present
                            if 'bbox_page' in ann:
                                bbox_display = self.bbox_page_to_display(ann['bbox_page'])
                            else:
                                xs = [p[0] for p in points_page]
                                ys = [p[1] for p in points_page]
                                bbox_page = (min(xs), min(ys), max(xs), max(ys))
                                bbox_display = self.bbox_page_to_display(bbox_page)
                            
                            cx = bbox_display[0] + 8
                            cy = bbox_display[1] + 8
                            draw.ellipse([cx - 6, cy - 6, cx + 6, cy + 6], fill=(0, 128, 0, 200))
                        
                        if ann_type == 'highlight':
                            highlight_count += 1
                        else:
                            error_count += 1

                # -------- PEN STROKES --------
                elif ann_type == 'pen' and 'points' in ann:
                    points_page = ann['points']
                    if len(points_page) >= 2:
                        points_display = self.page_to_display_coords(points_page)
                        stroke_width = max(2, int(3 * self.zoom_level))
                        for i in range(len(points_display) - 1):
                            x1, y1 = points_display[i]
                            x2, y2 = points_display[i + 1]
                            draw.line([x1, y1, x2, y2], fill='red', width=stroke_width)
                        pen_count += 1

                # -------- TEXT ANNOTATIONS --------
                elif ann_type == 'text' and 'pos_page' in ann:
                    pos_page = ann['pos_page']
                    pos_display = self.page_to_display_coords(pos_page)
                    text = ann.get('text', '')
                    if text:
                        # Draw text background for visibility
                        try:
                            bbox = draw.textbbox(pos_display, text, font=font)
                            padding = 2
                            draw.rectangle(
                                [bbox[0] - padding, bbox[1] - padding,
                                 bbox[2] + padding, bbox[3] + padding],
                                fill=(255, 255, 200, 200)
                            )
                        except:
                            pass
                        draw.text(pos_display, text, fill='red', font=font)
                        text_count += 1
                
                # -------- BOX ANNOTATIONS - REMOVED (counting for debugging only) --------
                elif ann_type == 'box':
                    box_count += 1
                    print(f"  ⚠️ Skipping box annotation (boxes are disabled)")

            print(f"Rendered annotations:")
            print(f"  🖍️ Highlights: {highlight_count}")
            print(f"  ❌ Errors (as highlights): {error_count}")
            print(f"  ✏️ Pen strokes: {pen_count}")
            print(f"  🅰️ Text: {text_count}")
            if box_count > 0:
                print(f"  📦 Boxes (skipped): {box_count}")
            print(f"{'='*40}\n")

            self.photo = ImageTk.PhotoImage(img)
            self.canvas.delete("all")
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)
            self.canvas.config(scrollregion=self.canvas.bbox(tk.ALL))
            self.page_label.config(text=f"Page: {self.current_page + 1}/{len(self.pdf_document)}")
            self.syncmgrstatsonly()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to display page: {e}")
            import traceback
            traceback.print_exc()

    # ================================================================
    # COORDINATE CONVERSION HELPERS
    # ================================================================
    
    def getnextsr(self):
        try:
            if not self.excel_file or not os.path.exists(self.excel_file):
                return 1
            
            wb = load_workbook(self.excel_file, read_only=True)
            ws = wb[self.punch_sheet_name] if self.punch_sheet_name in wb.sheetnames else wb.active
            
            last_sr_no = 0
            row_num = 9
            
            while row_num <= ws.max_row + 5:
                val = self.read_cell(ws, row_num, self.punch_cols['sr_no'])
                if val is None:
                    break
                try:
                    last_sr_no = int(val)
                except:
                    pass
                row_num += 1
            
            wb.close()
            return last_sr_no + 1
        except Exception:
            return 1
    
    def page_to_display_scale(self):
        return 2.0 * self.zoom_level
    
    def display_to_page_coords(self, pts):
        """Convert display-space coordinates to page-space coordinates."""
        scale = self.page_to_display_scale()
        
        # Handle single point tuple
        if isinstance(pts, tuple) and len(pts) == 2:
            if not isinstance(pts[0], (list, tuple)):
                return (pts[0] / scale, pts[1] / scale)
        
        # Handle list of points
        return [(x / scale, y / scale) for x, y in pts]
    
    def page_to_display_coords(self, pts):
        """Convert page coords to display coords"""
        scale = self.page_to_display_scale()
        
        # Handle single point tuple
        if isinstance(pts, tuple) and len(pts) == 2:
            if not isinstance(pts[0], (list, tuple)):
                return (pts[0] * scale, pts[1] * scale)
        
        # Handle list of points
        return [(x * scale, y * scale) for x, y in pts]
    
    def bbox_page_to_display(self, bbox_page):
        scale = self.page_to_display_scale()
        x1, y1, x2, y2 = bbox_page
        return (x1 * scale, y1 * scale, x2 * scale, y2 * scale)
    
    def bbox_display_to_page(self, bbox_display):
        scale = self.page_to_display_scale()
        x1, y1, x2, y2 = bbox_display
        return (x1 / scale, y1 / scale, x2 / scale, y2 / scale)
    
    # ================================================================
    # ROTATION TRANSFORMATION METHODS FOR PDF EXPORT
    # ================================================================
    
    def transform_bbox_for_rotation(self, rect, page):
        """Transform bbox for page rotation (for rectangle annotations)"""
        r = page.rotation
        w = page.rect.width
        h = page.rect.height
        x1, y1, x2, y2 = rect

        if r == 0:
            return fitz.Rect(x1, y1, x2, y2)
        if r == 90:
            return fitz.Rect(y1, w - x2, y2, w - x1)
        if r == 180:
            return fitz.Rect(w - x2, h - y2, w - x1, h - y1)
        if r == 270:
            return fitz.Rect(h - y2, x1, h - y1, x2)

        return fitz.Rect(x1, y1, x2, y2)

    def transform_point_for_rotation(self, point, page):
        """Transform a single point (x, y) for page rotation
        
        Used for:
        - Pen stroke points
        - Text annotation positions
        """
        r = page.rotation
        w = page.rect.width
        h = page.rect.height
        x, y = point

        if r == 0:
            return fitz.Point(x, y)
        elif r == 90:
            return fitz.Point(y, w - x)
        elif r == 180:
            return fitz.Point(w - x, h - y)
        elif r == 270:
            return fitz.Point(h - y, x)
        
        return fitz.Point(x, y)

    def transform_highlight_points_for_rotation(self, points, page):
        """Transform highlighter stroke points for page rotation
        
        Highlighters store a list of (x, y) tuples representing the stroke path.
        Each point needs to be individually transformed based on page rotation.
        
        Args:
            points: List of (x, y) tuples representing the highlight stroke
            page: PyMuPDF page object with rotation info
            
        Returns:
            List of fitz.Point objects, transformed for the page rotation
        """
        r = page.rotation
        w = page.rect.width
        h = page.rect.height
        
        transformed_points = []
        
        for point in points:
            x, y = point
            
            if r == 0:
                transformed_points.append(fitz.Point(x, y))
            elif r == 90:
                transformed_points.append(fitz.Point(y, w - x))
            elif r == 180:
                transformed_points.append(fitz.Point(w - x, h - y))
            elif r == 270:
                transformed_points.append(fitz.Point(h - y, x))
            else:
                transformed_points.append(fitz.Point(x, y))
        
        return transformed_points
    
    def zoomin(self, canvas_x, canvas_y, zoom_delta):
        if not self.pdf_document:
            return
        
        old_zoom = self.zoom_level
        new_zoom = max(0.5, min(3.0, old_zoom + zoom_delta))
        
        if new_zoom == old_zoom:
            return
        
        self.zoom_level = new_zoom
        self.display()
        
        scale = new_zoom / old_zoom
        bbox = self.canvas.bbox("all")
        if not bbox:
            return
        
        self.canvas.xview_moveto((canvas_x * scale) / max(1, bbox[2]))
        self.canvas.yview_moveto((canvas_y * scale) / max(1, bbox[3]))
    
    def doubleclick(self, event):
        self.drawing = False
        self.temp_highlight_id = None
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        self.zoomin(x, y, +0.25)
    
    def doubleright(self, event):
        self.drawing = False
        self.temp_highlight_id = None
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        self.zoomin(x, y, -0.25)
    
    def prev(self):
        if self.pdf_document and self.current_page > 0:
            self.current_page -= 1
            self.display()
    
    def next(self):
        if self.pdf_document and self.current_page < len(self.pdf_document) - 1:
            self.current_page += 1
            self.display()
    
    def zoom(self):
        if self.zoom_level < 3.0:
            self.zoom_level += 0.25
            self.display()
    
    def zoomout(self):
        if self.zoom_level > 0.5:
            self.zoom_level -= 0.25
            self.display()

    # ================================================================
    # SESSION MANAGEMENT - HIGHLIGHTER COMPATIBLE
    # ================================================================
    
    def getsesspathforpdf(self):
        """Get session path for current PDF"""
        if not self.current_pdf_path or not self.cabinet_id:
            return None
        
        if hasattr(self, 'storage_location') and self.storage_location:
            project_folder = os.path.join(
                self.storage_location,
                self.project_name.replace(' ', '_')
            )
            cabinet_root = os.path.join(
                project_folder,
                self.cabinet_id.replace(' ', '_')
            )
            session_path = os.path.join(
                cabinet_root,
                "Sessions",
                f"{self.cabinet_id}_annotations.json"
            )
            
            return session_path if os.path.exists(session_path) else None
        
        return None
    
    def savesess(self):
        """Save current session to JSON file with all annotation types"""
        if not self.pdf_document:
            print("⚠️ No PDF loaded - cannot save session")
            return
        
        if not hasattr(self, 'storage_location') or not self.storage_location:
            print("⚠️ Storage location not set - cannot save session")
            return
        
        # Determine save path
        project_folder = os.path.join(
            self.storage_location,
            self.project_name.replace(' ', '_')
        )
        cabinet_root = os.path.join(
            project_folder,
            self.cabinet_id.replace(' ', '_')
        )
        sessions_dir = os.path.join(cabinet_root, "Sessions")
        
        # Ensure sessions directory exists
        os.makedirs(sessions_dir, exist_ok=True)
        
        save_path = os.path.join(
            sessions_dir,
            f"{self.cabinet_id}_annotations.json"
        )
        
        data = {
            'project_name': self.project_name,
            'sales_order_no': self.sales_order_no,
            'cabinet_id': self.cabinet_id,
            'pdf_path': self.current_pdf_path,
            'current_page': self.current_page,
            'zoom_level': self.zoom_level,
            'current_sr_no': self.current_sr_no,
            'session_refs': list(self.session_refs),
            'annotations': [],
            'undo_stack_size': len(self.undo_stack) if hasattr(self, 'undo_stack') else 0,
            'save_timestamp': datetime.now().isoformat()
        }
        
        # Process all annotation types
        for ann in self.annotations:
            entry = ann.copy()
            
            # ===== HIGHLIGHTER ANNOTATIONS - Convert tuples to lists =====
            if 'points_page' in entry:
                entry['points_page'] = [[float(x), float(y)] for x, y in entry['points_page']]
            
            # ===== BBOX for highlights =====
            if 'bbox_page' in entry:
                entry['bbox_page'] = [float(x) for x in entry['bbox_page']]
            
            # ===== PEN STROKES - Convert tuples to lists =====
            if 'points' in entry:
                entry['points'] = [[float(x), float(y)] for x, y in entry['points']]
            
            # ===== TEXT ANNOTATIONS - Convert tuple to list =====
            if 'pos_page' in entry:
                pos = entry['pos_page']
                entry['pos_page'] = [float(pos[0]), float(pos[1])]
            
            # Ensure text content is saved
            if 'text' in entry:
                entry['text'] = str(entry['text'])
            
            data['annotations'].append(entry)
        
        try:
            with open(save_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            
            # Count annotation types for feedback
            highlight_count = len([a for a in self.annotations if a.get('type') == 'highlight'])
            error_count = len([a for a in self.annotations if a.get('type') == 'error'])
            pen_count = len([a for a in self.annotations if a.get('type') == 'pen'])
            text_count = len([a for a in self.annotations if a.get('type') == 'text'])
            
            print(f"\n✓ Session saved to: {save_path}")
            print(f"Total annotations: {len(self.annotations)}")
            if highlight_count > 0:
                print(f"  🖍️ Highlights: {highlight_count}")
            if error_count > 0:
                print(f"  ❌ Errors: {error_count}")
            if pen_count > 0:
                print(f"  ✏️ Pen strokes: {pen_count}")
            if text_count > 0:
                print(f"  🅰️ Text annotations: {text_count}")
            
        except Exception as e:
            print(f"❌ Failed to save session: {e}")
            import traceback
            traceback.print_exc()
    
    def loadsessfrompath(self, path):
        """Load annotation session - FULL HIGHLIGHTER SUPPORT"""
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror("Session Load Error", f"Failed to load session:\n{e}")
            return
        
        self.project_name = data.get('project_name', self.project_name)
        self.sales_order_no = data.get('sales_order_no', self.sales_order_no)
        self.cabinet_id = data.get('cabinet_id', getattr(self, "cabinet_id", ""))
        self.current_page = data.get('current_page', 0)
        self.zoom_level = data.get('zoom_level', 1.0)
        self.current_sr_no = data.get('current_sr_no', self.current_sr_no)
        
        # Restore session refs
        self.annotations = []
        self.session_refs = set(data.get('session_refs', []))
        
        highlight_count = 0
        error_count = 0
        pen_count = 0
        text_count = 0
        box_count = 0
        
        for entry in data.get('annotations', []):
            ann = entry.copy()
            ann_type = ann.get('type')
            
            # ===== HIGHLIGHTER ANNOTATIONS - points_page =====
            if 'points_page' in ann:
                ann['points_page'] = [(float(p[0]), float(p[1])) for p in ann['points_page']]
                if ann_type == 'highlight':
                    highlight_count += 1
                elif ann_type == 'error':
                    error_count += 1
            
            # ===== BBOX - Convert list to tuple =====
            if 'bbox_page' in ann:
                ann['bbox_page'] = tuple(float(x) for x in ann['bbox_page'])
            
            # ===== PEN STROKES - points =====
            if 'points' in ann:
                ann['points'] = [(float(p[0]), float(p[1])) for p in ann['points']]
                pen_count += 1
            
            # ===== TEXT ANNOTATIONS - pos_page =====
            if 'pos_page' in ann:
                pos = ann['pos_page']
                ann['pos_page'] = (float(pos[0]), float(pos[1]))
                text_count += 1
            
            # ===== BOX ANNOTATIONS - Count but skip =====
            if ann_type == 'box':
                box_count += 1
                print(f"  ⚠️ Skipping box annotation (type='box') - boxes are disabled")
                continue  # Skip box annotations
            
            # Ensure text content is restored
            if 'text' in ann:
                ann['text'] = str(ann['text'])
            
            self.annotations.append(ann)
            
            # Add ref_no to session refs
            if ann.get('ref_no'):
                self.session_refs.add(str(ann['ref_no']).strip())
        
        self.display()
        
        print(f"\n✓ Session loaded from: {path}")
        print(f"Total annotations loaded: {len(self.annotations)}")
        print(f"  🖍️ Highlights: {highlight_count}")
        print(f"  ❌ Errors (as highlights): {error_count}")
        print(f"  ✏️ Pen strokes: {pen_count}")
        print(f"  🅰️ Text annotations: {text_count}")
        if box_count > 0:
            print(f"  📦 Box annotations (skipped): {box_count}")
        
        types_loaded = {}
        for ann in self.annotations:
            ann_type = ann.get('type', 'unknown')
            types_loaded[ann_type] = types_loaded.get(ann_type, 0) + 1
        print(f"Annotation types loaded: {types_loaded}\n")


def main():
    root = tk.Tk()
    app = ProductionTool(root)
    root.mainloop()


if __name__ == "__main__":
    main()
