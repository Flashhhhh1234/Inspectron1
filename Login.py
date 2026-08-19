import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys
import subprocess
import runpy

# Make sibling modules importable in frozen one-file builds.
if getattr(sys, "frozen", False):
    bundle_dir = getattr(sys, "_MEIPASS", "")
    if bundle_dir:
        bundled_pages_dir = os.path.join(bundle_dir, "pages")
        if os.path.isdir(bundled_pages_dir) and bundled_pages_dir not in sys.path:
            sys.path.insert(0, bundled_pages_dir)

from credentials_store_pg import load_users_from_postgres, save_users_to_postgres


# ======================================================
# APP BASE DIR (Portable)
# ======================================================
def get_app_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR = get_app_base_dir()


def get_asset_path(filename: str) -> str:
    """Resolve image path in source mode and frozen PyInstaller mode."""
    bundle_dir = getattr(sys, "_MEIPASS", "")
    if bundle_dir:
        bundled_path = os.path.join(bundle_dir, "assets", filename)
        if os.path.exists(bundled_path):
            return bundled_path

    if getattr(sys, "frozen", False):
        return os.path.join(BASE_DIR, "assets", filename)

    return os.path.join(os.path.dirname(BASE_DIR), "assets", filename)


# ======================================================
# CREDENTIAL HELPERS
# ======================================================
def load_credentials():
    """Load user credentials from PostgreSQL credential table."""
    try:
        users = load_users_from_postgres("inspection_tool")
        return {"users": users}
    except Exception as e:
        print(f"[ERROR] Failed to load credentials from PostgreSQL: {e}")
        return {"users": {}}


def save_credentials(credentials):
    """Save user credentials into PostgreSQL credential table."""
    users = credentials.get("users", {}) if isinstance(credentials, dict) else {}
    save_users_to_postgres(users, "inspection_tool")


def authenticate_user(username, password, credentials):
    """Authenticate user and return role and full name."""
    users = credentials.get("users", {})
    if username in users:
        if users[username]["password"] == password:
            return users[username]["role"], users[username].get("full_name", username)
    return None, None


# ======================================================
# ROUTER - PASS USERNAME AND FULL_NAME TO MODULES
# ======================================================
def route_to_role(username, full_name, role):
    """Route to appropriate module with username and full_name as command-line arguments."""
    module_by_role = {
        "Quality": "quality",
        "Manager": "manager",
        "Production": "production",
    }

    module_name = module_by_role.get(role)
    if not module_name:
        messagebox.showerror("Routing Error", f"Role '{role}' is not enabled in this login screen.")
        return False

    launch_args = ["--module", module_name, username, full_name]

    if getattr(sys, "frozen", False):
        subprocess.Popen([sys.executable] + launch_args)
        return True

    python_exec = sys.executable or "python"
    login_script = os.path.join(BASE_DIR, "Login.py")
    subprocess.Popen([python_exec, login_script] + launch_args)
    return True


def _resolve_pages_dir() -> str:
    """Resolve pages directory in both source and PyInstaller one-file modes."""
    bundle_dir = getattr(sys, "_MEIPASS", "")
    if bundle_dir:
        bundled_pages = os.path.join(bundle_dir, "pages")
        if os.path.isdir(bundled_pages):
            return bundled_pages
    return BASE_DIR


def _run_module_entry(module_name: str, username: str, full_name: str) -> bool:
    """Run one of the role modules by executing its script file as __main__."""
    script_name_by_module = {
        "quality": "quality.py",
        "manager": "manager.py",
        "production": "production.py",
    }
    script_name = script_name_by_module.get(module_name.lower())
    if not script_name:
        return False

    pages_dir = _resolve_pages_dir()
    script_path = os.path.join(pages_dir, script_name)
    if not os.path.exists(script_path):
        print(f"[ERROR] Module script not found: {script_path}")
        return False

    original_argv = list(sys.argv)
    inserted_path = False
    try:
        if pages_dir not in sys.path:
            sys.path.insert(0, pages_dir)
            inserted_path = True

        # Keep argv shape compatible with existing module code.
        sys.argv = [script_path, username, full_name]
        runpy.run_path(script_path, run_name="__main__")
        return True
    finally:
        sys.argv = original_argv
        if inserted_path:
            try:
                sys.path.remove(pages_dir)
            except ValueError:
                pass


def dispatch_from_args() -> bool:
    """Dispatch to a role module when running as launcher process."""
    if "--module" not in sys.argv:
        return False

    idx = sys.argv.index("--module")
    module_name = sys.argv[idx + 1] if idx + 1 < len(sys.argv) else ""
    username = sys.argv[idx + 2] if idx + 2 < len(sys.argv) else ""
    full_name = sys.argv[idx + 3] if idx + 3 < len(sys.argv) else ""

    ran = _run_module_entry(module_name, username, full_name)
    if not ran:
        messagebox.showerror("Launch Error", f"Unknown module: {module_name}")
    return True


# ======================================================
# ADMIN PANEL - SECTION-WISE USER TABLES
# ======================================================
class AdminPanel:
    BG = "#f4f6f8"
    SURFACE = "#ffffff"
    TEXT = "#1f2937"
    MUTED = "#64748b"
    BORDER = "#d9e0e7"
    PRIMARY = "#2563eb"
    SUCCESS = "#16803c"
    DANGER = "#c93737"

    def __init__(self, parent, credentials):
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.title("TRACE - User Management")
        self.window.geometry("1080x700")
        self.window.minsize(900, 600)
        self.window.configure(bg=self.BG)
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)

        self.credentials = credentials
        self.roles = ["Admin", "Manager", "Quality", "Production"]
        self.trees = {}
        self.new_row_counter = 0
        self.row_passwords = {}
        self.cell_editor = None
        self.editor_ctx = None

        self._configure_admin_styles()

        header = tk.Frame(self.window, bg=self.SURFACE, height=76,
                          highlightthickness=1, highlightbackground=self.BORDER)
        header.pack(fill=tk.X)
        header.pack_propagate(False)

        title_wrap = tk.Frame(header, bg=self.SURFACE)
        title_wrap.pack(side=tk.LEFT, padx=24, pady=14)
        tk.Label(title_wrap, text="TRACE", font=("Segoe UI", 17, "bold"),
                 bg=self.SURFACE, fg=self.TEXT).pack(anchor="w")
        tk.Label(title_wrap, text="User management", font=("Segoe UI", 9),
                 bg=self.SURFACE, fg=self.MUTED).pack(anchor="w")

        tk.Button(header, text="Close", command=self.on_close, bg=self.SURFACE,
                  fg=self.MUTED, activebackground="#eef2f6", relief=tk.FLAT,
                  font=("Segoe UI", 9), cursor="hand2", padx=12, pady=6
                  ).pack(side=tk.RIGHT, padx=24)

        body = tk.Frame(self.window, bg=self.BG)
        body.pack(fill=tk.BOTH, expand=True, padx=22, pady=20)

        self.notebook = ttk.Notebook(body, style="Admin.TNotebook")
        self.notebook.pack(fill=tk.BOTH, expand=True)
        for section in ["All"] + self.roles:
            self._create_section_tab(section)

        status_bar = tk.Frame(self.window, bg=self.SURFACE, height=38,
                              highlightthickness=1, highlightbackground=self.BORDER)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)
        status_bar.pack_propagate(False)
        self.status_label = tk.Label(status_bar, text="Ready", font=("Segoe UI", 9),
                                     bg=self.SURFACE, fg=self.MUTED, anchor="w")
        self.status_label.pack(fill=tk.BOTH, expand=True, padx=22)

        self.refresh_users()
        self.set_status("Double-click a cell to edit. Use the buttons above the table for row actions.")

    def _configure_admin_styles(self):
        style = ttk.Style(self.window)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Admin.TNotebook", background=self.BG, borderwidth=0)
        style.configure("Admin.TNotebook.Tab", padding=(15, 9), font=("Segoe UI", 9),
                        background="#e7ebf0", foreground=self.MUTED)
        style.map("Admin.TNotebook.Tab",
                  background=[("selected", self.SURFACE)],
                  foreground=[("selected", self.TEXT)])
        style.configure("Admin.Treeview", background=self.SURFACE,
                        fieldbackground=self.SURFACE, foreground=self.TEXT,
                        rowheight=32, borderwidth=0, font=("Segoe UI", 9))
        style.configure("Admin.Treeview.Heading", background="#eef2f6",
                        foreground=self.TEXT, relief=tk.FLAT,
                        font=("Segoe UI", 9, "bold"), padding=(8, 9))
        style.map("Admin.Treeview", background=[("selected", "#dbeafe")],
                  foreground=[("selected", self.TEXT)])

    def _make_action_button(self, parent, text, command, bg, fg="white"):
        button = tk.Button(parent, text=text, command=command, bg=bg, fg=fg,
                           activebackground=bg, activeforeground=fg,
                           font=("Segoe UI", 9, "bold"), padx=13, pady=7,
                           relief=tk.FLAT, bd=0, cursor="hand2")
        button.pack(side=tk.LEFT, padx=(0, 8))
        return button

    def _create_section_tab(self, section):
        tab = tk.Frame(self.notebook, bg=self.SURFACE)
        self.notebook.add(tab, text=section)

        toolbar = tk.Frame(tab, bg=self.SURFACE)
        toolbar.pack(fill=tk.X, padx=14, pady=(14, 10))

        add_title = f"Add {section} user" if section != "All" else "Add user"
        self._make_action_button(toolbar, add_title,
                                 lambda s=section: self.new_from_section(s), self.SUCCESS)
        self._make_action_button(toolbar, "Save row",
                                 lambda s=section: self.save_selected_row(s), self.PRIMARY)
        self._make_action_button(toolbar, "Delete",
                                 lambda s=section: self.delete_selected(s), self.DANGER)
        self._make_action_button(toolbar, "Refresh", self.refresh_users,
                                 "#e7ebf0", self.TEXT)

        tk.Label(toolbar, text="Double-click any cell to edit", bg=self.SURFACE,
                 fg=self.MUTED, font=("Segoe UI", 9)).pack(side=tk.RIGHT)

        table_wrap = tk.Frame(tab, bg=self.BORDER, padx=1, pady=1)
        table_wrap.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 14))

        columns = ("Username", "Full Name", "Role", "Password")
        tree = ttk.Treeview(table_wrap, columns=columns, show="headings",
                            style="Admin.Treeview", height=14)
        widths = {"Username": 180, "Full Name": 310, "Role": 160, "Password": 160}
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=widths[col], minwidth=110,
                        anchor="w", stretch=True)

        y_scroll = ttk.Scrollbar(table_wrap, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=y_scroll.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        y_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        tree.bind("<Double-1>", lambda e, s=section: self.start_inline_edit(e, s))
        tree.bind("<F2>", lambda _e, s=section: self.edit_selected_first_cell(s))
        self.trees[section] = tree

    def current_section(self):
        tab_id = self.notebook.select()
        title = self.notebook.tab(tab_id, "text")
        return title.replace(" Users", "")

    def _get_selected_item(self, section):
        tree = self.trees[section]
        selected = tree.selection()
        if not selected:
            return None
        return selected[0]

    def _is_draft_item(self, section, item_id):
        tree = self.trees[section]
        tags = tree.item(item_id, "tags")
        return "draft" in tags

    def _close_cell_editor(self, commit=True):
        if not self.cell_editor:
            return True

        if commit:
            ok = self._commit_cell_edit()
            if not ok:
                return False

        try:
            self.cell_editor.destroy()
        except tk.TclError:
            pass
        self.cell_editor = None
        self.editor_ctx = None
        return True

    def _commit_cell_edit(self):
        if not self.cell_editor or not self.editor_ctx:
            return True

        ctx = self.editor_ctx
        value = self.cell_editor.get().strip()
        return self._apply_cell_value(ctx, value)

    def _apply_cell_value(self, ctx, value):
        section = ctx["section"]
        tree = ctx["tree"]
        item_id = ctx["item_id"]
        col_idx = ctx["col_idx"]

        values = list(tree.item(item_id, "values"))
        old_username = values[0]
        draft = self._is_draft_item(section, item_id)

        if col_idx == 2 and value not in self.roles:
            messagebox.showerror("Validation", "Invalid role selected.")
            return False

        if col_idx in (0, 1, 2) and not value:
            messagebox.showerror("Validation", "This field cannot be empty.")
            return False

        if col_idx == 3 and not value:
            messagebox.showerror("Validation", "Password cannot be empty.")
            return False

        if draft:
            if col_idx == 3:
                self.row_passwords[item_id] = value
                values[3] = "******"
            else:
                values[col_idx] = value
            tree.item(item_id, values=values)
            self.set_status("Draft row updated. Click Save Selected Row to commit.", color="#facc15")
            return True

        users = self.credentials.setdefault("users", {})
        if old_username not in users:
            messagebox.showerror("Error", "Selected user not found. Refresh and try again.")
            return False

        record = users[old_username]
        new_username = old_username

        if col_idx == 0:
            new_username = value
            if new_username != old_username and new_username in users:
                messagebox.showerror("Duplicate", "Username already exists.")
                return False

        full_name = record.get("full_name", old_username)
        role = record.get("role", "Quality")
        password = record.get("password", "")

        if col_idx == 1:
            full_name = value
        elif col_idx == 2:
            role = value
        elif col_idx == 3:
            password = value

        if new_username != old_username:
            users[new_username] = record
            del users[old_username]
            record = users[new_username]

        record["full_name"] = full_name
        record["role"] = role
        record["password"] = password

        save_credentials(self.credentials)
        self._sync_user_rows(old_username, new_username, record)
        self.set_status(f"Saved row: {new_username}", color="#4ade80")
        return True

    def _sync_user_rows(self, old_username, new_username, record):
        display = (new_username, record.get("full_name", new_username), record.get("role", ""), "******")
        role = record.get("role", "")

        for section, tree in self.trees.items():
            found_item = None
            for item in tree.get_children():
                vals = tree.item(item, "values")
                if vals and vals[0] == old_username:
                    found_item = item
                    break

            should_exist = section == "All" or section == role

            if found_item and not should_exist:
                tree.delete(found_item)
                continue

            if found_item and should_exist:
                tree.item(found_item, values=display)
                continue

            if (not found_item) and should_exist:
                tree.insert("", tk.END, values=display)

    def _ask_save_active_editor(self, action_text):
        if not self.cell_editor or not self.editor_ctx:
            return True

        current = self.cell_editor.get().strip()
        if current == self.editor_ctx["initial"]:
            return self._close_cell_editor(commit=False)

        ans = messagebox.askyesnocancel(
            "Unsaved Cell Edit",
            f"You have an unsaved cell edit. Save before {action_text}?",
        )
        if ans is None:
            return False
        if ans:
            return self._close_cell_editor(commit=True)
        return self._close_cell_editor(commit=False)

    def edit_selected_first_cell(self, section):
        item_id = self._get_selected_item(section)
        if not item_id:
            messagebox.showwarning("Warning", "Please select a row first")
            return
        self._begin_cell_editor(section, self.trees[section], item_id, 0)

    def start_inline_edit(self, event, section):
        tree = self.trees[section]
        row_id = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if not row_id or not col:
            return

        if col not in ("#1", "#2", "#3", "#4"):
            return

        col_idx = int(col[1:]) - 1
        self._begin_cell_editor(section, tree, row_id, col_idx)

    def _begin_cell_editor(self, section, tree, item_id, col_idx):
        if not self._close_cell_editor(commit=True):
            return

        col = f"#{col_idx + 1}"
        bbox = tree.bbox(item_id, col)
        if not bbox:
            return

        x, y, width, height = bbox
        values = list(tree.item(item_id, "values"))
        initial = values[col_idx] if col_idx < len(values) else ""

        if col_idx == 3:
            if self._is_draft_item(section, item_id):
                initial = self.row_passwords.get(item_id, "")
            else:
                username = values[0]
                initial = self.credentials.get("users", {}).get(username, {}).get("password", "")

        if col_idx == 2:
            editor = ttk.Combobox(tree, values=self.roles, state="readonly")
            editor.set(initial if initial in self.roles else "Quality")
        else:
            editor = tk.Entry(tree)
            editor.insert(0, initial)

        editor.place(x=x, y=y, width=width, height=height)
        editor.focus_set()

        if isinstance(editor, tk.Entry):
            editor.select_range(0, tk.END)

        self.cell_editor = editor
        self.editor_ctx = {
            "section": section,
            "tree": tree,
            "item_id": item_id,
            "col_idx": col_idx,
            "initial": initial,
        }

        editor.bind("<Return>", lambda _e: self._close_cell_editor(commit=True))
        editor.bind("<Escape>", lambda _e: self._close_cell_editor(commit=False))
        editor.bind("<FocusOut>", lambda _e: self._close_cell_editor(commit=True))

    def set_status(self, message, color="#93c5fd"):
        self.status_label.config(text=message, fg=color)

    def new_from_section(self, section):
        if not self._ask_save_active_editor("adding a new row"):
            return

        tree = self.trees[section]
        self.new_row_counter += 1
        role = section if section in self.roles else "Quality"
        draft_username = f"new_user_{self.new_row_counter}"
        item_id = tree.insert("", tk.END, values=(draft_username, "", role, ""), tags=("draft",))
        self.row_passwords[item_id] = ""

        tree.selection_set(item_id)
        tree.focus(item_id)
        tree.see(item_id)
        self.set_status("Draft row added. Double-click cells to edit and save.", color="#facc15")

        # Start inline edit in Username cell.
        self._begin_cell_editor(section, tree, item_id, 0)

    def refresh_users(self):
        if not self._close_cell_editor(commit=True):
            return

        self.credentials = load_credentials()
        users = self.credentials.get("users", {})
        self.row_passwords = {}

        for section, tree in self.trees.items():
            for item in tree.get_children():
                tree.delete(item)

            for username in sorted(users.keys(), key=lambda u: u.lower()):
                data = users[username]
                full_name = data.get("full_name", username)
                role = data.get("role", "")

                if section != "All" and role != section:
                    continue

                tree.insert("", tk.END, values=(username, full_name, role, "******"))

    def save_selected_row(self, section):
        if not self._close_cell_editor(commit=True):
            return False

        item_id = self._get_selected_item(section)
        if not item_id:
            messagebox.showwarning("Warning", "Please select a row first")
            return False

        tree = self.trees[section]
        if not self._is_draft_item(section, item_id):
            self.set_status("Selected row is already saved.", color="#93c5fd")
            return True

        values = list(tree.item(item_id, "values"))
        username, full_name, role = values[0].strip(), values[1].strip(), values[2].strip()
        password = self.row_passwords.get(item_id, "").strip()

        if not username or not full_name or role not in self.roles or not password:
            messagebox.showerror(
                "Validation",
                "Draft row must have Username, Full Name, Role, and Password before save.",
            )
            return False

        users = self.credentials.setdefault("users", {})
        if username in users:
            messagebox.showerror("Duplicate", "Username already exists.")
            return False

        users[username] = {
            "password": password,
            "role": role,
            "full_name": full_name,
        }
        save_credentials(self.credentials)

        try:
            tree.delete(item_id)
        except tk.TclError:
            pass
        self.row_passwords.pop(item_id, None)

        self._sync_user_rows(username, username, users[username])
        self.set_status(f"Draft saved: {username}", color="#4ade80")
        return True

    def delete_selected(self, section):
        item_id = self._get_selected_item(section)
        if not item_id:
            messagebox.showwarning("Warning", "Please select a user row to delete")
            return

        if not self._ask_save_active_editor("deleting the selected row"):
            return

        tree = self.trees[section]

        if self._is_draft_item(section, item_id):
            answer = messagebox.askyesnocancel(
                "Draft Row",
                "Save this draft row before delete?",
            )

            if answer is None:
                return

            if answer:
                if not self.save_selected_row(section):
                    return
                # Continue with delete flow after saving.
                item_id = self._get_selected_item(section)
                if not item_id:
                    return

            if not messagebox.askyesno("Delete Draft", "Delete this draft row?"):
                return

            if self._is_draft_item(section, item_id):
                tree.delete(item_id)
                self.row_passwords.pop(item_id, None)
                self.set_status("Draft row deleted.", color="#f87171")
                return

        username = tree.item(item_id)["values"][0]

        if username == "admin":
            messagebox.showerror("Protected", "Cannot delete admin user.")
            return

        if not messagebox.askyesno("Confirm Delete", f"Delete user '{username}'?"):
            return

        users = self.credentials.setdefault("users", {})
        if username in users:
            del users[username]
            save_credentials(self.credentials)
            self.refresh_users()
            self.set_status(f"Row deleted: {username}", color="#f87171")

    def delete_from_current_section(self):
        self.delete_selected(self.current_section())

    def on_close(self):
        self._close_cell_editor(commit=False)
        self.parent.destroy()


# ======================================================
# TRACE LOGIN UI
# ======================================================
class LoginPage:
    BG = "#eef2f6"
    CARD = "#ffffff"
    TEXT = "#1f2937"
    MUTED = "#6b7280"
    BORDER = "#cfd8e3"
    PRIMARY = "#2563eb"
    PRIMARY_HOVER = "#1d4ed8"
    DANGER = "#c93737"

    def __init__(self, root):
        self.root = root
        self.root.title("TRACE - Login")
        self.root.geometry("500x560")
        self.root.resizable(False, False)
        self.root.configure(bg=self.BG)
        self.credentials = load_credentials()
        self.password_visible = False

        self._configure_styles()
        self._build_ui()
        self._center_window()
        self.root.bind("<Return>", lambda _e: self.validate_login())
        self.user_entry.focus_set()

    def _configure_styles(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Trace.TEntry", fieldbackground="#ffffff",
                        foreground=self.TEXT, insertcolor=self.TEXT,
                        bordercolor=self.BORDER, lightcolor=self.BORDER,
                        darkcolor=self.BORDER, padding=(12, 10),
                        font=("Segoe UI", 11))
        style.map("Trace.TEntry", bordercolor=[("focus", self.PRIMARY)],
                  lightcolor=[("focus", self.PRIMARY)],
                  darkcolor=[("focus", self.PRIMARY)])

    def _build_ui(self):
        card = tk.Frame(self.root, bg=self.CARD, highlightthickness=1,
                        highlightbackground="#dde3ea")
        card.pack(fill=tk.BOTH, expand=True, padx=36, pady=32)

        header = tk.Frame(card, bg=self.CARD)
        header.pack(fill=tk.X, padx=42, pady=(42, 24))

        logo = tk.Canvas(header, width=42, height=42, bg=self.CARD,
                         highlightthickness=0)
        logo.pack()
        logo.create_oval(3, 3, 39, 39, fill=self.PRIMARY, outline="")
        logo.create_text(21, 21, text="T", fill="white",
                         font=("Segoe UI", 18, "bold"))

        tk.Label(header, text="TRACE", bg=self.CARD, fg=self.TEXT,
                 font=("Segoe UI", 22, "bold")).pack(pady=(10, 3))
        tk.Label(header, text="Sign in to continue", bg=self.CARD, fg=self.MUTED,
                 font=("Segoe UI", 10)).pack()

        form = tk.Frame(card, bg=self.CARD)
        form.pack(fill=tk.X, padx=42)

        self._label(form, "Username")
        self.user_entry = ttk.Entry(form, style="Trace.TEntry")
        self.user_entry.pack(fill=tk.X, pady=(6, 17))

        self._label(form, "Password")
        password_row = tk.Frame(form, bg=self.CARD)
        password_row.pack(fill=tk.X, pady=(6, 4))
        self.pwd_entry = ttk.Entry(password_row, style="Trace.TEntry", show="*")
        self.pwd_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.show_btn = tk.Button(password_row, text="Show", command=self.toggle_password,
                                  bg=self.CARD, fg=self.PRIMARY, activebackground=self.CARD,
                                  activeforeground=self.PRIMARY_HOVER, relief=tk.FLAT,
                                  bd=0, cursor="hand2", font=("Segoe UI", 9), padx=8)
        self.show_btn.pack(side=tk.RIGHT)

        self.status_label = tk.Label(form, text="", bg=self.CARD, fg=self.DANGER,
                                     anchor="w", font=("Segoe UI", 9))
        self.status_label.pack(fill=tk.X, pady=(7, 4))

        self.login_btn = tk.Button(form, text="Sign in", command=self.validate_login,
                                   bg=self.PRIMARY, fg="white",
                                   activebackground=self.PRIMARY_HOVER,
                                   activeforeground="white", font=("Segoe UI", 10, "bold"),
                                   relief=tk.FLAT, bd=0, cursor="hand2")
        self.login_btn.pack(fill=tk.X, ipady=10, pady=(10, 0))

        tk.Label(card, text="Authorized users only", bg=self.CARD, fg="#9ca3af",
                 font=("Segoe UI", 8)).pack(side=tk.BOTTOM, pady=20)

    def _label(self, parent, text):
        tk.Label(parent, text=text, bg=self.CARD, fg=self.TEXT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")

    def _center_window(self):
        self.root.update_idletasks()
        width, height = self.root.winfo_width(), self.root.winfo_height()
        x = max(0, (self.root.winfo_screenwidth() - width) // 2)
        y = max(0, (self.root.winfo_screenheight() - height) // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def toggle_password(self):
        self.password_visible = not self.password_visible
        self.pwd_entry.config(show="" if self.password_visible else "*")
        self.show_btn.config(text="Hide" if self.password_visible else "Show")
        self.pwd_entry.focus_set()

    def _set_busy(self, busy):
        self.login_btn.config(state=tk.DISABLED if busy else tk.NORMAL,
                              text="Signing in..." if busy else "Sign in")
        self.root.update_idletasks()

    def validate_login(self):
        username = self.user_entry.get().strip()
        password = self.pwd_entry.get()
        self.status_label.config(text="")

        if not username or not password:
            self.status_label.config(text="Enter your username and password.")
            (self.user_entry if not username else self.pwd_entry).focus_set()
            return

        self._set_busy(True)
        try:
            self.credentials = load_credentials()
            role, full_name = authenticate_user(username, password, self.credentials)
            if not role:
                self.status_label.config(text="Invalid username or password.")
                self.pwd_entry.delete(0, tk.END)
                self.pwd_entry.focus_set()
                return
            if role == "Admin":
                self.root.withdraw()
                AdminPanel(self.root, self.credentials)
                return
            if route_to_role(username, full_name, role):
                self.root.withdraw()
        except Exception as exc:
            print(f"[ERROR] Login failed: {exc}")
            self.status_label.config(text="Unable to sign in. Please try again.")
        finally:
            if self.root.winfo_exists():
                self._set_busy(False)


# ======================================================
# RUN APP
# ======================================================
if __name__ == "__main__":
    if dispatch_from_args():
        sys.exit(0)

    root = tk.Tk()
    app = LoginPage(root)
    root.mainloop()
