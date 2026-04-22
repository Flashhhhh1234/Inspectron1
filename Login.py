import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys
import subprocess
import runpy
from PIL import Image, ImageTk

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
    def __init__(self, parent, credentials):
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.title("Admin Panel - User Management")
        self.window.geometry("1020x690")
        self.window.configure(bg="#111827")
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)

        self.credentials = credentials
        self.roles = ["Admin", "Manager", "Quality", "Production"]
        self.trees = {}
        self.new_row_counter = 0
        self.row_passwords = {}

        self.cell_editor = None
        self.editor_ctx = None

        tk.Label(
            self.window,
            text="User Management",
            font=("Segoe UI", 18, "bold"),
            bg="#111827",
            fg="#f9fafb",
        ).pack(pady=(18, 8))

        tk.Label(
            self.window,
            text="Inline table editing: double-click a cell to edit, add row inline, save/delete by row.",
            font=("Segoe UI", 10),
            bg="#111827",
            fg="#93c5fd",
        ).pack(pady=(0, 12))

        notebook_wrap = tk.Frame(self.window, bg="#111827")
        notebook_wrap.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 10))

        self.notebook = ttk.Notebook(notebook_wrap)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        for section in ["All"] + self.roles:
            self._create_section_tab(section)

        self.status_label = tk.Label(
            self.window,
            text="Ready",
            font=("Segoe UI", 9),
            bg="#111827",
            fg="#93c5fd",
        )
        self.status_label.pack(fill=tk.X, padx=16, pady=(0, 10), anchor="w")

        self.refresh_users()
        self.set_status("Double-click cells to edit inline. Use section buttons for row actions.")

    def _create_section_tab(self, section):
        tab = tk.Frame(self.notebook, bg="#111827")
        self.notebook.add(tab, text=f"{section} Users")

        table_wrap = tk.Frame(tab, bg="#111827")
        table_wrap.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 8))

        columns = ("Username", "Full Name", "Role", "Password")
        tree = ttk.Treeview(table_wrap, columns=columns, show="headings", height=12)
        for col in columns:
            tree.heading(col, text=col)
            if col == "Full Name":
                tree.column(col, width=260)
            elif col == "Username":
                tree.column(col, width=150)
            elif col == "Role":
                tree.column(col, width=130)
            else:
                tree.column(col, width=120)

        y_scroll = ttk.Scrollbar(table_wrap, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=y_scroll.set)

        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        y_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        tree.bind("<Double-1>", lambda e, s=section: self.start_inline_edit(e, s))
        tree.bind("<F2>", lambda _e, s=section: self.edit_selected_first_cell(s))
        self.trees[section] = tree

        section_btn_row = tk.Frame(tab, bg="#111827")
        section_btn_row.pack(fill=tk.X, padx=10, pady=(0, 10))

        add_title = f"Add Row in {section}" if section != "All" else "Add Row"
        tk.Button(
            section_btn_row,
            text=add_title,
            command=lambda s=section: self.new_from_section(s),
            bg="#16a34a",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            padx=12,
            pady=6,
            relief=tk.FLAT,
            cursor="hand2",
        ).pack(side=tk.LEFT, padx=4)

        tk.Button(
            section_btn_row,
            text="Save Selected Row",
            command=lambda s=section: self.save_selected_row(s),
            bg="#2563eb",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            padx=12,
            pady=6,
            relief=tk.FLAT,
            cursor="hand2",
        ).pack(side=tk.LEFT, padx=4)

        tk.Button(
            section_btn_row,
            text="Delete Selected",
            command=lambda s=section: self.delete_selected(s),
            bg="#dc2626",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            padx=12,
            pady=6,
            relief=tk.FLAT,
            cursor="hand2",
        ).pack(side=tk.LEFT, padx=4)

        tk.Button(
            section_btn_row,
            text="Refresh",
            command=self.refresh_users,
            bg="#475569",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            padx=12,
            pady=6,
            relief=tk.FLAT,
            cursor="hand2",
        ).pack(side=tk.LEFT, padx=4)

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
# LOGIN UI - WITH EMERSON LOGO
# ======================================================
class LoginPage:
    def __init__(self, root):
        self.root = root
        self.root.title("Inprocess Tool - Login")
        self.root.geometry("540x620")
        self.root.resizable(False, False)
        self.root.configure(bg="#0f172a")
        self.credentials = load_credentials()

        bg_frame = tk.Frame(root, bg="#0f172a")
        bg_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        card = tk.Frame(bg_frame, bg="#111827", highlightthickness=1, highlightbackground="#1f2937")
        card.pack(fill=tk.BOTH, expand=True)

        header_frame = tk.Frame(card, bg="#111827", height=160)
        header_frame.pack(fill=tk.X, padx=30, pady=(28, 12))
        header_frame.pack_propagate(False)

        try:
            logo_path = get_asset_path("EmersonLogo.png")
            if os.path.exists(logo_path):
                logo_img = Image.open(logo_path)
                logo_img.thumbnail((230, 120), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(logo_img)
                tk.Label(header_frame, image=self.logo_photo, bg="#111827").pack(pady=(4, 8))
            else:
                tk.Label(
                    header_frame,
                    text="INSPECTRON",
                    font=("Segoe UI", 24, "bold"),
                    bg="#111827",
                    fg="#22d3ee",
                ).pack(pady=(24, 8))
        except Exception as e:
            tk.Label(
                header_frame,
                text="INSPECTRON",
                font=("Segoe UI", 24, "bold"),
                bg="#111827",
                fg="#22d3ee",
            ).pack(pady=(24, 8))
            print(f"[WARN] Error loading logo: {e}")

        tk.Label(
            header_frame,
            text="Sign in to continue",
            font=("Segoe UI", 11),
            bg="#111827",
            fg="#94a3b8",
        ).pack()

        container = tk.Frame(card, bg="#111827")
        container.pack(fill=tk.BOTH, expand=True, padx=50, pady=(0, 28))

        tk.Label(
            container,
            text="Username",
            font=("Segoe UI", 10, "bold"),
            bg="#111827",
            fg="#e5e7eb",
        ).pack(anchor="w", pady=(10, 6))

        self.user_entry = tk.Entry(
            container,
            font=("Segoe UI", 12),
            bg="#1f2937",
            fg="#f9fafb",
            relief=tk.FLAT,
            insertbackground="#f9fafb",
            highlightthickness=1,
            highlightbackground="#334155",
            highlightcolor="#06b6d4",
            bd=0,
        )
        self.user_entry.pack(fill=tk.X, ipady=11)

        tk.Label(
            container,
            text="Password",
            font=("Segoe UI", 10, "bold"),
            bg="#111827",
            fg="#e5e7eb",
        ).pack(anchor="w", pady=(20, 6))

        self.pwd_entry = tk.Entry(
            container,
            font=("Segoe UI", 12),
            show="*",
            bg="#1f2937",
            fg="#f9fafb",
            relief=tk.FLAT,
            insertbackground="#f9fafb",
            highlightthickness=1,
            highlightbackground="#334155",
            highlightcolor="#06b6d4",
            bd=0,
        )
        self.pwd_entry.pack(fill=tk.X, ipady=11)

        self.login_btn = tk.Button(
            container,
            text="Sign In",
            command=self.validate_login,
            bg="#0891b2",
            fg="white",
            font=("Segoe UI", 12, "bold"),
            relief=tk.FLAT,
            cursor="hand2",
            activebackground="#0e7490",
            activeforeground="white",
            bd=0,
        )
        self.login_btn.pack(fill=tk.X, pady=(30, 10), ipady=12)

        self.hint_label = tk.Label(
            container,
            text=" ",
            font=("Segoe UI", 9),
            bg="#111827",
            fg="#94a3b8",
        )
        self.hint_label.pack(pady=(6, 0))

        self.root.bind("<Return>", lambda _e: self.validate_login())
        self.user_entry.focus()

    def validate_login(self):
        username = self.user_entry.get().strip()
        password = self.pwd_entry.get()

        if not username or not password:
            messagebox.showerror("Error", "Please enter username and password!")
            return

        self.credentials = load_credentials()
        role, full_name = authenticate_user(username, password, self.credentials)

        if role:
            if role == "Admin":
                self.root.withdraw()
                AdminPanel(self.root, self.credentials)
                return

            launched = route_to_role(username, full_name, role)
            if launched:
                self.root.withdraw()
        else:
            messagebox.showerror("Login Failed", "Invalid username or password!")
            self.pwd_entry.delete(0, tk.END)


# ======================================================
# RUN APP
# ======================================================
if __name__ == "__main__":
    if dispatch_from_args():
        sys.exit(0)

    root = tk.Tk()
    app = LoginPage(root)
    root.mainloop()
