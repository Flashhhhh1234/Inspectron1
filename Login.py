import tkinter as tk
from tkinter import ttk, messagebox
import os, sys
import subprocess
import runpy
from PIL import Image, ImageTk

# Make sibling modules importable in frozen one-file builds.
if getattr(sys, 'frozen', False):
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
    if getattr(sys, 'frozen', False):
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

    if getattr(sys, 'frozen', False):
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
    """Authenticate user and return role and full name"""
    users = credentials.get("users", {})
    if username in users:
        if users[username]["password"] == password:
            return users[username]["role"], users[username].get("full_name", username)
    return None, None

# ====================================================== 
# ROUTER - PASS USERNAME AND FULL_NAME TO MODULES
# ====================================================== 
def route_to_role(username, full_name, role):
    """Route to appropriate module with username and full_name as command-line arguments"""
    module_by_role = {
        "Quality": "quality",
        "Manager": "manager",
        "Production": "production",
    }

    if role == "Admin":
        messagebox.showinfo("Admin", "Admin panel opened!")
        return

    module_name = module_by_role.get(role)
    if not module_name:
        messagebox.showerror("Routing Error", f"Page for '{role}' not implemented yet!")
        return

    launch_args = ["--module", module_name, username, full_name]

    if getattr(sys, 'frozen', False):
        subprocess.Popen([sys.executable] + launch_args)
        return

    python_exec = sys.executable or "python"
    login_script = os.path.join(BASE_DIR, "Login.py")
    subprocess.Popen([python_exec, login_script] + launch_args)


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
# ADMIN PANEL
# ====================================================== 
class AdminPanel:
    def __init__(self, parent, credentials):
        self.window = tk.Toplevel(parent)
        self.window.title("Admin Panel - User Management")
        self.window.geometry("700x550")
        self.window.configure(bg="#2b2b2b")
        self.credentials = credentials
        
        # Title
        tk.Label(self.window, text="User Management", font=("Segoe UI", 18, "bold"), 
                bg="#2b2b2b", fg="#ffffff").pack(pady=20)
        
        # User List
        frame = tk.Frame(self.window, bg="#2b2b2b")
        frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        columns = ("Username", "Full Name", "Role", "Password")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", height=12)
        
        for col in columns:
            self.tree.heading(col, text=col)
            if col == "Full Name":
                self.tree.column(col, width=200)
            elif col == "Username":
                self.tree.column(col, width=120)
            elif col == "Role":
                self.tree.column(col, width=100)
            else:
                self.tree.column(col, width=100)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Load users
        self.refresh_users()
        
        # Buttons
        btn_frame = tk.Frame(self.window, bg="#2b2b2b")
        btn_frame.pack(pady=15)
        
        tk.Button(btn_frame, text="Add User", command=self.add_user, 
                 bg="#4CAF50", fg="white", font=("Segoe UI", 10, "bold"),
                 padx=15, pady=8, relief=tk.FLAT, cursor="hand2").pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_frame, text="Edit User", command=self.edit_user,
                 bg="#2196F3", fg="white", font=("Segoe UI", 10, "bold"),
                 padx=15, pady=8, relief=tk.FLAT, cursor="hand2").pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_frame, text="Delete User", command=self.delete_user,
                 bg="#f44336", fg="white", font=("Segoe UI", 10, "bold"),
                 padx=15, pady=8, relief=tk.FLAT, cursor="hand2").pack(side=tk.LEFT, padx=5)
    
    def refresh_users(self):
        self.credentials = load_credentials()

        for item in self.tree.get_children():
            self.tree.delete(item)
        
        users = self.credentials.get("users", {})
        for username, data in users.items():
            full_name = data.get("full_name", username)
            self.tree.insert("", tk.END, values=(username, full_name, data["role"], "••••••"))
    
    def add_user(self):
        AddEditUserDialog(self.window, self.credentials, self.refresh_users)
    
    def edit_user(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a user to edit")
            return
        
        username = self.tree.item(selected[0])["values"][0]
        AddEditUserDialog(self.window, self.credentials, self.refresh_users, username)
    
    def delete_user(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a user to delete")
            return
        
        username = self.tree.item(selected[0])["values"][0]
        
        if username == "admin":
            messagebox.showerror("Error", "Cannot delete admin user!")
            return
        
        if messagebox.askyesno("Confirm", f"Delete user '{username}'?"):
            del self.credentials["users"][username]
            save_credentials(self.credentials)
            self.refresh_users()
            messagebox.showinfo("Success", "User deleted successfully!")

# ====================================================== 
# ADD/EDIT USER DIALOG
# ====================================================== 
class AddEditUserDialog:
    def __init__(self, parent, credentials, refresh_callback, username=None):
        self.window = tk.Toplevel(parent)
        self.window.title("Add User" if not username else "Edit User")
        self.window.geometry("450x350")
        self.window.configure(bg="#2b2b2b")
        self.credentials = credentials
        self.refresh_callback = refresh_callback
        self.edit_username = username
        
        # Title
        title = "Add New User" if not username else f"Edit User: {username}"
        tk.Label(self.window, text=title, font=("Segoe UI", 14, "bold"),
                bg="#2b2b2b", fg="#ffffff").pack(pady=15)
        
        # Form
        form_frame = tk.Frame(self.window, bg="#2b2b2b")
        form_frame.pack(pady=10, padx=30)
        
        # Username
        tk.Label(form_frame, text="Username:", bg="#2b2b2b", fg="#ffffff",
                font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w", pady=8)
        self.username_entry = tk.Entry(form_frame, width=25, font=("Segoe UI", 10))
        self.username_entry.grid(row=0, column=1, pady=8, padx=10)
        
        # Full Name
        tk.Label(form_frame, text="Full Name:", bg="#2b2b2b", fg="#ffffff",
                font=("Segoe UI", 10)).grid(row=1, column=0, sticky="w", pady=8)
        self.fullname_entry = tk.Entry(form_frame, width=25, font=("Segoe UI", 10))
        self.fullname_entry.grid(row=1, column=1, pady=8, padx=10)
        
        # Password
        tk.Label(form_frame, text="Password:", bg="#2b2b2b", fg="#ffffff",
                font=("Segoe UI", 10)).grid(row=2, column=0, sticky="w", pady=8)
        self.password_entry = tk.Entry(form_frame, width=25, show="*", font=("Segoe UI", 10))
        self.password_entry.grid(row=2, column=1, pady=8, padx=10)
        
        # Role
        tk.Label(form_frame, text="Role:", bg="#2b2b2b", fg="#ffffff",
                font=("Segoe UI", 10)).grid(row=3, column=0, sticky="w", pady=8)
        self.role_var = tk.StringVar(value="Quality")
        self.role_combo = ttk.Combobox(form_frame, textvariable=self.role_var,
                                      values=["Admin", "Manager", "Quality", "Production"],
                                      state="readonly", width=23, font=("Segoe UI", 10))
        self.role_combo.grid(row=3, column=1, pady=8, padx=10)
        
        # Load existing data if editing
        if username:
            self.username_entry.insert(0, username)
            self.username_entry.config(state="disabled")
            user_data = self.credentials["users"][username]
            self.fullname_entry.insert(0, user_data.get("full_name", username))
            self.password_entry.insert(0, user_data["password"])
            self.role_var.set(user_data["role"])
        
        # Buttons
        btn_frame = tk.Frame(self.window, bg="#2b2b2b")
        btn_frame.pack(pady=20)
        
        tk.Button(btn_frame, text="Save", command=self.save_user,
                 bg="#4CAF50", fg="white", font=("Segoe UI", 10, "bold"),
                 padx=25, pady=8, relief=tk.FLAT, cursor="hand2").pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_frame, text="Cancel", command=self.window.destroy,
                 bg="#757575", fg="white", font=("Segoe UI", 10, "bold"),
                 padx=25, pady=8, relief=tk.FLAT, cursor="hand2").pack(side=tk.LEFT, padx=5)
    
    def save_user(self):
        username = self.username_entry.get().strip()
        full_name = self.fullname_entry.get().strip()
        password = self.password_entry.get()
        role = self.role_var.get()
        
        if not username or not password or not full_name:
            messagebox.showerror("Error", "Username, full name, and password are required!")
            return
        
        if not self.edit_username and username in self.credentials["users"]:
            messagebox.showerror("Error", "Username already exists!")
            return
        
        self.credentials["users"][username] = {
            "password": password, 
            "role": role,
            "full_name": full_name
        }
        save_credentials(self.credentials)
        self.refresh_callback()
        messagebox.showinfo("Success", "User saved successfully!")
        self.window.destroy()

# ====================================================== 
# LOGIN UI - WITH EMERSON LOGO
# ====================================================== 
class LoginPage:
    def __init__(self, root):
        self.root = root
        self.root.title("Inprocess Tool - Login")
        self.root.geometry("450x550")
        self.root.resizable(False, False)
        self.root.configure(bg="#2b2b2b")
        self.credentials = load_credentials()
        
        # Header with Logo
        header_frame = tk.Frame(root, bg="#2b2b2b", height=120)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        # Load and display Emerson logo
        try:
            logo_path = get_asset_path("EmersonLogo.png")
            if os.path.exists(logo_path):
                logo_img = Image.open(logo_path)
                # Resize logo to fit nicely in header (maintain aspect ratio)
                # Adjust these dimensions as needed for your logo
                logo_img.thumbnail((200, 150), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(logo_img)
                
                logo_label = tk.Label(header_frame, image=self.logo_photo,bg="#2b2b2b")
                logo_label.pack(pady=10)
                print(f"[OK] Logo loaded from: {logo_path}")
            else:
                # Fallback to text if logo not found
                tk.Label(header_frame, text="INPROCESS TOOL", 
                        font=("Segoe UI", 26, "bold"), 
                        bg="#1e1e1e", fg="#00bcd4").pack(pady=30)
                print(f"[WARN] Logo not found at: {logo_path}")
                print("   Could not load bundled EmersonLogo.png")
        except Exception as e:
            # Fallback to text if error loading logo
            tk.Label(header_frame, text="INPROCESS TOOL", 
                    font=("Segoe UI", 26, "bold"), 
                    bg="#1e1e1e", fg="#00bcd4").pack(pady=30)
            print(f"[WARN] Error loading logo: {e}")
        
        # Main container
        container = tk.Frame(root, bg="#2b2b2b")
        container.pack(expand=True, fill=tk.BOTH, padx=40, pady=30)
        
        tk.Label(container, text="Login to your account", 
                font=("Segoe UI", 12), bg="#2b2b2b", fg="#b0b0b0").pack(pady=(0, 25))
        
        # Username
        tk.Label(container, text="Username", font=("Segoe UI", 10, "bold"),
                bg="#2b2b2b", fg="#ffffff").pack(anchor="w", pady=(10, 5))
        
        self.user_entry = tk.Entry(container, font=("Segoe UI", 12), 
                                   bg="#3a3a3a", fg="#ffffff", 
                                   relief=tk.FLAT, insertbackground="white")
        self.user_entry.pack(fill=tk.X, ipady=10)
        
        # Password
        tk.Label(container, text="Password", font=("Segoe UI", 10, "bold"),
                bg="#2b2b2b", fg="#ffffff").pack(anchor="w", pady=(20, 5))
        
        self.pwd_entry = tk.Entry(container, font=("Segoe UI", 12), show="●",
                                  bg="#3a3a3a", fg="#ffffff",
                                  relief=tk.FLAT, insertbackground="white")
        self.pwd_entry.pack(fill=tk.X, ipady=10)
        
        # Login Button
        self.login_btn = tk.Button(container, text="LOGIN", command=self.validate_login,
                                   bg="#00bcd4", fg="white", font=("Segoe UI", 12, "bold"),
                                   relief=tk.FLAT, cursor="hand2", activebackground="#00acc1")
        self.login_btn.pack(fill=tk.X, pady=(30, 10), ipady=12)
        
        # Admin Button
        admin_btn = tk.Button(container, text="Admin Panel", command=self.open_admin,
                             bg="#3a3a3a", fg="#ffffff", font=("Segoe UI", 9),
                             relief=tk.FLAT, cursor="hand2", activebackground="#4a4a4a")
        admin_btn.pack(pady=(10, 0))
        
        # Bind Enter key
        self.root.bind('<Return>', lambda e: self.validate_login())
        
        # Focus username
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
            messagebox.showinfo("Success", f"Welcome, {full_name}!\nRole: {role}")
            if role != "Admin":
                # Hide the login window instead of destroying it
                self.root.withdraw()
                # Route to the appropriate module
                route_to_role(username, full_name, role)
                # Login window stays hidden but keeps the process alive
            else:
                AdminPanel(self.root, self.credentials)
        else:
            messagebox.showerror("Login Failed", "Invalid username or password!")
            self.pwd_entry.delete(0, tk.END)
    
    def open_admin(self):
        # Quick admin access for demo purposes
        self.credentials = load_credentials()
        AdminPanel(self.root, self.credentials)

# ====================================================== 
# RUN APP
# ====================================================== 
if __name__ == "__main__":
    if dispatch_from_args():
        sys.exit(0)

    root = tk.Tk()
    app = LoginPage(root)
    root.mainloop()
