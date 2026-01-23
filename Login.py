import tkinter as tk
from tkinter import ttk, messagebox
import os, sys, json
from datetime import datetime
import subprocess
from PIL import Image, ImageTk

# ====================================================== 
# APP BASE DIR (Portable)
# ====================================================== 
def get_app_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_app_base_dir()
# Assets folder is one level up from current folder
ASSETS_DIR = os.path.join(os.path.dirname(BASE_DIR), "assets")
CRED_FILE = os.path.join(ASSETS_DIR, "credentials.json")

# ====================================================== 
# CREDENTIAL HELPERS
# ====================================================== 
def load_credentials():
    """Load user credentials from assets/credentials.json"""
    if not os.path.exists(CRED_FILE):
        os.makedirs(ASSETS_DIR, exist_ok=True)
        default_creds = {
            "users": {
                "admin": {"password": "admin123", "role": "Admin", "full_name": "Administrator"},
                "manager1": {"password": "mgr@2024", "role": "Manager", "full_name": "Manager User"},
                "quality1": {"password": "qc@2024", "role": "Quality", "full_name": "Kshitij Palshikar"},
                "prod1": {"password": "prod@2024", "role": "Production", "full_name": "Kshitij Palshikar"}
            }
        }
        with open(CRED_FILE, "w") as f:
            json.dump(default_creds, f, indent=4)
        print(f"✓ Created default credentials at: {CRED_FILE}")
        return default_creds
    
    with open(CRED_FILE, "r") as f:
        return json.load(f)

def save_credentials(credentials):
    """Save user credentials to assets/credentials.json"""
    with open(CRED_FILE, "w") as f:
        json.dump(credentials, f, indent=4)

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
    if role == "Quality":
        # Pass both username and full_name to quality module
        quality_path = os.path.join(BASE_DIR, "quality.py")
        subprocess.Popen(["python", quality_path, username, full_name])
    elif role == "Manager":
        manager_path = os.path.join(BASE_DIR, "manager.py")
        subprocess.Popen(["python", manager_path, username, full_name])
    elif role == "Production":
        # Pass both username and full_name to production module
        production_path = os.path.join(BASE_DIR, "production.py")
        subprocess.Popen(["python", production_path, username, full_name])
    elif role == "Admin":
        messagebox.showinfo("Admin", "Admin panel opened!")
    else:
        messagebox.showerror("Routing Error", f"Page for '{role}' not implemented yet!")

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
            logo_path = os.path.join(ASSETS_DIR, "EmersonLogo.png")
            if os.path.exists(logo_path):
                logo_img = Image.open(logo_path)
                # Resize logo to fit nicely in header (maintain aspect ratio)
                # Adjust these dimensions as needed for your logo
                logo_img.thumbnail((200, 150), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(logo_img)
                
                logo_label = tk.Label(header_frame, image=self.logo_photo,bg="#2b2b2b")
                logo_label.pack(pady=10)
                print(f"✓ Logo loaded from: {logo_path}")
            else:
                # Fallback to text if logo not found
                tk.Label(header_frame, text="INPROCESS TOOL", 
                        font=("Segoe UI", 26, "bold"), 
                        bg="#1e1e1e", fg="#00bcd4").pack(pady=30)
                print(f"⚠️ Logo not found at: {logo_path}")
                print(f"   Please place EmersonLogo.png in the assets folder")
        except Exception as e:
            # Fallback to text if error loading logo
            tk.Label(header_frame, text="INPROCESS TOOL", 
                    font=("Segoe UI", 26, "bold"), 
                    bg="#1e1e1e", fg="#00bcd4").pack(pady=30)
            print(f"⚠️ Error loading logo: {e}")
        
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
        AdminPanel(self.root, self.credentials)

# ====================================================== 
# RUN APP
# ====================================================== 
if __name__ == "__main__":
    root = tk.Tk()
    app = LoginPage(root)
    root.mainloop()
