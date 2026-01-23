# Inspectron1

# Inspectron1 - Comprehensive Quality Inspection Tool

## Overview

Inspectron1 is a comprehensive quality inspection and production management system designed for manufacturing environments. It provides integrated tools for quality inspectors, production teams, and managers to conduct detailed inspections, track defects, manage rework cycles, and maintain complete audit trails of all manufacturing processes.

The system combines PDF annotation capabilities with Excel-based punch sheets, optical character recognition (OCR) for automated data extraction, and intelligent workflow management to streamline the quality assurance process.

---

## Table of Contents

- [System Architecture](#system-architecture)
- [Installation & Setup](#installation--setup)
- [Core Modules](#core-modules)
- [Database Structure](#database-structure)
- [User Workflows](#user-workflows)
- [Key Features](#key-features)
- [API Reference](#api-reference)
- [Configuration](#configuration)
- [Troubleshooting](#troubleshooting)

---

## System Architecture

### Technology Stack

- **Language:** Python 3.7+
- **GUI Framework:** Tkinter (standard Python GUI library)
- **PDF Processing:** PyMuPDF (fitz) for PDF manipulation
- **Database:** SQLite3 for persistent data storage
- **Data Format:** JSON for configuration and structured data storage
- **Excel Processing:** openpyxl for spreadsheet management
- **OCR Engine:** Tesseract for optical character recognition
- **Image Processing:** OpenCV and PIL for image handling
- **Visualization:** Matplotlib for analytics and charts

### Module Organization

```
Inspectron1/
├── login.py                        # Login system and user management
├── manager.py                      # Manager analytics and reporting interface
├── quality.py                      # Main quality inspection tool
├── production.py                   # Production rework interface
├── database_manager.py             # SQLite database operations
├── handover_database.py            # Quality-Production workflow management
├── categories.json                 # Defect type definitions
├── credentials.json                # User credentials storage
├── Emerson.xlsx                    # Master template for Excel files
├── inspection_tool.db              # Quality inspection records
├── manager.db                      # Management and analytics data
└── handover.db                     # Handover transaction logs
```

### System Flow

```
User Authentication (user_authentication.py)
    |
    +-- Quality Inspector --> quality.py (Annotation & Defect Logging)
    |                            |
    |                            +-- Highlight defects with OCR extraction
    |                            +-- Log punches to Excel
    |                            +-- Export annotated PDFs
    |                            +-- Handover to Production
    |
    +-- Production Team --> production.py (Rework Management)
    |                          |
    |                          +-- Review defects highlighted by quality
    |                          +-- Implement fixes
    |                          +-- Mark punches as completed
    |                          +-- Handback to Quality
    |
    +-- Manager --> manager.py (Analytics & Oversight)
                        |
                        +-- View cabinet statistics
                        +-- Pareto analysis of defects
                        +-- Category library management
                        +-- Template Excel configuration
```

---

## Installation & Setup

### Prerequisites

- Python 3.7 or higher
- pip (Python package manager)
- SQLite3 (included with Python)
- Tesseract OCR engine

### Dependencies Installation

```bash
# Install required Python packages
pip install openpyxl
pip install pillow
pip install pymupdf
pip install opencv-python
pip install pytesseract
pip install matplotlib
pip install numpy

# Install Tesseract OCR (Windows)
# Download installer from: https://github.com/UB-Mannheim/tesseract/wiki

# Or for Linux:
sudo apt-get install tesseract-ocr

# Or for macOS:
brew install tesseract
```

### Project Setup Steps

1. **Clone the Repository**
   ```bash
   git clone https://github.com/Flashhhhh1234/Inspectron1.git
   cd Inspectron1
   ```

2. **Create Assets Directory**
   ```bash
   mkdir assets
   ```

3. **Configure Credentials**
   - Create `assets/credentials.json` with user accounts:
   ```json
   {
     "users": {
       "admin": {
         "password": "admin123",
         "role": "Admin",
         "full_name": "Administrator"
       },
       "inspector1": {
         "password": "password123",
         "role": "Quality",
         "full_name": "Quality Inspector Name"
       }
     }
   }
   ```

4. **Set Tesseract Path** (Windows)
   - In `quality.py`, update the path:
   ```python
   TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
   if os.path.exists(TESSERACT_PATH):
       pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
   ```

5. **Launch Application**
   ```bash
   python user_authentication.py
   ```

---

## Core Modules

### 1. user_authentication.py - Authentication & User Management

**Purpose:** Manages user login, credential validation, and role-based access control.

**Key Classes:**

#### LoginPage
- **Purpose:** Main login interface
- **Methods:**
  - `validate_login()` - Checks credentials against stored database
  - `open_admin()` - Opens admin panel for user management

#### AdminPanel
- **Purpose:** User management interface
- **Methods:**
  - `refresh_users()` - Reloads user list from credentials
  - `add_user()` - Creates new user account
  - `edit_user()` - Modifies existing user
  - `delete_user()` - Removes user account (except admin)

#### AddEditUserDialog
- **Purpose:** Dialog for adding or editing users
- **Methods:**
  - `save_user()` - Persists user data to credentials.json

**Key Functions:**

```python
def load_credentials():
    """Load user credentials from assets/credentials.json
    
    Returns:
        dict: Dictionary with users and their properties
    """
    
def save_credentials(credentials):
    """Save credentials to file with proper JSON formatting
    
    Args:
        credentials (dict): Dictionary to save
    """
    
def authenticate_user(username, password, credentials):
    """Validate username and password combination
    
    Args:
        username (str): Username to validate
        password (str): Password to validate
        credentials (dict): Credentials database
        
    Returns:
        tuple: (role, full_name) if valid, (None, None) if invalid
    """
    
def route_to_role(username, full_name, role):
    """Launch appropriate application module based on user role
    
    Args:
        username (str): Logged-in username
        full_name (str): User's full name
        role (str): User's role (Quality, Production, Manager, Admin)
    """
```

**Supported Roles:**
- **Admin:** Full system access, user management
- **Quality:** Quality inspection module access
- **Production:** Production rework module access
- **Manager:** Dashboard and analytics access

---

### 2. quality.py - Quality Inspection Tool

**Purpose:** Primary tool for quality inspectors to annotate PDFs, log defects, and manage the inspection workflow.

**Key Classes:**

#### CircuitInspector
- **Purpose:** Main application controller for quality inspection
- **Major Methods:**

```python
def load_pdf():
    """Open and load a PDF file for inspection
    
    Prompts user for project details, creates directory structure,
    and initializes the working Excel file.
    """

def display_page():
    """Render current PDF page with all annotations
    
    Handles:
    - Conversion of page to displayable image
    - Rendering of highlighter strokes
    - Rendering of pen annotations
    - Rendering of text annotations
    - Coordinate transformation for zoom and scroll
    """

def extract_text_from_highlight_area(annotation):
    """Extract text from highlighted region using OCR
    
    Args:
        annotation (dict): Highlight annotation with bbox_page
        
    Returns:
        str: Extracted text in uppercase, or None if extraction fails
        
    Process:
    1. Crop image to highlight boundaries with padding
    2. Upscale image 2x for better OCR accuracy
    3. Convert to grayscale and apply threshold
    4. Run Tesseract OCR
    5. Validate text is all capitals
    """

def on_left_press(event):
    """Handle mouse button down events
    
    Initiates annotation based on active tool:
    - Highlighter: Start collecting points for highlight stroke
    - Pen: Start collecting points for pen drawing
    - Text: Record position for text annotation
    """

def on_left_drag(event):
    """Handle mouse movement while button held
    
    For highlighter and pen tools, draws visual feedback line
    and accumulates points for the stroke.
    """

def on_left_release_with_ocr(event):
    """Complete annotation after mouse button release
    
    For highlighter (orange only):
    - Straightens path
    - Extracts text from highlighted area
    - Shows category menu for logging punch
    
    For pen and text:
    - Adds annotation to collection
    - Triggers display refresh
    """

def save_session():
    """Save all annotations to JSON file
    
    Serializes annotations with coordinate conversion:
    - Highlighter points to list format
    - Bounding boxes for error highlights
    - Text positions for annotations
    - Metadata (sr_no, ref_no, category, etc.)
    """

def load_session_from_path(path):
    """Load previously saved annotation session
    
    Deserializes JSON and reconstructs all annotations,
    maintaining proper coordinate format and metadata.
    """

def export_annotated_pdf():
    """Create annotated PDF with all marks
    
    Adds to PDF:
    - Ink annotations for highlighter strokes
    - Line drawings for pen strokes
    - Text annotations with timestamps
    - SR numbers for closed punches
    """

def review_checklist_now():
    """Open Interphase checklist review dialog
    
    Displays items needing status (OK/NOK/N/A),
    allows user to mark progress and add remarks.
    """

def punch_closing_mode():
    """Interactive mode for closing logged defects
    
    Workflow:
    1. Load open punches from Excel
    2. Display each punch for review
    3. Allow adding quality remarks
    4. Mark punch as closed
    5. Convert orange highlight to green
    """

def handover_to_production():
    """Hand cabinet to production for rework
    
    Validates:
    - Checklist completion
    - Saves current session
    - Creates handover record
    - Updates cabinet status
    """

def auto_finalize_if_complete():
    """Automatically finalize cabinet when all work complete
    
    Conditions checked:
    - Zero open punches
    - Checklist fully reviewed
    - Then: saves Excel, exports PDF, marks closed
    """
```

**Highlighter System:**

The tool uses color-coded highlighters for different purposes:

- **Orange Highlighter:** Marks defects requiring action
  - Supports OCR text extraction
  - Automatically extracts text from highlighted area
  - Shows menu to classify defect
  - Converts to green when punch is closed

- **Green Highlighter:** Marks approved/resolved items
  - Used for quality-approved areas
  - Shows SR number in exported PDF

- **Yellow Highlighter:** General marking or wiring notes
  - No OCR required
  - Used for informational marking

**Defect Logging Workflow:**

```
Orange Highlight (OCR extracts text)
    |
    +-- Select Category from categories.json
    |       |
    |       +-- Template Category: Auto-fills punch text
    |       |
    |       +-- Parent Category: Select subcategory
    |       |       |
    |       |       +-- Run subcategory template with OCR text pre-fill
    |       |
    |       +-- Wiring Selector: Select wiring type
    |               |
    |               +-- Select specific wiring defect
    |
    +-- Template execution with OCR-extracted text as first input
    |
    +-- Log punch to Excel (Auto-increment SR No, Ref No)
    |
    +-- Update Interphase status for Reference
    |
    +-- Store annotation with metadata
```

**Excel Integration:**

Reads and writes to Punch Sheet:
- Column A: SR No (serial number, auto-incremented)
- Column B: Ref No (reference number, classification)
- Column C: Description (defect description)
- Column D: Category (defect category)
- Column E: Checked Name (inspector who logged)
- Column F: Checked Date (timestamp of logging)
- Column G: Implemented Name (production implementer)
- Column H: Implemented Date (rework completion timestamp)
- Column I: Closed Name (quality inspector who closed)
- Column J: Closed Date (final closure timestamp)

---

### 3. production.py - Production Rework Tool

**Purpose:** Allows production team to review quality findings and implement fixes.

**Key Classes:**

#### ProductionTool
- **Purpose:** Interface for production rework workflow

**Key Methods:**

```python
def load_from_handover_queue():
    """Display list of items handed to production
    
    Shows:
    - Cabinet ID, Project Name
    - Number of punches
    - Who handed over and when
    - Allows selection to load item
    """

def load_handover_item(item):
    """Load a handover item and auto-open production mode
    
    Process:
    1. Verify PDF and Excel files exist
    2. Load PDF document
    3. Load session from quality tool
    4. Auto-open production mode dialog
    """

def production_mode():
    """Interactive mode for reviewing and completing punches
    
    Workflow:
    1. Load open (non-implemented) punches
    2. Display each punch with quality remarks
    3. Allow adding implementation notes
    4. Mark punch as implemented
    5. Auto-save session
    6. Move to next punch
    """

def complete_rework_handback():
    """Finalize rework and hand back to quality
    
    Validates:
    - All punches marked as implemented
    - Session auto-saved
    - Handback record created
    - Status updated to "being_closed_by_quality"
    """

def navigate_to_punch_highlighter(sr_no, punch_text):
    """Highlight annotation location on current page
    
    Displays:
    - Dashed box around defect location
    - Arrow pointing to defect
    - SR number label
    """

def sync_manager_stats():
    """Update manager dashboard with current punch counts"""
```

**Production Workflow:**

```
Load from Handover Queue
    |
    +-- Select Cabinet
    |
    +-- Load PDF, Excel, Session
    |
    +-- Auto-open Production Mode
    |
    +-- For Each Open Punch:
    |   |
    |   +-- Display punch details and quality remarks
    |   |
    |   +-- Navigate to highlighted area on PDF
    |   |
    |   +-- Add implementation remarks (optional)
    |   |
    |   +-- Mark as Implemented
    |   |
    |   +-- Auto-save session
    |
    +-- All Punches Completed
    |
    +-- Complete & Handback to Quality
    |
    +-- Auto-save session
    |
    +-- Return to Quality Tool for Verification
```

---

### 4. manager.py - Management Analytics

**Purpose:** Provides managers and supervisors with project overview and analytics.

**Key Classes:**

#### ManagerDatabase
- **Purpose:** Database operations for manager statistics

**Key Methods:**

```python
def init_database():
    """Initialize manager.db with cabinet and category tables"""

def count_punches_from_excel(excel_path):
    """Count punch statistics directly from Excel file
    
    Returns:
        tuple: (total_punches, implemented_punches, closed_punches)
    """

def get_status_from_interphase(excel_path):
    """Determine cabinet status from Interphase worksheet
    
    Reads Interphase column D (Status) for each reference number,
    determines highest completed reference, and assigns status:
    - Refs 1-2: project_info_sheet
    - Refs 3-9: mechanical_assembly
    - Refs 10-18: component_assembly
    - Refs 19-26: final_assembly
    - Refs 27+: final_documentation
    
    Returns:
        str: Status string or None
    """

def get_all_projects():
    """Retrieve all projects with cabinet counts
    
    Returns:
        list: Project dictionaries with last_updated timestamps
    """

def get_cabinets_by_project(project_name):
    """Get all cabinets in a project with real-time statistics
    
    Returns:
        list: Cabinet data with punch counts and status
    """

def get_category_stats(start_date=None, end_date=None, project_name=None):
    """Retrieve defect category statistics with date and project filtering
    
    Args:
        start_date (str): ISO format date (optional)
        end_date (str): ISO format date (optional)
        project_name (str): Filter by project (optional)
        
    Returns:
        list: Category occurrences sorted by frequency
    """
```

#### ManagerUI
- **Purpose:** Dashboard user interface

**Key Methods:**

```python
def show_dashboard():
    """Display project overview with statistics cards
    
    Shows:
    - Daily cabinet count
    - Weekly cabinet count
    - Monthly cabinet count
    - Financial year cabinet count
    - Project list with expandable cabinet details
    """

def show_analytics():
    """Display Pareto chart with category analysis
    
    Features:
    - Search by project name
    - Date range filtering (today, month, quarter, year, custom)
    - View by category or subcategory
    - Filter to show only "problematic" (80% cumulative) items
    - Interactive tooltips and legends
    - Export to Excel
    """

def show_defect_library():
    """Display and manage defect category definitions
    
    Allows:
    - Add new defect types
    - Edit existing categories
    - Add subcategories
    - Manage wiring selector categories
    - View special subcategories
    """

def show_template_editor():
    """Interface for managing master Excel template
    
    Operations:
    - Open current template
    - Replace with new template
    - Export template copy
    - Verify template structure
    """

def export_excel_filtered():
    """Export analytics as formatted Excel file
    
    Generates workbooks with:
    - Category analysis with Pareto ranking
    - Problematic items flagged in red
    - Month-wise or project-wise breakdown
    - Percentage and cumulative percentage columns
    """
```

**Dashboard Features:**

- **Statistics Cards:** Daily, weekly, monthly, and annual cabinet counts
- **Project Cards:** Expandable project view with cabinet details
- **Cabinet Information:**
  - Total punches, implemented, closed counts
  - Current status (workflow stage)
  - Clickable cabinet IDs to open Excel files
- **Analytics:**
  - Pareto chart showing defect frequency
  - 80% cumulative threshold visualization
  - Project and date range filtering
  - Problematic item highlighting

---

### 5. database_manager.py - SQLite Database Operations

**Purpose:** Centralized database manager for data persistence and retrieval.

**Key Classes:**

#### DatabaseManager
- **Purpose:** Handle all database operations

**Key Methods:**

```python
def add_project(project_data):
    """Add new project to database
    
    Args:
        project_data (dict): Project information including:
        - cabinet_id, project_name, sales_order_no
        - storage_location, pdf_path, excel_path
        
    Returns:
        bool: Success status
    """

def update_project(cabinet_id, updates):
    """Update existing project information
    
    Args:
        cabinet_id (str): Project identifier
        updates (dict): Fields to update
        
    Returns:
        bool: Success status
    """

def get_project(cabinet_id):
    """Retrieve project by cabinet ID
    
    Returns:
        dict: Project data or None
    """

def get_recent_projects(limit=20):
    """Get recently accessed projects
    
    Returns:
        list: Project dictionaries ordered by access time
    """

def search_projects(search_term):
    """Search projects by name, cabinet ID, or sales order
    
    Returns:
        list: Matching project dictionaries
    """
```

**Database Schema:**

The system uses multiple SQLite databases:

- **inspection_tool.db** (DatabaseManager)
  - projects: Project metadata and file paths
  - recent_projects: Access tracking
  - quality_handovers: Handover workflow tracking

- **manager.db** (ManagerDatabase)
  - cabinets: Cabinet statistics and workflow status
  - category_occurrences: Defect frequency tracking

---

### 6. handover_database.py (HandoverDB) - Workflow Management

**Purpose:** Manages the Quality-Production handover workflow using JSON storage.

**Key Classes:**

#### HandoverDB
- **Purpose:** Track items between Quality and Production queues

**Key Methods:**

```python
def add_quality_handover(handover_data):
    """Create new handover from Quality to Production
    
    Args:
        handover_data (dict): Cabinet and punch information
        
    Returns:
        bool: Success status
    """

def get_pending_production_items():
    """Retrieve items waiting in production queue
    
    Returns:
        list: Handover items in pending/in_progress status
    """

def add_production_handback(handback_data):
    """Create handback from Production to Quality
    
    Marks quality handover as completed,
    creates new production_to_quality record
    """

def get_pending_quality_items():
    """Retrieve items waiting for quality verification
    
    Returns:
        list: Items with status='pending' in production_to_quality
    """

def verify_production_item(cabinet_id, verified_by, notes, mark_as_closed):
    """Mark production item as verified or closed
    
    Args:
        cabinet_id (str): Cabinet identifier
        verified_by (str): Verifying user
        notes (str): Verification notes
        mark_as_closed (bool): If True, sets to 'closed', else 'verified'
        
    Returns:
        bool: Success status
    """
```

**Handover State Machine:**

```
Quality Inspection
    |
    v
add_quality_handover() --> pending_production
    |
    v
Production Receives --> in_production
    |
    v
Production Completes --> add_production_handback()
    |
    v
pending_quality_verification
    |
    v
verify_production_item() --> closed
    |
    v
Cabinet Finalized
```

---



## User Workflows

### Quality Inspector Workflow

```
1. Login with Quality role credentials
   |
2. Open PDF file for inspection
   |
3. Select or create project details
   |
4. Working directory structure created automatically
   |
5. For each defect found:
   |
   +-- Use orange highlighter to mark area
   |
   +-- OCR extracts text from highlighted region
   |
   +-- Select defect category/subcategory
   |
   +-- Template automatically generates punch description
   |
   +-- OCR text pre-fills first input field
   |
   +-- Log punch to Excel (auto-increment SR No)
   |
   +-- Update Interphase status for reference
   |
   +-- Annotation saved with metadata
   |
6. Review checklist items (Interphase sheet)
   |
   +-- Mark each reference as OK/NOK/N/A
   |
   +-- Add remarks for N/A items
   |
7. Handover to production
   |
   +-- Session auto-saved
   |
   +-- Handover record created
   |
8. Monitor production handback
   |
   +-- Load returned item for verification
   |
   +-- Review punch closing marks
   |
   +-- Add quality remarks
   |
   +-- Mark punches as closed
   |
   +-- Auto-finalize when complete
```

### Production Team Workflow

```
1. Login with Production role credentials
   |
2. View pending handover queue
   |
   +-- Display cabinets handed to production
   |
3. Select cabinet to rework
   |
   +-- Load PDF, Excel, and session
   |
   +-- Auto-open production mode
   |
4. For each open punch:
   |
   +-- Display punch description
   |
   +-- Navigate to annotated location on PDF
   |
   +-- Review quality remarks
   |
   +-- Add implementation notes
   |
   +-- Mark as implemented
   |
5. Complete all punch implementations
   |
   +-- Auto-save session
   |
   +-- Return to quality for verification
```

### Manager Workflow

```
1. Login with Manager role credentials
   |
2. Access Manager Dashboard
   |
   +-- View project overview
   |
   +-- Statistics cards (daily, weekly, monthly, yearly)
   |
   +-- Expandable project list with cabinets
   |
3. Perform analytics
   |
   +-- View Pareto chart of defects
   |
   +-- Filter by project, date range
   |
   +-- Identify problematic (80% cumulative) categories
   |
   +-- Export to Excel for reporting
   |
4. Manage defect library
   |
   +-- Add/edit defect categories
   |
   +-- Configure subcategories
   |
   +-- Set up wiring selector types
   |
5. Manage template
   |
   +-- Review template structure
   |
   +-- Replace with new template
   |
   +-- Export template copies
```

---

## Key Features

### 1. PDF Annotation System

- **Highlighter Tool:** Color-coded highlighting with three colors
  - Orange: Defects requiring action (supports OCR)
  - Green: Approved/resolved items
  - Yellow: General markup and notes

- **Pen Tool:** Freehand drawing for additional marks

- **Text Tool:** Add text annotations with timestamps

- **Coordinate System:** Automatic scaling for zoom and pan

### 2. Optical Character Recognition (OCR)

- **Automatic Text Extraction:** Extracts text from orange-highlighted regions
- **Multi-Orientation Support:** Handles text at 0°, 90°, 180°, 270°
- **Confidence Scoring:** Uses Tesseract confidence metrics
- **Validation:** Ensures extracted text is all capitals
- **Pre-fill Capability:** OCR text automatically pre-fills first template input

### 3. Excel Integration

- **Punch Sheet Management:** Automatic row insertion and numbering
- **Interphase Tracking:** Reference-based status management
- **Auto-increment:** SR numbers automatically assigned
- **Timestamp Tracking:** Records when checks, implementations, and closes occurred
- **Merged Cell Handling:** Properly reads project details from merged cells

### 4. Workflow Management

- **Quality to Production Handover:** Tracks items handed for rework
- **Production to Quality Handback:** Monitors returned items
- **Status Tracking:** Workflow-based status updates
- **Audit Trail:** Complete history of all transitions

### 5. Analytics & Reporting

- **Pareto Analysis:** Identifies defect categories causing 80% of issues
- **Category Statistics:** Tracks defect frequency by category
- **Project Filtering:** Analyze by specific project
- **Date Range Filtering:** Daily, weekly, monthly, quarterly, yearly views
- **Export to Excel:** Formatted reports with color-coding and formulas

### 6. Session Management

- **Auto-save:** Periodic automatic session saving
- **Session Recovery:** Load previous inspection sessions
- **Portable Sessions:** JSON-based for easy sharing
- **Complete Annotation Serialization:** All mark types preserved

---

## API Reference

### Key Function Signatures

#### User Authentication

```python
def load_credentials() -> dict
    """Load user credentials from JSON file"""

def authenticate_user(username: str, password: str, credentials: dict) -> tuple
    """Validate credentials, return (role, full_name) or (None, None)"""

def route_to_role(username: str, full_name: str, role: str) -> None
    """Launch appropriate module based on role"""
```

#### PDF and Annotation

```python
def display_page() -> None
    """Render current page with all annotations"""

def save_session() -> None
    """Serialize annotations to JSON file"""

def load_session_from_path(path: str) -> None
    """Deserialize annotations from JSON file"""

def export_annotated_pdf() -> None
    """Create PDF with all annotations embedded"""
```

#### Excel Operations

```python
def read_cell(ws: Worksheet, row: int, col: str) -> Any
    """Read cell value handling merged cells"""

def write_cell(ws: Worksheet, row: int, col: str, value: Any) -> None
    """Write cell value handling merged cells"""

def get_next_sr_no() -> int
    """Get next serial number for punch"""
```

#### OCR and Extraction

```python
def extract_text_from_highlight_area(annotation: dict) -> str
    """Extract text from highlight using OCR"""

def extract_text_from_pdf_page(pdf_path: str, page_number: int) -> str
    """Extract all text from PDF page using OCR"""

def clean_ocr_text(text: str) -> str
    """Clean and validate OCR output"""
```

#### Database Operations

```python
def add_project(project_data: dict) -> bool
    """Add project to database"""

def update_project(cabinet_id: str, updates: dict) -> bool
    """Update project information"""

def get_project(cabinet_id: str) -> dict
    """Retrieve project by cabinet ID"""

def get_recent_projects(limit: int = 20) -> list
    """Get recently accessed projects"""
```

#### Handover Management

```python
def add_quality_handover(handover_data: dict) -> bool
    """Create Quality to Production handover"""

def get_pending_production_items() -> list
    """Get items in production queue"""

def add_production_handback(handback_data: dict) -> bool
    """Create Production to Quality handback"""

def verify_production_item(cabinet_id: str, verified_by: str, notes: str) -> bool
    """Mark production item as verified"""
```

---

## Configuration

### categories.json Structure

```json
[
  {
    "name": "Electrical Defects",
    "ref_number": "1",
    "mode": "parent",
    "subcategories": [
      {
        "name": "Wire Connection",
        "ref_number": "1.1",
        "inputs": [
          {"name": "location", "label": "Where is the connection?"},
          {"name": "issue", "label": "What is the issue?"}
        ],
        "template": "Wire connection issue at {location}: {issue}"
      }
    ]
  },
  {
    "name": "Wiring Selector",
    "ref_number": "2",
    "mode": "wiring_selector",
    "wiring_types": [
      {
        "type": "Power Wiring",
        "ref_number": "2.1",
        "subcategories": [...]
      }
    ]
  }
]
```

### credentials.json Structure

```json
{
  "users": {
    "username": {
      "password": "hashed_password",
      "role": "Quality|Production|Manager|Admin",
      "full_name": "User Full Name"
    }
  }
}
```

### Directory Structure (Auto-created)

```
storage_location/
├── Project_Name/
│   └── CABINET_ID/
│       ├── Working_Excel/
│       │   └── CABINET_ID_Working.xlsx
│       ├── Annotated_Drawings/
│       │   └── CABINET_ID_Annotated.pdf
│       ├── Interphase_Export/
│       │   └── CABINET_ID_Interphase.xlsx
│       └── Sessions/
│           └── CABINET_ID_annotations.json
```



## Performance Considerations

- **Large PDFs:** Zoom level affects rendering performance. Use 1.0 zoom for fastest response.
- **OCR Processing:** Text extraction adds 2-5 seconds per highlight. Reduce image upscaling for faster processing.
- **Database Queries:** Use indexed columns for faster lookups (cabinet_id, status).
- **File Storage:** Store projects on local fast storage, not network drives.

---

## Security Notes

- Credentials stored in JSON (not encrypted). Use environment variables in production.
- No authentication between production and quality tools. Use network security.
- Session files contain all annotation data. Control access to storage locations.

---

## Future Enhancements

- Database encryption for credentials
- Exclusive database management
- Multi-user concurrent access with locking
- Cloud storage integration
- Real-time notifications for handovers
- Advanced analytics and ML-based defect prediction

**Document Version:** 1.0.0  
**Last Updated:** January 23, 2026  
**Maintained By:** Development Team


