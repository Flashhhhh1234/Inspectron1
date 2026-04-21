# Inspectron1

# Inspectron1 - Comprehensive Quality Inspection Tool

## Overview

Inspectron1 is a desktop quality-inspection platform for electrical/control cabinet workflows. It provides a role-driven process for:

- Quality inspection and punch creation on PDF drawings
- Production rework execution and handback
- Manager-level dashboarding and Pareto analytics
- Admin user maintenance

The application is implemented as a Tkinter-based multi-module system where Login dispatches users to role-specific interfaces. The latest codebase has moved core persistence behavior to a PostgreSQL-backed compatibility layer, while retaining sqlite-style data-access code for minimal disruption.

---

- [Overview](#overview)
- [System Architecture](#system-architecture)
  - [Technology Stack](#technology-stack)
  - [Runtime Architecture](#runtime-architecture)
  - [Data Architecture](#data-architecture)
  - [Recent Codebase Changes](#recent-codebase-changes)
- [Installation & Setup](#installation--setup)
  - [Prerequisites](#prerequisites)
  - [Dependencies Installation](#dependencies-installation)
  - [Project Setup Steps](#project-setup-steps)
  - [Directory Structure](#directory-structure)
  - [Configure PostgreSQL and Base Path](#configure-postgresql-and-base-path)
  - [Set Tesseract Path (Windows)](#set-tesseract-path-windows)
  - [Launch Application](#launch-application)
- [Core Modules](#core-modules)
  1. [Login.py - Authentication, Routing, and Admin UI](#1-loginpy---authentication-routing-and-admin-ui)
  2. [quality.py - Quality Inspection Workbench](#2-qualitypy---quality-inspection-workbench)
  3. [production.py - Production Rework Workbench](#3-productionpy---production-rework-workbench)
  4. [manager.py - Dashboard, Analytics, and Category Governance](#4-managerpy---dashboard-analytics-and-category-governance)
  5. [database_manager.py - Project and Handover Data Access](#5-database_managerpy---project-and-handover-data-access)
  6. [handover_database.py - Queue-Oriented Handover Lifecycle](#6-handover_databasepy---queue-oriented-handover-lifecycle)
  7. [pg_sqlite_compat.py - sqlite-to-PostgreSQL Compatibility Layer](#7-pg_sqlite_compatpy---sqlite-to-postgresql-compatibility-layer)
  8. [category_store_pg.py and category_catalog_format.py - Defect Library Persistence](#8-category_store_pgpy-and-category_catalog_formatpy---defect-library-persistence)
  9. [credentials_store_pg.py - Credential Table Persistence](#9-credentials_store_pgpy---credential-table-persistence)
  10. [path_policy.py and filedialog_compat.py - Runtime Resilience Utilities](#10-path_policypy-and-filedialog_compatpy---runtime-resilience-utilities)
- [User Workflows](#user-workflows)
  - [Quality Inspector Workflow](#quality-inspector-workflow)
  - [Production Team Workflow](#production-team-workflow)
  - [Manager Workflow](#manager-workflow)
  - [Admin Workflow](#admin-workflow)
- [Key Features](#key-features)
  - [1. Drawing Annotation and Punch Capture](#1-drawing-annotation-and-punch-capture)
  - [2. OCR-Assisted Description Building](#2-ocr-assisted-description-building)
  - [3. Excel-Driven Execution Model](#3-excel-driven-execution-model)
  - [4. Quality-Production Queue Lifecycle](#4-quality-production-queue-lifecycle)
  - [5. Management Analytics and Pareto Reporting](#5-management-analytics-and-pareto-reporting)
  - [6. Frozen-Build and Path Compatibility](#6-frozen-build-and-path-compatibility)
- [API Reference](#api-reference)
- [Performance Considerations](#performance-considerations)
- [Security Notes](#security-notes)
- [Future Enhancements](#future-enhancements)
- [Document Metadata](#document-metadata)

---

## System Architecture

### Technology Stack

- Language: Python 3.x
- GUI: Tkinter
- PDF engine: PyMuPDF (fitz)
- Spreadsheet engine: openpyxl
- OCR stack: pytesseract + OpenCV + PIL
- Charting and visualization: matplotlib
- Data layer: PostgreSQL accessed through sqlite-style compatibility wrapper
- Data serialization: JSON (sessions/config)

### Runtime Architecture

Inspectron1 is launched from a single entry point and routed by role:

```text
Login.py
  |
  +-- Admin      -> In-process AdminPanel and user management dialogs
  +-- Quality    -> quality.py
  +-- Production -> production.py
  +-- Manager    -> manager.py
```

Key runtime behavior:

- Login dispatches role scripts by passing username and full name as argv data
- In frozen distribution mode, bundled pages and assets are resolved via _MEIPASS-aware logic
- In source mode, role modules are launched through the same login script with module arguments

### Data Architecture

Inspectron1 uses three logical database keys, all routed through the compatibility layer:

- inspection_tool: project records, recent projects, quality handovers, credential/category lookups
- manager: cabinet snapshots and category occurrence analytics
- handover_db: queue lifecycle between quality and production

Important implementation detail:

- pg_sqlite_compat.py keeps sqlite-like connect/cursor/execute semantics while translating SQL and placeholders for PostgreSQL
- Logical db keys are interpreted as PostgreSQL schema names for search_path routing when non-public
- category and credentials modules load schema hints from assets/postgres.json and use fully qualified table names where required

### Recent Codebase Changes

The current repository state reflects these notable updates:

- Migration from sqlite storage assumptions to PostgreSQL-backed compatibility usage
- Dedicated credential persistence in credentials_store_pg.py
- Dedicated category catalog persistence in category_store_pg.py using normalized tables
- Centralized storage path normalization with path_policy.py for base-path-safe relative persistence
- File dialog fallback wrapper in filedialog_compat.py for environments where tkinter.filedialog is unavailable
- Enhanced frozen-app handling for module loading and asset resolution in Login.py and role modules

---

## Installation & Setup

### Prerequisites

- Python 3.9+
- pip
- PostgreSQL server reachable by the application runtime
- Tesseract OCR engine installed on workstation

### Dependencies Installation

```bash
pip install pillow
pip install pymupdf
pip install openpyxl
pip install numpy
pip install opencv-python
pip install pytesseract
pip install matplotlib
pip install psycopg2-binary
```

Tesseract installation:

- Windows: install from UB Mannheim Tesseract build or official package
- Linux: install tesseract-ocr from distribution repositories
- macOS: install with Homebrew

### Project Setup Steps

1. Clone or copy the repository into a local working directory.
2. Ensure PostgreSQL is accessible from the machine running the app.
3. Verify that required database tables exist in the expected schema(s).
4. Prepare runtime assets that are referenced by the UI but may not be committed in this repository snapshot:
   - Emerson.xlsx (master template in application base directory)
   - pen_icon.png, text_icon.png, undo_icon.png (assets folder)
5. Confirm OCR executable availability (PATH, TESSERACT_CMD, or default install path).
6. Start from Login.py and authenticate with a valid role.

### Directory Structure

```text
Inspectron1/
├── __init__.py
├── Login.py
├── quality.py
├── production.py
├── manager.py
├── database_manager.py
├── handover_database.py
├── pg_sqlite_compat.py
├── category_store_pg.py
├── category_catalog_format.py
├── credentials_store_pg.py
├── path_policy.py
├── filedialog_compat.py
├── README.md
└── assets/
    ├── credentials.json
    └── postgres.json
```

### Configure PostgreSQL and Base Path

The application reads assets/postgres.json for schema and shared path metadata in supporting modules.

Example:

```json
{
  "postgres": {
    "host": "localhost",
    "port": 5432,
    "dbname": "inspection_tool",
    "user": "postgres",
    "password": "your_password",
    "schema": "public",
    "base_path": "\\\\server\\share\\projects"
  }
}
```

Notes:

- path_policy.py resolves persisted relative paths against base_path
- values outside base_path are rejected when path-policy conversion is applied
- pg_sqlite_compat.py currently contains fixed connection constants in code; align these with your deployment policy before production use

### Set Tesseract Path (Windows)

quality.py resolves OCR executable path in this order:

1. TESSERACT_CMD environment variable
2. TESSERACT_PATH environment variable
3. tesseract found on PATH
4. common install locations

If needed, set an environment variable before launch:

```powershell
setx TESSERACT_CMD "C:\Program Files\Tesseract-OCR\tesseract.exe"
```

### Launch Application

```bash
python Login.py
```

---

## Core Modules

### 1. Login.py - Authentication, Routing, and Admin UI

Purpose:

- Loads user credentials from PostgreSQL-backed credential storage
- Authenticates users and routes them to role modules
- Provides Admin UI for add/edit/delete user operations
- Supports both source execution and frozen executable dispatch

Key functions:

- load_credentials()
- save_credentials(credentials)
- authenticate_user(username, password, credentials)
- route_to_role(username, full_name, role)
- dispatch_from_args()

Primary classes:

- LoginPage: login form and role routing trigger
- AdminPanel: user list management surface
- AddEditUserDialog: user creation/edit modal

### 2. quality.py - Quality Inspection Workbench

Purpose:

- Drives drawing-based quality inspection on PDF documents
- Creates and tracks punch entries in the Punch Sheet worksheet
- Maintains annotation sessions, handover transitions, and manager sync

Major capability groups:

- Multi-tool annotation: highlighter, pen, text
- OCR-assisted text extraction from highlighted defects
- Dynamic defect categorization from PostgreSQL-backed category catalog
- Interphase status update and checklist-style review operations
- Session save/load and annotated PDF export
- Quality to Production handover queue creation

Important classes:

- ManagerDB: manager schema update helpers (status, cabinet stats, category occurrences)
- CircuitInspector: end-to-end quality UI controller and workflow orchestrator

### 3. production.py - Production Rework Workbench

Purpose:

- Loads pending handover items from queue
- Guides rework completion punch-by-punch
- Records implementation details and returns cabinets to quality verification

Major capability groups:

- Queue intake from quality_to_production
- Annotation navigation to exact defect locations
- Implemented/closed punch tracking through Excel
- Complete-and-handback action into production_to_quality queue
- Manager snapshot sync for progress visibility

Important classes:

- ManagerDB: cabinet status and stat update bridge for manager schema
- ProductionTool: production UI and rework execution controller

### 4. manager.py - Dashboard, Analytics, and Category Governance

Purpose:

- Provides management-level project visibility and defect analytics
- Allows category library maintenance and export workflows
- Supports template workbook governance flows

Major capability groups:

- Dashboard cards for daily/weekly/monthly/financial-year cabinet counts
- Project search and cabinet expansion with real-time punch data read from Excel
- Pareto-style category/subcategory analytics
- Date and project-based filtering for report context
- Category CRUD via PostgreSQL-backed normalized category tables

Important classes:

- ManagerDatabase: analytics and cabinet data-access layer
- ManagerUI: navigation, dashboard, analytics, and configuration surfaces

### 5. database_manager.py - Project and Handover Data Access

Purpose:

- Central manager for project metadata persistence and retrieval
- Tracks recent projects
- Maintains quality_handover records and status transitions

Key methods include:

- add_project(project_data)
- update_project(cabinet_id, updates)
- get_project(cabinet_id)
- get_all_projects(status=None)
- get_recent_projects(limit=20)
- add_quality_handover(handover_data)
- update_production_received(...)
- update_production_completed(...)
- update_quality_verification(...)

### 6. handover_database.py - Queue-Oriented Handover Lifecycle

Purpose:

- Owns explicit quality_to_production and production_to_quality queue operations
- Applies lightweight migration updates (for example, verification_notes)
- Exposes helper methods for queue checks, verification, and queue removal

Key methods include:

- add_quality_handover(handover_data)
- get_pending_production_items()
- update_production_status(cabinet_id, status, user=None)
- add_production_handback(handback_data)
- get_pending_quality_items()
- verify_production_item(...)
- remove_from_rework_queue(...)

### 7. pg_sqlite_compat.py - sqlite-to-PostgreSQL Compatibility Layer

Purpose:

- Lets legacy sqlite3-style code run against PostgreSQL
- Preserves familiar usage patterns while translating syntax

Notable behavior:

- qmark placeholder translation from ? to %s
- SQL rewrites for INSERT OR REPLACE and INSERT OR IGNORE
- optional ALTER TABLE ADD COLUMN IF NOT EXISTS compatibility rewrite
- Row mapping class that supports dict-style and index-style access
- Connection/Cursor wrappers preserving expected sqlite-like semantics

### 8. category_store_pg.py and category_catalog_format.py - Defect Library Persistence

Purpose:

- Loads and saves category catalogs from/to normalized PostgreSQL tables
- Supports hierarchical category model:
  - category types
  - wiring types
  - subcategory templates
  - template inputs
- Serializes categories to an envelope format with metadata and postgres seed structure

### 9. credentials_store_pg.py - Credential Table Persistence

Purpose:

- Reads active users from credential table
- Upserts modified users and deletes removed entries
- Provides a clean mapping format consumed by Login.py

### 10. path_policy.py and filedialog_compat.py - Runtime Resilience Utilities

Purpose:

- path_policy.py:
  - centralizes shared base path resolution
  - stores persisted paths in relative form
  - reconstructs absolute paths safely at runtime
- filedialog_compat.py:
  - uses tkinter.filedialog when available
  - falls back to Tcl/Tk dialog commands in constrained/frozen environments

---

## User Workflows

### Quality Inspector Workflow

```text
1. Login as Quality
2. Load drawing PDF
3. Resolve or create project/cabinet context
4. Annotate defects with highlighter/pen/text tools
5. Use OCR-assisted category/template flow to generate punch descriptions
6. Write punch entries into Punch Sheet and update Interphase progress
7. Save session and optionally export annotated PDF
8. Handover cabinet to production queue
9. Receive handback and verify closure when returned
```

### Production Team Workflow

```text
1. Login as Production
2. Open pending items from production queue
3. Load PDF + session + workbook for selected cabinet
4. Navigate to each open punch location
5. Record implementation details and mark implemented rows
6. Complete rework and hand back to quality
7. Trigger manager sync updates throughout progression
```

### Manager Workflow

```text
1. Login as Manager
2. View dashboard cards and project summaries
3. Drill into cabinets and read live punch metrics from Excel
4. Open analytics to inspect category/subcategory Pareto trends
5. Filter by date range or project as required
6. Maintain defect library and template configuration paths
7. Export analysis reports for stakeholders
```

### Admin Workflow

```text
1. Login as Admin
2. Open AdminPanel
3. Add/edit/deactivate users through dialog actions
4. Persist updates back to PostgreSQL credential storage
```

---

## Key Features

### 1. Drawing Annotation and Punch Capture

- Multi-color highlighter semantics for defect states
- Pen and text overlays for visual context and instructions
- Coordinate conversion utilities for zoom-accurate drawing placement
- Undo stack and session-safe persistence

### 2. OCR-Assisted Description Building

- OCR extraction from highlighted regions
- Tesseract path auto-detection with environment variable support
- Template-driven defect description generation using captured text

### 3. Excel-Driven Execution Model

- Punch Sheet and Interphase worksheet integration
- Merged-cell-safe read/write helpers
- Auto-increment style punch progression and lifecycle timestamps

### 4. Quality-Production Queue Lifecycle

- Quality handover queue with pending/in_progress/completed states
- Production handback queue with pending/verified/closed outcomes
- Verification notes and rework queue controls for exception handling

### 5. Management Analytics and Pareto Reporting

- Category occurrence aggregation for root-cause concentration analysis
- Time-windowed cabinet productivity visibility
- Filterable project-level reporting surfaces

### 6. Frozen-Build and Path Compatibility

- _MEIPASS-aware resource and module lookup
- Relative path persistence and policy-validated reconstruction
- Dialog fallback path for constrained Tk environments

---

## API Reference

Representative function signatures and integration points:

```python
# Login and routing
load_credentials() -> dict
save_credentials(credentials: dict) -> None
authenticate_user(username: str, password: str, credentials: dict) -> tuple
route_to_role(username: str, full_name: str, role: str) -> None
dispatch_from_args() -> bool

# Database manager
DatabaseManager.add_project(project_data: dict) -> bool
DatabaseManager.update_project(cabinet_id: str, updates: dict) -> bool
DatabaseManager.get_project(cabinet_id: str) -> dict | None
DatabaseManager.get_recent_projects(limit: int = 20) -> list[dict]

# Handover queues
HandoverDB.add_quality_handover(handover_data: dict) -> bool
HandoverDB.get_pending_production_items() -> list[dict]
HandoverDB.add_production_handback(handback_data: dict) -> bool
HandoverDB.get_pending_quality_items() -> list[dict]
HandoverDB.verify_production_item(...) -> bool

# Category and credentials persistence
load_categories_from_postgres(db_key: str = "inspection_tool") -> list[dict]
save_categories_to_postgres(categories: list[dict], db_key: str = "inspection_tool") -> None
load_users_from_postgres(db_key: str = "inspection_tool") -> dict
save_users_to_postgres(users: dict, db_key: str = "inspection_tool") -> None

# Path policy helpers
get_base_path(force_refresh: bool = False) -> str
to_relative_path(path: str | None) -> str | None
to_absolute_path(path: str | None) -> str | None
resolve_storage_location(stored_value: str | None) -> str
```

---

## Performance Considerations

- PDF rendering cost scales with page complexity and zoom level
- OCR is CPU intensive; batch/high-frequency extraction can increase latency
- openpyxl workbook operations can be slow on very large punch sheets
- manager dashboard currently recalculates live punch counts from Excel for accuracy; this favors consistency over raw speed
- network-path latency can impact storage and workbook access when base_path points to remote shares

---

## Security Notes

- pg_sqlite_compat.py currently includes fixed PostgreSQL host/user/password values in code; move these to secure runtime secrets for production
- credentials are managed in database tables; enforce strong password policy and role governance externally
- audit-sensitive session and workbook files should be stored on controlled-access storage
- apply least-privilege database roles for runtime users

---

## Future Enhancements

- remove hardcoded database connection values in favor of environment-backed secret management
- introduce role action auditing with immutable event logs
- add optimistic locking or transactional safeguards for concurrent workbook edits
- provide background OCR workers for smoother UI responsiveness
- expand manager analytics with trend decomposition and first-pass-yield views
- improve packaging docs for frozen deployments including required external assets

---

## Document Metadata

Document Version: 2.0.0
Last Updated: April 21, 2026
Maintained By: Inspectron1 Contributors