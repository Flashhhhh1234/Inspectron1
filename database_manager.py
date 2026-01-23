import sqlite3
import json
import os
from datetime import datetime
from typing import List, Dict, Optional, Tuple

class DatabaseManager:
    """Centralized SQLite database manager for the Quality Inspection Tool"""
    
    def __init__(self, db_path: str):
        """Initialize database connection and create tables if needed"""
        self.db_path = db_path
        self.conn = None
        self.cursor = None
        self._connect()
        self._create_tables()
    
    def _connect(self):
        """Establish database connection"""
        self.conn = sqlite3.connect(self.db_path)
        self.conn.row_factory = sqlite3.Row  # Enable column access by name
        self.cursor = self.conn.cursor()

    def get_project_location(self, project_name):
        """Get storage location for an existing project name"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT storage_location FROM projects 
                WHERE project_name = ? 
                LIMIT 1
            ''', (project_name,))
            
            result = cursor.fetchone()
            conn.close()
            
            return result[0] if result else None
        except Exception as e:
            print(f"Error getting project location: {e}")
            return None
    
    def _create_tables(self):
        """Create all required tables"""
        
        # Projects table - stores all projects ever created
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_name TEXT NOT NULL,
                sales_order_no TEXT,
                cabinet_id TEXT NOT NULL UNIQUE,
                storage_location TEXT NOT NULL,
                created_date TEXT NOT NULL,
                last_accessed TEXT NOT NULL,
                pdf_path TEXT,
                excel_path TEXT,
                session_path TEXT,
                status TEXT DEFAULT 'active',
                notes TEXT
            )
        """)
        
        # Recent projects table - tracks recently accessed projects
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS recent_projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cabinet_id TEXT NOT NULL,
                last_accessed TEXT NOT NULL,
                FOREIGN KEY (cabinet_id) REFERENCES projects(cabinet_id)
            )
        """)
        
        # Quality handovers table
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS quality_handovers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cabinet_id TEXT NOT NULL UNIQUE,
                project_name TEXT NOT NULL,
                sales_order_no TEXT,
                pdf_path TEXT,
                excel_path TEXT,
                session_path TEXT,
                total_punches INTEGER DEFAULT 0,
                open_punches INTEGER DEFAULT 0,
                closed_punches INTEGER DEFAULT 0,
                handed_over_by TEXT NOT NULL,
                handed_over_date TEXT NOT NULL,
                status TEXT DEFAULT 'pending_production',
                production_received_by TEXT,
                production_received_date TEXT,
                rework_completed_by TEXT,
                rework_completed_date TEXT,
                production_remarks TEXT,
                quality_verified_by TEXT,
                quality_verified_date TEXT,
                FOREIGN KEY (cabinet_id) REFERENCES projects(cabinet_id)
            )
        """)
        
        # Create indexes for better performance
        self.cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_recent_accessed 
            ON recent_projects(last_accessed DESC)
        """)
        
        self.cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_project_cabinet 
            ON projects(cabinet_id)
        """)
        
        self.cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_handover_status 
            ON quality_handovers(status)
        """)
        
        self.conn.commit()
    
    # ================================================================
    # PROJECT MANAGEMENT
    # ================================================================
    
    def add_project(self, project_data: Dict) -> bool:
        """Add a new project to the database"""
        try:
            self.cursor.execute("""
                INSERT INTO projects (
                    project_name, sales_order_no, cabinet_id, storage_location,
                    created_date, last_accessed, pdf_path, excel_path, 
                    session_path, status, notes
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                project_data.get('project_name'),
                project_data.get('sales_order_no'),
                project_data.get('cabinet_id'),
                project_data.get('storage_location'),
                project_data.get('created_date', datetime.now().isoformat()),
                project_data.get('last_accessed', datetime.now().isoformat()),
                project_data.get('pdf_path'),
                project_data.get('excel_path'),
                project_data.get('session_path'),
                project_data.get('status', 'active'),
                project_data.get('notes')
            ))
            self.conn.commit()
            
            # Add to recent projects
            self._add_to_recent(project_data.get('cabinet_id'))
            return True
        except sqlite3.IntegrityError:
            # Project already exists, try updating instead
            try:
                return self.update_project(
                    project_data.get('cabinet_id'),
                    {k: v for k, v in project_data.items() if k != 'cabinet_id'}
                )
            except:
                return False
        except Exception as e:
            print(f"Error adding project: {e}")
            import traceback
            traceback.print_exc()
            self.conn.rollback()
            return False
    
    def update_project(self, cabinet_id: str, updates: Dict) -> bool:
        """Update project information
        
        FIXED: Properly matches number of placeholders with values
        """
        try:
            # Build dynamic UPDATE query
            set_clause = ", ".join([f"{key} = ?" for key in updates.keys()])
            values = list(updates.values())
            
            # Add last_accessed and cabinet_id (for WHERE clause)
            values.append(datetime.now().isoformat())
            values.append(cabinet_id)
            
            self.cursor.execute(f"""
                UPDATE projects 
                SET {set_clause}, last_accessed = ?
                WHERE cabinet_id = ?
            """, values)
            
            self.conn.commit()
            
            # Update recent projects
            self._add_to_recent(cabinet_id)
            return True
        except Exception as e:
            print(f"Error updating project: {e}")
            import traceback
            traceback.print_exc()
            self.conn.rollback()
            return False
    
    def get_project(self, cabinet_id: str) -> Optional[Dict]:
        """Get project by cabinet ID"""
        self.cursor.execute("""
            SELECT * FROM projects WHERE cabinet_id = ?
        """, (cabinet_id,))
        
        row = self.cursor.fetchone()
        if row:
            return dict(row)
        return None
    
    def get_all_projects(self, status: Optional[str] = None) -> List[Dict]:
        """Get all projects, optionally filtered by status"""
        if status:
            self.cursor.execute("""
                SELECT * FROM projects 
                WHERE status = ?
                ORDER BY last_accessed DESC
            """, (status,))
        else:
            self.cursor.execute("""
                SELECT * FROM projects 
                ORDER BY last_accessed DESC
            """)
        
        return [dict(row) for row in self.cursor.fetchall()]
    
    def search_projects(self, search_term: str) -> List[Dict]:
        """Search projects by name, cabinet ID, or sales order"""
        search_pattern = f"%{search_term}%"
        self.cursor.execute("""
            SELECT * FROM projects 
            WHERE project_name LIKE ? 
               OR cabinet_id LIKE ? 
               OR sales_order_no LIKE ?
            ORDER BY last_accessed DESC
        """, (search_pattern, search_pattern, search_pattern))
        
        return [dict(row) for row in self.cursor.fetchall()]
    
    def project_exists(self, cabinet_id: str) -> bool:
        """Check if project exists"""
        self.cursor.execute("""
            SELECT COUNT(*) FROM projects WHERE cabinet_id = ?
        """, (cabinet_id,))
        return self.cursor.fetchone()[0] > 0
    
    def get_storage_location(self, cabinet_id: str) -> Optional[str]:
        """Get storage location for a project"""
        self.cursor.execute("""
            SELECT storage_location FROM projects WHERE cabinet_id = ?
        """, (cabinet_id,))
        
        row = self.cursor.fetchone()
        return row[0] if row else None
    
    # ================================================================
    # RECENT PROJECTS
    # ================================================================
    
    def _add_to_recent(self, cabinet_id: str):
        """Add or update project in recent projects"""
        # Remove old entry if exists
        self.cursor.execute("""
            DELETE FROM recent_projects WHERE cabinet_id = ?
        """, (cabinet_id,))
        
        # Add new entry
        self.cursor.execute("""
            INSERT INTO recent_projects (cabinet_id, last_accessed)
            VALUES (?, ?)
        """, (cabinet_id, datetime.now().isoformat()))
        
        # Keep only last 20 recent projects
        self.cursor.execute("""
            DELETE FROM recent_projects 
            WHERE id NOT IN (
                SELECT id FROM recent_projects 
                ORDER BY last_accessed DESC 
                LIMIT 20
            )
        """)
        
        self.conn.commit()
    
    def get_recent_projects(self, limit: int = 20) -> List[Dict]:
        """Get recent projects with full details"""
        self.cursor.execute("""
            SELECT p.* 
            FROM projects p
            INNER JOIN recent_projects r ON p.cabinet_id = r.cabinet_id
            ORDER BY r.last_accessed DESC
            LIMIT ?
        """, (limit,))
        
        return [dict(row) for row in self.cursor.fetchall()]
    
    def clear_old_recent_projects(self, days: int = 7):
        """Clear recent projects older than specified days"""
        cutoff_date = datetime.now().timestamp() - (days * 24 * 60 * 60)
        cutoff_iso = datetime.fromtimestamp(cutoff_date).isoformat()
        
        self.cursor.execute("""
            DELETE FROM recent_projects 
            WHERE last_accessed < ?
        """, (cutoff_iso,))
        
        self.conn.commit()
    
    # ================================================================
    # QUALITY HANDOVERS
    # ================================================================
    
    def add_quality_handover(self, handover_data: Dict) -> bool:
        """Add quality handover to production"""
        try:
            self.cursor.execute("""
                INSERT INTO quality_handovers (
                    cabinet_id, project_name, sales_order_no, pdf_path,
                    excel_path, session_path, total_punches, open_punches,
                    closed_punches, handed_over_by, handed_over_date, status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                handover_data.get('cabinet_id'),
                handover_data.get('project_name'),
                handover_data.get('sales_order_no'),
                handover_data.get('pdf_path'),
                handover_data.get('excel_path'),
                handover_data.get('session_path'),
                handover_data.get('total_punches', 0),
                handover_data.get('open_punches', 0),
                handover_data.get('closed_punches', 0),
                handover_data.get('handed_over_by'),
                handover_data.get('handed_over_date', datetime.now().isoformat()),
                'pending_production'
            ))
            self.conn.commit()
            return True
        except sqlite3.IntegrityError:
            # Already handed over
            return False
        except Exception as e:
            print(f"Error adding handover: {e}")
            self.conn.rollback()
            return False
    
    def update_production_received(self, cabinet_id: str, user: str, 
                                   remarks: Optional[str] = None) -> bool:
        """Mark item as received by production"""
        try:
            self.cursor.execute("""
                UPDATE quality_handovers 
                SET production_received_by = ?,
                    production_received_date = ?,
                    production_remarks = ?,
                    status = 'in_production'
                WHERE cabinet_id = ?
            """, (user, datetime.now().isoformat(), remarks, cabinet_id))
            
            self.conn.commit()
            return True
        except Exception as e:
            print(f"Error updating production received: {e}")
            self.conn.rollback()
            return False
    
    def update_production_completed(self, cabinet_id: str, user: str, 
                                    remarks: Optional[str] = None) -> bool:
        """Mark production work as completed"""
        try:
            self.cursor.execute("""
                UPDATE quality_handovers 
                SET rework_completed_by = ?,
                    rework_completed_date = ?,
                    production_remarks = ?,
                    status = 'pending_quality_verification'
                WHERE cabinet_id = ?
            """, (user, datetime.now().isoformat(), remarks, cabinet_id))
            
            self.conn.commit()
            return True
        except Exception as e:
            print(f"Error updating production completed: {e}")
            self.conn.rollback()
            return False
    
    def update_quality_verification(self, cabinet_id: str, status: str, 
                                    user: str) -> bool:
        """Update quality verification status"""
        try:
            self.cursor.execute("""
                UPDATE quality_handovers 
                SET quality_verified_by = ?,
                    quality_verified_date = ?,
                    status = ?
                WHERE cabinet_id = ?
            """, (user, datetime.now().isoformat(), status, cabinet_id))
            
            self.conn.commit()
            return True
        except Exception as e:
            print(f"Error updating quality verification: {e}")
            self.conn.rollback()
            return False
    
    def get_pending_production_items(self) -> List[Dict]:
        """Get items pending production receipt"""
        self.cursor.execute("""
            SELECT * FROM quality_handovers 
            WHERE status = 'pending_production'
            ORDER BY handed_over_date DESC
        """)
        
        return [dict(row) for row in self.cursor.fetchall()]
    
    def get_pending_quality_items(self) -> List[Dict]:
        """Get items pending quality verification"""
        self.cursor.execute("""
            SELECT * FROM quality_handovers 
            WHERE status = 'pending_quality_verification'
            ORDER BY rework_completed_date DESC
        """)
        
        return [dict(row) for row in self.cursor.fetchall()]
    
    def get_handover_by_cabinet(self, cabinet_id: str) -> Optional[Dict]:
        """Get handover record by cabinet ID"""
        self.cursor.execute("""
            SELECT * FROM quality_handovers WHERE cabinet_id = ?
        """, (cabinet_id,))
        
        row = self.cursor.fetchone()
        return dict(row) if row else None
    
    # ================================================================
    # UTILITY METHODS
    # ================================================================
    
    def close(self):
        """Close database connection"""
        if self.conn:
            self.conn.close()
    
    def __enter__(self):
        """Context manager entry"""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit"""
        self.close()
