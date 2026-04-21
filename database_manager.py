import pg_sqlite_compat as sqlite3
import json
import os
from datetime import datetime
from typing import List, Dict, Optional, Tuple
from path_policy import (
    resolve_storage_location,
    to_absolute_path,
    to_relative_path,
    to_relative_storage_location,
)

class DatabaseManager:
    """Centralized PostgreSQL database manager for the Quality Inspection Tool"""
    _PATH_FIELDS = ("pdf_path", "excel_path", "session_path")
    
    def __init__(self, db_path: str):
        """Initialize database connection."""
        self.db_path = db_path
        self.conn = None
        self.cursor = None
        self._connect()
    
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

    def _serialize_project_data(self, data: Dict) -> Dict:
        normalized = dict(data or {})

        if 'storage_location' in normalized:
            normalized['storage_location'] = to_relative_storage_location(normalized.get('storage_location'))

        for field in self._PATH_FIELDS:
            if field in normalized:
                normalized[field] = to_relative_path(normalized.get(field))

        return normalized

    def _serialize_handover_data(self, data: Dict) -> Dict:
        normalized = dict(data or {})
        for field in self._PATH_FIELDS:
            if field in normalized:
                normalized[field] = to_relative_path(normalized.get(field))
        return normalized

    def _resolve_project_record(self, record: Dict) -> Dict:
        resolved = dict(record or {})
        resolved['storage_location'] = resolve_storage_location(resolved.get('storage_location'))
        for field in self._PATH_FIELDS:
            if field in resolved:
                resolved[field] = to_absolute_path(resolved.get(field))
        return resolved

    def _resolve_handover_record(self, record: Dict) -> Dict:
        resolved = dict(record or {})
        for field in self._PATH_FIELDS:
            if field in resolved:
                resolved[field] = to_absolute_path(resolved.get(field))
        return resolved
    
    # ================================================================
    # PROJECT MANAGEMENT
    # ================================================================
    
    def add_project(self, project_data: Dict) -> bool:
        """Add a new project to the database"""
        try:
            project_data = self._serialize_project_data(project_data)

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
                project_data.get('storage_location') or '.',
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
            updates = self._serialize_project_data(updates)

            # Build dynamic UPDATE query while assigning last_accessed only once.
            set_parts = []
            values = []

            for key, value in updates.items():
                if key == "last_accessed":
                    continue
                set_parts.append(f"{key} = ?")
                values.append(value)

            requested_last_accessed = updates.get("last_accessed")
            if requested_last_accessed in (None, ""):
                requested_last_accessed = datetime.now().isoformat()

            set_parts.append("last_accessed = ?")
            values.append(requested_last_accessed)
            values.append(cabinet_id)
            set_clause = ", ".join(set_parts)
            
            self.cursor.execute(f"""
                UPDATE projects 
                SET {set_clause}
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
            return self._resolve_project_record(dict(row))
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
        
        return [self._resolve_project_record(dict(row)) for row in self.cursor.fetchall()]
    
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
        
        return [self._resolve_project_record(dict(row)) for row in self.cursor.fetchall()]
    
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
        return resolve_storage_location(row[0]) if row else None
    
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
        
        return [self._resolve_project_record(dict(row)) for row in self.cursor.fetchall()]
    
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
            handover_data = self._serialize_handover_data(handover_data)

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
        
        return [self._resolve_handover_record(dict(row)) for row in self.cursor.fetchall()]
    
    def get_pending_quality_items(self) -> List[Dict]:
        """Get items pending quality verification"""
        self.cursor.execute("""
            SELECT * FROM quality_handovers 
            WHERE status = 'pending_quality_verification'
            ORDER BY rework_completed_date DESC
        """)
        
        return [self._resolve_handover_record(dict(row)) for row in self.cursor.fetchall()]
    
    def get_handover_by_cabinet(self, cabinet_id: str) -> Optional[Dict]:
        """Get handover record by cabinet ID"""
        self.cursor.execute("""
            SELECT * FROM quality_handovers WHERE cabinet_id = ?
        """, (cabinet_id,))
        
        row = self.cursor.fetchone()
        return self._resolve_handover_record(dict(row)) if row else None
    
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
