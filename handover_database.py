"""
Handover Database Manager - SQLite Version
Manages Quality <-> Production handover workflow using SQLite
Integrated with inspection_tool.db
"""

import sqlite3
from datetime import datetime
from typing import List, Dict, Optional
import os


class HandoverDB:
    """Manages handover records between Quality and Production using SQLite"""
    
    def __init__(self, db_path: str = None):
        """Initialize database at specified path
        
        Args:
            db_path: Path to SQLite database file. If None, uses inspection_tool.db
        """
        if db_path is None:
            # Default to inspection_tool.db in same directory as this script
            db_path = os.path.join(os.path.dirname(__file__), "inspection_tool.db")
        
        self.db_path = db_path
        self._init_tables()
        self._migrate_database()  # Apply any pending migrations
    
    def _init_tables(self):
        """Create handover tables if they don't exist"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Quality to Production handover table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS quality_to_production (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cabinet_id TEXT NOT NULL,
                project_name TEXT NOT NULL,
                sales_order_no TEXT,
                pdf_path TEXT,
                excel_path TEXT,
                session_path TEXT,
                total_punches INTEGER DEFAULT 0,
                open_punches INTEGER DEFAULT 0,
                closed_punches INTEGER DEFAULT 0,
                handed_over_by TEXT,
                handed_over_date TEXT,
                status TEXT DEFAULT 'pending',
                received_by TEXT,
                received_date TEXT,
                completed_by TEXT,
                completed_date TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Production to Quality handback table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS production_to_quality (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cabinet_id TEXT NOT NULL,
                project_name TEXT NOT NULL,
                sales_order_no TEXT,
                pdf_path TEXT,
                excel_path TEXT,
                session_path TEXT,
                rework_completed_by TEXT,
                rework_completed_date TEXT,
                production_remarks TEXT,
                status TEXT DEFAULT 'pending',
                verified_by TEXT,
                verified_date TEXT,
                verification_notes TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Create indexes for faster lookups
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_qtp_cabinet 
            ON quality_to_production(cabinet_id)
        ''')
        
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_qtp_status 
            ON quality_to_production(status)
        ''')
        
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_ptq_cabinet 
            ON production_to_quality(cabinet_id)
        ''')
        
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_ptq_status 
            ON production_to_quality(status)
        ''')
        
        conn.commit()
        conn.close()
        print("✓ Handover database tables initialized")
    
    def _migrate_database(self):
        """Apply database migrations for schema updates"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        migrations_applied = []
        
        # Migration 1: Add verification_notes column to production_to_quality
        try:
            cursor.execute("SELECT verification_notes FROM production_to_quality LIMIT 1")
        except sqlite3.OperationalError:
            # Column doesn't exist, add it
            try:
                cursor.execute('''
                    ALTER TABLE production_to_quality 
                    ADD COLUMN verification_notes TEXT
                ''')
                migrations_applied.append("Added verification_notes column to production_to_quality")
            except Exception as e:
                print(f"⚠️ Migration error (verification_notes): {e}")
        
        # Future migrations can be added here as needed
        # Example:
        # Migration 2: Add another column
        # try:
        #     cursor.execute("SELECT new_column FROM production_to_quality LIMIT 1")
        # except sqlite3.OperationalError:
        #     cursor.execute("ALTER TABLE production_to_quality ADD COLUMN new_column TEXT")
        #     migrations_applied.append("Added new_column")
        
        if migrations_applied:
            conn.commit()
            for msg in migrations_applied:
                print(f"✓ Migration: {msg}")
        
        conn.close()
    
    # ================================================================
    # QUALITY TO PRODUCTION HANDOVER
    # ================================================================
    
    def add_quality_handover(self, handover_data: dict) -> bool:
        """
        Add a new Quality -> Production handover
        
        Args:
            handover_data: Dict containing:
                - cabinet_id: str
                - project_name: str
                - sales_order_no: str
                - pdf_path: str
                - excel_path: str
                - session_path: str
                - total_punches: int
                - open_punches: int
                - closed_punches: int
                - handed_over_by: str
                - handed_over_date: str (ISO format)
        
        Returns:
            bool: True if successful, False if already exists
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Check if already handed over (pending or in_progress)
            cursor.execute('''
                SELECT id FROM quality_to_production
                WHERE cabinet_id = ? AND status IN ('pending', 'in_progress')
            ''', (handover_data['cabinet_id'],))
            
            existing = cursor.fetchone()
            
            if existing:
                conn.close()
                return False  # Already in production queue
            
            # Insert new handover
            cursor.execute('''
                INSERT INTO quality_to_production (
                    cabinet_id, project_name, sales_order_no,
                    pdf_path, excel_path, session_path,
                    total_punches, open_punches, closed_punches,
                    handed_over_by, handed_over_date, status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending')
            ''', (
                handover_data['cabinet_id'],
                handover_data['project_name'],
                handover_data.get('sales_order_no', ''),
                handover_data.get('pdf_path', ''),
                handover_data.get('excel_path', ''),
                handover_data.get('session_path', ''),
                handover_data.get('total_punches', 0),
                handover_data.get('open_punches', 0),
                handover_data.get('closed_punches', 0),
                handover_data.get('handed_over_by', ''),
                handover_data.get('handed_over_date', datetime.now().isoformat())
            ))
            
            conn.commit()
            conn.close()
            print(f"✓ Quality handover added: {handover_data['cabinet_id']}")
            return True
            
        except Exception as e:
            print(f"Error adding quality handover: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def get_pending_production_items(self) -> List[Dict]:
        """Get all items pending in production (pending or in_progress)"""
        try:
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT * FROM quality_to_production
                WHERE status IN ('pending', 'in_progress')
                ORDER BY handed_over_date DESC
            ''')
            
            rows = cursor.fetchall()
            conn.close()
            
            return [dict(row) for row in rows]
            
        except Exception as e:
            print(f"Error getting pending production items: {e}")
            return []
    
    def update_production_status(self, cabinet_id: str, status: str, user: str = None) -> bool:
        """Update production status for an item
        
        Args:
            cabinet_id: Cabinet identifier
            status: New status ('in_progress', 'completed')
            user: Username (optional)
        
        Returns:
            bool: True if successful
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Build update query based on status
            if status == 'in_progress':
                cursor.execute('''
                    UPDATE quality_to_production
                    SET status = ?,
                        received_by = COALESCE(received_by, ?),
                        received_date = COALESCE(received_date, ?),
                        updated_at = ?
                    WHERE cabinet_id = ?
                ''', (status, user, datetime.now().isoformat(), 
                      datetime.now().isoformat(), cabinet_id))
            
            elif status == 'completed':
                cursor.execute('''
                    UPDATE quality_to_production
                    SET status = ?,
                        completed_by = ?,
                        completed_date = ?,
                        updated_at = ?
                    WHERE cabinet_id = ?
                ''', (status, user, datetime.now().isoformat(), 
                      datetime.now().isoformat(), cabinet_id))
            
            else:
                cursor.execute('''
                    UPDATE quality_to_production
                    SET status = ?,
                        updated_at = ?
                    WHERE cabinet_id = ?
                ''', (status, datetime.now().isoformat(), cabinet_id))
            
            conn.commit()
            affected = cursor.rowcount
            conn.close()
            
            if affected > 0:
                print(f"✓ Production status updated: {cabinet_id} -> {status}")
                return True
            return False
            
        except Exception as e:
            print(f"Error updating production status: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    # ================================================================
    # PRODUCTION TO QUALITY HANDBACK
    # ================================================================
    
    def add_production_handback(self, handback_data: dict) -> bool:
        """
        Add a Production -> Quality handback
        
        Args:
            handback_data: Dict containing:
                - cabinet_id: str
                - project_name: str
                - sales_order_no: str
                - pdf_path: str
                - excel_path: str
                - session_path: str
                - rework_completed_by: str
                - rework_completed_date: str
                - production_remarks: str (optional)
        
        Returns:
            bool: True if successful
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Mark quality handover as completed
            cursor.execute('''
                UPDATE quality_to_production
                SET status = 'completed',
                    updated_at = ?
                WHERE cabinet_id = ?
            ''', (datetime.now().isoformat(), handback_data['cabinet_id']))
            
            # Insert handback record
            cursor.execute('''
                INSERT INTO production_to_quality (
                    cabinet_id, project_name, sales_order_no,
                    pdf_path, excel_path, session_path,
                    rework_completed_by, rework_completed_date,
                    production_remarks, status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending')
            ''', (
                handback_data['cabinet_id'],
                handback_data['project_name'],
                handback_data.get('sales_order_no', ''),
                handback_data.get('pdf_path', ''),
                handback_data.get('excel_path', ''),
                handback_data.get('session_path', ''),
                handback_data.get('rework_completed_by', ''),
                handback_data.get('rework_completed_date', datetime.now().isoformat()),
                handback_data.get('production_remarks', '')
            ))
            
            conn.commit()
            conn.close()
            print(f"✓ Production handback added: {handback_data['cabinet_id']}")
            return True
            
        except Exception as e:
            print(f"Error adding production handback: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def get_pending_quality_items(self) -> List[Dict]:
        """Get all items pending quality verification"""
        try:
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT * FROM production_to_quality
                WHERE status = 'pending'
                ORDER BY rework_completed_date DESC
            ''')
            
            rows = cursor.fetchall()
            conn.close()
            
            return [dict(row) for row in rows]
            
        except Exception as e:
            print(f"Error getting pending quality items: {e}")
            return []
    
    def is_in_rework_queue(self, cabinet_id: str) -> bool:
        """Check if a cabinet is currently in the rework verification queue
        
        Args:
            cabinet_id: Cabinet identifier
            
        Returns:
            bool: True if in pending rework queue
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT id FROM production_to_quality
                WHERE cabinet_id = ? AND status = 'pending'
            ''', (cabinet_id,))
            
            result = cursor.fetchone()
            conn.close()
            
            return result is not None
            
        except Exception as e:
            print(f"Error checking rework queue: {e}")
            return False
    
    def verify_production_item(self, cabinet_id: str, verified_by: str = None, 
                               verification_notes: str = None, mark_as_closed: bool = False) -> bool:
        """Mark a production handback as verified (or remove from queue)
        
        Args:
            cabinet_id: Cabinet identifier
            verified_by: Username who verified
            verification_notes: Optional notes about verification
            mark_as_closed: If True, marks as 'closed', otherwise 'verified'
        
        Returns:
            bool: True if successful
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            status = 'closed' if mark_as_closed else 'verified'
            
            cursor.execute('''
                UPDATE production_to_quality
                SET status = ?,
                    verified_by = ?,
                    verified_date = ?,
                    verification_notes = ?,
                    updated_at = ?
                WHERE cabinet_id = ? AND status = 'pending'
            ''', (status, verified_by, datetime.now().isoformat(), 
                  verification_notes, datetime.now().isoformat(), cabinet_id))
            
            conn.commit()
            affected = cursor.rowcount
            conn.close()
            
            if affected > 0:
                print(f"✓ Production item verified: {cabinet_id} -> {status}")
                return True
            else:
                print(f"⚠️ No pending item found for {cabinet_id}")
                return False
            
        except Exception as e:
            print(f"Error verifying production item: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def remove_from_rework_queue(self, cabinet_id: str, removed_by: str = None, 
                                 reason: str = None) -> bool:
        """Remove a cabinet from the rework verification queue
        
        This is useful when:
        - Cabinet needs to be re-handed over to production
        - Cabinet was incorrectly added to rework queue
        - Cabinet needs to return to quality inspection
        
        Args:
            cabinet_id: Cabinet identifier
            removed_by: Username who removed it
            reason: Reason for removal
            
        Returns:
            bool: True if successful
        """
        notes = f"Removed by {removed_by or 'Unknown'}"
        if reason:
            notes += f" - Reason: {reason}"
        
        return self.verify_production_item(
            cabinet_id, 
            verified_by=removed_by,
            verification_notes=notes,
            mark_as_closed=False
        )
    
    def update_quality_verification(self, cabinet_id: str, status: str, user: str = None) -> bool:
        """Update quality verification status
        
        Args:
            cabinet_id: Cabinet identifier
            status: New status ('verified', 'closed')
            user: Username (optional)
        
        Returns:
            bool: True if successful
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                UPDATE production_to_quality
                SET status = ?,
                    verified_by = ?,
                    verified_date = ?,
                    updated_at = ?
                WHERE cabinet_id = ?
            ''', (status, user, datetime.now().isoformat(), 
                  datetime.now().isoformat(), cabinet_id))
            
            conn.commit()
            affected = cursor.rowcount
            conn.close()
            
            if affected > 0:
                print(f"✓ Quality verification updated: {cabinet_id} -> {status}")
                return True
            return False
            
        except Exception as e:
            print(f"Error updating quality verification: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    # ================================================================
    # UTILITY FUNCTIONS
    # ================================================================
    
    def get_item_by_cabinet_id(self, cabinet_id: str, queue: str = "quality_to_production") -> Optional[Dict]:
        """Get handover item by cabinet ID
        
        Args:
            cabinet_id: Cabinet identifier
            queue: Which queue to search ('quality_to_production' or 'production_to_quality')
        
        Returns:
            Dict with item data or None if not found
        """
        try:
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            
            cursor.execute(f'''
                SELECT * FROM {queue}
                WHERE cabinet_id = ?
                ORDER BY created_at DESC
                LIMIT 1
            ''', (cabinet_id,))
            
            row = cursor.fetchone()
            conn.close()
            
            return dict(row) if row else None
            
        except Exception as e:
            print(f"Error getting item by cabinet ID: {e}")
            return None
    
    def get_handover_by_cabinet(self, cabinet_id: str) -> Optional[Dict]:
        """Get most recent production handback for a cabinet
        
        Args:
            cabinet_id: Cabinet identifier
        
        Returns:
            Dict with handback data or None
        """
        return self.get_item_by_cabinet_id(cabinet_id, "production_to_quality")
    
    def get_all_handovers(self) -> Dict:
        """Get complete handover data from both tables
        
        Returns:
            Dict with 'quality_to_production' and 'production_to_quality' lists
        """
        try:
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            
            # Get quality to production
            cursor.execute('SELECT * FROM quality_to_production ORDER BY handed_over_date DESC')
            qtp_rows = cursor.fetchall()
            
            # Get production to quality
            cursor.execute('SELECT * FROM production_to_quality ORDER BY rework_completed_date DESC')
            ptq_rows = cursor.fetchall()
            
            conn.close()
            
            return {
                'quality_to_production': [dict(row) for row in qtp_rows],
                'production_to_quality': [dict(row) for row in ptq_rows]
            }
            
        except Exception as e:
            print(f"Error getting all handovers: {e}")
            return {
                'quality_to_production': [],
                'production_to_quality': []
            }
    
    def cleanup_completed(self, days_old: int = 30):
        """Remove completed items older than specified days
        
        Args:
            days_old: Number of days (items older than this will be deleted)
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cutoff_date = datetime.now().timestamp() - (days_old * 24 * 60 * 60)
            cutoff_iso = datetime.fromtimestamp(cutoff_date).isoformat()
            
            # Clean quality_to_production (completed items)
            cursor.execute('''
                DELETE FROM quality_to_production
                WHERE status = 'completed' AND completed_date < ?
            ''', (cutoff_iso,))
            
            qtp_deleted = cursor.rowcount
            
            # Clean production_to_quality (closed and verified items)
            cursor.execute('''
                DELETE FROM production_to_quality
                WHERE status IN ('closed', 'verified') AND verified_date < ?
            ''', (cutoff_iso,))
            
            ptq_deleted = cursor.rowcount
            
            conn.commit()
            conn.close()
            
            print(f"✓ Cleanup: Removed {qtp_deleted} from quality_to_production, "
                  f"{ptq_deleted} from production_to_quality")
            
        except Exception as e:
            print(f"Error during cleanup: {e}")
            import traceback
            traceback.print_exc()
