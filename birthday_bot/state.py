"""Idempotency state tracking."""

import logging
import csv
from pathlib import Path
from datetime import datetime
from typing import Set, Tuple, Optional

logger = logging.getLogger(__name__)


class StateError(Exception):
    """State management error."""
    pass


class SentLog:
    """Track sent emails for idempotency."""
    
    def __init__(self, log_path: str = "state/sent_log.csv"):
        """
        Initialize sent log.
        
        Args:
            log_path: Path to CSV log file
        """
        self.log_path = Path(log_path)
        self.log_path.parent.mkdir(parents=True, exist_ok=True)
        self._load_sent()
    
    def _load_sent(self):
        """Load sent records from CSV."""
        self.sent = set()  # Set of (date, member_id, email) tuples
        
        if not self.log_path.exists():
            return
        
        try:
            with open(self.log_path, 'r', encoding='utf-8', newline='') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    if row.get('date') and row.get('member_id') and row.get('email'):
                        key = (row['date'], row['member_id'], row['email'])
                        self.sent.add(key)
            logger.debug(f"Loaded {len(self.sent)} sent records")
        except Exception as e:
            logger.warning(f"Failed to load sent log: {e}")
    
    def is_sent(self, date: str, member_id: str, email: str) -> bool:
        """
        Check if email was already sent.
        
        Args:
            date: Date in YYYY-MM-DD format
            member_id: Member ID
            email: Email address
            
        Returns:
            True if already sent, False otherwise
        """
        return (date, member_id, email) in self.sent
    
    def mark_sent(self, date: str, member_id: str, email: str, sent_at: Optional[datetime] = None):
        """
        Mark email as sent.
        
        Args:
            date: Date in YYYY-MM-DD format
            member_id: Member ID
            email: Email address
            sent_at: Timestamp (default: now)
        """
        if sent_at is None:
            sent_at = datetime.now()
        
        key = (date, member_id, email)
        self.sent.add(key)
        
        # Append to CSV
        self._append_record(date, member_id, email, sent_at)
    
    def _append_record(self, date: str, member_id: str, email: str, sent_at: datetime):
        """Append record to CSV file."""
        try:
            file_exists = self.log_path.exists()
            with open(self.log_path, 'a', encoding='utf-8', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=['date', 'member_id', 'email', 'sent_at'])
                if not file_exists:
                    writer.writeheader()
                writer.writerow({
                    'date': date,
                    'member_id': member_id,
                    'email': email,
                    'sent_at': sent_at.isoformat()
                })
        except Exception as e:
            logger.error(f"Failed to write sent log: {e}")
            raise StateError(f"Failed to record sent email: {e}")
    
    def get_sent_today(self, date: str) -> Set[str]:
        """
        Get set of member IDs sent for date.
        
        Args:
            date: Date in YYYY-MM-DD format
            
        Returns:
            Set of member IDs
        """
        return {member_id for (d, member_id, _) in self.sent if d == date}
