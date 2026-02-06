"""
State tracking for Office Automation Foundation.

Provides persistent state tracking to prevent duplicate operations:
- Track sent emails by Message-ID or content hash
- Track processed items to allow resume after crash
- Generic state storage for custom tracking needs

State is stored in JSON files for human readability and easy debugging.

Usage:
    from core.state import StateTracker
    
    tracker = StateTracker(state_dir=Path("./state"), state_name="email_send")
    
    if not tracker.is_processed(email_id):
        send_email(email)
        tracker.mark_processed(email_id, metadata={"recipient": "user@example.com"})
    
    tracker.save()  # Persist to disk
"""

import hashlib
import json
import logging
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional, Set

logger = logging.getLogger(__name__)


class StateTracker:
    """
    Persistent state tracker for preventing duplicate operations.
    
    Stores processed item IDs and optional metadata in a JSON file.
    Designed to be resilient to crashes - data is persisted on save().
    
    Attributes:
        state_file: Path to the JSON state file.
        processed_ids: Set of processed item identifiers.
        metadata: Dictionary of metadata for processed items.
    """
    
    def __init__(
        self,
        state_dir: Path,
        state_name: str,
        auto_save: bool = True,
        auto_save_interval: int = 10
    ):
        """
        Initialize state tracker.
        
        Args:
            state_dir: Directory for state files.
            state_name: Name for this state (used in filename).
            auto_save: If True, automatically save after auto_save_interval changes.
            auto_save_interval: Number of changes between auto-saves.
        """
        self.state_dir = Path(state_dir)
        self.state_name = state_name
        self.auto_save = auto_save
        self.auto_save_interval = auto_save_interval
        
        self.state_file = self.state_dir / f"{state_name}_state.json"
        
        # Internal state
        self.processed_ids: Set[str] = set()
        self.metadata: Dict[str, Dict[str, Any]] = {}
        self._changes_since_save = 0
        self._created_at: Optional[str] = None
        self._last_modified: Optional[str] = None
        
        # Ensure directory exists
        self.state_dir.mkdir(parents=True, exist_ok=True)
        
        # Load existing state if available
        self._load()
    
    def _load(self) -> None:
        """Load state from disk if file exists."""
        if not self.state_file.exists():
            self._created_at = datetime.now().isoformat()
            logger.debug(f"Creating new state file: {self.state_file}")
            return
        
        try:
            with open(self.state_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            
            self.processed_ids = set(data.get("processed_ids", []))
            self.metadata = data.get("metadata", {})
            self._created_at = data.get("created_at")
            self._last_modified = data.get("last_modified")
            
            logger.info(
                f"Loaded state from {self.state_file}: "
                f"{len(self.processed_ids)} processed items"
            )
        
        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse state file {self.state_file}: {e}")
            # Backup corrupt file and start fresh
            backup_path = self.state_file.with_suffix(".json.corrupt")
            self.state_file.rename(backup_path)
            logger.warning(f"Backed up corrupt state to: {backup_path}")
            self._created_at = datetime.now().isoformat()
        
        except Exception as e:
            logger.error(f"Failed to load state file: {e}")
            self._created_at = datetime.now().isoformat()
    
    def save(self) -> None:
        """Persist current state to disk."""
        self._last_modified = datetime.now().isoformat()
        
        data = {
            "state_name": self.state_name,
            "created_at": self._created_at,
            "last_modified": self._last_modified,
            "total_processed": len(self.processed_ids),
            "processed_ids": sorted(self.processed_ids),
            "metadata": self.metadata,
        }
        
        # Write to temp file first for atomicity
        temp_file = self.state_file.with_suffix(".json.tmp")
        
        try:
            with open(temp_file, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            
            # Replace original with temp file
            temp_file.replace(self.state_file)
            self._changes_since_save = 0
            
            logger.debug(f"State saved: {len(self.processed_ids)} items to {self.state_file}")
        
        except Exception as e:
            logger.error(f"Failed to save state: {e}")
            if temp_file.exists():
                temp_file.unlink()
            raise
    
    def _maybe_auto_save(self) -> None:
        """Auto-save if enabled and interval reached."""
        if self.auto_save and self._changes_since_save >= self.auto_save_interval:
            self.save()
    
    def is_processed(self, item_id: str) -> bool:
        """
        Check if an item has been processed.
        
        Args:
            item_id: Unique identifier for the item.
            
        Returns:
            True if the item was previously processed.
        """
        return str(item_id) in self.processed_ids
    
    def mark_processed(
        self,
        item_id: str,
        metadata: Optional[Dict[str, Any]] = None
    ) -> None:
        """
        Mark an item as processed.
        
        Args:
            item_id: Unique identifier for the item.
            metadata: Optional metadata to store (e.g., timestamp, recipient).
        """
        item_id = str(item_id)
        
        if item_id in self.processed_ids:
            logger.debug(f"Item already marked as processed: {item_id}")
            return
        
        self.processed_ids.add(item_id)
        
        if metadata:
            self.metadata[item_id] = {
                **metadata,
                "processed_at": datetime.now().isoformat(),
            }
        else:
            self.metadata[item_id] = {
                "processed_at": datetime.now().isoformat(),
            }
        
        self._changes_since_save += 1
        self._maybe_auto_save()
        
        logger.debug(f"Marked as processed: {item_id}")
    
    def unmark_processed(self, item_id: str) -> bool:
        """
        Remove an item from the processed set.
        
        Useful for allowing re-processing of items.
        
        Args:
            item_id: Unique identifier for the item.
            
        Returns:
            True if the item was removed, False if it wasn't processed.
        """
        item_id = str(item_id)
        
        if item_id not in self.processed_ids:
            return False
        
        self.processed_ids.discard(item_id)
        self.metadata.pop(item_id, None)
        
        self._changes_since_save += 1
        self._maybe_auto_save()
        
        logger.debug(f"Unmarked as processed: {item_id}")
        return True
    
    def get_metadata(self, item_id: str) -> Optional[Dict[str, Any]]:
        """
        Get metadata for a processed item.
        
        Args:
            item_id: Unique identifier for the item.
            
        Returns:
            Metadata dictionary, or None if item not found.
        """
        return self.metadata.get(str(item_id))
    
    def get_processed_count(self) -> int:
        """Return the number of processed items."""
        return len(self.processed_ids)
    
    def get_all_processed_ids(self) -> Set[str]:
        """Return a copy of all processed item IDs."""
        return self.processed_ids.copy()
    
    def clear(self) -> None:
        """Clear all state (use with caution)."""
        self.processed_ids.clear()
        self.metadata.clear()
        self._changes_since_save += 1
        logger.warning(f"State cleared: {self.state_name}")


class ContentHashTracker(StateTracker):
    """
    State tracker that uses content hashing for duplicate detection.
    
    Useful when items don't have stable unique IDs, but you can
    compute a hash from their content.
    
    Example:
        tracker = ContentHashTracker(state_dir, "email_send")
        
        content_hash = tracker.compute_hash(f"{recipient}|{subject}|{body}")
        if not tracker.is_processed(content_hash):
            send_email(recipient, subject, body)
            tracker.mark_processed(content_hash)
    """
    
    @staticmethod
    def compute_hash(*content_parts: str) -> str:
        """
        Compute a SHA-256 hash from content parts.
        
        Args:
            content_parts: Strings to include in the hash.
            
        Returns:
            Hex string of the hash.
        """
        combined = "|".join(str(part) for part in content_parts)
        return hashlib.sha256(combined.encode("utf-8")).hexdigest()
    
    def is_content_processed(self, *content_parts: str) -> bool:
        """
        Check if content has been processed.
        
        Args:
            content_parts: Strings to include in the hash.
            
        Returns:
            True if matching content was previously processed.
        """
        content_hash = self.compute_hash(*content_parts)
        return self.is_processed(content_hash)
    
    def mark_content_processed(
        self,
        *content_parts: str,
        metadata: Optional[Dict[str, Any]] = None
    ) -> str:
        """
        Mark content as processed.
        
        Args:
            content_parts: Strings to include in the hash.
            metadata: Optional metadata to store.
            
        Returns:
            The computed hash.
        """
        content_hash = self.compute_hash(*content_parts)
        self.mark_processed(content_hash, metadata)
        return content_hash


class RunStateTracker:
    """
    Track state within a single run for resume capability.
    
    Stores progress through a list of items so processing can
    resume from where it left off if interrupted.
    
    Example:
        tracker = RunStateTracker(state_dir, "batch_send")
        
        start_index = tracker.get_resume_index()
        for i, item in enumerate(items[start_index:], start=start_index):
            process(item)
            tracker.update_progress(i)
    """
    
    def __init__(self, state_dir: Path, run_name: str):
        """
        Initialize run state tracker.
        
        Args:
            state_dir: Directory for state files.
            run_name: Name for this run.
        """
        self.state_dir = Path(state_dir)
        self.run_name = run_name
        self.state_file = self.state_dir / f"{run_name}_run.json"
        
        self.state_dir.mkdir(parents=True, exist_ok=True)
        
        self._data: Dict[str, Any] = {
            "run_name": run_name,
            "started_at": None,
            "last_updated": None,
            "current_index": 0,
            "total_items": 0,
            "completed": False,
            "custom": {},
        }
        
        self._load()
    
    def _load(self) -> None:
        """Load existing run state if available."""
        if not self.state_file.exists():
            return
        
        try:
            with open(self.state_file, "r", encoding="utf-8") as f:
                self._data = json.load(f)
            logger.info(
                f"Resuming run '{self.run_name}' from index {self._data['current_index']}"
            )
        except Exception as e:
            logger.warning(f"Could not load run state: {e}")
    
    def _save(self) -> None:
        """Save run state to disk."""
        self._data["last_updated"] = datetime.now().isoformat()
        
        temp_file = self.state_file.with_suffix(".json.tmp")
        with open(temp_file, "w", encoding="utf-8") as f:
            json.dump(self._data, f, indent=2)
        temp_file.replace(self.state_file)
    
    def start(self, total_items: int) -> None:
        """
        Start a new run or resume existing.
        
        Args:
            total_items: Total number of items to process.
        """
        if not self._data["started_at"]:
            self._data["started_at"] = datetime.now().isoformat()
        self._data["total_items"] = total_items
        self._data["completed"] = False
        self._save()
    
    def get_resume_index(self) -> int:
        """Get the index to resume from."""
        return self._data["current_index"]
    
    def update_progress(self, current_index: int) -> None:
        """
        Update progress to current index.
        
        Args:
            current_index: Index of the last successfully processed item.
        """
        self._data["current_index"] = current_index + 1  # Next to process
        self._save()
    
    def complete(self) -> None:
        """Mark the run as complete."""
        self._data["completed"] = True
        self._save()
        logger.info(f"Run '{self.run_name}' completed")
    
    def is_complete(self) -> bool:
        """Check if the run was completed."""
        return self._data.get("completed", False)
    
    def reset(self) -> None:
        """Reset run state for a fresh start."""
        if self.state_file.exists():
            self.state_file.unlink()
        self._data = {
            "run_name": self.run_name,
            "started_at": None,
            "last_updated": None,
            "current_index": 0,
            "total_items": 0,
            "completed": False,
            "custom": {},
        }
    
    def set_custom(self, key: str, value: Any) -> None:
        """Store custom data in run state."""
        self._data["custom"][key] = value
        self._save()
    
    def get_custom(self, key: str, default: Any = None) -> Any:
        """Retrieve custom data from run state."""
        return self._data["custom"].get(key, default)
