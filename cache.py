#!/usr/bin/env python3
"""Disk-based cache system for PPTX thumbnails to reduce memory usage."""

import hashlib
import shutil
import tempfile
import time
from pathlib import Path
from typing import Optional, Dict, Any
import json

class DiskCache:
    """Disk-based cache for PPTX thumbnails with LRU eviction."""
    
    def __init__(self, cache_dir: Optional[Path] = None, max_size_mb: int = 200):
        self.cache_dir = cache_dir or Path(tempfile.gettempdir()) / "pptx_cache"
        self.max_size_mb = max_size_mb
        self.max_size_bytes = max_size_mb * 1024 * 1024
        self.metadata_file = self.cache_dir / "cache_metadata.json"
        self.metadata: Dict[str, Dict[str, Any]] = {}
        
        # Ensure cache directory exists
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        
        # Load existing metadata
        self._load_metadata()
        
        # Cleanup if needed
        self._enforce_size_limit()
    
    def _load_metadata(self) -> None:
        """Load cache metadata from disk."""
        if self.metadata_file.exists():
            try:
                with open(self.metadata_file, 'r', encoding='utf-8') as f:
                    self.metadata = json.load(f)
            except (json.JSONDecodeError, IOError):
                self.metadata = {}
    
    def _save_metadata(self) -> None:
        """Save cache metadata to disk."""
        try:
            with open(self.metadata_file, 'w', encoding='utf-8') as f:
                json.dump(self.metadata, f, indent=2)
        except IOError:
            pass  # Metadata save failure is not critical
    
    def _get_cache_key(self, pptx_path: Path, slide_index: int, dpi: int) -> str:
        """Generate cache key for a specific slide."""
        # Use file hash, slide index, and DPI for cache key
        file_hash = self._get_file_hash(pptx_path)
        return f"{file_hash}_{slide_index:03d}_{dpi}"
    
    def _get_file_hash(self, file_path: Path) -> str:
        """Get hash of file for cache key."""
        # Use file size and modification time for quick hash
        stat = file_path.stat()
        hash_input = f"{file_path.name}_{stat.st_size}_{stat.st_mtime}"
        return hashlib.md5(hash_input.encode()).hexdigest()[:12]
    
    def _get_cache_path(self, cache_key: str) -> Path:
        """Get full path for cached file."""
        return self.cache_dir / f"{cache_key}.png"
    
    def _enforce_size_limit(self) -> None:
        """Remove oldest files to stay within size limit."""
        total_size = self._get_total_size()
        
        if total_size <= self.max_size_bytes:
            return
        
        # Sort by last access time and remove oldest
        items_by_time = sorted(
            self.metadata.items(),
            key=lambda x: x[1].get('last_access', 0)
        )
        
        for cache_key, meta in items_by_time:
            cache_path = self._get_cache_path(cache_key)
            if cache_path.exists():
                try:
                    cache_path.unlink()
                    del self.metadata[cache_key]
                    total_size -= meta.get('size', 0)
                    if total_size <= self.max_size_bytes * 0.8:  # Leave 20% headroom
                        break
                except OSError:
                    continue
        
        self._save_metadata()
    
    def _get_total_size(self) -> int:
        """Calculate total cache size in bytes."""
        total = 0
        for cache_key, meta in self.metadata.items():
            cache_path = self._get_cache_path(cache_key)
            if cache_path.exists():
                total += cache_path.stat().st_size
            else:
                # Remove metadata for missing files
                del self.metadata[cache_key]
        return total
    
    def get(self, pptx_path: Path, slide_index: int, dpi: int) -> Optional[Path]:
        """Get cached thumbnail if available."""
        cache_key = self._get_cache_key(pptx_path, slide_index, dpi)
        cache_path = self._get_cache_path(cache_key)
        
        if not cache_path.exists():
            return None
        
        # Update access time
        if cache_key in self.metadata:
            self.metadata[cache_key]['last_access'] = time.time()
            self._save_metadata()
        
        return cache_path
    
    def put(self, pptx_path: Path, slide_index: int, dpi: int, image_path: Path) -> Path:
        """Store thumbnail in cache."""
        cache_key = self._get_cache_key(pptx_path, slide_index, dpi)
        cache_path = self._get_cache_path(cache_key)
        
        # Copy image to cache
        shutil.copy2(image_path, cache_path)
        
        # Update metadata
        file_size = cache_path.stat().st_size
        self.metadata[cache_key] = {
            'pptx_name': pptx_path.name,
            'slide_index': slide_index,
            'dpi': dpi,
            'size': file_size,
            'created': time.time(),
            'last_access': time.time()
        }
        
        self._save_metadata()
        
        # Enforce size limit
        self._enforce_size_limit()
        
        return cache_path
    
    def clear(self) -> None:
        """Clear all cached files."""
        for cache_key in list(self.metadata.keys()):
            cache_path = self._get_cache_path(cache_key)
            if cache_path.exists():
                try:
                    cache_path.unlink()
                except OSError:
                    continue
        
        self.metadata.clear()
        self._save_metadata()
    
    def get_stats(self) -> Dict[str, Any]:
        """Get cache statistics."""
        total_size = self._get_total_size()
        file_count = len(self.metadata)
        
        return {
            'total_size_mb': round(total_size / 1024 / 1024, 2),
            'file_count': file_count,
            'max_size_mb': self.max_size_mb,
            'usage_percent': round((total_size / self.max_size_bytes) * 100, 1)
        }

# Global cache instance
_global_cache: Optional[DiskCache] = None

def get_cache() -> DiskCache:
    """Get global cache instance."""
    global _global_cache
    if _global_cache is None:
        _global_cache = DiskCache()
    return _global_cache

def cleanup_cache() -> None:
    """Cleanup global cache."""
    global _global_cache
    if _global_cache is not None:
        _global_cache.clear()
        _global_cache = None
