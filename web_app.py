#!/usr/bin/env python3
"""Web API for PowerPoint script slide generation - Render.com compatible."""

from __future__ import annotations

import os
import tempfile
import shutil
import asyncio
from collections import deque
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, List, Dict
import zipfile

from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
import uvicorn
import gc
import psutil
import time
import logging

# Import disk cache for memory optimization
try:
    from cache import get_cache, cleanup_cache
    CACHE_AVAILABLE = True
except ImportError:
    CACHE_AVAILABLE = False
    def get_cache():
        return None
    def cleanup_cache():
        pass

from app import generate_script_slides

app = FastAPI(
    title="PPTX Script Slides API",
    description="Convert PowerPoint notes to large text slides",
    version="1.0.0"
)


@app.on_event("startup")
async def _init_queue_workers() -> None:
    global queue_event, queue_lock, worker_task
    if queue_event is None:
        queue_event = asyncio.Event()
    if queue_lock is None:
        queue_lock = asyncio.Lock()
    if worker_task is None:
        worker_task = asyncio.create_task(_conversion_worker())

# Mount static files
app.mount("/static", StaticFiles(directory="static"), name="static")

# CORS for development
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class ConversionStatus(BaseModel):
    task_id: str
    status: str
    message: str
    download_url: Optional[str] = None
    logs: List[str] = Field(default_factory=list)
    queue_position: Optional[int] = None

# In-memory storage for demo (use Redis/database in production)
task_status: Dict[str, Dict[str, object]] = {}


@dataclass
class ConversionJob:
    task_id: str
    input_path: Path
    output_path: Path
    temp_dir: Path


# Sequential conversion queue (single worker)
conversion_queue: deque[ConversionJob] = deque()
queue_event: Optional[asyncio.Event] = None
queue_lock: Optional[asyncio.Lock] = None
worker_task: Optional[asyncio.Task] = None

LOGGER = logging.getLogger(__name__)


def _ensure_queue_primitives_initialized() -> None:
    """Ensure asyncio synchronization primitives exist."""
    global queue_event, queue_lock
    if queue_event is None:
        queue_event = asyncio.Event()
    if queue_lock is None:
        queue_lock = asyncio.Lock()


def _update_queue_positions_locked() -> None:
    """Update waiting job positions (requires queue_lock)."""
    for idx, job in enumerate(conversion_queue):
        status = task_status.get(job.task_id)
        if status is not None:
            status["queue_position"] = idx


async def _enqueue_job(job: ConversionJob) -> int:
    """Add job to queue and return its position."""
    _ensure_queue_primitives_initialized()
    assert queue_lock is not None
    async with queue_lock:
        conversion_queue.append(job)
        _update_queue_positions_locked()
        status = task_status.get(job.task_id)
        position = status.get("queue_position", len(conversion_queue) - 1) if status else len(conversion_queue) - 1
        if queue_event:
            queue_event.set()
    return position


async def _conversion_worker() -> None:
    """Background worker that processes conversion jobs sequentially."""
    while True:
        if queue_event is None or queue_lock is None:
            await asyncio.sleep(0.1)
            continue

        await queue_event.wait()

        while True:
            async with queue_lock:
                if conversion_queue:
                    job = conversion_queue.popleft()
                    _update_queue_positions_locked()
                    if not conversion_queue:
                        queue_event.clear()
                else:
                    queue_event.clear()
                    job = None

            if job is None:
                break

            await _run_conversion_job(job)


async def _run_conversion_job(job: ConversionJob) -> None:
    """Process a single conversion job in a worker thread."""
    status = task_status.get(job.task_id)
    if not status:
        shutil.rmtree(job.temp_dir, ignore_errors=True)
        return

    status["status"] = "processing"
    status["message"] = "Processing slides..."
    status["queue_position"] = 0

    try:
        await asyncio.to_thread(_execute_conversion_job, job)
    except Exception as exc:  # noqa: BLE001
        LOGGER.exception("Conversion job %s failed", job.task_id, exc_info=exc)
    finally:
        entry = task_status.get(job.task_id)
        if entry:
            entry["queue_position"] = None

        if queue_lock is not None:
            async with queue_lock:
                _update_queue_positions_locked()


def _execute_conversion_job(job: ConversionJob) -> None:
    """Perform the conversion work (runs in a thread)."""
    task_id = job.task_id
    input_path = job.input_path
    output_path = job.output_path
    temp_dir = job.temp_dir

    timestamp = lambda: time.strftime('%Y-%m-%d %H:%M:%S')  # noqa: E731

    try:
        def log_callback(message: str) -> None:
            entry = task_status.get(task_id)
            if not entry:
                return
            entry["message"] = message
            logs = entry.setdefault("logs", [])
            logs.append(f"{timestamp()} {message}")

        result_path = generate_script_slides(
            input_path,
            temp_dir,
            log_callback,
        )

        final_path = result_path
        if result_path != output_path:
            shutil.copy2(result_path, output_path)
            final_path = output_path

        entry = task_status.get(task_id)
        if entry is not None:
            entry.update(
                {
                    "status": "completed",
                    "message": "Conversion completed successfully!",
                    "download_url": f"/download/{task_id}",
                    "file_path": str(final_path),
                }
            )
            entry.setdefault("logs", []).append(
                f"{timestamp()} Conversion completed successfully!"
            )

    except Exception as exc:  # noqa: BLE001
        entry = task_status.get(task_id)
        if entry is not None:
            entry.update(
                {
                    "status": "failed",
                    "message": f"Conversion failed: {exc}",
                    "download_url": None,
                }
            )
            entry.setdefault("logs", []).append(
                f"{timestamp()} Conversion failed: {exc}"
            )
        LOGGER.exception("Conversion job %s raised an exception", task_id, exc_info=exc)
    finally:
        force_garbage_collection()
        entry = task_status.get(task_id)
        if entry is None or entry.get("status") != "completed":
            shutil.rmtree(temp_dir, ignore_errors=True)
# Memory monitoring and cleanup
MAX_MEMORY_MB = 400  # Maximum memory usage before forced cleanup
CLEANUP_INTERVAL = 300  # Check every 5 minutes
last_cleanup = time.time()

def _process_tree_memory_mb(proc: psutil.Process) -> float:
    """Return RSS usage of a process and all its children in MB."""
    try:
        total_rss = proc.memory_info().rss
        for child in proc.children(recursive=True):
            try:
                total_rss += child.memory_info().rss
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        return total_rss / 1024 / 1024
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        return 0.0


def get_memory_usage() -> float:
    """Get current process tree memory usage in MB."""
    try:
        process = psutil.Process()
        return _process_tree_memory_mb(process)
    except Exception:
        return 0.0

def force_garbage_collection() -> None:
    """Force garbage collection to free memory."""
    try:
        gc.collect()
    except Exception:
        pass

def auto_cleanup_if_needed() -> None:
    """Automatically cleanup if memory usage is too high."""
    global last_cleanup
    current_time = time.time()
    
    # Check memory usage periodically
    if current_time - last_cleanup > CLEANUP_INTERVAL:
        memory_mb = get_memory_usage()
        if memory_mb > MAX_MEMORY_MB:
            # Clean up old completed tasks
            old_tasks = []
            for task_id, status in task_status.items():
                if status.get("status") == "completed" and status.get("created_at", 0) < current_time - 3600:
                    old_tasks.append(task_id)
            
            for task_id in old_tasks:
                cleanup_task(task_id)
            
            # Force garbage collection
            force_garbage_collection()
        
        last_cleanup = current_time

@app.api_route("/", methods=["GET", "HEAD"])
async def root():
    """Serve the main web interface (HEAD returns headers only)."""
    return FileResponse("static/index.html")

@app.get("/health")
async def health():
    """Health check endpoint with memory info."""
    auto_cleanup_if_needed()
    memory_mb = get_memory_usage()
    
    # Include cache statistics if available
    cache_stats = {}
    if CACHE_AVAILABLE:
        cache = get_cache()
        if cache:
            cache_stats = cache.get_stats()
    
    return {
        "status": "healthy", 
        "service": "PPTX Script Slides API",
        "memory_usage_mb": round(memory_mb, 2),
        "active_tasks": len(task_status),
        "cache": cache_stats
    }

@app.post("/convert")
async def convert_pptx(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(..., description="PowerPoint file to convert")
):
    """Upload and convert PowerPoint file to script slides."""
    
    # Check memory usage before accepting new task
    auto_cleanup_if_needed()
    memory_mb = get_memory_usage()
    if memory_mb > MAX_MEMORY_MB:
        raise HTTPException(
            status_code=503, 
            detail=f"Server memory usage too high ({memory_mb:.1f}MB). Please try again later."
        )
    
    # Validate file type
    if not file.filename or not file.filename.lower().endswith('.pptx'):
        raise HTTPException(status_code=400, detail="Only .pptx files are supported")
    
    # Generate task ID
    import uuid
    task_id = str(uuid.uuid4())
    
    # Create temporary directory
    temp_dir = Path(tempfile.mkdtemp())
    input_path = temp_dir / file.filename
    output_path = temp_dir / "スクリプトスライド_自動生成.pptx"
    
    try:
        # Save uploaded file in streaming fashion to avoid large memory spikes
        chunk_size = 4 * 1024 * 1024  # 4MB
        with open(input_path, "wb") as buffer:
            while True:
                chunk = await file.read(chunk_size)
                if not chunk:
                    break
                buffer.write(chunk)
        
        job = ConversionJob(
            task_id=task_id,
            input_path=input_path,
            output_path=output_path,
            temp_dir=temp_dir,
        )

        _ensure_queue_primitives_initialized()
        queue_len = len(conversion_queue)
        if queue_len == 0 and queue_lock is not None:
            # optimistic check if queue is empty; worker will set status to processing
            pass

        task_status[task_id] = {
            "status": "queued",
            "message": "Waiting for available converter...",
            "download_url": None,
            "created_at": time.time(),
            "logs": [
                f"{time.strftime('%Y-%m-%d %H:%M:%S')} Conversion request accepted. Added to queue."
            ],
            "file_path": None,
            "queue_position": queue_len,
        }

        position = await _enqueue_job(job)
        status_entry = task_status.get(task_id)
        if status_entry is not None:
            status_entry["queue_position"] = position
            status_entry.setdefault("logs", []).append(
                f"{time.strftime('%Y-%m-%d %H:%M:%S')} Queue position: {position}"
            )

        return ConversionStatus(
            task_id=task_id,
            status="queued" if position > 0 else "processing",
            message="Queued for processing" if position > 0 else "Started processing",
            logs=task_status[task_id]["logs"],
            queue_position=position if position > 0 else None,
        )
        
    except Exception as e:
        # Cleanup on error
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"Failed to process file: {str(e)}")

@app.get("/status/{task_id}")
async def get_status(task_id: str):
    """Get conversion status."""
    if task_id not in task_status:
        raise HTTPException(status_code=404, detail="Task not found")
    
    status_data = task_status[task_id]
    return ConversionStatus(
        task_id=task_id,
        **status_data
    )

@app.get("/download/{task_id}")
async def download_file(task_id: str, background_tasks: BackgroundTasks):
    """Download converted file."""
    if task_id not in task_status:
        raise HTTPException(status_code=404, detail="Task not found")
    
    status_data = task_status[task_id]
    if status_data["status"] != "completed":
        raise HTTPException(status_code=400, detail="Conversion not completed")
    
    if not status_data.get("file_path"):
        raise HTTPException(status_code=404, detail="File not available")
    
    file_path = Path(status_data["file_path"])
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="File not found on disk")
    
    background_tasks.add_task(_cleanup_task_files, task_id)

    return FileResponse(
        path=file_path,
        filename="スクリプトスライド_自動生成.pptx",
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        background=background_tasks
    )

async def process_conversion(
    task_id: str,
    input_path: Path,
    output_path: Path,
    temp_dir: Path
):
    """Backward compatibility shim for legacy call sites (unused)."""
    job = ConversionJob(task_id=task_id, input_path=input_path, output_path=output_path, temp_dir=temp_dir)
    await _enqueue_job(job)

def _cleanup_task_files(task_id: str) -> None:
    """Background-safe cleanup for task artifacts."""
    status_data = task_status.pop(task_id, None)
    if not status_data:
        return

    file_path_value = status_data.get("file_path")
    if file_path_value:
        file_path = Path(file_path_value)
        try:
            if file_path.exists():
                file_path.unlink()
        except Exception:
            pass

        try:
            shutil.rmtree(file_path.parent, ignore_errors=True)
        except Exception:
            pass

    force_garbage_collection()


@app.delete("/cleanup/{task_id}")
async def cleanup_task(task_id: str):
    """Clean up task files."""
    if task_id not in task_status:
        raise HTTPException(status_code=404, detail="Task not found")

    _cleanup_task_files(task_id)

    return {"message": "Task cleaned up successfully"}

if __name__ == "__main__":
    port = int(os.getenv("PORT", 8000))
    uvicorn.run("web_app:app", host="0.0.0.0", port=port, reload=False)
