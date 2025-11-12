#!/usr/bin/env python3
"""Web API for PowerPoint script slide generation - Render.com compatible."""

from __future__ import annotations

import os
import tempfile
import shutil
from pathlib import Path
from typing import Optional, List
import zipfile

from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import uvicorn
import gc
import psutil
import time

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

# In-memory storage for demo (use Redis/database in production)
task_status = {}

# Memory monitoring and cleanup
MAX_MEMORY_MB = 400  # Maximum memory usage before forced cleanup
CLEANUP_INTERVAL = 300  # Check every 5 minutes
last_cleanup = time.time()

def get_memory_usage() -> float:
    """Get current memory usage in MB."""
    try:
        process = psutil.Process()
        return process.memory_info().rss / 1024 / 1024
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
        
        # Initialize task status with creation time
        task_status[task_id] = {
            "status": "processing",
            "message": "Conversion started...",
            "download_url": None,
            "created_at": time.time()
        }
        
        # Process in background
        background_tasks.add_task(
            process_conversion,
            task_id,
            input_path,
            output_path,
            temp_dir
        )
        
        return ConversionStatus(
            task_id=task_id,
            status="processing",
            message="Conversion started. Check status endpoint."
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
    """Background task to process PowerPoint conversion."""
    try:
        # Update status
        task_status[task_id]["message"] = "Processing slides..."
        
        # Convert using existing app logic
        def log_callback(message: str):
            task_status[task_id]["message"] = message
        
        result_path = generate_script_slides(
            input_path,
            temp_dir,
            log_callback
        )
        
        # Use the result path directly if it's already the expected output path
        final_path = result_path
        if result_path != output_path:
            # Copy result to expected output path only if different
            shutil.copy2(result_path, output_path)
            final_path = output_path
        
        # Update status
        task_status[task_id].update({
            "status": "completed",
            "message": "Conversion completed successfully!",
            "download_url": f"/download/{task_id}",
            "file_path": str(final_path)
        })
        
    except Exception as e:
        # Update status with error
        task_status[task_id].update({
            "status": "failed",
            "message": f"Conversion failed: {str(e)}",
            "download_url": None
        })
    finally:
        # Force garbage collection to free memory
        force_garbage_collection()

        # Cleanup workspace only when conversion failed
        status = task_status.get(task_id, {})
        if status.get("status") != "completed":
            shutil.rmtree(temp_dir, ignore_errors=True)

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
