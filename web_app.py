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
from pydantic import BaseModel
import uvicorn

from app import generate_script_slides

app = FastAPI(
    title="PPTX Script Slides API",
    description="Convert PowerPoint notes to large text slides",
    version="1.0.0"
)

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

@app.get("/")
async def root():
    """Health check endpoint."""
    return {"status": "healthy", "service": "PPTX Script Slides API"}

@app.post("/convert")
async def convert_pptx(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(..., description="PowerPoint file to convert")
):
    """Upload and convert PowerPoint file to script slides."""
    
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
        # Save uploaded file
        with open(input_path, "wb") as buffer:
            content = await file.read()
            buffer.write(content)
        
        # Initialize task status
        task_status[task_id] = {
            "status": "processing",
            "message": "Conversion started...",
            "download_url": None
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
async def download_file(task_id: str):
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
    
    return FileResponse(
        path=file_path,
        filename="スクリプトスライド_自動生成.pptx",
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
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
        
        # Copy result to expected output path
        shutil.copy2(result_path, output_path)
        
        # Update status
        task_status[task_id].update({
            "status": "completed",
            "message": "Conversion completed successfully!",
            "file_path": str(output_path),
            "download_url": f"/download/{task_id}"
        })
        
    except Exception as e:
        # Update status with error
        task_status[task_id].update({
            "status": "failed",
            "message": f"Conversion failed: {str(e)}",
            "download_url": None
        })
        
        # Cleanup on error
        shutil.rmtree(temp_dir, ignore_errors=True)

@app.delete("/cleanup/{task_id}")
async def cleanup_task(task_id: str):
    """Clean up task files."""
    if task_id not in task_status:
        raise HTTPException(status_code=404, detail="Task not found")
    
    status_data = task_status[task_id]
    if status_data.get("file_path"):
        file_path = Path(status_data["file_path"])
        if file_path.exists():
            # Try to cleanup parent temp directory
            try:
                shutil.rmtree(file_path.parent, ignore_errors=True)
            except:
                pass
    
    # Remove from memory
    del task_status[task_id]
    
    return {"message": "Task cleaned up successfully"}

if __name__ == "__main__":
    port = int(os.getenv("PORT", 8000))
    uvicorn.run("web_app:app", host="0.0.0.0", port=port, reload=False)
