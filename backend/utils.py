import logging
import os
from fastapi import HTTPException
from config import Config

logger = logging.getLogger(__name__)

def validate_file(filename: str):
    ext = os.path.splitext(filename)[1].lower()
    if ext not in Config.ALLOWED_EXTENSIONS:
        raise HTTPException(status_code=400, detail=f"Invalid file extension: {ext}. Only .docx is allowed.")
