import os
import logging

class Config:
    UPLOAD_DIR = "uploads"
    OUTPUT_DIR = "outputs"
    MODEL_NAME = "all-MiniLM-L6-v2"
    LOG_LEVEL = logging.INFO
    ALLOWED_EXTENSIONS = {".docx"}

    @staticmethod
    def setup():
        os.makedirs(Config.UPLOAD_DIR, exist_ok=True)
        os.makedirs(Config.OUTPUT_DIR, exist_ok=True)
        logging.basicConfig(
            level=Config.LOG_LEVEL,
            format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
        )
