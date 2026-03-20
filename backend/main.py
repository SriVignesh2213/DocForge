from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import os
import uuid
import logging

from config import Config
from utils import validate_file
from template_parser import TemplateParser
from document_parser import DocumentParser
from section_mapper import SectionMapper
from formatter import Formatter
from style_extractor import StyleExtractor

Config.setup()
logger = logging.getLogger(__name__)

app = FastAPI(title="AI Document Template Converter", version="1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

section_mapper = SectionMapper(Config.MODEL_NAME)

@app.post("/convert")
async def convert_document(template_file: UploadFile = File(...), article_file: UploadFile = File(...)):
    validate_file(template_file.filename)
    validate_file(article_file.filename)
    
    unique_id = str(uuid.uuid4())
    temp_template_path = os.path.join(Config.UPLOAD_DIR, f"{unique_id}_template.docx")
    temp_article_path = os.path.join(Config.UPLOAD_DIR, f"{unique_id}_article.docx")
    output_path = os.path.join(Config.OUTPUT_DIR, f"{unique_id}_output.docx")
    
    try:
        with open(temp_template_path, "wb") as f:
            f.write(await template_file.read())
        with open(temp_article_path, "wb") as f:
            f.write(await article_file.read())
            
        logger.info("Parsing template...")
        style_extractor = StyleExtractor(temp_template_path)
        template_parser = TemplateParser(temp_template_path)
        template_sections = template_parser.get_sections()
        
        logger.info("Parsing article...")
        article_parser = DocumentParser(temp_article_path)
        article_sections = article_parser.get_sections()
        
        logger.info("Mapping sections using AI embeddings...")
        mapping = section_mapper.map_sections(article_sections, template_sections)
        
        logger.info("Formatting new document...")
        formatter = Formatter(temp_template_path, temp_article_path, output_path)
        formatter.apply_styles_and_build(article_sections, mapping)
        
        return FileResponse(output_path, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename="formatted_document.docx")
    except Exception as e:
        logger.error(f"Error compiling document: {e}")
        raise HTTPException(status_code=500, detail=str(e))
