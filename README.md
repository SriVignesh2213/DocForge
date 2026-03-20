# DocForge

AI Document Template Converter that intelligently analyzes, extracts, and maps raw manuscripts into complex target architecture templates natively without losing table references or image layouts.

## Architecture Stack
- **Backend:** FastAPI, python-docx, docxcompose, sentence-transformers, scikit-learn
- **Frontend:** React, Tailwind CSS, Vite

## Setup

1. **Backend:**
   ```bash
   cd backend
   python -m venv .venv
   .\.venv\Scripts\activate   # Windows
   pip install -r requirements.txt
   uvicorn main:app --reload
   ```
2. **Frontend:**
   ```bash
   cd frontend
   npm install
   npm run dev
   ```
