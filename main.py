import os
import json
import shutil
import uvicorn
import fitz  # PyMuPDF
from openai import OpenAI
from uuid import uuid4
from datetime import datetime, timedelta

from concurrent.futures import ThreadPoolExecutor, as_completed
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse
from dotenv import load_dotenv
from pydantic import BaseModel, Field
from typing import Dict, List, Optional
from langchain_community.document_loaders import Docx2txtLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

DOCUMENT_SESSIONS = {}


# Load environment variables
load_dotenv()

# Initialize OpenAI client
try:
    client = OpenAI(api_key=os.getenv("llm_open_ai"))
except Exception as e:
    raise RuntimeError(f"Failed to initialize OpenAI client: {str(e)}")

app = FastAPI()

class GuidelinesSchema(BaseModel):
    grammar: Optional[str] = Field(None, description="Check spelling, punctuation, and syntax.")
    sentence_structure: Optional[str] = Field(None, description="Avoid incomplete or run-on sentences.")
    clarity: Optional[str] = Field(None, description="Ensure ideas are well-expressed and clear.")
    adherence_to_writing_rules: Optional[str] = Field(None, description="Follow formal style, tone, and structure.")

class ChunkAssessment(BaseModel):
    chunk_text: str = Field(..., description="The original chunk of text being assessed.")
    corrected_text: str = Field(..., description="AI-generated corrected version") 
    violations: Dict[str, List[str]] = Field(..., description="Detailed violations per category")

class DocumentAssessmentResponse(BaseModel):
    filename: str
    download_url: str
    session_id: str
    expires_at: datetime

def load_document(file_path: str) -> str:
    """Load and extract text from supported document types"""
    try:
        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".pdf":
            with fitz.open(file_path) as doc:
                return "\n".join(page.get_text("text") for page in doc)
        elif ext in [".docx", ".doc"]:
            return Docx2txtLoader(file_path).load()[0].page_content
        raise ValueError(f"Unsupported file format: {ext}")
    except Exception as e:
        raise RuntimeError(f"Document loading failed: {str(e)}")

def split_text(text: str) -> List[str]:
    """Improved text splitting with sentence awareness"""
    splitter = RecursiveCharacterTextSplitter(
        chunk_size=800,  # Reduced for better focus
        chunk_overlap=10,
        separators=["\n\n", ". ", "! ", "? ", "\n", " ", ""],
        keep_separator=True
    )
    return splitter.split_text(text)

def process_chunk(chunk: str) -> ChunkAssessment:
    """Process individual text chunks with GPT-4"""
    guidelines = {
        "grammar": "Check for spelling, punctuation, and grammatical errors",
        "sentence_structure": "Identify run-on sentences, fragments, and awkward constructions",
        "clarity": "Flag ambiguous phrases and unclear references",
        "adherence": "Check against formal writing standards and style guidelines"
    }
    
    prompt = f"""Analyze and correct this text chunk. Return JSON with:
    - "violations": writing issues found
    - "corrected_text": improved version
    
    {{
        "violations": {{
            "grammar": ["..."],
            "sentence_structure": ["..."],
            "clarity": ["..."],
            "adherence": ["..."]
        }},
        "corrected_text": "Improved text version here"
    }}
    
    Text chunk: {chunk}"""  # Truncate to prevent token limit issues
    
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": "You are a writing quality analyst. Return JSON only."},
            {"role": "user", "content": prompt},
        ],
        temperature=0.1,  # More deterministic output
        seed=42,  # Enable deterministic sampling
        max_tokens=1000,  # Allow more space for complete JSON
        top_p=0.95,
    )
    result = json.loads(response.choices[0].message.content)
    return ChunkAssessment(
        chunk_text=chunk,
        corrected_text=result.get("corrected_text", chunk),  # Fallback to original
        violations=result.get("violations", {})
        )

def rebuild_document(assessments: List[ChunkAssessment], original_path: str) -> str:
    """PDF reconstruction with robust text measurement"""
    try:
        ext = os.path.splitext(original_path)[1].lower()
        corrected_content = "".join([a.corrected_text for a in sorted(
            assessments,
            key=lambda x: assessments.index(x))  # Preserve original order
        ])
        output_path = f"corrected_{os.path.basename(original_path)}"

        if ext == ".pdf":

            # Configuration constants
            MARGIN = 72  # 1 inch margins (72 points = 1 inch)
            FONT_SIZE = 11
            LINE_HEIGHT = FONT_SIZE * 1.2
            FONT_NAME = "helv"  # Helvetica

            doc = fitz.open()
            font = fitz.Font(FONT_NAME)
            
            # Create first page
            page = doc.new_page(width=612, height=792)  # Letter size (8.5" x 11")
            page_rect = fitz.Rect(
                MARGIN, MARGIN, 
                page.rect.width - MARGIN, 
                page.rect.height - MARGIN
            )
            
            y_position = page_rect.y0  # Start at top margin

            def get_text_width(text):
                """Get text width using font metrics"""
                return font.text_length(text, fontsize=FONT_SIZE)

            # Process content paragraph by paragraph
            paragraphs = corrected_content.split('\n')
            
            for para in paragraphs:
                if not para.strip():
                    y_position += LINE_HEIGHT * 1.5  # Empty line spacing
                    continue

                # Split paragraph into words
                words = para.split()
                current_line = []
                
                while words:
                    # Test line width
                    test_line = ' '.join(current_line + [words[0]])
                    text_width = get_text_width(test_line)
                    
                    if text_width < page_rect.width:
                        current_line.append(words.pop(0))
                    else:
                        # Add completed line
                        line_text = ' '.join(current_line)
                        page.insert_text(
                            (page_rect.x0, y_position),
                            line_text,
                            fontname=FONT_NAME,
                            fontsize=FONT_SIZE
                        )
                        y_position += LINE_HEIGHT
                        current_line = [words.pop(0)]
                        
                        # Check page overflow
                        if y_position + LINE_HEIGHT > page_rect.y1:
                            # New page
                            page = doc.new_page(width=612, height=792)
                            page_rect = fitz.Rect(
                                MARGIN, MARGIN, 
                                page.rect.width - MARGIN, 
                                page.rect.height - MARGIN
                            )
                            y_position = page_rect.y0

                # Add remaining line
                if current_line:
                    line_text = ' '.join(current_line)
                    page.insert_text(
                        (page_rect.x0, y_position),
                        line_text,
                        fontname=FONT_NAME,
                        fontsize=FONT_SIZE
                    )
                    y_position += LINE_HEIGHT

                # Paragraph spacing
                y_position += LINE_HEIGHT * 0.5

                # Check page overflow after paragraph
                if y_position + LINE_HEIGHT > page_rect.y1:
                    page = doc.new_page(width=612, height=792)
                    page_rect = fitz.Rect(
                        MARGIN, MARGIN, 
                        page.rect.width - MARGIN, 
                        page.rect.height - MARGIN
                    )
                    y_position = page_rect.y0

            doc.save(output_path)

        elif ext in [".docx", ".doc"]:

            doc = Document()
            style = doc.styles['Normal']
            style.font.name = 'Arial'
            style.font.size = Pt(11)
            style.paragraph_format.space_after = Pt(12)
            
            for para in corrected_content.split('\n'):
                if para.strip():
                    p = doc.add_paragraph(para)
                    p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            doc.save(output_path)

        else:  # Plain text
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(corrected_content)

        return output_path
        

    except Exception as e:
        raise RuntimeError(f"Document rebuild failed: {str(e)}")


def assess_document(text_chunks: List[str]) -> List[ChunkAssessment]:
    """Process chunks in parallel but preserve original order"""
    with ThreadPoolExecutor(max_workers=None) as executor:
        # Submit tasks with index tracking
        future_to_index = {
            executor.submit(process_chunk, chunk): idx
            for idx, chunk in enumerate(text_chunks)
        }
        
        # Initialize ordered results list
        results = [None] * len(text_chunks)
        
        # Collect results as they complete
        for future in as_completed(future_to_index):
            idx = future_to_index[future]
            results[idx] = future.result()
            
    return [res for res in results if res is not None]

@app.post("/upload/")
async def auto_correct_document(file: UploadFile = File(...)) -> DocumentAssessmentResponse:
    session_id = str(uuid4())
    file_path = f"temp_{session_id}_{file.filename}"
    
    # Save and process file
    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    
    text = load_document(file_path)
    chunks = split_text(text)
    assessments = assess_document(chunks)
    
    # Generate corrected document
    corrected_path = rebuild_document(assessments, file_path)
    
    # Store session
    DOCUMENT_SESSIONS[session_id] = {
        "original_path": file_path,
        "corrected_path": corrected_path,
        "expires": datetime.now() + timedelta(hours=1)
    }
    
    return DocumentAssessmentResponse(
        filename=file.filename,
        download_url=f"/download/{session_id}",
        session_id=session_id,
        expires_at=DOCUMENT_SESSIONS[session_id]["expires"]
    )

# Simplified download endpoint
@app.get("/download/{session_id}")
async def download_corrected(session_id: str):
    session = DOCUMENT_SESSIONS.get(session_id)
    if not session or datetime.now() > session["expires"]:
        raise HTTPException(404, "Session expired or invalid")
    
    return FileResponse(
        session["corrected_path"],
        filename=f"Corrected_{os.path.basename(session['original_path'])}"
    )

if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8000)