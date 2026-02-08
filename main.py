import os
import shutil
import uuid
import json
from typing import List, Dict, Optional
from pathlib import Path
from fastapi import FastAPI, Request, UploadFile, File, Form, Cookie, HTTPException, Response
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.middleware.cors import CORSMiddleware
import uvicorn
import zipfile
from io import BytesIO

from utils.document_processor import DocumentProcessor
from utils.data_handler import DataHandler


from utils.email_handler import EmailHandler

app = FastAPI(title="AutoDispatch")

# Add CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Setup storage
import tempfile
UPLOAD_DIR = Path(tempfile.gettempdir()) / "print_automation_uploads"
UPLOAD_DIR.mkdir(exist_ok=True)
GENERATED_DIR = Path(tempfile.gettempdir()) / "print_automation_generated"
GENERATED_DIR.mkdir(exist_ok=True)

# Mount static files
BASE_DIR = Path(__file__).resolve().parent
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")

# Setup templates
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

# Cleanup helper
def get_session_dir(session_id: str) -> Path:
    session_dir = UPLOAD_DIR / session_id
    session_dir.mkdir(exist_ok=True)
    return session_dir

@app.get("/health")
async def health_check():
    """Health check endpoint for Render."""
    return {"status": "ok"}

@app.head("/")
async def head_home():
    """Explicitly handle HEAD requests for Render health checks."""
    return Response(status_code=200)

@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    """Render the home page (Upload Template)."""
    response = templates.TemplateResponse("index.html", {"request": request})
    # Ensure session cookie exists
    if not request.cookies.get("session_id"):
        session_id = str(uuid.uuid4())
        response.set_cookie(key="session_id", value=session_id)
    return response

@app.post("/upload-template")
async def upload_template(
    request: Request,
    template_file: UploadFile = File(...),
    session_id: Optional[str] = Cookie(None)
):
    """Handle template upload."""
    if not session_id:
        session_id = str(uuid.uuid4())
    
    session_dir = get_session_dir(session_id)
    
    # Save template
    template_path = session_dir / "template.docx"
    with open(template_path, "wb") as buffer:
        shutil.copyfileobj(template_file.file, buffer)
        
    # Validate template
    with open(template_path, "rb") as f:
        content = f.read()
        processor = DocumentProcessor(content)
        placeholders = processor.get_placeholders()
        
    if not placeholders:
        return templates.TemplateResponse(
            "index.html", 
            {"request": request, "error": "No placeholders found in template. Please use {placeholder} format."}
        )
        
    # Save placeholders to session file
    with open(session_dir / "placeholders.json", "w") as f:
        json.dump(placeholders, f)
        
    response = RedirectResponse(url="/upload-data", status_code=303)
    response.set_cookie(key="session_id", value=session_id)
    return response

@app.get("/upload-data", response_class=HTMLResponse)
async def upload_data_page(request: Request, session_id: Optional[str] = Cookie(None)):
    """Render data upload page."""
    if not session_id or not (get_session_dir(session_id) / "template.docx").exists():
        return RedirectResponse(url="/")
        
    return templates.TemplateResponse("upload_data.html", {"request": request})

@app.post("/upload-data")
async def upload_data(
    request: Request,
    data_file: UploadFile = File(...),
    session_id: Optional[str] = Cookie(None)
):
    """Handle data file upload."""
    if not session_id:
        return RedirectResponse(url="/")
        
    session_dir = get_session_dir(session_id)
    
    # Save data file
    filename = data_file.filename
    file_path = session_dir / filename
    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(data_file.file, buffer)
        
    # Validate data file
    with open(file_path, "rb") as f:
        file_bytes = f.read()
        try:
            handler = DataHandler(file_bytes, filename)
            columns = handler.get_columns()
            row_count = handler.get_row_count()
        except Exception as e:
             return templates.TemplateResponse(
                "upload_data.html", 
                {"request": request, "error": f"Error loading data: {str(e)}"}
            )

    # Save metadata
    with open(session_dir / "data_meta.json", "w") as f:
        json.dump({
            "filename": filename,
            "columns": columns,
            "row_count": row_count
        }, f)
        
    return RedirectResponse(url="/map-columns", status_code=303)

@app.get("/map-columns", response_class=HTMLResponse)
async def map_columns_page(request: Request, session_id: Optional[str] = Cookie(None)):
    """Render column mapping page."""
    if not session_id:
        return RedirectResponse(url="/")
        
    session_dir = get_session_dir(session_id)
    if not (session_dir / "template.docx").exists():
        return RedirectResponse(url="/")
    if not (session_dir / "data_meta.json").exists():
        return RedirectResponse(url="/upload-data")
        
    # Load placeholders and columns
    with open(session_dir / "placeholders.json", "r") as f:
        placeholders = json.load(f)
        
    with open(session_dir / "data_meta.json", "r") as f:
        data_meta = json.load(f)
        columns = data_meta["columns"]
        
    # Auto-mapping logic
    auto_mapping = {}
    for p in placeholders:
        for c in columns:
            if p.lower() == c.lower():
                auto_mapping[p] = c
                break
            elif p.lower() in c.lower() or c.lower() in p.lower():
                auto_mapping[p] = c
                
    return templates.TemplateResponse(
        "map_columns.html", 
        {
            "request": request, 
            "placeholders": placeholders, 
            "columns": columns, 
            "auto_mapping": auto_mapping
        }
    )

@app.post("/save-mapping")
async def save_mapping(request: Request, session_id: Optional[str] = Cookie(None)):
    """Save mapping configuration to session."""
    if not session_id:
        return RedirectResponse(url="/")
    
    session_dir = get_session_dir(session_id)
    form_data = await request.form()
    
    mapping = {}
    # Reconstruct mapping from form data (excludes other fields)
    for key, value in form_data.items():
        if key.startswith("mapping_"):
            placeholder = key.replace("mapping_", "")
            if value:
                mapping[placeholder] = value
                
    filename_col = form_data.get("filename_column")
    
    with open(session_dir / "mapping_config.json", "w") as f:
        json.dump({"mapping": mapping, "filename_column": filename_col}, f)
        
    return {"status": "success"}

@app.post("/generate")
async def generate_documents(
    request: Request,
    session_id: Optional[str] = Cookie(None)
):
    """Generate documents and return ZIP."""
    if not session_id:
        return RedirectResponse(url="/")
        
    session_dir = get_session_dir(session_id)
    form_data = await request.form()
    
    # Save mapping first
    mapping = {}
    with open(session_dir / "placeholders.json", "r") as f:
        placeholders = json.load(f)
        
    for p in placeholders:
        val = form_data.get(f"mapping_{p}")
        if val:
            mapping[p] = val
            
    filename_col = form_data.get("filename_column")
    
    # Save mapping to session for email step
    with open(session_dir / "mapping_config.json", "w") as f:
        json.dump({"mapping": mapping, "filename_column": filename_col}, f)
    
    # Load data handler
    with open(session_dir / "data_meta.json", "r") as f:
        data_meta = json.load(f)
        data_filename = data_meta["filename"]
        
    with open(session_dir / data_filename, "rb") as f:
        file_bytes = f.read()
        handler = DataHandler(file_bytes, data_filename)
        
    # Load template processor
    with open(session_dir / "template.docx", "rb") as f:
        template_bytes = f.read()
        processor = DocumentProcessor(template_bytes)
        
    # Generate
    data_rows = handler.get_data_as_dicts(mapping)
    documents = processor.generate_documents(data_rows, filename_column=filename_col)
    
    # Create ZIP
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for filename, doc_bytes in documents:
            zip_file.writestr(filename, doc_bytes)
            
    zip_buffer.seek(0)
    
    headers = {
        'Content-Disposition': 'attachment; filename="generated_documents.zip"'
    }
    return HTMLResponse(
        content=zip_buffer.getvalue(), 
        headers=headers, 
        media_type="application/zip"
    )

# --- Email Endpoints ---

@app.get("/email-config", response_class=HTMLResponse)
async def email_config_page(request: Request, session_id: Optional[str] = Cookie(None)):
    """Render email configuration page."""
    if not session_id:
        return RedirectResponse(url="/")
    return templates.TemplateResponse("email_config.html", {"request": request})

@app.post("/email-config")
async def email_config(
    request: Request,
    smtp_server: str = Form(...),
    smtp_port: int = Form(...),
    sender_email: str = Form(...),
    sender_password: str = Form(...),
    session_id: Optional[str] = Cookie(None)
):
    """Validate SMTP and save config."""
    if not session_id:
        return RedirectResponse(url="/")
    
    session_dir = get_session_dir(session_id)
    
    # Test connection
    handler = EmailHandler(smtp_server, smtp_port, sender_email, sender_password)
    success, msg = handler.test_connection()
    
    if not success:
        return templates.TemplateResponse(
            "email_config.html", 
            {"request": request, "error": msg}
        )
        
    # Save config (WARNING: storing password in plain text in temp file - acceptable for local single-user tool but not prod)
    config = {
        "smtp_server": smtp_server,
        "smtp_port": smtp_port,
        "sender_email": sender_email,
        "sender_password": sender_password
    }
    with open(session_dir / "email_config.json", "w") as f:
        json.dump(config, f)
        
    return RedirectResponse(url="/email-compose", status_code=303)

@app.get("/email-compose", response_class=HTMLResponse)
async def email_compose_page(request: Request, session_id: Optional[str] = Cookie(None)):
    """Render email composition page."""
    if not session_id:
        return RedirectResponse(url="/")
    
    session_dir = get_session_dir(session_id)
    if not (session_dir / "data_meta.json").exists():
        return RedirectResponse(url="/upload-data")
        
    with open(session_dir / "data_meta.json", "r") as f:
        data_meta = json.load(f)
        columns = data_meta["columns"]
        
    return templates.TemplateResponse("email_compose.html", {"request": request, "columns": columns})


@app.post("/email-prepare")
async def email_prepare(
    request: Request,
    email_column: str = Form(...),
    subject: str = Form(...),
    body: str = Form(...),
    cc_emails: str = Form(None),
    bcc_emails: str = Form(None),
    session_id: Optional[str] = Cookie(None)
):
    """Prepare emails and save to session."""
    if not session_id:
        return RedirectResponse(url="/")
        
    session_dir = get_session_dir(session_id)
    
    # Load configurations
    with open(session_dir / "mapping_config.json", "r") as f:
        map_cfg = json.load(f)
        mapping = map_cfg["mapping"]
        filename_col = map_cfg["filename_column"]
        
    # Load data and template
    with open(session_dir / "data_meta.json", "r") as f:
        data_meta = json.load(f)
        data_filename = data_meta["filename"]
        
    with open(session_dir / data_filename, "rb") as f:
        handler = DataHandler(f.read(), data_filename)
        
    with open(session_dir / "template.docx", "rb") as f:
        processor = DocumentProcessor(f.read())
        
    # Prepare data
    data_rows = handler.get_data_as_dicts(mapping)
    documents = processor.generate_documents(data_rows, filename_column=filename_col)
    
    # Render Templates
    email_handler = EmailHandler("dummy", 0, "dummy", "dummy") # Just for rendering
    
    cc_list = [e.strip() for e in cc_emails.split(",")] if cc_emails else []
    bcc_list = [e.strip() for e in bcc_emails.split(",")] if bcc_emails else []
    
    email_queue = []
    
    for idx, (row, (doc_filename, doc_bytes)) in enumerate(zip(data_rows, documents)):
        recipient = row.get(email_column)
        if not recipient:
            continue
            
        rendered_subject = email_handler.render_template(subject, row)
        rendered_body = email_handler.render_template(body, row)
        
        # Save document to temp file for retrieval
        doc_path = session_dir / f"doc_{idx}.docx"
        with open(doc_path, "wb") as f:
            f.write(doc_bytes)
            
        email_queue.append({
            "index": idx,
            "to_email": recipient,
            "subject": rendered_subject,
            "body": rendered_body,
            "attachment_filename": doc_filename,
            "doc_path": f"doc_{idx}.docx",
            "cc_emails": cc_list,
            "bcc_emails": bcc_list,
            "status": "Pending"
        })
        
    # Save queue
    with open(session_dir / "email_queue.json", "w") as f:
        json.dump(email_queue, f)
        
    return RedirectResponse(url="/email-dashboard", status_code=303)

@app.get("/email-dashboard", response_class=HTMLResponse)
async def email_dashboard(request: Request, session_id: Optional[str] = Cookie(None)):
    """Render email dashboard."""
    if not session_id:
        return RedirectResponse(url="/")
    
    session_dir = get_session_dir(session_id)
    if not (session_dir / "email_queue.json").exists():
        return RedirectResponse(url="/email-compose")
        
    with open(session_dir / "email_queue.json", "r") as f:
        email_queue = json.load(f)
        
    return templates.TemplateResponse("email_dashboard.html", {"request": request, "emails": email_queue})

@app.post("/send-single/{index}")
async def send_single_email(index: int, session_id: Optional[str] = Cookie(None)):
    """Send a single email from the queue."""
    if not session_id:
        raise HTTPException(status_code=403, detail="No session")
        
    session_dir = get_session_dir(session_id)
    
    # Load config and queue
    with open(session_dir / "email_config.json", "r") as f:
        email_cfg = json.load(f)
        
    # We load queue each time to be stateless, 
    # but in full app we'd want a DB or valid persistent state.
    # Here we just read the specific item needed.
    with open(session_dir / "email_queue.json", "r") as f:
        email_queue = json.load(f)
        
    # Find item
    # Since queue is list, index should match if not filtered.
    # But let's find by checking idx just in case.
    item = None
    item_idx_in_list = -1
    for i, e in enumerate(email_queue):
        if i == index:
            item = e
            item_idx_in_list = i
            break
            
    if not item:
        return {"status": "error", "message": "Email not found"}
        
    # Load attachment
    with open(session_dir / item["doc_path"], "rb") as f:
        doc_bytes = f.read()

    # Send
    handler = EmailHandler(
        email_cfg["smtp_server"],
        email_cfg["smtp_port"],
        email_cfg["sender_email"],
        email_cfg["sender_password"]
    )
    
    success, msg = handler.send_personalized_email(
        to_email=item["to_email"],
        subject=item["subject"],
        body=item["body"],
        attachment_filename=item["attachment_filename"],
        attachment_data=doc_bytes,
        cc_emails=item["cc_emails"],
        bcc_emails=item["bcc_emails"]
    )
    
    # Update status (optional, for persistent record)
    item["status"] = "Sent" if success else "Failed"
    item["error"] = msg
    
    # Note: Writing back to JSON concurrently might be racey in high load,
    # but fine for single user sequential JS loop.
    email_queue[item_idx_in_list] = item
    with open(session_dir / "email_queue.json", "w") as f:
        json.dump(email_queue, f)
        
    if success:
        return {"status": "success"}
    else:
        return {"status": "error", "message": msg}

@app.post("/skip-single/{index}")
async def skip_single_email(index: int, session_id: Optional[str] = Cookie(None)):
    """Mark email as skipped."""
    if not session_id:
        raise HTTPException(status_code=403, detail="No session")
        
    session_dir = get_session_dir(session_id)
    
    with open(session_dir / "email_queue.json", "r") as f:
        email_queue = json.load(f)
        
    if 0 <= index < len(email_queue):
        email_queue[index]["status"] = "Skipped"
        with open(session_dir / "email_queue.json", "w") as f:
            json.dump(email_queue, f)
            
    return {"status": "success"}

if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)


