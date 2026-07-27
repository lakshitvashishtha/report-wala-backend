import json
import mimetypes
import os
import threading
import traceback
import uuid
import hashlib
import secrets
import time
import builtins
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse, parse_qs
import hmac
import re

import google.genai as genai
from google.genai import types as genai_types
from google.oauth2 import id_token
from google.auth.transport import requests

from engines import get_engine_for_report
from storage import ReportStorage

URL_SECRET_KEY = os.environ.get("URL_SECRET_KEY", os.urandom(32).hex())

def generate_signed_url(job_id, kind, expiration_seconds=900):
    expires = int(time.time()) + expiration_seconds
    msg = f"{job_id}:{kind}:{expires}"
    sig = hmac.new(URL_SECRET_KEY.encode(), msg.encode(), hashlib.sha256).hexdigest()
    return f"/api/download/{job_id}/{kind}?expires={expires}&sig={sig}"

_ORIGINAL_PRINT = builtins.print


def print(*args, **kwargs):
    try:
        _ORIGINAL_PRINT(*args, **kwargs)
    except OSError:
        pass


ROOT = Path(__file__).resolve().parent
UPLOAD_ROOT = ROOT / "uploaded_references"
HOST = os.getenv("HOST", "0.0.0.0")
PORT = int(os.getenv("PORT", "8000"))
ENGINE_VERSION = "universal-report-workspace-v9"
MAX_UPLOAD_BYTES = 50 * 1024 * 1024
MAX_FILE_BYTES = 20 * 1024 * 1024
MAX_REFERENCE_FILES = 10
ALLOWED_REFERENCE_SUFFIXES = {".pdf", ".docx", ".doc", ".txt", ".md", ".csv", ".xlsx", ".pptx"}
ALLOWED_PROJECT_SUFFIXES = {".pdf", ".docx", ".doc", ".txt", ".md", ".csv", ".xlsx", ".pptx", ".jpg", ".jpeg", ".png"}
ALLOWED_CUSTOM_PAGE_SUFFIXES = {".pdf", ".jpg", ".jpeg", ".png"}

JOBS = {}
JOBS_LOCK = threading.Lock()
STORAGE = ReportStorage(UPLOAD_ROOT)
FRONTEND_ORIGINS = {
    origin.strip().rstrip("/")
    for origin in os.getenv(
        "FRONTEND_ORIGINS",
        "*,https://report-ai-beryl.vercel.app,http://localhost:3000,http://localhost:8000,http://127.0.0.1:3000,http://127.0.0.1:8000,capacitor://localhost,http://localhost",
    ).split(",")
    if origin.strip()
}


def json_bytes(payload):
    return json.dumps(payload, indent=2).encode("utf-8")


def safe_filename(name):
    # Sanitize filename for Windows - remove invalid characters
    if not name:
        return "reference_file"
    
    # Remove invalid Windows characters: < > : " / \ | ? *
    invalid_chars = '<>:"/\\|?*'
    cleaned = name
    for char in invalid_chars:
        cleaned = cleaned.replace(char, '_')
    
    # Remove null bytes and all control characters. Windows rejects these in paths.
    cleaned = ''.join(c for c in cleaned if ord(c) >= 32)
    
    # Get just the filename (no path)
    cleaned = Path(cleaned).name
    
    # Limit length to avoid path issues
    if len(cleaned) > 100:
        name_part, ext = os.path.splitext(cleaned)
        cleaned = name_part[:90] + ext
    
    cleaned = cleaned.strip().strip(".")
    reserved_names = {
        "CON", "PRN", "AUX", "NUL",
        "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
    }
    stem, ext = os.path.splitext(cleaned)
    if stem.upper() in reserved_names:
        cleaned = f"{stem}_file{ext}"

    return cleaned or "reference_file"


def parse_multipart(headers, body):
    content_type = headers.get("Content-Type", "")
    message = BytesParser(policy=default).parsebytes(
        f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8") + body
    )

    fields = {}
    files = []

    if not message.is_multipart():
        return fields, files

    for part in message.iter_parts():
        field_name = part.get_param("name", header="content-disposition")
        if not field_name:
            continue

        filename = part.get_filename()
        content = part.get_payload(decode=True) or b""

        if filename:
            files.append({
                "field": field_name,
                "filename": safe_filename(filename),
                "content": content,
            })
        else:
            charset = part.get_content_charset() or "utf-8"
            fields[field_name] = content.decode(charset, errors="ignore").strip()

    return fields, files


def update_job(job_id, **values):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if job:
            job.update(values)
            snapshot = dict(job)
        else:
            snapshot = None
    if snapshot:
        STORAGE.save_job(job_id, snapshot)


def append_log(job_id, message):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if job:
            job.setdefault("logs", []).append(message)
            snapshot = dict(job)
        else:
            snapshot = None
    if snapshot:
        STORAGE.save_job(job_id, snapshot)


def run_report_job(job_id, metadata):
    try:
        reference_dir = STORAGE.materialize_references(job_id)
        project_detail_dir = STORAGE.materialize_project_detail_files(job_id)
        custom_page_files = STORAGE.materialize_custom_page_files(job_id)
        metadata["refFolder"] = str(reference_dir) if reference_dir else None
        metadata["projectDetailsFolder"] = str(project_detail_dir) if project_detail_dir else None
        if custom_page_files.get("title"):
            metadata["customTitleFile"] = str(custom_page_files["title"])
        if custom_page_files.get("certificate"):
            metadata["customCertificateFile"] = str(custom_page_files["certificate"])
        update_job(job_id, status="running", message="Designing your report structure.", progress=2)

        def report_ai_progress(call_count, message):
            calculated_progress = min(94, 5 + int(call_count * 9.8))
            message_lower = message.lower()
            if "table of contents" in message_lower:
                friendly_message = "Report structure and chapter plan ready"
            elif "abstract" in message_lower:
                friendly_message = "Concise abstract ready"
            elif "chapter" in message_lower:
                chapter_number = next(
                    (number for number in range(1, 21) if f"chapter {number}" in message_lower),
                    None,
                )
                sec_match = re.search(r"section\s+([\d\.]+)", message_lower)
                sec_id = sec_match.group(1) if sec_match else None
                if "intro" in message_lower:
                    friendly_message = f"Chapter {chapter_number} introduction written"
                elif sec_id:
                    friendly_message = f"Chapter {chapter_number} section {sec_id} written"
                elif "repair" in message_lower:
                    friendly_message = f"Chapter {chapter_number} refined" if chapter_number else "Refining chapter"
                else:
                    friendly_message = f"Chapter {chapter_number} written" if chapter_number else "Chapter content ready"
            elif "rate limit" in message_lower:
                friendly_message = message
            else:
                friendly_message = "Report content updated"
            with JOBS_LOCK:
                current_progress = (JOBS.get(job_id) or {}).get("progress", 0)
            append_log(job_id, friendly_message)
            update_job(
                job_id,
                progress=max(current_progress, calculated_progress),
                message=friendly_message,
            )

        metadata["progress_callback"] = report_ai_progress
        job_dir = STORAGE.local_root / job_id
        job_dir.mkdir(parents=True, exist_ok=True)
        engine = get_engine_for_report(metadata.get("reportType", ""), metadata, job_dir=job_dir)
        msg = "Writing the abstract and chapters." if metadata.get("includeFrontMatter", True) else "Writing the chapters."
        update_job(job_id, plan=engine.plan, message=msg, progress=16)

        build_result = engine.execute(15)
        if build_result.get("status") != "success":
            raise RuntimeError(build_result.get("message", "Document assembly failed."))
            
        if (job_dir / "cancel.flag").exists():
            raise InterruptedError("Generation cancelled by user")
            
        update_job(job_id, message="Completing final quality checks.", progress=98)
        output = engine.execute(16)
        if output.get("status") != "complete":
            raise RuntimeError("Generation ended before a final report was returned.")
            
        # Automatic formatting validation before download
        try:
            import docx
            docx_path = output.get("finalReport")
            if not docx_path or not Path(docx_path).exists():
                raise RuntimeError("Generated document file is missing.")
            doc_test = docx.Document(docx_path)
            if len(doc_test.paragraphs) < 10:
                raise RuntimeError("The generated document is too short or malformed.")
        except Exception as e:
            raise RuntimeError(f"Formatting validation failed: {str(e)}")

        if output.get("pdfFile"):
            append_log(job_id, "Word and PDF files created")
        else:
            append_log(job_id, "Word report created")
            if output.get("pdfWarning"):
                append_log(job_id, "PDF export was skipped or failed; Word download is still available")

        update_job(
            job_id,
            status="complete",
            message="Report is ready." if output.get("pdfFile") else "Word report is ready. PDF export was unavailable.",
            progress=100,
            result=output,
        )
        append_log(job_id, "Report ready to download")

        # Persist output files to MySQL so they survive server restarts (24hr recovery)
        try:
            docx_path = output.get("finalReport")
            if docx_path and Path(docx_path).exists():
                STORAGE.save_output_file(job_id, "docx", Path(docx_path).name, Path(docx_path).read_bytes())
            pdf_path = output.get("pdfFile")
            if pdf_path and Path(pdf_path).exists():
                STORAGE.save_output_file(job_id, "pdf", Path(pdf_path).name, Path(pdf_path).read_bytes())
            md_path = output.get("metadata", {}).get("markdownPath")
            if md_path and Path(md_path).exists():
                STORAGE.save_output_file(job_id, "md", Path(md_path).name, Path(md_path).read_bytes())
        except Exception as _save_exc:
            print(f"Warning: could not persist output files to DB: {_save_exc}")

    except Exception as exc:
        import traceback

        traceback.print_exc()
        append_log(job_id, "The report needs attention before it can be completed")
        update_job(job_id, status="error", error=str(exc), message="The report could not be completed.")


class ReportRequestHandler(BaseHTTPRequestHandler):
    server_version = "ReportWalaAI/1.0"

    def do_GET(self):
        parsed = urlparse(self.path)
        path = unquote(parsed.path)

        if path == "/api/user/reports":
            self.handle_user_reports(parsed.query)
            return
        if path == "/api/samples":
            self.handle_samples()
            return
        if path == "/api/health":
            self.send_json({"status": "ok", "engine": ENGINE_VERSION, "storage": STORAGE.backend_name})
            return

        if path.startswith("/api/status/"):
            self.handle_status(path)
            return

        if path.startswith("/api/download/"):
            self.handle_download(self.path)
            return

        if path.startswith("/api/samples/download/"):
            filename = path.replace("/api/samples/download/", "").split("?")[0]
            self.handle_sample_download(filename)
            return

        if path.startswith("/api/samples/thumbnail/"):
            filename = path.replace("/api/samples/thumbnail/", "").split("?")[0]
            self.handle_sample_thumbnail(filename)
            return

        if path == "/":
            self.send_json({"status": "ok", "service": "ReportForge API", "engine": ENGINE_VERSION})
            return

        self.send_json({"error": "Not found"}, status=404)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization, Bypass-Tunnel-Reminder, bypass-tunnel-reminder")
        self.send_header("Access-Control-Max-Age", "86400")
        self.end_headers()

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/login":
            self.handle_login()
            return
            
        if parsed.path == "/api/logout":
            self.handle_logout()
            return
        if parsed.path == "/api/generate":
            self.handle_generate()
            return

        if parsed.path == "/api/suggest-description":
            self.handle_suggest_description()
            return
            
        if parsed.path == "/api/regenerate-chapter":
            self.handle_regenerate_chapter()
            return

        if parsed.path == "/api/suggest-outline":
            self.handle_suggest_outline()
            return

        if parsed.path.startswith("/api/cancel/"):
            job_id = parsed.path.split("/")[-1]
            self.handle_cancel(job_id)
            return

        self.send_json({"error": "Not found"}, status=404)

    def handle_cancel(self, job_id):
        if not re.match(r"^[a-fA-F0-9\-]+$", job_id):
            self.send_json({"error": "Invalid job ID format."}, status=400)
            return
            
        job_dir = STORAGE.local_root / job_id
        if not job_dir.exists():
            self.send_json({"error": "Job not found."}, status=404)
            return
            
        cancel_flag = job_dir / "cancel.flag"
        cancel_flag.touch()
        self.send_json({"success": True})

    def log_message(self, format, *args):
        super().log_message(format, *args)

    def end_headers(self):
        origin = (self.headers.get("Origin") or "").rstrip("/")
        if origin and ("*" in FRONTEND_ORIGINS or origin in FRONTEND_ORIGINS):
            self.send_header("Access-Control-Allow-Origin", origin)
            self.send_header("Access-Control-Allow-Credentials", "true")
            self.send_header("Vary", "Origin")
        super().end_headers()

    def send_json(self, payload, status=200):
        data = json_bytes(payload)
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def handle_generate(self):
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            self.send_json({"error": "Invalid request size."}, status=400)
            return
        if length <= 0 or length > MAX_UPLOAD_BYTES:
            self.send_json({"error": "The request is empty or larger than the 50 MB upload limit."}, status=413)
            return
        body = self.rfile.read(length)
        fields, files = parse_multipart(self.headers, body)

        required = [
            "api_key",
            "topic",
            "studentName",
            "collegeName",
        ]
        missing = [field for field in required if not fields.get(field)]
        if missing:
            self.send_json({"error": "Missing required fields: " + ", ".join(missing)}, status=400)
            return

        token = self.headers.get("Authorization", "").replace("Bearer ", "").strip()
        user_id = None
        if token:
            session = STORAGE.get_session(token)
            if session:
                user_id = session["user_id"]
        
        job_id = uuid.uuid4().hex

        references = [item for item in files if item["field"] == "references" and item["content"]]
        project_files = [item for item in files if item["field"] == "projectFiles" and item["content"]]
        custom_title_files = [item for item in files if item["field"] == "customTitleFile" and item["content"]]
        custom_certificate_files = [item for item in files if item["field"] == "customCertificateFile" and item["content"]]
        if len(references) > MAX_REFERENCE_FILES:
            self.send_json({"error": f"Add no more than {MAX_REFERENCE_FILES} reference files."}, status=400)
            return
        if len(custom_title_files) > 1 or len(custom_certificate_files) > 1:
            self.send_json({"error": "Upload only one title page file and one certificate page file."}, status=400)
            return
        if len(project_files) > MAX_REFERENCE_FILES:
            self.send_json({"error": f"Add no more than {MAX_REFERENCE_FILES} project detail files."}, status=400)
            return
        custom_chapters_json = fields.get("customChaptersJson") or ""
        custom_chapters_data = []
        if custom_chapters_json:
            try:
                custom_chapters_data = json.loads(custom_chapters_json)
            except Exception:
                pass
        
        if not custom_chapters_data:
            custom_chapter_text = (fields.get("customChapterTitles") or fields.get("customChapters") or "").strip()
            custom_chapter_lines = [line for line in custom_chapter_text.splitlines() if line.strip()]
            for line in custom_chapter_lines:
                custom_chapters_data.append({"title": line, "subchapters": []})
        
        if len(custom_chapters_data) > 20:
            self.send_json({"error": "You can provide a maximum of 20 custom chapters."}, status=400)
            return
        def validate_file_signature(content, suffix):
            if not content:
                return False
            if suffix == ".pdf":
                return content.startswith(b"%PDF-")
            if suffix in [".docx", ".xlsx", ".pptx"]:
                return content.startswith(b"PK\x03\x04")
            if suffix == ".png":
                return content.startswith(b"\x89PNG\r\n\x1a\n")
            if suffix in [".jpg", ".jpeg"]:
                return content.startswith(b"\xff\xd8\xff")
            if suffix in [".txt", ".md", ".csv"]:
                # Text files don't have a reliable magic number, but shouldn't be binary executables
                return not content.startswith(b"MZ") and not content.startswith(b"\x7fELF")
            return True

        for uploaded in references:
            suffix = Path(uploaded["filename"]).suffix.lower()
            if suffix not in ALLOWED_REFERENCE_SUFFIXES:
                self.send_json({"error": f"{uploaded['filename']} is not a supported reference file."}, status=400)
                return
            if len(uploaded["content"]) > MAX_FILE_BYTES:
                self.send_json({"error": f"{uploaded['filename']} is larger than the 20 MB per-file limit."}, status=413)
                return
            if not validate_file_signature(uploaded["content"], suffix):
                self.send_json({"error": f"{uploaded['filename']} failed MIME/signature validation."}, status=400)
                return
                
        for uploaded in project_files:
            suffix = Path(uploaded["filename"]).suffix.lower()
            if suffix not in ALLOWED_PROJECT_SUFFIXES:
                self.send_json({"error": f"{uploaded['filename']} is not a supported project detail file."}, status=400)
                return
            if len(uploaded["content"]) > MAX_FILE_BYTES:
                self.send_json({"error": f"{uploaded['filename']} is larger than the 20 MB per-file limit."}, status=413)
                return
            if not validate_file_signature(uploaded["content"], suffix):
                self.send_json({"error": f"{uploaded['filename']} failed MIME/signature validation."}, status=400)
                return
                
        for uploaded in custom_title_files + custom_certificate_files:
            suffix = Path(uploaded["filename"]).suffix.lower()
            if suffix not in ALLOWED_CUSTOM_PAGE_SUFFIXES:
                self.send_json(
                    {
                        "error": (
                            f"{uploaded['filename']} cannot be used as a custom title/certificate page here. "
                            "Upload PDF, JPG, or PNG so the exact page is preserved."
                        )
                    },
                    status=400,
                )
                return
            if len(uploaded["content"]) > MAX_FILE_BYTES:
                self.send_json({"error": f"{uploaded['filename']} is larger than the 20 MB per-file limit."}, status=413)
                return
            if not validate_file_signature(uploaded["content"], suffix):
                self.send_json({"error": f"{uploaded['filename']} failed MIME/signature validation."}, status=400)
                return

        include_fm = str(fields.get("includeFrontMatter", "false")).lower() == "true" or str(fields.get("includeFrontMatter", "")) == "on"
        
        logs = ["Details received"]
        if references:
            logs.append(f"{len(references)} reference file(s) added")
        else:
            logs.append("No reference files added")
            
        if project_files:
            logs.append(f"{len(project_files)} project detail file(s) added")
        else:
            logs.append("No project detail files added")
            
        if include_fm:
            if custom_title_files or custom_certificate_files:
                logs.append("Custom title/certificate file added")
            else:
                logs.append("Automatic title/certificate pages selected")
        else:
            logs.append("Front matter (TOC, Cover, etc.) disabled")
            
        if custom_chapters_data:
            logs.append(f"Using {len(custom_chapters_data)} custom chapter title(s)")
        else:
            logs.append("AI will plan chapter titles")
            
        logs.append("Designing the report structure")

        initial_job = {
            "status": "queued",
            "message": "Report job queued.",
            "progress": 0,
            "topic": fields.get("topic", ""),
            "reportType": fields.get("reportType", ""),
            "logs": logs,
            "result": None,
            "error": None,
        }
        with JOBS_LOCK:
            JOBS[job_id] = initial_job
        STORAGE.save_job(job_id, initial_job, user_id=user_id)

        saved_files = []
        for uploaded in references:
            if uploaded["field"] != "references" or not uploaded["content"]:
                continue
            try:
                content_type = mimetypes.guess_type(uploaded["filename"])[0]
                STORAGE.save_reference(
                    job_id,
                    uploaded["filename"],
                    content_type,
                    uploaded["content"],
                )
                saved_files.append(uploaded["filename"])
            except Exception as exc:
                print(f"Failed to store reference file {uploaded['filename']}: {exc}")
                traceback.print_exc()
                update_job(job_id, status="error", error=str(exc), message="A reference file could not be stored.")
                self.send_json({"error": f"Could not store {uploaded['filename']}."}, status=500)
                return
        for page_role, upload_group in (("title", custom_title_files), ("certificate", custom_certificate_files)):
            for uploaded in upload_group:
                try:
                    content_type = mimetypes.guess_type(uploaded["filename"])[0]
                    STORAGE.save_custom_page_file(
                        job_id,
                        page_role,
                        safe_filename(uploaded["filename"]),
                        content_type,
                        uploaded["content"],
                    )
                except Exception as exc:
                    print(f"Failed to store custom page file {uploaded['filename']}: {exc}")
                    traceback.print_exc()
                    update_job(job_id, status="error", error=str(exc), message="A title or certificate file could not be stored.")
                    self.send_json({"error": f"Could not store {uploaded['filename']}."}, status=500)
                    return
        for uploaded in project_files:
            try:
                content_type = mimetypes.guess_type(uploaded["filename"])[0]
                STORAGE.save_project_detail_file(
                    job_id,
                    safe_filename(uploaded["filename"]),
                    content_type,
                    uploaded["content"],
                )
            except Exception as exc:
                print(f"Failed to store project detail file {uploaded['filename']}: {exc}")
                traceback.print_exc()
                update_job(job_id, status="error", error=str(exc), message="A project detail file could not be stored.")
                self.send_json({"error": f"Could not store {uploaded['filename']}."}, status=500)
                return

        try:
            target_pages = int(fields.get("targetPages") or 70)
        except ValueError:
            target_pages = 70
        
        if target_pages < 5:
            target_pages = 5

        metadata = {
            "api_key": fields.get("api_key"),
            "topic": fields.get("topic"),
            "studentName": fields.get("studentName"),
            "enrollmentNumber": fields.get("enrollmentNumber"),
            "collegeName": fields.get("collegeName"),
            "department": fields.get("department"),
            "guideName": fields.get("guideName"),
            "session": fields.get("session"),
            "university": fields.get("university"),
            "companyName": fields.get("companyName"),
            "mentorName": fields.get("mentorName"),
            "trainingDuration": fields.get("trainingDuration"),
            "companyDepartment": fields.get("companyDepartment"),
            "companyLocation": fields.get("companyLocation"),
            "degreeName": fields.get("degreeName"),
            "researchArea": fields.get("researchArea"),
            "submissionPurpose": fields.get("submissionPurpose"),
            "jobTitle": fields.get("jobTitle"),
            "projectDomain": fields.get("projectDomain"),
            "clientOrUnit": fields.get("clientOrUnit"),
            "customTitlePage": fields.get("customTitlePage"),
            "customCertificateText": fields.get("customCertificateText"),
            "targetPages": target_pages,
            "includeFrontMatter": str(fields.get("includeFrontMatter", "false")).lower() == "true" or str(fields.get("includeFrontMatter", "")) == "on",
            "reportType": fields.get("reportType") or "Academic project report",
            "tone": fields.get("tone") or "Formal academic",
            "specialInstructions": fields.get("specialInstructions") or "",
            "projectDetails": fields.get("projectDetails") or "",
            "evidenceDetails": fields.get("evidenceDetails") or "",
            "authorRole": fields.get("authorRole") or "Student",
            "customChapters": custom_chapters_data,
            "refFolder": None,
            # PSD Station Fields
            "psdStation": fields.get("psdStation") or "",
            "psdProjectDomain": fields.get("psdProjectDomain") or "",
            "psdIndustryMentor": fields.get("psdIndustryMentor") or "",
            "psdFacultyMentor": fields.get("psdFacultyMentor") or "",
            "psdDuration": fields.get("psdDuration") or "",
            "psdDeliverables": fields.get("psdDeliverables") or "",
            # VSD VLSI Fields
            "vsdPdkNode": fields.get("vsdPdkNode") or "",
            "vsdToolchain": fields.get("vsdToolchain") or "",
            "vsdCoreIp": fields.get("vsdCoreIp") or "",
            "vsdSpecs": fields.get("vsdSpecs") or "",
            "vsdWaveforms": fields.get("vsdWaveforms") or "",
            # Literature Review Fields
            "litDatabases": fields.get("litDatabases") or "",
            "litKeywords": fields.get("litKeywords") or "",
            "litCriteria": fields.get("litCriteria") or "",
            "litThemes": fields.get("litThemes") or "",
            # Academic/BTech fields
            "academicTechStack": fields.get("academicTechStack") or "",
            "academicSysDesign": fields.get("academicSysDesign") or "",
            "academicModules": fields.get("academicModules") or "",
            "academicTesting": fields.get("academicTesting") or "",
            # Thesis fields
            "thesisMethodology": fields.get("thesisMethodology") or "",
            "thesisDataCollection": fields.get("thesisDataCollection") or "",
            "thesisStats": fields.get("thesisStats") or "",
            # Research paper fields
            "rpJournal": fields.get("rpJournal") or "",
            "rpDatasets": fields.get("rpDatasets") or "",
            "rpSetup": fields.get("rpSetup") or "",
            # Internship work details
            "internshipTasks": fields.get("internshipTasks") or "",
            # Technical report fields
            "techReqs": fields.get("techReqs") or "",
            "techInstall": fields.get("techInstall") or "",
            "techParams": fields.get("techParams") or "",
            "techTesting": fields.get("techTesting") or "",
            # Business report fields
            "bizContext": fields.get("bizContext") or "",
            "bizStakeholders": fields.get("bizStakeholders") or "",
            "bizCosts": fields.get("bizCosts") or "",
            "bizOptions": fields.get("bizOptions") or "",
            # Medical case fields
            "medPatient": fields.get("medPatient") or "",
            "medTests": fields.get("medTests") or "",
            "medTreatment": fields.get("medTreatment") or "",
            # Professional outcomes
            "profOutcome": fields.get("profOutcome") or "",
            # Policy report fields
            "policyIssue": fields.get("policyIssue") or "",
            "policyContext": fields.get("policyContext") or "",
            "policyOptions": fields.get("policyOptions") or "",
            # Seminar report fields
            "semTheme": fields.get("semTheme") or "",
            "semConcepts": fields.get("semConcepts") or "",
            "semTrends": fields.get("semTrends") or "",
            "semTakeaways": fields.get("semTakeaways") or "",
        }

        thread = threading.Thread(target=run_report_job, args=(job_id, metadata), daemon=True)
        thread.start()

        self.send_json({"jobId": job_id, "status": "queued"}, status=202)

    def handle_regenerate_chapter(self):
        length = int(self.headers.get("Content-Length", "0"))
        if length <= 0 or length > 100_000:
            self.send_json({"error": "Invalid request."}, status=400)
            return

        try:
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
        except (json.JSONDecodeError, UnicodeDecodeError):
            self.send_json({"error": "The request could not be read."}, status=400)
            return

        job_id = payload.get("jobId")
        chapter_id = payload.get("chapterId")
        
        if not job_id or not chapter_id:
            self.send_json({"error": "Missing jobId or chapterId"}, status=400)
            return
            
        with JOBS_LOCK:
            job = JOBS.get(job_id)
            if job:
                job.setdefault("logs", []).append(f"Regenerating chapter {chapter_id}...")
        
        # Here we would interface with ReportEngine to regenerate a specific chapter.
        # Since this is a placeholder implementation, we just return success.
        # A full implementation would require persisting the ReportEngine state.
        self.send_json({"status": "queued", "message": f"Chapter {chapter_id} regeneration queued"})

    def handle_suggest_description(self):
        length = int(self.headers.get("Content-Length", "0"))
        if length <= 0 or length > 100_000:
            self.send_json({"error": "Invalid request."}, status=400)
            return

        try:
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
        except (json.JSONDecodeError, UnicodeDecodeError):
            self.send_json({"error": "The request could not be read."}, status=400)
            return

        api_key = str(payload.get("api_key") or "").strip()
        topic = str(payload.get("topic") or "").strip()
        if not api_key or not topic:
            self.send_json({"error": "Add a topic and Gemini key first."}, status=400)
            return

        report_type = str(payload.get("reportType") or "Professional report")

        prompt = f"""You are a report assistant. Do two things:

1. Check if the topic "{topic}" has spelling errors. If yes, write "CORRECTION: <corrected topic>" on the first line. If no errors, skip this line.
2. Write a 2-sentence report brief for this topic (report type: {report_type}). Be concise and professional.

Output format:
[CORRECTION: corrected topic if needed]
BRIEF: <your 2-sentence brief here>"""

        try:
            from google.genai import types as genai_types
            client = genai.Client(
                api_key=api_key,
                http_options=genai_types.HttpOptions(
                    retry_options=genai_types.HttpRetryOptions(attempts=1),
                    timeout=15000,
                )
            )

            import re
            response = None
            last_call_error = None
            for model_name in ["gemini-1.5-flash", "gemini-2.0-flash", "gemini-1.5-pro"]:
                try:
                    response = client.models.generate_content(
                        model=model_name,
                        contents=prompt,
                        config=genai_types.GenerateContentConfig(
                            temperature=0.3,
                            max_output_tokens=200,
                        ),
                    )
                    break
                except Exception as e:
                    last_call_error = e
                    print(f"Model {model_name} failed: {e}")
                    continue

            if not response:
                if last_call_error is not None:
                    raise last_call_error
                raise RuntimeError("No models available.")

            raw = (response.text or "").strip()

            # Parse correction
            correction = ""
            corr_match = re.search(r"CORRECTION:\s*(.+)", raw, re.IGNORECASE)
            if corr_match:
                candidate = corr_match.group(1).strip().strip('"\'')
                if candidate.lower() != topic.lower():
                    correction = candidate

            # Parse brief
            brief_match = re.search(r"BRIEF:\s*(.+)", raw, re.IGNORECASE | re.DOTALL)
            if brief_match:
                description = brief_match.group(1).strip()
            else:
                # Fallback: remove the correction line and use remainder
                description = re.sub(r"CORRECTION:.*\n?", "", raw, flags=re.IGNORECASE).strip()

            if len(description) < 20:
                raise RuntimeError("The suggested brief was incomplete.")

            self.send_json({"description": description, "correction": correction})
        except Exception as exc:
            import traceback
            traceback.print_exc()
            error_text = str(exc).lower()
            if "reported as leaked" in error_text:
                message = "This key has been disabled by Google. Create a new Gemini key and try again."
            elif "429" in error_text or "resource_exhausted" in error_text:
                message = "The key has reached its current usage limit. Wait briefly and try again."
            elif "403" in error_text or "permission_denied" in error_text:
                message = "Google rejected this key. Check it or create a new Gemini key."
            else:
                message = f"Failed to generate brief: {str(exc)}"
            self.send_json({"error": message}, status=400)

    def handle_status(self, path):
        job_id = path.rsplit("/", 1)[-1]
        with JOBS_LOCK:
            job = JOBS.get(job_id)
            payload = dict(job) if job else None

        if not payload:
            payload = STORAGE.load_job(job_id)

        if not payload:
            self.send_json({"error": "Unknown job id"}, status=404)
            return
            
        if payload.get("status") == "complete":
            payload["docx_url"] = generate_signed_url(job_id, "docx")
            payload["pdf_url"] = generate_signed_url(job_id, "pdf")
            payload["md_url"] = generate_signed_url(job_id, "md")

        self.send_json(payload)

    def handle_download(self, path):
        parsed = urlparse(path)
        parts = parsed.path.strip("/").split("/")
        if len(parts) != 4:
            self.send_json({"error": "Invalid download URL"}, status=400)
            return

        _, _, job_id, kind = parts
        
        query_params = parse_qs(parsed.query)
        expires = query_params.get("expires", [""])[0]
        sig = query_params.get("sig", [""])[0]
        
        if not expires or not sig:
            self.send_json({"error": "Missing signature or expiration"}, status=403)
            return
            
        try:
            if int(expires) < time.time():
                self.send_json({"error": "Download link has expired. Please refresh the page."}, status=403)
                return
        except ValueError:
            self.send_json({"error": "Invalid expiration timestamp"}, status=400)
            return
            
        msg = f"{job_id}:{kind}:{expires}"
        expected_sig = hmac.new(URL_SECRET_KEY.encode(), msg.encode(), hashlib.sha256).hexdigest()
        
        if not hmac.compare_digest(expected_sig, sig):
            self.send_json({"error": "Invalid signature"}, status=403)
            return

        with JOBS_LOCK:
            job = JOBS.get(job_id)

        if not job:
            job = STORAGE.load_job(job_id)

        if not job or job.get("status") != "complete":
            self.send_json({"error": "Report is not ready"}, status=404)
            return

        result = job.get("result") or {}
        if kind == "docx":
            file_path = result.get("finalReport")
        elif kind == "pdf":
            file_path = result.get("pdfFile")
        elif kind == "md":
            file_path = result.get("metadata", {}).get("markdownPath")
        else:
            self.send_json({"error": "Unsupported download type"}, status=400)
            return

        # Try local file first; fall back to MySQL-stored content
        if file_path and Path(file_path).exists():
            target = Path(file_path)
            data = target.read_bytes()
            filename = target.name
        else:
            stored = STORAGE.get_output_file(job_id, kind)
            if not stored:
                self.send_json({"error": "Report file not found. It may have expired or the server was restarted."}, status=404)
                return
            filename, data = stored

        content_type = (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            if kind == "docx"
            else "application/pdf" if kind == "pdf" else "text/markdown"
        )

        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def handle_sample_download(self, filename):
        file_path = ROOT / "samples" / filename
        if file_path.exists() and file_path.is_file():
            content_type = "application/pdf" if filename.endswith(".pdf") else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            with open(file_path, "rb") as f:
                self.wfile.write(f.read())
        else:
            self.send_json({"error": "Sample not found"}, status=404)

    def handle_sample_thumbnail(self, filename):
        file_path = ROOT / "samples" / filename
        if file_path.exists() and file_path.is_file() and filename.endswith(".pdf"):
            try:
                import fitz
                doc = fitz.open(file_path)
                page = doc.load_page(0)
                pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5))
                img_data = pix.tobytes("jpeg")
                doc.close()
                self.send_response(200)
                self.send_header("Content-Type", "image/jpeg")
                self.send_header("Cache-Control", "public, max-age=86400")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(img_data)
                return
            except Exception as e:
                print(f"Thumbnail error: {e}")
        self.send_json({"error": "Thumbnail not found or could not be generated"}, status=404)

    def handle_suggest_outline(self):
        length = int(self.headers.get("Content-Length", "0"))
        if length <= 0 or length > 100_000:
            self.send_json({"error": "Invalid request."}, status=400)
            return

        try:
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
        except (json.JSONDecodeError, UnicodeDecodeError):
            self.send_json({"error": "The request could not be read."}, status=400)
            return

        api_key = str(payload.get("api_key") or os.getenv("GEMINI_API_KEY") or "").strip()
        topic = str(payload.get("topic") or "").strip()
        if not api_key or not topic:
            self.send_json({"error": "Add a topic first."}, status=400)
            return

        report_type = str(payload.get("reportType") or "Professional report")
        target_pages = int(payload.get("targetPages") or 70)
        custom_chapters = payload.get("customChapters", [])
        
        from engines import get_engine_for_report
        engine = get_engine_for_report(report_type, {})
        section_label = engine._get_section_label()
        report_profile = getattr(engine, "report_profile", "")

        if target_pages <= 5:
            num_chapters = 2
        elif target_pages <= 15:
            num_chapters = 3
        elif target_pages <= 25:
            num_chapters = 4
        elif target_pages <= 45:
            num_chapters = 5
        elif target_pages <= 70:
            num_chapters = 6
        else:
            num_chapters = 7

        if custom_chapters and len(custom_chapters) > num_chapters:
            num_chapters = len(custom_chapters)

        custom_directive = ""
        section_prefix = f"{section_label} " if section_label else ""
        
        if custom_chapters:
            explicit_chapters = []
            for i, ch in enumerate(custom_chapters):
                ch_title = ch.get("title", ch) if isinstance(ch, dict) else ch
                if ch_title and ch_title.strip() != "" and ch_title != "[AI Suggested Chapter]":
                    explicit_chapters.append(f"{section_prefix}{i+1}: {ch_title}")
                else:
                    if i == 0:
                        explicit_chapters.append(f"{section_prefix}1: Introduction")
                    elif i == num_chapters - 1:
                        explicit_chapters.append(f"{section_prefix}{num_chapters}: Conclusion")
                    else:
                        explicit_chapters.append(f"{section_prefix}{i+1}: [AI SHOULD GENERATE THIS SECTION]")
                        
            custom_directive = f"The user has provided constraints. You must include these exact sections at their positions:\n{json.dumps(explicit_chapters, indent=2)}\nFor sections marked [AI SHOULD GENERATE THIS SECTION], dynamically invent appropriate titles."
        else:
            custom_directive = f"Ensure Section 1 is 'Introduction'. Ensure Section {num_chapters} is 'Conclusion'."

        prompt = (
            f"Create a dynamic {num_chapters}-chapter outline for a report on: {topic}. "
            f"Report type: {report_type}.\n\n"
            f"Report Formatting Rules:\n{report_profile}\n\n"
            f"{custom_directive}\n"
            f"For each chapter, suggest 2 to 3 relevant subchapters/subsections. "
            f"Return a JSON object with a key 'chapters' which is an array of objects. "
            f"Each object must have two fields:\n"
            f"1. 'title': The chapter title string (do not include number prefixes like 'Chapter 1:')\n"
            f"2. 'subchapters': A JSON array of 2-3 suggested subchapter title strings for this chapter.\n\n"
            f"Example format:\n"
            f'{{\n  "chapters": [\n    {{\n      "title": "Introduction",\n      "subchapters": ["Background and Context", "Research Objectives"]\n    }}\n  ]\n}}'
        )

        try:
            from google.genai import types as genai_types
            client = genai.Client(
                api_key=api_key,
                http_options=genai_types.HttpOptions(
                    retry_options=genai_types.HttpRetryOptions(attempts=1),
                    timeout=15000,
                )
            )

            response = None
            for model_name in ["gemini-1.5-flash", "gemini-2.0-flash", "gemini-1.5-pro"]:
                try:
                    response = client.models.generate_content(
                        model=model_name,
                        contents=prompt,
                    )
                    break
                except Exception as e:
                    print(f"Call with {model_name} failed: {e}")
                    
            if not response:
                raise RuntimeError("All models failed.")
                
            raw_text = response.text
            json_start = raw_text.find("{")
            json_end = raw_text.rfind("}") + 1
            bundle = json.loads(raw_text[json_start:json_end])
            
            chapters = bundle.get("chapters", [])
            cleaned_chapters = []
            for i, ch in enumerate(chapters):
                if isinstance(ch, str):
                    cleaned_chapters.append({"title": ch, "subchapters": []})
                elif isinstance(ch, dict):
                    title = ch.get("title", f"Chapter {i+1}")
                    subchapters = ch.get("subchapters", [])
                    cleaned_chapters.append({"title": title, "subchapters": subchapters})
                    
            if cleaned_chapters:
                cleaned_chapters[0]["title"] = "Introduction"
                cleaned_chapters[-1]["title"] = "Conclusion"
                    
            self.send_json({"chapters": cleaned_chapters})
            
        except Exception as e:
            self.send_json({"error": str(e)}, status=500)

    def handle_login(self):
        try:
            length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(length)
            data = json.loads(body)
            id_token = data.get("credential", "").strip()
            
            if not id_token:
                self.send_json({"error": "No credential provided."}, status=400)
                return
                
            import urllib.request
            
            # Verify token
            req = urllib.request.Request(f"https://oauth2.googleapis.com/tokeninfo?id_token={id_token}")
            try:
                with urllib.request.urlopen(req) as response:
                    token_info = json.loads(response.read().decode())
            except Exception as e:
                self.send_json({"error": str(e)}, status=500)
                return
                
            if token_info.get("aud") != os.getenv("GOOGLE_CLIENT_ID", "513844870470-39ecvnte1gol2kqvcadcgf87lgda2fhf.apps.googleusercontent.com"):
                self.send_json({"error": "Invalid client ID."}, status=401)
                return
                
            email = token_info.get("email")
            if not email:
                self.send_json({"error": "No email in token."}, status=401)
                return
                
            user = STORAGE.get_user_by_email(email)
            if not user:
                user_id = uuid.uuid4().hex
                # Pass empty string for password hash since it's OAuth
                if not STORAGE.create_user(user_id, email, ""):
                    self.send_json({"error": "Database error creating user."}, status=500)
                    return
                user = {"id": user_id, "email": email}
                
            token = secrets.token_hex(32)
            expires_at = time.strftime('%Y-%m-%d %H:%M:%S', time.gmtime(time.time() + 30*86400))
            STORAGE.create_session(token, user["id"], expires_at)
            picture = token_info.get("picture", "")
            self.send_json({"status": "success", "token": token, "email": email, "picture": picture})
        except Exception as e:
            print(f"Login error: {e}")
            self.send_json({"error": "Invalid request or token verification failed."}, status=400)

    def handle_logout(self):
        # We don't delete from DB here, we just return success and frontend deletes token
        self.send_json({"status": "success"})

    def handle_samples(self):
        samples_dir = ROOT / "samples"
        samples = []
        if samples_dir.exists():
            for f in samples_dir.iterdir():
                if f.is_file():
                    stat = f.stat()
                    page_count = 0
                    if f.name.endswith(".pdf"):
                        try:
                            import fitz
                            doc = fitz.open(f)
                            page_count = len(doc)
                            doc.close()
                        except Exception:
                            pass
                    
                    samples.append({
                        "name": f.name,
                        "url": f"/api/samples/download/{f.name}",
                        "size_bytes": stat.st_size,
                        "created_at": stat.st_ctime,
                        "page_count": page_count,
                        "thumbnail_url": f"/api/samples/thumbnail/{f.name}" if f.name.endswith(".pdf") else None
                    })
        self.send_json({"samples": samples})

    def handle_user_reports(self, query):
        token = self.headers.get("Authorization", "").replace("Bearer ", "").strip()
        if not token:
            self.send_json({"error": "Unauthorized"}, status=401)
            return
            
        session = STORAGE.get_session(token)
        if not session:
            self.send_json({"error": "Invalid or expired token"}, status=401)
            return
            
        jobs = STORAGE.get_user_jobs(session["user_id"])
        
        # Merge with in-memory running jobs for real-time progress
        for j in jobs:
            mem = JOBS.get(j["id"])
            if mem:
                j["payload"] = {**j["payload"], **mem}
            
            # Inject signed URLs for history downloads
            if j["payload"].get("status") == "complete":
                j["payload"]["docx_url"] = generate_signed_url(j["id"], "docx")
                j["payload"]["pdf_url"] = generate_signed_url(j["id"], "pdf")
                j["payload"]["md_url"] = generate_signed_url(j["id"], "md")
                
        self.send_json({"status": "success", "reports": jobs})


def main():
    os.chdir(ROOT)
    server = ThreadingHTTPServer((HOST, PORT), ReportRequestHandler)
    try:
        print(f"ReportForge running at http://{HOST}:{PORT} using {STORAGE.backend_name} storage")
        print("Press Ctrl+C to stop.")
    except OSError:
        pass
    server.serve_forever()



def cleanup_thread():
    while True:
        try:
            STORAGE.cleanup_old_jobs(1)
        except Exception as e:
            print(f"Cleanup error: {e}")
        time.sleep(3600) # run every hour

t = threading.Thread(target=cleanup_thread, daemon=True)
t.start()

if __name__ == "__main__":
    main()
