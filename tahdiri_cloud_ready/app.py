import os
import uuid
import json
import base64
import io
import re
from datetime import datetime, timedelta
from typing import List, Optional

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from PIL import Image

import doc_builder
import ai_engine
import importlib

app = FastAPI(title="منصة تحضيري - المعلم الفلسطيني")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXPORTS_DIR = os.path.join(BASE_DIR, "exports")
STATIC_DIR = os.path.join(BASE_DIR, "static")
os.makedirs(EXPORTS_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)

STORE = {}

def compress_image_to_b64(content: bytes, max_dim: int = 1400) -> str:
    """Compress and resize image to prevent socket timeout and upload fast."""
    try:
        im = Image.open(io.BytesIO(content))
        if im.mode in ("RGBA", "P"):
            im = im.convert("RGB")
        im.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
        buf = io.BytesIO()
        im.save(buf, format="JPEG", quality=82, optimize=True)
        return base64.b64encode(buf.getvalue()).decode("utf-8")
    except Exception:
        return base64.b64encode(content).decode("utf-8")

def format_date_for_period(start_date_str, period_idx):
    arabic_days = ["الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت", "الأحد"]
    try:
        dt = datetime.strptime(start_date_str, "%Y-%m-%d")
    except Exception:
        dt = datetime.now()
    
    current = dt
    added = 0
    while added < period_idx:
        current += timedelta(days=1)
        if current.weekday() not in [4, 5]: # skip weekend
            added += 1
            
    day_name = arabic_days[current.weekday()]
    return f"{day_name} {current.strftime('%Y/%m/%d')}"

@app.post("/api/test_key")
async def test_key(key: str = Form("")):
    clean_key = key.strip()
    if not clean_key:
        return JSONResponse(status_code=400, content={"success": False, "error": "يرجى كتابة أو لصق المفتاح أولاً."})
    
    try:
        from google import genai
        client = genai.Client(api_key=clean_key)
        discovered = []
        pager = client.models.list()
        for m in pager:
            name = (m.name or "").replace("models/", "")
            actions = getattr(m, 'supported_actions', None) or []
            if not actions or 'generateContent' in actions:
                discovered.append(name)
        
        if not discovered:
            return JSONResponse(status_code=400, content={
                "success": False,
                "error": "المفتاح يعمل لكن لا توجد نماذج دردشة مفعلة في هذا المشروع. يُرجى إنشاء مفتاح جديد من Google AI Studio."
            })
            
        return {
            "success": True,
            "message": f"المفتاح يعمل بنجاح 100%! تم العثور على {len(discovered)} نموذج متاح.",
            "models": discovered[:6]
        }
    except Exception as e:
        err_str = str(e)
        if "API_KEY_INVALID" in err_str or "API key not valid" in err_str:
            msg = "مفتاح API غير صالح أو ناقص. تأكد من نسخه بشكل صحيح من Google AI Studio."
        elif "PERMISSION_DENIED" in err_str or "disabled" in err_str.lower():
            msg = "مشروع المفتاح لا يمتلك صلاحية Generative Language API. يرجى الدخول إلى aistudio.google.com والضغط على Create API Key."
        elif "location" in err_str.lower() or "region" in err_str.lower():
            msg = "خدمة Gemini محظورة في منطقتك الجغرافية الحالية. يلزم تشغيل تطبيق VPN لتجاوز الحظر."
        else:
            msg = f"خطأ أثناء فحص المفتاح: {err_str}"
        return JSONResponse(status_code=400, content={"success": False, "error": msg})

@app.post("/api/prepare")
async def prepare_lesson(
    grade: str = Form(...),
    subject: str = Form(...),
    periods_mode: str = Form("auto"),
    teacher_name: str = Form(""),
    school_name: str = Form(""),
    date_start: str = Form(""),
    lesson_title_hint: str = Form(""),
    pdf_page: str = Form(""),
    teacher_instructions: str = Form(""),
    gemini_api_key: str = Form(""),
    gemini_model: str = Form("gemini-2.5-flash"),
    images: List[UploadFile] = File([])
):
    try:
        importlib.reload(ai_engine)
        importlib.reload(doc_builder)
        images_b64 = []

        # Process uploaded files (Images or PDF)
        for img in images:
            content = await img.read()
            if not content:
                continue

            filename = img.filename.lower() if img.filename else ""
            
            # Handle PDF uploads
            if filename.endswith(".pdf") or (img.content_type and "pdf" in img.content_type):
                import pymupdf
                doc_pdf = pymupdf.open(stream=content, filetype="pdf")
                total_pages = len(doc_pdf)

                # Determine target page from pdf_page or teacher_instructions
                target_page_idx = None
                p_match = re.search(r'(?:صفحة|ص|page)\s*[:=]?\s*(\d+)', teacher_instructions + " " + pdf_page, re.IGNORECASE)
                if p_match:
                    page_num = int(p_match.group(1))
                    if 1 <= page_num <= total_pages:
                        target_page_idx = page_num - 1

                if target_page_idx is not None:
                    pix = doc_pdf[target_page_idx].get_pixmap(dpi=180)
                    b64 = compress_image_to_b64(pix.tobytes("jpeg"))
                    images_b64.append(b64)
                elif total_pages <= 5:
                    for p in doc_pdf:
                        pix = p.get_pixmap(dpi=180)
                        b64 = compress_image_to_b64(pix.tobytes("jpeg"))
                        images_b64.append(b64)
                else:
                    found_idx = None
                    if lesson_title_hint.strip():
                        hint = lesson_title_hint.strip().split()[0]
                        for idx, p in enumerate(doc_pdf):
                            if hint in p.get_text():
                                found_idx = idx
                                break
                    target_idx = found_idx if found_idx is not None else 0
                    pix = doc_pdf[target_idx].get_pixmap(dpi=180)
                    b64 = compress_image_to_b64(pix.tobytes("jpeg"))
                    images_b64.append(b64)
            else:
                # Regular image - compress to prevent timeouts
                b64 = compress_image_to_b64(content)
                images_b64.append(b64)

        api_key = gemini_api_key.strip() or os.environ.get("GEMINI_API_KEY", "").strip()

        # If user uploaded images but didn't provide API key:
        if images_b64 and not api_key:
            return JSONResponse(
                status_code=400,
                content={
                    "success": False,
                    "error": "لقراءة تفاصيل واستخراج نصوص الصور الجديدة بالذكاء الاصطناعي، يرجى إدخال مفتاح Gemini API (اضغط على زر المفتاح في الأعلى للحصول عليه مجاناً ولصقه)."
                }
            )
        
        # If API key is provided, use Gemini Vision
        if api_key:
            try:
                ai_result = ai_engine.generate_with_gemini(
                    api_key=api_key,
                    images_b64=images_b64,
                    grade=grade,
                    subject=subject,
                    periods_mode=periods_mode,
                    lesson_title_hint=lesson_title_hint,
                    teacher_instructions=teacher_instructions,
                    model=gemini_model
                )
            except Exception as e:
                # DO NOT silently swallow error! Report to user
                return JSONResponse(
                    status_code=400,
                    content={
                        "success": False,
                        "error": f"خطأ أثناء قراءة الصور عبر الذكاء الاصطناعي ({str(e)}). تأكد من صحة المفتاح واتصال الإنترنت."
                    }
                )
        else:
            # Fallback only when no images are uploaded and testing with a title
            ai_result = ai_engine.generate_smart_fallback(
                grade=grade,
                subject=subject,
                periods_mode=periods_mode,
                lesson_title_hint=lesson_title_hint,
                teacher_instructions=teacher_instructions
            )

        # Build full lesson plan object
        start_date = date_start or datetime.now().strftime("%Y-%m-%d")
        final_lessons = []
        for idx, les in enumerate(ai_result.get("lessons", [])):
            day_date = format_date_for_period(start_date, idx)
            final_lessons.append({
                "period_number": idx + 1,
                "day_date": day_date,
                "title": les.get("title", ""),
                "outcomes": les.get("outcomes", ""),
                "procedures": les.get("procedures", ""),
                "assessment": les.get("assessment", ""),
                "notes": les.get("notes", "")
            })

        plan_data = {
            "grade": grade,
            "subject": subject,
            "teacher_name": teacher_name,
            "school_name": school_name,
            "detected_lesson_title": ai_result.get("detected_lesson_title", lesson_title_hint),
            "lessons": final_lessons
        }

        # Generate unique ID and files
        plan_id = str(uuid.uuid4())[:8]
        docx_filename = f"تحضير_{subject}_{plan_id}.docx"
        pdf_filename = f"تحضير_{subject}_{plan_id}.pdf"
        
        docx_path = os.path.join(EXPORTS_DIR, docx_filename)
        pdf_path = os.path.join(EXPORTS_DIR, pdf_filename)
        
        emblem_path = os.path.join(BASE_DIR, "palestine_emblem_trans.png")
        doc_builder.generate_palestine_lesson_docx(plan_data, docx_path, emblem_path=emblem_path)
        
        # Convert to PDF
        has_pdf = False
        try:
            doc_builder.convert_docx_to_pdf(docx_path, pdf_path)
            has_pdf = True
        except Exception as err:
            print("PDF conversion error:", err)

        STORE[plan_id] = {
            "data": plan_data,
            "docx_path": docx_path,
            "pdf_path": pdf_path if has_pdf else None,
            "docx_filename": docx_filename,
            "pdf_filename": pdf_filename
        }

        return {
            "success": True,
            "id": plan_id,
            "data": plan_data,
            "has_pdf": has_pdf,
            "docx_url": f"/api/download/{plan_id}/docx",
            "pdf_url": f"/api/download/{plan_id}/pdf" if has_pdf else None
        }

    except Exception as e:
        import traceback
        traceback.print_exc()
        return JSONResponse(status_code=500, content={"success": False, "error": str(e)})

@app.post("/api/export_update")
async def export_update(payload: dict):
    try:
        plan_id = payload.get("id")
        plan_data = payload.get("data")
        if not plan_id or not plan_data:
            raise HTTPException(status_code=400, detail="Missing plan_id or data")

        docx_filename = f"تحضير_{plan_data.get('subject', 'درس')}_{plan_id}.docx"
        pdf_filename = f"تحضير_{plan_data.get('subject', 'درس')}_{plan_id}.pdf"
        
        docx_path = os.path.join(EXPORTS_DIR, docx_filename)
        pdf_path = os.path.join(EXPORTS_DIR, pdf_filename)
        
        emblem_path = os.path.join(BASE_DIR, "palestine_emblem_trans.png")
        doc_builder.generate_palestine_lesson_docx(plan_data, docx_path, emblem_path=emblem_path)

        has_pdf = False
        try:
            doc_builder.convert_docx_to_pdf(docx_path, pdf_path)
            has_pdf = True
        except Exception as err:
            print("PDF conversion error:", err)

        STORE[plan_id] = {
            "data": plan_data,
            "docx_path": docx_path,
            "pdf_path": pdf_path if has_pdf else None,
            "docx_filename": docx_filename,
            "pdf_filename": pdf_filename
        }

        return {
            "success": True,
            "id": plan_id,
            "has_pdf": has_pdf,
            "docx_url": f"/api/download/{plan_id}/docx",
            "pdf_url": f"/api/download/{plan_id}/pdf" if has_pdf else None
        }
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "error": str(e)})

@app.get("/api/download/{plan_id}/{fmt}")
async def download_file(plan_id: str, fmt: str):
    info = STORE.get(plan_id)
    if not info:
        target_docx = None
        target_pdf = None
        for f in os.listdir(EXPORTS_DIR):
            if plan_id in f and f.endswith(".docx"):
                target_docx = os.path.join(EXPORTS_DIR, f)
            if plan_id in f and f.endswith(".pdf"):
                target_pdf = os.path.join(EXPORTS_DIR, f)
        if target_docx or target_pdf:
            info = {
                "docx_path": target_docx,
                "pdf_path": target_pdf,
                "docx_filename": os.path.basename(target_docx) if target_docx else "lesson.docx",
                "pdf_filename": os.path.basename(target_pdf) if target_pdf else "lesson.pdf"
            }
        else:
            raise HTTPException(status_code=404, detail="File not found")

    if fmt == "docx":
        path = info["docx_path"]
        name = info["docx_filename"]
        media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    elif fmt == "pdf":
        path = info["pdf_path"]
        name = info["pdf_filename"]
        media_type = "application/pdf"
    else:
        raise HTTPException(status_code=400, detail="Unknown format")

    if not path or not os.path.exists(path):
        raise HTTPException(status_code=404, detail="Requested file does not exist")

    from urllib.parse import quote
    encoded_name = quote(name)
    headers = {
        "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_name}"
    }
    return FileResponse(path, media_type=media_type, headers=headers)

app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="127.0.0.1", port=8000, reload=True)
