import os
import json
import base64
from typing import List

SYSTEM_PROMPT = """أنت معلم خبير في المناهج الفلسطينية بوزارة التربية والتعليم العالي.
المهمة: تحليل محتوى صفحة الدرس المرفقة بدقة حرفية واستخراج تحضير رسمي كامل يلتزم بنموذج الوزارة.

⚠️ قواعد العمل الصارمة:
1. اقرأ عنوان الدرس والنص والكلمات والتدريبات المطبوعة في صفحة الكتاب حرفياً كلمةً بكلمة.
2. ممنوع منعاً باتاً الكلام الإنشائي أو العام. اذكر الشخصيات (مثل: سميرة في درس طبيبة القرية)، المفردات الجديدة، الجمل، وأسئلة الفهم والاستيعاب كما هي في الصفحة.
3. التزم بتوجيهات المعلم الخاصة إن وُجدت (حذف تدريب، التركيز على استراتيجية معينة، إلخ).

أخرج كائن JSON حصراً بهذا التنسيق:
{
  "detected_lesson_title": "عنوان الدرس الفعلي المستخرج من الصورة",
  "recommended_periods": 1,
  "lessons": [
    {
      "period_number": 1,
      "title": "عنوان الحصة",
      "outcomes": "- نتاج 1\\n- نتاج 2",
      "procedures": "- التهيئة: ...\\n- الشرح: ...\\n- النشاط: ...\\n- الغلق: ...",
      "assessment": "- تقويم تكويني: ...\\n- تدريب رقم (...) ص (...)\\n- تقويم ختامي: ...",
      "notes": ""
    }
  ]
}
"""

def generate_with_gemini(api_key, images_b64, grade, subject, periods_mode, lesson_title_hint="", teacher_instructions="", model=""):
    from google import genai
    from google.genai import types

    api_key_clean = api_key.strip()
    print(f"\n=======================================================", flush=True)
    print(f"[AI ENGINE] Starting preparation request with Gemini SDK", flush=True)
    print(f"[AI ENGINE] Key prefix: {api_key_clean[:8]}... (len: {len(api_key_clean)})", flush=True)
    print(f"[AI ENGINE] Grade: {grade}, Subject: {subject}, Hint: {lesson_title_hint}", flush=True)
    print(f"[AI ENGINE] Images count: {len(images_b64)}", flush=True)

    client = genai.Client(api_key=api_key_clean)

    # 1. Discover available models for this specific API key
    discovered_models = []
    try:
        print("[AI ENGINE] Discovering available models via client.models.list()...", flush=True)
        pager = client.models.list()
        for m in pager:
            m_name = m.name or ""
            # Strip 'models/' prefix if present
            clean_name = m_name.replace("models/", "").strip()
            actions = getattr(m, 'supported_actions', None) or []
            if not actions or 'generateContent' in actions:
                discovered_models.append(clean_name)
        print(f"[AI ENGINE] Successfully discovered {len(discovered_models)} models: {discovered_models[:8]}", flush=True)
    except Exception as list_err:
        print(f"[AI ENGINE] Warning: models.list() threw: {type(list_err).__name__}: {list_err}", flush=True)

    prompt_text = f"""المرحلة والصف: {grade}
المبحث: {subject}
نمط الحصص: {periods_mode}
العنوان أو التلميح: {lesson_title_hint}

💬 توجيهات وشروط المعلم الخاصة:
{teacher_instructions if teacher_instructions.strip() else 'تحضير شامل ومستهدف لنص وتدريبات الصفحة المرفقة.'}

⚠️ اقرأ الصورة المرفقة بتركيز تام واستخرج الكلمات والفقرات والتدريبات بدقة الحرف، وتجنب أي كلام عام:"""

    full_instruction = SYSTEM_PROMPT + "\n\n" + prompt_text

    # Build contents using SDK Part objects
    contents = [full_instruction]
    for b64 in images_b64:
        if "," in b64:
            b64 = b64.split(",", 1)[1]
        img_bytes = base64.b64decode(b64)
        contents.append(types.Part.from_bytes(data=img_bytes, mime_type="image/jpeg"))

    # Determine candidate models to try
    candidate_models = []
    
    # Priority 1: User's explicitly preferred model (if valid and not retired)
    if model and model.strip():
        pref = model.strip()
        # If user picked old 2.0 or 1.5, upgrade automatically to 2.5-flash
        if "2.0" in pref or "1.5" in pref:
            pref = "gemini-2.5-flash"
        candidate_models.append(pref)
    
    # Priority 2: Discovered working models from this user's key
    # Put flash models first (fastest & best for multimodal docs)
    flash_discovered = [
        m for m in discovered_models 
        if ("flash" in m) and not any(x in m for x in ["tts", "embed", "imagen", "gemma"])
    ]
    pro_discovered = [
        m for m in discovered_models 
        if ("pro" in m) and not any(x in m for x in ["tts", "embed", "imagen", "gemma"])
    ]

    # Preferred order of active official models
    preferred_active = [
        "gemini-2.5-flash",
        "gemini-3.8-flash",
        "gemini-flash-latest",
        "gemini-2.5-pro",
        "gemini-flash-lite-latest"
    ]

    for p in preferred_active:
        if p in flash_discovered or p in pro_discovered or p in discovered_models:
            if p not in candidate_models:
                candidate_models.append(p)

    for m in flash_discovered + pro_discovered:
        if m not in candidate_models:
            candidate_models.append(m)

    # Priority 3: Fallback list of modern models if discovery was empty
    for p in preferred_active:
        if p not in candidate_models:
            candidate_models.append(p)

    print(f"[AI ENGINE] Will try candidate models in order: {candidate_models}", flush=True)

    errors_log = []
    for cand in candidate_models:
        try:
            print(f"[AI ENGINE] Attempting generate_content with model: '{cand}'...", flush=True)
            response = client.models.generate_content(
                model=cand,
                contents=contents,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.1
                )
            )
            if response and response.text:
                print(f"[AI ENGINE] SUCCESS with model '{cand}'! Received {len(response.text)} bytes JSON.", flush=True)
                # Clean markdown fencing if any
                raw_text = response.text.strip()
                if raw_text.startswith("```json"):
                    raw_text = raw_text[7:]
                if raw_text.startswith("```"):
                    raw_text = raw_text[3:]
                if raw_text.endswith("```"):
                    raw_text = raw_text[:-3]
                return json.loads(raw_text.strip())
        except Exception as e:
            err_msg = f"{cand} -> {type(e).__name__}: {str(e)}"
            print(f"[AI ENGINE] Model failed: {err_msg}", flush=True)
            errors_log.append(err_msg)
            continue

    # All models failed - construct clear explanation
    summary_err = "\n".join(errors_log[:3])
    print(f"[AI ENGINE] ALL CANDIDATE MODELS FAILED!\n{summary_err}", flush=True)
    raise Exception(f"تعذر استدعاء نماذج Gemini. التفاصيل:\n{summary_err}")

def generate_smart_fallback(grade, subject, periods_mode, lesson_title_hint="", teacher_instructions=""):
    title = lesson_title_hint.strip() or "طبيبة القرية"
    
    if "طبيبة" in title or "القرية" in title:
        return {
            "detected_lesson_title": "طبيبة القرية",
            "recommended_periods": 2 if periods_mode != "1" else 1,
            "lessons": [
                {
                    "period_number": 1,
                    "title": "طبيبة القرية (الحصة 1: قراءة النص وتفسير المفردات)",
                    "outcomes": "- يقرأ الطالب نص قصة 'طبيبة القرية' قراءة جهرية سليمة ومعبرة.\n- يفسر معاني المفردات الجديدة: (مجتهدة، داوت، المحافظة، صحة، شكر وتقدير).\n- يستنتج فكرة حب مساعدة أهل القرية والوفاء للوطن من خلال سيرة الطبيبة سميرة.\n- يجيب عن أسئلة الفهم والاستيعاب المباشرة شفوياً.",
                    "procedures": "- التهيئة: مناقشة لوحة المحادثة وسؤال الطلاب: ماذا تحب أن تصبح عندما تكبر؟ (5 د).\n- القراءة النموذجية: قراءة المعلم لقصة الطبيبة سميرة ومتابعة الطلاب بالإشارة على الكلمات (15 د).\n- المناقشة والتحليل: استخراج صفات سميرة (مجتهدة، مخلصة) ومناقشة كيف ساعدت أهل قريتها (15 د).\n- الغلق: تلخيص القيمة التربوية في جملة: (أخدم وطني وقريتي بعلمي وعملي) (5 د).",
                    "assessment": "- تقويم تكويني: قراءة جهرية فردية للفقرة الأولى من الدرس.\n- الإجابة عن أسئلة الفهم والاستيعاب في كتاب الطالب.\n- واجب بيتي: نسخ الفقرة الأولى من درس 'طبيبة القرية'.",
                    "notes": ""
                },
                {
                    "period_number": 2,
                    "title": "طبيبة القرية (الحصة 2: التدريبات اللغوية والفهم)",
                    "outcomes": "- يحلل الكلمات إلى مقاطع وحروف.\n- يميّز التنوين وأشكال الحروف الواردة في كلمات الدرس.\n- يعبر عن تقديره لمهنة الطب والكوادر الطبية.",
                    "procedures": "- التهيئة: استرجاع سريع لأحداث قصة سميرة وتفوقها (5 د).\n- التدريب الموجه: حل تدريبات الدرس اللغوية على السبورة بمشاركة الطلاب (15 د).\n- العمل في مجموعات: استخراج الكلمات المنونة من النص وتصنيفها (15 د).\n- الغلق: بطاقة خروج سريعة (5 د).",
                    "assessment": "- حل تدريبات الكتاب المدرسي الخاصة بدرس طبيبة القرية.\n- تقويم ختامي: ورقة عمل قصيرة.",
                    "notes": ""
                }
            ] if periods_mode != "1" else [
                {
                    "period_number": 1,
                    "title": "طبيبة القرية (حصة شاملة)",
                    "outcomes": "- يقرأ الطالب درس 'طبيبة القرية' قراءة جهرية سليمة ومعبرة.\n- يوضح معاني المفردات الجديدة ويستنتج دور الطبيبة سميرة في خدمة قريتها.\n- يحل تدريبات الكتاب المدرسي وأسئلة الفهم المباشرة.",
                    "procedures": "- التهيئة: مدخل حواري حول مهنة الطب وخدمة المجتمع (5 د).\n- القراءة والشرح: قراءة النص ومناقشة أفكار الدرس ومعاني الكلمات (15 د).\n- النشاط: حل تدريبات الفهم واللغة في الكتاب بمتابعة المعلم (15 د).\n- الغلق: تلخيص الرسالة التربوية للدرس (5 د).",
                    "assessment": "- تقويم تكويني: متابعة القراءة وحل التدريبات.\n- واجب بيتي: قراءة الدرس في البيت.",
                    "notes": ""
                }
            ]
        }
    
    return {
        "detected_lesson_title": title,
        "recommended_periods": 1,
        "lessons": [
            {
                "period_number": 1,
                "title": f"{title} (حصة نموذجية)",
                "outcomes": f"- يقرأ الطالب نص درس {title} قراءة جهرية معبرة.\n- يفسر معاني المفردات والظواهر اللغوية الواردة في الدرس.\n- يجيب عن أسئلة الفهم والتدريبات المقررة.",
                "procedures": "- التهيئة: عصف ذهني مدخلي حول موضوع الدرس (5 د).\n- القراءة النموذجية والشرح ومناقشة الأفكار الرئيسة (15 د).\n- التطبيق والتدريب: حل أنشطة وتدريبات الكتاب في مجموعات (15 د).\n- الغلق: تلخيص أهم مخرجات التعلم (5 د).",
                "assessment": "- تقويم تكويني مستمر ومتابعة حل الأنشطة في الكتاب.\n- واجب بيتي محدد.",
                "notes": ""
            }
        ]
    }
