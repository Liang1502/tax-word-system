from fastapi import FastAPI
from fastapi.responses import JSONResponse, StreamingResponse
from pydantic import BaseModel
from docx import Document
import re
import uuid
import time
import asyncio
import io

app = FastAPI()

TEMPLATE_PATH = "templates/納保申請書_官方格式_可套填_v8.docx"

download_store = {}


class GenerateRequest(BaseModel):

    apply_year: str
    apply_month: str
    apply_day: str

    formal_statement: str

    case_category_suggestion: str
    tax_items_suggestion: str

    apply_method_suggestion: str
    reply_method_suggestion: str
    notify_method_suggestion: str

    notify_email: str = ""

    representative_name: str = ""
    representative_address: str = ""
    representative_phone: str = ""

    agent_name: str = ""
    agent_address: str = ""
    agent_phone: str = ""

    evidence_list: str = ""


def replace_text(paragraph, replacements):

    text = paragraph.text

    for key, value in replacements.items():
        if key in text:
            text = text.replace(key, value)

    paragraph.text = text


def replace_all(doc, replacements):

    for p in doc.paragraphs:
        replace_text(p, replacements)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_text(p, replacements)


def remove_email_label(doc):

    targets = ["{{notify_email}}", "信箱", "Email", "（信箱：）"]

    for p in doc.paragraphs:
        text = p.text
        for t in targets:
            text = text.replace(t, "")
        p.text = text

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    text = p.text
                    for t in targets:
                        text = text.replace(t, "")
                    p.text = text


def has_placeholder(doc):

    pattern = re.compile(r"\{\{.*?\}\}")

    for p in doc.paragraphs:
        if pattern.search(p.text):
            return True

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if pattern.search(p.text):
                        return True

    return False


@app.post("/generate-word")
def generate_word(data: GenerateRequest):

    try:

        doc = Document(TEMPLATE_PATH)

        replacements = {

            "{{apply_year}}": data.apply_year,
            "{{apply_month}}": data.apply_month,
            "{{apply_day}}": data.apply_day,

            "{{formal_statement}}": data.formal_statement,

            "{{case_category_suggestion}}": data.case_category_suggestion,
            "{{tax_items_suggestion}}": data.tax_items_suggestion,

            "{{apply_method_suggestion}}": data.apply_method_suggestion,
            "{{reply_method_suggestion}}": data.reply_method_suggestion,
            "{{notify_method_suggestion}}": data.notify_method_suggestion,

            "{{notify_email}}": data.notify_email,

            "{{representative_name}}": data.representative_name,
            "{{representative_address}}": data.representative_address,
            "{{representative_phone}}": data.representative_phone,

            "{{agent_name}}": data.agent_name,
            "{{agent_address}}": data.agent_address,
            "{{agent_phone}}": data.agent_phone,

            "{{evidence_list}}": data.evidence_list
        }

        replace_all(doc, replacements)

        if data.notify_email.strip() == "":
            remove_email_label(doc)

        if has_placeholder(doc):

            return JSONResponse(
                status_code=500,
                content={
                    "success": False,
                    "message": "產製失敗",
                    "reason": "Word 殘留 placeholder"
                }
            )

        memory_file = io.BytesIO()
        doc.save(memory_file)
        memory_file.seek(0)

        token = str(uuid.uuid4())

        filename = f"申請書_{data.apply_year}{data.apply_month}{data.apply_day}.docx"

        download_store[token] = {
            "file": memory_file,
            "filename": filename,
            "expire_time": time.time() + 1800,
            "used": False
        }

        return JSONResponse(
            content={
                "success": True,
                "filename": filename,
                "download_url": f"https://tax-word-system-production.up.railway.app/download/{token}"
            }
        )

    except Exception as e:

        return JSONResponse(
            status_code=500,
            content={
                "success": False,
                "message": "產製失敗",
                "reason": str(e)
            }
        )


@app.get("/download/{token}")
def download_file(token: str):

    if token not in download_store:
        return JSONResponse(status_code=404, content={"message": "file not found"})

    record = download_store[token]

    if time.time() > record["expire_time"]:
        del download_store[token]
        return JSONResponse(status_code=410, content={"message": "link expired"})

    if record["used"]:
        return JSONResponse(status_code=410, content={"message": "file already downloaded"})

    record["used"] = True

    record["file"].seek(0)

    response = StreamingResponse(
        record["file"],
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    response.headers["Content-Disposition"] = f'attachment; filename="{record["filename"]}"'

    del download_store[token]

    return response


async def cleanup_expired():

    while True:

        now = time.time()

        expired = []

        for token, record in download_store.items():
            if now > record["expire_time"]:
                expired.append(token)

        for token in expired:
            del download_store[token]

        await asyncio.sleep(300)


@app.on_event("startup")
async def startup_event():

    asyncio.create_task(cleanup_expired())
