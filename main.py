from fastapi import FastAPI
from fastapi.responses import JSONResponse, FileResponse
from pydantic import BaseModel
from docx import Document
import os
import re

app = FastAPI()

TEMPLATE_PATH = "templates/納保申請書_官方格式_可套填_v8.docx"
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)


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
    changed = False

    for key, value in replacements.items():
        if key in text:
            text = text.replace(key, value)
            changed = True

    if changed:
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

        safe_filename = f"application_{data.apply_year}{data.apply_month}{data.apply_day}.docx"
        display_filename = f"申請書_{data.apply_year}{data.apply_month}{data.apply_day}.docx"

        output_path = os.path.join(OUTPUT_DIR, safe_filename)

        doc.save(output_path)

        check_doc = Document(output_path)

        if has_placeholder(check_doc):
            return JSONResponse(
                status_code=500,
                content={
                    "success": False,
                    "message": "產製失敗",
                    "reason": "Word 殘留 placeholder"
                }
            )

        return JSONResponse(
            content={
                "success": True,
                "filename": display_filename,
                "download_url": f"https://tax-word-system-production.up.railway.app/download/{safe_filename}"
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


@app.get("/download/{filename}")
def download_file(filename: str):

    path = os.path.join(OUTPUT_DIR, filename)

    if not os.path.exists(path):
        return JSONResponse(
            status_code=404,
            content={
                "success": False,
                "message": "file not found"
            }
        )

    return FileResponse(
        path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
