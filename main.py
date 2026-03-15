from fastapi import FastAPI
from fastapi.responses import FileResponse
from pydantic import BaseModel
from datetime import datetime
from zoneinfo import ZoneInfo
from docx import Document
import uuid

app = FastAPI()

class StatementRequest(BaseModel):
    statement: str
    case_category: str
    tax_items: str
    apply_method: int
    reply_method: int
    notify_method: int
    notify_email: str = ""

TEMPLATE_PATH = "templates/納保申請書_官方格式_可套填_v8.docx"

apply_method_map = {
    1: "現場申請",
    2: "書面或傳真申請"
}

reply_method_map = {
    1: "現場答復",
    2: "公文",
    3: "電話"
}

notify_method_map = {
    1: "電子郵件",
    2: "簡訊",
    3: "無須通知"
}

@app.post("/generate-statement")
def generate_statement(req: StatementRequest):

    now = datetime.now(ZoneInfo("Asia/Taipei"))

    roc_year = now.year - 1911
    month = now.month
    day = now.day

    apply_method_text = apply_method_map.get(req.apply_method, "")
    reply_method_text = reply_method_map.get(req.reply_method, "")
    notify_method_text = notify_method_map.get(req.notify_method, "")

    if req.notify_method == 1 and req.notify_email == "":
        return {"error": "電子郵件通知必須提供 email"}
    
    doc = Document(TEMPLATE_PATH)

    replacements = {
        "{{apply_year}}": str(roc_year),
        "{{apply_month}}": str(month),
        "{{apply_day}}": str(day),

        "{{formal_statement}}": req.statement,

        "{{case_category_suggestion}}": req.case_category,
        "{{tax_items_suggestion}}": req.tax_items,

        "{{apply_method_suggestion}}": apply_method_text,
        "{{reply_method_suggestion}}": reply_method_text,
        "{{notify_method_suggestion}}": notify_method_text,

        "{{notify_email}}": req.notify_email,

        "{{representative_name}}": "",
        "{{representative_address}}": "",
        "{{representative_phone}}": "",

        "{{agent_name}}": "",
        "{{agent_address}}": "",
        "{{agent_phone}}": "",

        "{{evidence_list}}": ""
    }

    for p in doc.paragraphs:
        for key, val in replacements.items():
            if key in p.text:
                p.text = p.text.replace(key, val)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, val in replacements.items():
                        if key in p.text:
                            p.text = p.text.replace(key, val)

    for p in doc.paragraphs:
        if "{{" in p.text:
            return {"error": "Word仍殘留placeholder"}

    filename = f"申請書_{roc_year}{month:02d}{day:02d}.docx"

    output_path = f"/tmp/{uuid.uuid4()}.docx"

    doc.save(output_path)

    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=filename
    )
