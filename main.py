from fastapi import FastAPI
from fastapi.responses import FileResponse
from pydantic import BaseModel
from docxtpl import DocxTemplate
import os

app = FastAPI()

TEMPLATE_PATH = "templates/納保申請書_官方格式_可套填_v8.docx"
OUTPUT_DIR = "output"

os.makedirs(OUTPUT_DIR, exist_ok=True)


class RequestData(BaseModel):
    apply_year: str
    apply_month: str
    apply_day: str

    formal_statement: str

    case_category_suggestion: str
    tax_items_suggestion: str

    apply_method_suggestion: str
    reply_method_suggestion: str
    notify_method_suggestion: str

    notify_email: str
    evidence_list: str


@app.post("/generate-word")
def generate_word(data: RequestData):

    doc = DocxTemplate(TEMPLATE_PATH)

    context = {
        "apply_year": data.apply_year,
        "apply_month": data.apply_month,
        "apply_day": data.apply_day,

        "formal_statement": data.formal_statement,

        "case_category_suggestion": data.case_category_suggestion,
        "tax_items_suggestion": data.tax_items_suggestion,

        "apply_method_suggestion": data.apply_method_suggestion,
        "reply_method_suggestion": data.reply_method_suggestion,
        "notify_method_suggestion": data.notify_method_suggestion,

        "notify_email": data.notify_email,
        "evidence_list": data.evidence_list
    }

    doc.render(context)

    filename = f"申請書_{data.apply_year}{data.apply_month}{data.apply_day}.docx"
    filepath = os.path.join(OUTPUT_DIR, filename)

    doc.save(filepath)

    return FileResponse(
        path=filepath,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
