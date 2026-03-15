from fastapi import FastAPI
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel
from docxtpl import DocxTemplate
from docx import Document
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


@app.get("/")
def root():
    return {"message": "tax-word-system running"}


@app.post("/generate-word")
def generate_word(data: RequestData):

    if not os.path.exists(TEMPLATE_PATH):
        return JSONResponse(
            status_code=500,
            content={
                "message": "產製失敗",
                "reason": f"找不到模板檔案：{TEMPLATE_PATH}"
            }
        )

    try:

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

        filename = f"申請書_{data.apply_year}{data.apply_month}{data.apply_day}.docx"
        filepath = os.path.join(OUTPUT_DIR, filename)

        doc.render(context)
        doc.save(filepath)

        # 再開一次檢查 placeholder
        check_doc = Document(filepath)

        for p in check_doc.paragraphs:
            if "{{" in p.text and "}}" in p.text:
                return JSONResponse(
                    status_code=500,
                    content={
                        "message": "產製失敗",
                        "reason": "文件仍殘留 placeholder"
                    }
                )

        for table in check_doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if "{{" in p.text and "}}" in p.text:
                            return JSONResponse(
                                status_code=500,
                                content={
                                    "message": "產製失敗",
                                    "reason": "文件仍殘留 placeholder"
                                }
                            )

        return FileResponse(
            path=filepath,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:

        return JSONResponse(
            status_code=500,
            content={
                "message": "產製失敗",
                "reason": str(e)
            }
        )
