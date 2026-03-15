from fastapi import FastAPI
from fastapi.responses import FileResponse
from pydantic import BaseModel
from docxtpl import DocxTemplate
import uuid
import os

app = FastAPI()

TEMPLATE_PATH = "templates/納保申請書_官方格式_可套填_v8.docx"
OUTPUT_DIR = "output"

os.makedirs(OUTPUT_DIR, exist_ok=True)


class RequestData(BaseModel):
    applicant_name: str
    applicant_address: str
    contact_phone: str
    email: str
    case_description: str


@app.post("/generate-word")
def generate_word(data: RequestData):

    doc = DocxTemplate(TEMPLATE_PATH)

    context = {
        "applicant_name": data.applicant_name,
        "applicant_address": data.applicant_address,
        "contact_phone": data.contact_phone,
        "email": data.email,
        "case_description": data.case_description
    }

    doc.render(context)

    filename = f"{uuid.uuid4()}.docx"
    filepath = os.path.join(OUTPUT_DIR, filename)

    doc.save(filepath)

    return FileResponse(
        path=filepath,
        filename="納保申請書.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
