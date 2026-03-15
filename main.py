from fastapi import FastAPI
from fastapi.responses import Response
from pydantic import BaseModel
from docx import Document
import io

app = FastAPI()


class ApplicationData(BaseModel):
    case_description: str
    applicant_name: str


@app.post("/generate-word")
async def generate_word(data: ApplicationData):

    doc = Document()

    doc.add_heading("納稅者權利保護事項申請書", level=1)
    doc.add_paragraph(f"申請人：{data.applicant_name}")
    doc.add_paragraph("案件事實與理由：")
    doc.add_paragraph(data.case_description)

    buffer = io.BytesIO()
    doc.save(buffer)

    return Response(
        content=buffer.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": "attachment; filename=tax_application.docx"
        }
    )
