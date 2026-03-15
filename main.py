from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from docx import Document
import io

app = FastAPI()


class ApplicationData(BaseModel):
    application_item: str | None = ""
    tax_type: str | None = ""
    case_description: str | None = ""
    applicant_name: str | None = ""
    applicant_id: str | None = ""
    applicant_address: str | None = ""
    contact_phone: str | None = ""
    email: str | None = ""
    reply_method: str | None = ""
    notification_method: str | None = ""
    agent_name: str | None = ""
    agent_phone: str | None = ""
    evidence_list: str | None = ""


def create_word_document(data: ApplicationData):

    doc = Document()

    doc.add_heading("納稅者權利保護事項申請書", level=1)

    doc.add_paragraph(f"申請事項：{data.application_item}")
    doc.add_paragraph(f"稅目別：{data.tax_type}")
    doc.add_paragraph(f"案件事實與理由：{data.case_description}")

    doc.add_paragraph("")

    doc.add_paragraph(f"申請人姓名：{data.applicant_name}")
    doc.add_paragraph(f"身分證字號：{data.applicant_id}")
    doc.add_paragraph(f"地址：{data.applicant_address}")
    doc.add_paragraph(f"聯絡電話：{data.contact_phone}")
    doc.add_paragraph(f"Email：{data.email}")

    doc.add_paragraph("")

    doc.add_paragraph(f"回覆方式：{data.reply_method}")
    doc.add_paragraph(f"通知方式：{data.notification_method}")

    doc.add_paragraph("")

    doc.add_paragraph(f"代理人：{data.agent_name}")
    doc.add_paragraph(f"代理人電話：{data.agent_phone}")

    doc.add_paragraph("")

    doc.add_paragraph(f"證明文件：{data.evidence_list}")

    return doc


@app.post("/generate-word")
async def generate_word(data: ApplicationData):

    doc = create_word_document(data)

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)

    return StreamingResponse(
        file_stream,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": "attachment; filename=tax_application.docx"
        }
    )
