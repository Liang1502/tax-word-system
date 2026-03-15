from fastapi import FastAPI
from fastapi.responses import JSONResponse, FileResponse
from pydantic import BaseModel
from docx import Document
import os

app = FastAPI()

TEMPLATE = "納保申請書_官方格式_可套填_v8.docx"
OUTPUT = "output.docx"


class ApplicationData(BaseModel):
    case_description: str


@app.post("/generate-word")
async def generate_word(data: ApplicationData):

    try:

        doc = Document(TEMPLATE)

        for p in doc.paragraphs:
            if "{{formal_statement}}" in p.text:
                p.text = p.text.replace("{{formal_statement}}", data.case_description)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if "{{formal_statement}}" in p.text:
                            p.text = p.text.replace("{{formal_statement}}", data.case_description)

        doc.save(OUTPUT)

        return JSONResponse(
            {
                "success": True,
                "filename": OUTPUT,
                "download_url": "https://tax-word-system-production.up.railway.app/download/output.docx"
            }
        )

    except Exception as e:

        return JSONResponse(
            {
                "success": False,
                "message": "產製失敗",
                "reason": str(e)
            }
        )


@app.get("/download/{filename}")
def download_file(filename: str):

    return FileResponse(
        filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
