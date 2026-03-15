from fastapi import FastAPI
from fastapi.responses import JSONResponse, FileResponse
from pydantic import BaseModel
from docx import Document
import io
import os

app = FastAPI()

OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)


class ApplicationData(BaseModel):
    case_description: str


@app.get("/")
def root():
    return {"message": "ok"}


@app.post("/generate-word")
async def generate_word(data: ApplicationData):
    try:
        filename = "tax_application.docx"
        filepath = os.path.join(OUTPUT_DIR, filename)

        doc = Document()
        doc.add_heading("納稅者權利保護事項申請書", level=1)
        doc.add_paragraph("案件事實與理由：")
        doc.add_paragraph(data.case_description)
        doc.save(filepath)

        return JSONResponse(
            content={
                "success": True,
                "filename": filename,
                "download_url": f"https://tax-word-system-production.up.railway.app/download/{filename}"
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
    filepath = os.path.join(OUTPUT_DIR, filename)

    if not os.path.exists(filepath):
        return JSONResponse(
            status_code=404,
            content={"success": False, "message": "file not found"}
        )

    return FileResponse(
        path=filepath,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
