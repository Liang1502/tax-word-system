from fastapi import FastAPI
from fastapi.responses import JSONResponse, FileResponse
from pydantic import BaseModel
from docx import Document
import os

app = FastAPI()

# Word 模板位置
TEMPLATE_PATH = "templates/納保申請書_官方格式_可套填_v8.docx"

# 輸出資料夾
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)


class GenerateRequest(BaseModel):
    case_description: str


@app.get("/")
def root():
    return {"status": "running"}


@app.post("/generate-word")
def generate_word(data: GenerateRequest):

    try:

        # 讀取模板
        doc = Document(TEMPLATE_PATH)

        # 替換 placeholder
        for p in doc.paragraphs:
            if "{{formal_statement}}" in p.text:
                for run in p.runs:
                    run.text = run.text.replace(
                        "{{formal_statement}}",
                        data.case_description
                    )

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if "{{formal_statement}}" in p.text:
                            for run in p.runs:
                                run.text = run.text.replace(
                                    "{{formal_statement}}",
                                    data.case_description
                                )

        filename = "tax_application.docx"
        output_path = os.path.join(OUTPUT_DIR, filename)

        doc.save(output_path)

        return JSONResponse(
            {
                "success": True,
                "filename": filename,
                "download_url": f"https://tax-word-system-production.up.railway.app/download/{filename}"
            }
        )

    except Exception as e:

        return JSONResponse(
            {
                "success": False,
                "error": str(e)
            }
        )


@app.get("/download/{filename}")
def download_file(filename: str):

    path = os.path.join(OUTPUT_DIR, filename)

    if not os.path.exists(path):
        return JSONResponse(
            {
                "success": False,
                "error": "file not found"
            }
        )

    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=filename
    )
