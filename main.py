from fastapi import FastAPI, Query
from fastapi.responses import JSONResponse, FileResponse, HTMLResponse
from pydantic import BaseModel
from docx import Document
import re
import uuid
import time
import asyncio
import os

app = FastAPI()

TEMPLATE_PATH = "templates/納保申請書_官方格式_可套填_v8.docx"

OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

download_store = {}


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

    for key, value in replacements.items():
        if key in text:
            text = text.replace(key, value)

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

        if has_placeholder(doc):

            return JSONResponse(
                status_code=500,
                content={
                    "success": False,
                    "message": "產製失敗",
                    "reason": "Word 殘留 placeholder"
                }
            )

        filename = f"申請書_{data.apply_year}{int(data.apply_month):02d}{int(data.apply_day):02d}.docx"

        unique_name = f"{uuid.uuid4()}_{filename}"

        file_path = os.path.join(OUTPUT_DIR, unique_name)

        doc.save(file_path)

        token = str(uuid.uuid4())

        download_store[token] = {
            "path": file_path,
            "filename": filename,
            "expire_time": time.time() + 600,
            "used": False
        }

        return JSONResponse(
            content={
                "success": True,
                "filename": filename,
                "download_url": f"https://tax-word-system-production.up.railway.app/download/{token}"
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


@app.get("/form", response_class=HTMLResponse)
def form_page(
    apply_year: str = Query(default=""),
    apply_month: str = Query(default=""),
    apply_day: str = Query(default=""),
    case_category_suggestion: str = Query(default=""),
    tax_items_suggestion: str = Query(default=""),
    apply_method_suggestion: str = Query(default=""),
    reply_method_suggestion: str = Query(default=""),
    notify_method_suggestion: str = Query(default=""),
):
    def esc(s):
        return s.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;").replace(">", "&gt;")

    html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>納保申請書產製</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Noto Sans TC',sans-serif;background:#f0f4f8;padding:20px;min-height:100vh}}
.container{{max-width:680px;margin:0 auto;background:#fff;border-radius:12px;padding:2rem;box-shadow:0 2px 12px rgba(0,0,0,.1)}}
h1{{font-size:1.3rem;color:#2c3e50;border-bottom:2px solid #3498db;padding-bottom:.6rem;margin-bottom:1.2rem}}
.field{{margin-bottom:1rem}}
label{{display:block;font-weight:600;margin-bottom:.3rem;color:#444;font-size:.9rem}}
input,textarea{{width:100%;padding:.55rem .75rem;border:1px solid #ccc;border-radius:6px;font-size:.95rem;font-family:inherit}}
textarea{{height:200px;resize:vertical;line-height:1.6}}
.row{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:.8rem}}
.hint.warn{{color:#e67e22;font-weight:600;font-size:.8rem;margin-top:.25rem}}
button{{background:#2980b9;color:#fff;border:none;padding:.85rem;border-radius:8px;font-size:1rem;cursor:pointer;width:100%;margin-top:.5rem;font-weight:600}}
button:disabled{{background:#aaa;cursor:not-allowed}}
#spinner{{display:none;text-align:center;padding:.8rem;color:#555}}
.error{{color:#c0392b;background:#fdf0f0;border:1px solid #f5c6cb;padding:.75rem;border-radius:6px;margin-top:.8rem}}
.success{{color:#1e8449;background:#eafaf1;border:1px solid #a9dfbf;padding:.75rem;border-radius:6px;margin-top:.8rem}}
.divider{{border:none;border-top:1px solid #eee;margin:1.2rem 0}}
</style>
</head>
<body>
<div class="container">
  <h1>📄 納保申請書產製</h1>
  <p style="color:#666;font-size:.88rem;margin-bottom:1.2rem">由納保申請助理轉介。請將助理產出的「申請事由」貼入下方欄位，其他欄位已自動填入，確認後點「產製申請書」即可下載。</p>
  <div class="row">
    <div class="field"><label>申請年份（民國）</label><input id="apply_year" value="{esc(apply_year)}" placeholder="114"></div>
    <div class="field"><label>月份</label><input id="apply_month" value="{esc(apply_month)}" placeholder="3"></div>
    <div class="field"><label>日期</label><input id="apply_day" value="{esc(apply_day)}" placeholder="15"></div>
  </div>
  <div class="field"><label>案件類型建議</label><input id="case_category_suggestion" value="{esc(case_category_suggestion)}"></div>
  <div class="field"><label>稅目建議</label><input id="tax_items_suggestion" value="{esc(tax_items_suggestion)}"></div>
  <hr class="divider">
  <div class="field">
    <label>✏️ 申請事由（請從助理對話中複製並貼入）</label>
    <textarea id="formal_statement" placeholder="請將 ChatGPT 助理產出的「申請事由」完整貼入此處…"></textarea>
    <div class="hint warn">※ 此欄需手動貼上，其餘欄位已自動填入</div>
  </div>
  <hr class="divider">
  <div class="field"><label>申請方式</label><input id="apply_method_suggestion" value="{esc(apply_method_suggestion)}"></div>
  <div class="field"><label>回覆方式</label><input id="reply_method_suggestion" value="{esc(reply_method_suggestion)}"></div>
  <div class="field"><label>通知方式</label><input id="notify_method_suggestion" value="{esc(notify_method_suggestion)}"></div>
  <button id="btn" onclick="generate()">產製申請書</button>
  <div id="spinner">⏳ 正在產製，請稍候…</div>
  <div id="result"></div>
</div>
<script>
async function generate() {{
  const btn = document.getElementById('btn');
  const spinner = document.getElementById('spinner');
  const result = document.getElementById('result');
  const formal = document.getElementById('formal_statement').value.trim();
  if (!formal) {{ result.innerHTML = '<div class="error">⚠️ 請先貼入申請事由！</div>'; return; }}
  btn.disabled = true; spinner.style.display = 'block'; result.innerHTML = '';
  const payload = {{
    apply_year: document.getElementById('apply_year').value.trim(),
    apply_month: document.getElementById('apply_month').value.trim(),
    apply_day: document.getElementById('apply_day').value.trim(),
    formal_statement: formal,
    case_category_suggestion: document.getElementById('case_category_suggestion').value.trim(),
    tax_items_suggestion: document.getElementById('tax_items_suggestion').value.trim(),
    apply_method_suggestion: document.getElementById('apply_method_suggestion').value.trim(),
    reply_method_suggestion: document.getElementById('reply_method_suggestion').value.trim(),
    notify_method_suggestion: document.getElementById('notify_method_suggestion').value.trim(),
    notify_email: '', evidence_list: ''
  }};
  try {{
    const res = await fetch('/generate-word', {{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify(payload)}});
    const json = await res.json();
    if (json.success && json.download_url) {{
      result.innerHTML = '<div class="success">✅ 申請書已產製！若下載未自動開始，<a href="' + json.download_url + '" target="_blank">點此下載</a></div>';
      setTimeout(() => {{ window.location.href = json.download_url; }}, 500);
    }} else {{
      result.innerHTML = '<div class="error">❌ 產製失敗：' + (json.reason || json.message || '未知錯誤') + '</div>';
    }}
  }} catch(e) {{ result.innerHTML = '<div class="error">❌ 網路錯誤：' + e.message + '</div>'; }}
  finally {{ btn.disabled = false; spinner.style.display = 'none'; }}
}}
</script>
</body>
</html>"""
    return HTMLResponse(content=html)
@app.get("/download/{token}")
def download_file(token: str):

    if token not in download_store:
        return JSONResponse(status_code=404, content={"message": "file not found"})

    record = download_store[token]

    if time.time() > record["expire_time"]:

        if os.path.exists(record["path"]):
            os.remove(record["path"])

        del download_store[token]

        return JSONResponse(status_code=410, content={"message": "link expired"})

    if record["used"]:
        return JSONResponse(status_code=410, content={"message": "file already downloaded"})

    record["used"] = True

    return FileResponse(
        record["path"],
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=record["filename"]
    )


async def cleanup_expired():

    while True:

        now = time.time()

        expired = []

        for token, record in download_store.items():
            if now > record["expire_time"]:
                expired.append(token)

        for token in expired:

            record = download_store[token]

            if os.path.exists(record["path"]):
                os.remove(record["path"])

            del download_store[token]

        await asyncio.sleep(60)


@app.on_event("startup")
async def startup_event():

    asyncio.create_task(cleanup_expired())
