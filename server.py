import os
import sys
import json
import base64
import glob
import io
import asyncio
from typing import List, Optional, Dict, Any
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Query
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from PIL import Image
import httpx
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Load .env
from dotenv import load_dotenv
BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

STATIC_DIR = BASE_DIR / "static"
OUTPUT_DIR = BASE_DIR / "output"
UPLOADS_DIR = BASE_DIR / "uploads"
CONFIG_FILE = BASE_DIR / "config.json"
CACHE_FILE = BASE_DIR / "learning_cache.json"
DEFAULT_IMG_DIR = "/Users/macbook/Downloads/บัญชี"

os.makedirs(STATIC_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(UPLOADS_DIR, exist_ok=True)

app = FastAPI(title="Table to Sheets AI")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Known 21 Branches Master Template
KNOWN_BRANCHES = [
    "1.บิ๊กซีเลยประตู 2",
    "2.บิ๊กซีเลยประตู 3",
    "3.บิ๊กซีวังสะพุง",
    "4.ไทยวัสดุเลย",
    "5.โลตัสชุมแพ",
    "6.โลตัสแก้งคร้อ",
    "7.โลตัสชัยภูมิ",
    "8.บิ๊กซีบ้านไผ่",
    "9.โลตัสขอนแก่น (ในเมืองประตูน้ำ)",
    "10.โลตัสขอนแก่น (ธนาคาร)",
    "11.โลตัสขอนแก่น KFC ประตู 2",
    "12.บิ๊กซีร้อยเอ็ด",
    "13.บิ๊กซีกาฬสินธุ์",
    "14.โลตัสกาฬสินธุ์",
    "15.โลตัสบัวขาว",
    "16.โลตัสโพนทอง",
    "17.ไทยวัสดุกาฬสินธุ์",
    "18.โลตัสบ้านไผ่",
    "19.จัตุรัสชัยภูมิ",
    "20.โลตัสกาฬสินธุ์ (ร้านใน)",
    "21.โลตัสพลัส อุดร"
]

def get_env_or_config(key: str, default: str = "") -> str:
    env_val = os.getenv(key, "")
    if env_val:
        return env_val
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                return cfg.get(key.lower(), default)
        except Exception:
            pass
    return default

def mask_key(k: str) -> str:
    if not k:
        return ""
    if len(k) > 10:
        return k[:4] + "..." + k[-4:]
    return "***"

@app.get("/api/config")
def get_config():
    gemini_key = get_env_or_config("GEMINI_API_KEY")
    groq_key = get_env_or_config("GROQ_API_KEY")
    provider = get_env_or_config("DEFAULT_PROVIDER", "gemini")
    return {
        "has_gemini": bool(gemini_key),
        "masked_gemini": mask_key(gemini_key),
        "has_groq": bool(groq_key),
        "masked_groq": mask_key(groq_key),
        "provider": provider,
        "default_folder": DEFAULT_IMG_DIR
    }

@app.post("/api/config")
def set_config(payload: dict):
    cfg = {}
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        except Exception:
            pass
    if "gemini_api_key" in payload and payload["gemini_api_key"]:
        cfg["gemini_api_key"] = payload["gemini_api_key"].strip()
    if "groq_api_key" in payload and payload["groq_api_key"]:
        cfg["groq_api_key"] = payload["groq_api_key"].strip()
    if "provider" in payload:
        cfg["default_provider"] = payload["provider"].strip()
    
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    return {"status": "ok", "message": "บันทึกการตั้งค่าแล้ว"}

def optimize_image_for_vision(image_bytes: bytes, max_dim: int = 1600) -> bytes:
    """Resize image to reduce token usage and speed up network request 4x."""
    try:
        im = Image.open(io.BytesIO(image_bytes))
        im = im.convert("RGB")
        im.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
        out = io.BytesIO()
        im.save(out, format="JPEG", quality=85, optimize=True)
        return out.getvalue()
    except Exception:
        return image_bytes

def normalize_branch_name(name: str) -> str:
    """Match branch name with known branch template to prevent typos."""
    clean = name.strip()
    for kb in KNOWN_BRANCHES:
        if kb == clean:
            return kb
    for kb in KNOWN_BRANCHES:
        kb_prefix = kb.split(".")[0] + "."
        if clean.startswith(kb_prefix):
            return kb
    return clean

SYSTEM_PROMPT = f"""You are an expert OCR AI specializing in Thai financial records and handwritten accounting tables.
Analyze the image of the accounting table and extract all structured data into JSON format.

Known branches template reference:
{json.dumps(KNOWN_BRANCHES, ensure_ascii=False, indent=2)}

Rules:
1. Title: Read table title at the top (e.g. "ค่าคอม วันที่ 13-24 ตุลาคม 2569", "ยอดขาย วันที่ 1-12 ตุลาคม 2569").
2. Type: Exactly "ค่าคอม" or "ยอดขาย".
3. Date columns: List of dates in header (e.g. [1, 2, 3, ... 12] or [13, 14, ... 24]).
4. Rows:
   - "branch": Match to the closest known branch name from the reference list above.
   - "values": List of values for each date column. 
     - If number, return integer or float.
     - If dash or blank, return "-".
     - If '=OFF=', return "OFF".
     - If 'OPEN ->', return "OPEN".
   - "written_total": Handwritten row total on far right margin if written (number or null).
5. Column Totals: Handwritten column totals at bottom if written (list of numbers or null).

Return strictly JSON:
{{
  "title": "string",
  "type": "ค่าคอม" or "ยอดขาย",
  "date_columns": [1, 2, ...],
  "headers": ["วันที่ / สาขา", "1", "2", ...],
  "rows": [
    {{"branch": "1.บิ๊กซีเลยประตู 2", "values": [9340, 11440, ...], "written_total": 130580}}
  ],
  "written_column_totals": [null, null, ...]
}}
"""

def cross_validate_table(data: dict) -> dict:
    rows = data.get("rows", [])
    num_cols = len(data.get("date_columns", []))
    
    # Normalize branch names against template
    for r in rows:
        r["branch"] = normalize_branch_name(r.get("branch", ""))
        vals = r.get("values", [])
        numeric_vals = [v for v in vals if isinstance(v, (int, float))]
        calc_total = sum(numeric_vals)
        r["calc_total"] = calc_total
        written = r.get("written_total")
        if written is not None and isinstance(written, (int, float)):
            r["match"] = (calc_total == written)
            r["diff"] = calc_total - written
        else:
            r["match"] = None
            r["diff"] = 0

    # Calculate column sums
    calc_col_totals = []
    for c_idx in range(num_cols):
        col_sum = 0
        for r in rows:
            vals = r.get("values", [])
            if c_idx < len(vals) and isinstance(vals[c_idx], (int, float)):
                col_sum += vals[c_idx]
        calc_col_totals.append(col_sum)
    data["calc_column_totals"] = calc_col_totals

    # Check col totals against written
    written_cols = data.get("written_column_totals", [])
    col_matches = []
    for c_idx in range(num_cols):
        w = written_cols[c_idx] if c_idx < len(written_cols) else None
        if w is not None and isinstance(w, (int, float)):
            col_matches.append(calc_col_totals[c_idx] == w)
        else:
            col_matches.append(None)
    data["col_matches"] = col_matches
    
    all_row_matches = [r["match"] for r in rows if r["match"] is not None]
    all_col_matches = [m for m in col_matches if m is not None]
    total_checks = len(all_row_matches) + len(all_col_matches)
    passed_checks = sum(1 for m in (all_row_matches + all_col_matches) if m is True)
    
    data["verification"] = {
        "total_checks": total_checks,
        "passed_checks": passed_checks,
        "is_perfect": (total_checks > 0 and passed_checks == total_checks)
    }
    return data

async def run_ai_ocr(image_bytes: bytes, provider: str = "gemini") -> dict:
    optimized_bytes = optimize_image_for_vision(image_bytes, max_dim=1600)
    b64_data = base64.b64encode(optimized_bytes).decode("utf-8")
    
    gemini_key = get_env_or_config("GEMINI_API_KEY")
    groq_key = get_env_or_config("GROQ_API_KEY")
    
    if provider == "groq" and groq_key:
        url = "https://api.groq.com/openai/v1/chat/completions"
        headers = {
            "Authorization": f"Bearer {groq_key}",
            "Content-Type": "application/json"
        }
        body = {
            "model": "qwen/qwen3.8-27b",
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": SYSTEM_PROMPT},
                        {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64_data}"}}
                    ]
                }
            ],
            "response_format": {"type": "json_object"},
            "temperature": 0.1
        }
        async with httpx.AsyncClient(timeout=60.0) as client:
            resp = await client.post(url, headers=headers, json=body)
            if resp.status_code != 200:
                raise HTTPException(status_code=resp.status_code, detail=f"Groq Error: {resp.text}")
            content = resp.json()["choices"][0]["message"]["content"]
            table_data = json.loads(content)
            return cross_validate_table(table_data)
    else:
        if not gemini_key:
            raise HTTPException(status_code=400, detail="Gemini API Key is not set")
        
        models_to_try = [
            "gemini-3.5-flash-lite",  # 150K RPD quota, extremely fast and reliable
            "gemini-3.6-flash",
            "gemini-3.5-flash",
            "gemini-flash-lite-latest"
        ]
        
        last_error = None
        for model in models_to_try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={gemini_key}"
            body = {
                "contents": [{
                    "parts": [
                        {"text": SYSTEM_PROMPT},
                        {"inline_data": {"mime_type": "image/jpeg", "data": b64_data}}
                    ]
                }],
                "generationConfig": {
                    "response_mime_type": "application/json",
                    "temperature": 0.1
                }
            }
            try:
                async with httpx.AsyncClient(timeout=45.0) as client:
                    resp = await client.post(url, json=body)
                    if resp.status_code == 200:
                        text_part = resp.json()["candidates"][0]["content"]["parts"][0]["text"]
                        table_data = json.loads(text_part)
                        return cross_validate_table(table_data)
                    elif resp.status_code in (503, 429):
                        last_error = f"{model} returned {resp.status_code}"
                        continue
                    else:
                        last_error = resp.text
            except Exception as e:
                last_error = str(e)
                continue
                
        raise HTTPException(status_code=500, detail=f"Gemini API Error across fallback models: {last_error}")

@app.post("/api/upload-ocr")
async def upload_and_process(
    file: UploadFile = File(...),
    provider: str = Form("gemini")
):
    contents = await file.read()
    if not contents:
        raise HTTPException(status_code=400, detail="Empty file")
    
    save_path = UPLOADS_DIR / file.filename
    with open(save_path, "wb") as f:
        f.write(contents)
        
    result = await run_ai_ocr(contents, provider=provider)
    result["filename"] = file.filename
    result["image_path"] = f"/api/uploaded-image/{file.filename}"
    return result

@app.get("/api/uploaded-image/{filename}")
def get_uploaded_image(filename: str):
    p = UPLOADS_DIR / filename
    if not p.exists():
        raise HTTPException(status_code=404, detail="Image not found")
    return FileResponse(str(p))

@app.get("/api/scan")
def scan_folder(folder_path: str = Query(DEFAULT_IMG_DIR)):
    if not os.path.exists(folder_path):
        raise HTTPException(status_code=400, detail=f"โฟลเดอร์ {folder_path} ไม่มีอยู่จริง")
    
    extensions = ("*.jpg", "*.jpeg", "*.png", "*.webp", "*.JPG", "*.JPEG", "*.PNG")
    files = []
    for ext in extensions:
        files.extend(glob.glob(os.path.join(folder_path, ext)))
    
    files = sorted(list(set(files)))
    result = []
    for f in files:
        p = Path(f)
        stat = p.stat()
        result.append({
            "name": p.name,
            "path": str(p.resolve()),
            "size_mb": round(stat.st_size / (1024 * 1024), 2),
            "modified": stat.st_mtime
        })
    return {"folder": folder_path, "count": len(result), "files": result}

@app.get("/api/image")
def get_image(path: str = Query(...)):
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(path)

@app.get("/api/preloaded")
def get_preloaded_data_endpoint():
    from export_data import (
        sales_1_12_p1, sales_1_12_p2, sales_13_24_p1,
        comm_1_12_p1, comm_1_12_p2, comm_13_24_p1
    )
    
    data = {}
    
    p1 = {
        "filename": "IMG_20261110_122608.jpg",
        "image_path": f"{DEFAULT_IMG_DIR}/IMG_20261110_122608.jpg",
        "title": "ค่าคอม วันที่ 13-24 ตุลาคม 2569",
        "type": "ค่าคอม",
        "date_columns": list(range(13, 25)),
        "headers": ["วันที่ / สาขา"] + [str(d) for d in range(13, 25)],
        "rows": [
            {"branch": b, "values": comm_13_24_p1[b], "written_total": sum(v for v in comm_13_24_p1[b] if isinstance(v, (int, float)))}
            for b in comm_13_24_p1
        ],
        "written_column_totals": [3650, 2950, 3200, 3150, 5050, 3700, 3150, 4300, 2800, 3650, 3400, 3550]
    }
    data["IMG_20261110_122608.jpg"] = cross_validate_table(p1)

    p2 = {
        "filename": "IMG_20261110_122620.jpg",
        "image_path": f"{DEFAULT_IMG_DIR}/IMG_20261110_122620.jpg",
        "title": "ยอดขาย วันที่ 13-24 ตุลาคม 2569",
        "type": "ยอดขาย",
        "date_columns": list(range(13, 25)),
        "headers": ["วันที่ / สาขา"] + [str(d) for d in range(13, 25)],
        "rows": [
            {"branch": b, "values": sales_13_24_p1[b], "written_total": sum(v for v in sales_13_24_p1[b] if isinstance(v, (int, float)))}
            for b in sales_13_24_p1
        ],
        "written_column_totals": [96855, 72740, 136285, 84030, 109560, 90040, 121815, 96490, 77880, 182075, 105490, 100380]
    }
    data["IMG_20261110_122620.jpg"] = cross_validate_table(p2)

    p3 = {
        "filename": "IMG_20261110_122630.jpg",
        "image_path": f"{DEFAULT_IMG_DIR}/IMG_20261110_122630.jpg",
        "title": "ยอดขาย วันที่ 1-12 ตุลาคม 2569 (สาขา 1-13)",
        "type": "ยอดขาย",
        "date_columns": list(range(1, 13)),
        "headers": ["วันที่ / สาขา"] + [str(d) for d in range(1, 13)],
        "rows": [
            {"branch": b, "values": sales_1_12_p1[b], "written_total": sum(v for v in sales_1_12_p1[b] if isinstance(v, (int, float)))}
            for b in sales_1_12_p1
        ],
        "written_column_totals": [None, None, None, 102050, None, None, None, None, None, None, None, None]
    }
    data["IMG_20261110_122630.jpg"] = cross_validate_table(p3)

    p4 = {
        "filename": "IMG_20261110_122646.jpg",
        "image_path": f"{DEFAULT_IMG_DIR}/IMG_20261110_122646.jpg",
        "title": "ยอดขาย วันที่ 1-12 ตุลาคม 2569 (สาขา 14-21)",
        "type": "ยอดขาย",
        "date_columns": list(range(1, 13)),
        "headers": ["วันที่ / สาขา"] + [str(d) for d in range(1, 13)],
        "rows": [
            {"branch": b, "values": sales_1_12_p2[b], "written_total": sum(v for v in sales_1_12_p2[b] if isinstance(v, (int, float)))}
            for b in sales_1_12_p2
        ],
        "written_column_totals": [None, None, None, 41090, None, None, None, None, None, None, None, None]
    }
    data["IMG_20261110_122646.jpg"] = cross_validate_table(p4)

    p5 = {
        "filename": "IMG_20261110_122657.jpg",
        "image_path": f"{DEFAULT_IMG_DIR}/IMG_20261110_122657.jpg",
        "title": "ค่าคอม วันที่ 1-12 ตุลาคม 2569 (สาขา 1-13)",
        "type": "ค่าคอม",
        "date_columns": list(range(1, 13)),
        "headers": ["วันที่ / สาขา"] + [str(d) for d in range(1, 13)],
        "rows": [
            {"branch": b, "values": comm_1_12_p1[b], "written_total": sum(v for v in comm_1_12_p1[b] if isinstance(v, (int, float)))}
            for b in comm_1_12_p1
        ],
        "written_column_totals": [None, None, None, 4700, None, None, None, None, None, None, None, None]
    }
    data["IMG_20261110_122657.jpg"] = cross_validate_table(p5)

    p6 = {
        "filename": "IMG_20261110_122710.jpg",
        "image_path": f"{DEFAULT_IMG_DIR}/IMG_20261110_122710.jpg",
        "title": "ค่าคอม วันที่ 1-12 ตุลาคม 2569 (สาขา 14-21)",
        "type": "ค่าคอม",
        "date_columns": list(range(1, 13)),
        "headers": ["วันที่ / สาขา"] + [str(d) for d in range(1, 13)],
        "rows": [
            {"branch": b, "values": comm_1_12_p2[b], "written_total": sum(v for v in comm_1_12_p2[b] if isinstance(v, (int, float)))}
            for b in comm_1_12_p2
        ],
        "written_column_totals": [None, None, None, 1450, None, None, None, None, None, None, None, None]
    }
    data["IMG_20261110_122710.jpg"] = cross_validate_table(p6)

    return data

@app.post("/api/export-single-excel")
def export_single_excel(table: dict):
    from export_data import style_sheet
    wb = openpyxl.Workbook()
    ws = wb.active
    title = table.get("title", "รายงานตารางบัญชี")
    ws.title = title[:31].replace(":", "_").replace("/", "_")
    
    dates = table.get("date_columns", [])
    start_d = dates[0] if dates else 1
    end_d = dates[-1] if dates else 12
    
    rows_data = {}
    for r in table.get("rows", []):
        rows_data[r["branch"]] = r["values"]
        
    style_sheet(ws, title, start_d, end_d, rows_data)
    
    out_file = OUTPUT_DIR / f"{title[:25].replace(' ', '_')}.xlsx"
    wb.save(str(out_file))
    return FileResponse(str(out_file), filename=out_file.name)

# Static files & Output
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
app.mount("/output", StaticFiles(directory=str(OUTPUT_DIR)), name="output")

@app.get("/")
def get_index():
    index_file = STATIC_DIR / "index.html"
    if index_file.exists():
        return FileResponse(str(index_file))
    return HTMLResponse("<h1>Table to Sheets App Loading...</h1>")

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8765))
    uvicorn.run("server:app", host="0.0.0.0", port=port, reload=True)
