"""
ZION Contract Review — AgentBase Web Server
--------------------------------------------
FastAPI server wrapping agent.py cho AgentBase Runtime.

Endpoints:
  GET  /health        → 200 OK (platform health check)
  POST /invocations   → nhận 2 file hợp đồng, trả về báo cáo .docx

Request (multipart/form-data):
  contract1: file (.pdf hoặc .docx) — HĐ Mẫu gốc CTV-1
  contract2: file (.pdf hoặc .docx) — HĐ Thực tế cần check

Response:
  Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document
  Content-Disposition: attachment; filename="<name>_REVIEWED_<date>_v1.docx"
"""

import os
import tempfile
import logging
from datetime import datetime
from pathlib import Path

import uvicorn
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse, JSONResponse

# Import business logic từ agent.py
from agent import extract_text, call_gemini, build_prompt, build_docx_report, CONFIG

# ─────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# App
# ─────────────────────────────────────────────

app = FastAPI(
    title="ZION Contract Review Agent",
    description="Rà soát và so sánh 2 hợp đồng, xuất báo cáo .docx tiếng Việt",
    version="1.0.0",
)

SUPPORTED_EXTENSIONS = {".docx", ".pdf"}


def validate_extension(filename: str) -> str:
    ext = Path(filename).suffix.lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise HTTPException(
            status_code=400,
            detail=f"Định dạng '{ext}' không được hỗ trợ. Chỉ chấp nhận .docx và .pdf.",
        )
    return ext


# ─────────────────────────────────────────────
# Routes
# ─────────────────────────────────────────────

@app.get("/health")
async def health():
    """Platform health check — phải trả về 200 để runtime ACTIVE."""
    return {"status": "ok", "agent": CONFIG["agent"]["name"]}


@app.post("/invocations")
async def invocations(
    contract1: UploadFile = File(..., description="HĐ Mẫu gốc CTV-1 (.docx hoặc .pdf)"),
    contract2: UploadFile = File(..., description="HĐ Thực tế cần check (.docx hoặc .pdf)"),
):
    """
    Nhận 2 file hợp đồng, chạy ZION review pipeline, trả về báo cáo .docx.
    """
    logger.info(f"Invocation: contract1={contract1.filename!r}, contract2={contract2.filename!r}")

    validate_extension(contract1.filename)
    validate_extension(contract2.filename)

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)

        # Lưu file upload tạm thời
        path1 = tmp / f"contract1{Path(contract1.filename).suffix.lower()}"
        path2 = tmp / f"contract2{Path(contract2.filename).suffix.lower()}"

        path1.write_bytes(await contract1.read())
        path2.write_bytes(await contract2.read())

        logger.info("Đọc nội dung hợp đồng...")
        try:
            text1 = extract_text(path1)
            text2 = extract_text(path2)
        except Exception as e:
            logger.error(f"Lỗi đọc file: {e}")
            raise HTTPException(status_code=422, detail=f"Không thể đọc file hợp đồng: {e}")

        logger.info(f"Hợp đồng 1: {len(text1):,} ký tự | Hợp đồng 2: {len(text2):,} ký tự")

        # Gọi Gemini
        logger.info("Gọi Gemini API...")
        try:
            sections = CONFIG["review"]["sections"]
            prompt = build_prompt(text1, text2, sections)
            review_text = call_gemini(prompt)
        except EnvironmentError as e:
            raise HTTPException(status_code=500, detail=str(e))
        except Exception as e:
            logger.error(f"Lỗi Gemini: {e}")
            raise HTTPException(status_code=502, detail=f"Lỗi khi gọi Gemini API: {e}")

        logger.info(f"Nhận được {len(review_text):,} ký tự phân tích")

        # Tạo báo cáo .docx
        date_str = datetime.now().strftime("%Y-%m-%d")
        base_name = Path(contract2.filename).stem
        output_filename = f"{base_name}_REVIEWED_{date_str}_v1.docx"
        output_path = tmp / output_filename

        try:
            build_docx_report(review_text, contract1.filename, contract2.filename, output_path, sections)
        except Exception as e:
            logger.error(f"Lỗi tạo docx: {e}")
            raise HTTPException(status_code=500, detail=f"Lỗi khi tạo báo cáo: {e}")

        logger.info(f"Trả về file: {output_filename}")

        # Trả file — copy ra ngoài tmpdir trước khi context manager đóng
        final_path = Path(tempfile.gettempdir()) / output_filename
        final_path.write_bytes(output_path.read_bytes())

    return FileResponse(
        path=str(final_path),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=output_filename,
        background=None,
    )


@app.exception_handler(Exception)
async def global_exception_handler(request, exc):
    logger.error(f"Unhandled exception: {exc}", exc_info=True)
    return JSONResponse(status_code=500, content={"detail": "Internal server error"})


# ─────────────────────────────────────────────
# Entrypoint
# ─────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.getenv("PORT", "8080"))
    logger.info(f"Starting ZION Contract Review Agent on port {port}")
    uvicorn.run(app, host="0.0.0.0", port=port, log_level="info")
