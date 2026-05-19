"""
dev_server.py - Serveur de développement local.
Simule Netlify: sert public/index.html + expose POST /.netlify/functions/process
"""
import sys
sys.path.insert(0, "netlify/functions/process")

from fastapi import FastAPI, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response, FileResponse
from processor import process_pdf

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=[
        "X-Stats-Total", "X-Stats-Signed", "X-Stats-Unsigned",
        "X-Stats-Unknowns", "X-Stats-Pointages",
    ],
)


@app.get("/")
async def index():
    return FileResponse("public/index.html")


@app.post("/.netlify/functions/process")
async def process(
    pdf: UploadFile = File(...),
    excel: UploadFile = File(None),
    site: str = Form(None),
):
    pdf_bytes = await pdf.read()
    excel_bytes = await excel.read() if excel else None

    zip_bytes, stats, _unknowns = process_pdf(pdf_bytes, excel_bytes, site or None)

    return Response(
        content=zip_bytes,
        media_type="application/zip",
        headers={
            "Content-Disposition": "attachment; filename=bons_traites.zip",
            "X-Stats-Total":     str(stats["total"]),
            "X-Stats-Signed":    str(stats["signes"]),
            "X-Stats-Unsigned":  str(stats["non_signes"]),
            "X-Stats-Unknowns":  str(stats["unknowns"]),
            "X-Stats-Pointages": str(stats["pointages"] or ""),
        },
    )
