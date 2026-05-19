"""
process.py - Netlify Function handler for traiter_bons.
Endpoint: POST /.netlify/functions/process
Body: multipart/form-data with fields: pdf (file), excel (file, optional), site (string, optional)
"""
import base64
import json
from email.parser import BytesParser
from email import policy as email_policy

from processor import process_pdf


def _parse_multipart(body_bytes: bytes, content_type: str) -> tuple:
    """Parse multipart/form-data. Returns (files: dict, fields: dict).

    Uses compat32 policy for reliable binary payload handling.
    """
    raw = f"Content-Type: {content_type}\r\n\r\n".encode() + body_bytes
    msg = BytesParser(policy=email_policy.compat32).parsebytes(raw)

    files = {}
    fields = {}

    for part in msg.get_payload():
        if not hasattr(part, "get_payload"):
            continue

        disp = part.get("content-disposition", "")
        params = {}
        for item in disp.split(";"):
            item = item.strip()
            if "=" in item:
                k, v = item.split("=", 1)
                params[k.strip()] = v.strip().strip('"')

        name = params.get("name")
        if name is None:
            continue

        filename = params.get("filename")
        payload = part.get_payload(decode=True)

        if filename:
            files[name] = (filename, payload or b"")
        else:
            fields[name] = (payload or b"").decode("utf-8").strip()

    return files, fields


def _cors_headers() -> dict:
    return {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Headers": "Content-Type",
        "Access-Control-Allow-Methods": "POST, OPTIONS",
        "Access-Control-Expose-Headers": (
            "X-Stats-Total, X-Stats-Signed, X-Stats-Unsigned, "
            "X-Stats-Unknowns, X-Stats-Pointages"
        ),
    }


def handler(event, context):
    http_method = event.get("httpMethod", "")

    if http_method == "OPTIONS":
        return {"statusCode": 204, "headers": _cors_headers(), "body": ""}

    if http_method != "POST":
        return {
            "statusCode": 405,
            "headers": {**_cors_headers(), "Content-Type": "application/json"},
            "body": json.dumps({"detail": "Méthode non autorisée."}),
        }

    try:
        body = event.get("body") or ""
        if event.get("isBase64Encoded"):
            body_bytes = base64.b64decode(body)
        else:
            body_bytes = body.encode("latin-1")

        content_type = ""
        for k, v in (event.get("headers") or {}).items():
            if k.lower() == "content-type":
                content_type = v
                break

        if "multipart/form-data" not in content_type:
            return {
                "statusCode": 400,
                "headers": {**_cors_headers(), "Content-Type": "application/json"},
                "body": json.dumps({"detail": "Requête multipart/form-data attendue."}),
            }

        files, fields = _parse_multipart(body_bytes, content_type)

        if "pdf" not in files:
            return {
                "statusCode": 400,
                "headers": {**_cors_headers(), "Content-Type": "application/json"},
                "body": json.dumps({"detail": "Champ 'pdf' manquant dans le formulaire."}),
            }

        pdf_filename, pdf_bytes = files["pdf"]
        if not pdf_filename.lower().endswith(".pdf"):
            return {
                "statusCode": 400,
                "headers": {**_cors_headers(), "Content-Type": "application/json"},
                "body": json.dumps({"detail": "Le fichier doit être un PDF (.pdf)."}),
            }

        excel_bytes = files["excel"][1] if "excel" in files else None
        site_key = fields.get("site") or None

        zip_bytes, stats, _unknowns = process_pdf(pdf_bytes, excel_bytes, site_key)

        response_headers = {
            **_cors_headers(),
            "Content-Type": "application/zip",
            "Content-Disposition": "attachment; filename=bons_traites.zip",
            "X-Stats-Total": str(stats["total"]),
            "X-Stats-Signed": str(stats["signes"]),
            "X-Stats-Unsigned": str(stats["non_signes"]),
            "X-Stats-Unknowns": str(stats["unknowns"]),
            "X-Stats-Pointages": str(stats["pointages"] or ""),
        }

        return {
            "statusCode": 200,
            "headers": response_headers,
            "body": base64.b64encode(zip_bytes).decode("ascii"),
            "isBase64Encoded": True,
        }

    except Exception as exc:
        return {
            "statusCode": 500,
            "headers": {**_cors_headers(), "Content-Type": "application/json"},
            "body": json.dumps({"detail": f"Erreur de traitement : {exc}"}),
        }
