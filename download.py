#!/usr/bin/env python3
"""Baixa a planilha do Google Drive como XLSX para /tmp/planilha.xlsx.

Autentica com Service Account. As credenciais (JSON) podem vir de:
  - var de ambiente GOOGLE_CREDENTIALS_JSON (conteúdo do JSON em string)
  - arquivo apontado por GOOGLE_APPLICATION_CREDENTIALS
  - arquivo ./service-account.json no diretório atual

ID da planilha vem de:
  - var de ambiente SHEET_ID
  - argumento --sheet-id
"""
import argparse
import io
import json
import os
import sys
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]


def load_credentials():
    raw = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if raw:
        info = json.loads(raw)
        return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS") or "./service-account.json"
    if Path(path).exists():
        return service_account.Credentials.from_service_account_file(path, scopes=SCOPES)
    raise SystemExit(
        "Credenciais não encontradas. Defina GOOGLE_CREDENTIALS_JSON (env)\n"
        "ou GOOGLE_APPLICATION_CREDENTIALS apontando para service-account.json\n"
        "ou coloque service-account.json no diretório atual."
    )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--sheet-id", default=os.environ.get("SHEET_ID"),
                        help="ID da planilha no Drive (ou via env SHEET_ID)")
    parser.add_argument("--out", default="/tmp/planilha.xlsx",
                        help="Path do XLSX de saída (padrão /tmp/planilha.xlsx)")
    args = parser.parse_args()

    if not args.sheet_id:
        sys.exit("Faltou --sheet-id ou env SHEET_ID")

    creds = load_credentials()
    service = build("drive", "v3", credentials=creds, cache_discovery=False)
    request = service.files().export_media(fileId=args.sheet_id, mimeType=XLSX_MIME)

    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()

    out_path = Path(args.out)
    out_path.write_bytes(buf.getvalue())
    print(f"Baixado: {out_path} ({out_path.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
