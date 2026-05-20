#!/usr/bin/env python3
"""Baixa a planilha do Google Drive como XLSX para /tmp/planilha.xlsx.

Usa o token OAuth gerenciado pelo rclone (~/.config/rclone/rclone.conf,
remote "gdrive"). Não precisa de Service Account nem GCP project — usa
sua autenticação pessoal já feita via `rclone config`.

Antes do download, força o rclone a renovar o token se necessário.
"""
import json
import os
import subprocess
import sys
from pathlib import Path

SHEET_ID = os.environ.get(
    "SHEET_ID",
    "1YQAYw3wpjRt_GArgqKmbp1CICLeH_7edET-bdbTC75U",
)
OUT = Path("/tmp/planilha.xlsx")
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
RCLONE_CONF = Path.home() / ".config/rclone/rclone.conf"
RCLONE_REMOTE = "gdrive"


def refresh_rclone_token():
    """Força o rclone a renovar o token OAuth se estiver expirado."""
    try:
        subprocess.run(
            ["rclone", "about", f"{RCLONE_REMOTE}:"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
            check=True,
            timeout=30,
        )
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired, FileNotFoundError) as e:
        sys.exit(f"rclone falhou ao renovar token: {e}")


def read_access_token():
    if not RCLONE_CONF.exists():
        sys.exit(f"rclone config não encontrado em {RCLONE_CONF}")
    text = RCLONE_CONF.read_text()
    in_section = False
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith("[") and stripped.endswith("]"):
            in_section = stripped == f"[{RCLONE_REMOTE}]"
            continue
        if in_section and stripped.startswith("token"):
            _, _, raw = stripped.partition("=")
            token = json.loads(raw.strip())
            return token["access_token"]
    sys.exit(f"Token não encontrado para remote '{RCLONE_REMOTE}' em {RCLONE_CONF}")


def download(sheet_id, access_token, out_path):
    url = f"https://www.googleapis.com/drive/v3/files/{sheet_id}/export?mimeType={XLSX_MIME}"
    result = subprocess.run(
        [
            "curl", "-sSfL",
            "-H", f"Authorization: Bearer {access_token}",
            "-o", str(out_path),
            "-w", "%{http_code}",
            url,
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )
    if result.returncode != 0:
        sys.exit(f"curl falhou ({result.returncode}): {result.stderr or result.stdout}")
    code = result.stdout.strip().splitlines()[-1] if result.stdout else "?"
    if code != "200":
        sys.exit(f"HTTP {code} ao baixar planilha. Verifique se a planilha está compartilhada com sua conta.")


def main():
    refresh_rclone_token()
    token = read_access_token()
    download(SHEET_ID, token, OUT)
    size = OUT.stat().st_size
    print(f"Baixado: {OUT} ({size:,} bytes)")


if __name__ == "__main__":
    main()
