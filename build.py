#!/usr/bin/env python3
"""Inject data.json into template.html and write index.html.

If --password is given, the data is encrypted with AES-GCM using a key derived
from the password via PBKDF2 (SHA-256, 250 000 iterations). The HTML keeps only
the ciphertext + salt + iv, and asks the user for the password before showing
the dashboard. Decryption happens entirely client-side via the Web Crypto API.

Usage:
    python3 build.py                         # plaintext (no protection)
    python3 build.py --password <senha>      # encrypted
"""
import argparse
import base64
import json
import os
import secrets
from hashlib import pbkdf2_hmac
from pathlib import Path

try:
    from Crypto.Cipher import AES  # pycryptodome (already on most macs via brew)
except ImportError:
    AES = None

ROOT = Path(__file__).parent
DATA = ROOT / "data.json"
TEMPLATE = ROOT / "template.html"
OUT = ROOT / "index.html"

PBKDF2_ITER = 250_000
KEY_LEN = 32  # AES-256
SALT_LEN = 16
IV_LEN = 12

ap = argparse.ArgumentParser()
ap.add_argument("--password", help="Encrypt the embedded data with this password")
args = ap.parse_args()

raw = DATA.read_text(encoding="utf-8")
obj = json.loads(raw)
compact = json.dumps(obj, ensure_ascii=False, separators=(",", ":"))

if args.password:
    if AES is None:
        # Fallback: use built-in pyca/cryptography
        try:
            from cryptography.hazmat.primitives.ciphers.aead import AESGCM
            backend = "cryptography"
        except ImportError:
            raise SystemExit(
                "Need 'pycryptodome' or 'cryptography'. Install with:\n"
                "  pip3 install pycryptodome   # or\n"
                "  pip3 install cryptography"
            )
    else:
        backend = "pycryptodome"

    salt = secrets.token_bytes(SALT_LEN)
    iv = secrets.token_bytes(IV_LEN)
    key = pbkdf2_hmac("sha256", args.password.encode("utf-8"), salt, PBKDF2_ITER, KEY_LEN)
    plaintext = compact.encode("utf-8")

    if backend == "pycryptodome":
        cipher = AES.new(key, AES.MODE_GCM, nonce=iv)
        ciphertext, tag = cipher.encrypt_and_digest(plaintext)
        # WebCrypto expects ciphertext || tag (default for AES-GCM)
        blob = ciphertext + tag
    else:
        aesgcm = AESGCM(key)
        blob = aesgcm.encrypt(iv, plaintext, None)  # already returns ct||tag

    payload = {
        "encrypted": True,
        "alg": "AES-256-GCM",
        "kdf": "PBKDF2-SHA256",
        "iter": PBKDF2_ITER,
        "salt": base64.b64encode(salt).decode(),
        "iv": base64.b64encode(iv).decode(),
        "data": base64.b64encode(blob).decode(),
    }
    placeholder_value = json.dumps(payload)
    print(f"Encrypted payload: {len(placeholder_value):,} chars (was {len(compact):,} plaintext)")
else:
    placeholder_value = compact
    print(f"Plaintext payload: {len(compact):,} chars (no password protection)")

html = TEMPLATE.read_text(encoding="utf-8")
html = html.replace("__DATA_PLACEHOLDER__", placeholder_value)

OUT.write_text(html, encoding="utf-8")
print(f"Wrote {OUT} ({len(html):,} chars)")
