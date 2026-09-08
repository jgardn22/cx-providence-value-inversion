#!/usr/bin/env python3
"""Encrypt an HTML resource into a passcode-gated page using gate_template.html.

Uses pycryptodome (AES-GCM) + hashlib (PBKDF2-HMAC-SHA256, 250k iterations) to
match the browser-side SubtleCrypto contract in gate_template.html.

Usage:
  python encrypt_resource.py --in raw/decision-canvas.html \
      --out Centene-CRM-Decision-Canvas.html \
      --code CENT0622 \
      --title "Centene CRM Decision Canvas"
"""
import argparse, base64, hashlib, os, pathlib, sys
from Crypto.Cipher import AES

GATE_TEMPLATE = pathlib.Path(__file__).parent / "gate_template.html"
ITER = 250_000

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in",  dest="src",  required=True, help="raw HTML body to encrypt")
    ap.add_argument("--out", dest="dest", required=True, help="output gated HTML file")
    ap.add_argument("--code", required=True, help="access code (e.g. CENT0622)")
    ap.add_argument("--title", required=True, help="resource title shown on gate")
    args = ap.parse_args()

    plaintext = pathlib.Path(args.src).read_bytes()
    salt = os.urandom(16)
    iv   = os.urandom(12)
    key  = hashlib.pbkdf2_hmac("sha256", args.code.encode("utf-8"), salt, ITER, dklen=32)
    cipher = AES.new(key, AES.MODE_GCM, nonce=iv)
    ct, tag = cipher.encrypt_and_digest(plaintext)
    # WebCrypto AES-GCM appends the auth tag to the ciphertext; match that.
    ct_with_tag = ct + tag

    tpl = GATE_TEMPLATE.read_text(encoding="utf-8")
    out = (tpl
        .replace("{{RESOURCE_TITLE}}", args.title)
        .replace("{{CIPHERTEXT_BASE64}}", base64.b64encode(ct_with_tag).decode())
        .replace("{{SALT_BASE64}}",       base64.b64encode(salt).decode())
        .replace("{{IV_BASE64}}",         base64.b64encode(iv).decode()))
    pathlib.Path(args.dest).write_text(out, encoding="utf-8")
    print(f"Wrote {args.dest}  ({len(ct_with_tag)} bytes ciphertext)")

if __name__ == "__main__":
    main()
