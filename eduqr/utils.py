import io
import re
from typing import Optional
from .models import ClassEntry


def parse_classes(text: str, suffix: str = "") -> list[ClassEntry]:
    entries = []
    lines = [l.strip() for l in text.strip().splitlines() if l.strip()]
    i = 0
    while i < len(lines):
        line = lines[i]
        if re.match(r"https?://", line):
            i += 1
            continue
        if i + 1 < len(lines) and re.match(r"https?://", lines[i + 1]):
            entries.append(ClassEntry(code=line, link=lines[i + 1], suffix=suffix))
            i += 2
        else:
            i += 1
    return entries


def generate_qr_bytes(url: str, logo_bytes: Optional[bytes] = None) -> bytes:
    import qrcode
    from PIL import Image

    error_correction = (
        qrcode.constants.ERROR_CORRECT_H if logo_bytes
        else qrcode.constants.ERROR_CORRECT_M
    )
    qr = qrcode.QRCode(
        version=None,
        error_correction=error_correction,
        box_size=10,
        border=4,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white").convert("RGBA")

    if logo_bytes:
        try:
            logo = Image.open(io.BytesIO(logo_bytes)).convert("RGBA")
            qr_w, qr_h = img.size
            max_logo = int(min(qr_w, qr_h) * 0.25)
            ratio = logo.width / logo.height
            logo_w = max_logo if ratio >= 1 else int(max_logo * ratio)
            logo_h = int(max_logo / ratio) if ratio >= 1 else max_logo
            logo = logo.resize((logo_w, logo_h), Image.LANCZOS)
            px = (qr_w - logo_w) // 2
            py = (qr_h - logo_h) // 2
            pad = 6
            bg = Image.new("RGBA", (logo_w + pad * 2, logo_h + pad * 2), (255, 255, 255, 255))
            img.paste(bg, (px - pad, py - pad))
            img.paste(logo, (px, py), logo)
        except Exception:
            pass

    buf = io.BytesIO()
    img.convert("RGB").save(buf, format="PNG")
    buf.seek(0)
    return buf.read()


def generate_qr_pil(url: str, size_px: int, logo_bytes: Optional[bytes] = None):
    from PIL import Image
    raw = generate_qr_bytes(url, logo_bytes)
    img = Image.open(io.BytesIO(raw)).convert("RGB")
    return img.resize((size_px, size_px), Image.LANCZOS)
