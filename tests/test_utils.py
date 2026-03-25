import pytest
from eduqr.utils import parse_classes, generate_qr_bytes
from eduqr.models import ClassEntry


class TestParseClasses:
    def test_basic(self):
        text = "101\nhttps://chat.whatsapp.com/abc"
        result = parse_classes(text)
        assert len(result) == 1
        assert result[0].code == "101"
        assert result[0].link == "https://chat.whatsapp.com/abc"

    def test_multiple_classes(self):
        text = "101\nhttps://chat.whatsapp.com/abc\n202\nhttps://chat.whatsapp.com/xyz"
        result = parse_classes(text)
        assert len(result) == 2
        assert result[0].code == "101"
        assert result[1].code == "202"

    def test_suffix_applied(self):
        text = "101\nhttps://chat.whatsapp.com/abc"
        result = parse_classes(text, suffix="AI")
        assert result[0].display_name == "101 AI"

    def test_empty_suffix(self):
        text = "101\nhttps://chat.whatsapp.com/abc"
        result = parse_classes(text, suffix="")
        assert result[0].display_name == "101"

    def test_suffix_stripped(self):
        text = "101\nhttps://chat.whatsapp.com/abc"
        result = parse_classes(text, suffix="  AI  ")
        assert result[0].display_name == "101 AI"

    def test_orphan_link_ignored(self):
        text = "https://chat.whatsapp.com/orphan\n101\nhttps://chat.whatsapp.com/abc"
        result = parse_classes(text)
        assert len(result) == 1
        assert result[0].code == "101"

    def test_empty_text(self):
        assert parse_classes("") == []

    def test_blank_lines_ignored(self):
        text = "\n101\nhttps://chat.whatsapp.com/abc\n\n"
        result = parse_classes(text)
        assert len(result) == 1

    def test_default_quantity(self):
        text = "101\nhttps://chat.whatsapp.com/abc"
        result = parse_classes(text)
        assert result[0].quantity == 10

    def test_http_variant(self):
        text = "101\nhttp://chat.whatsapp.com/abc"
        result = parse_classes(text)
        assert len(result) == 1


class TestGenerateQrBytes:
    def test_returns_bytes(self):
        result = generate_qr_bytes("https://chat.whatsapp.com/abc")
        assert isinstance(result, bytes)
        assert len(result) > 0

    def test_is_valid_png(self):
        result = generate_qr_bytes("https://chat.whatsapp.com/abc")
        assert result[:8] == b"\x89PNG\r\n\x1a\n"

    def test_with_logo(self):
        from PIL import Image
        import io
        img = Image.new("RGB", (50, 50), color=(255, 0, 0))
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        logo_bytes = buf.getvalue()

        result = generate_qr_bytes("https://chat.whatsapp.com/abc", logo_bytes)
        assert isinstance(result, bytes)
        assert result[:8] == b"\x89PNG\r\n\x1a\n"

    def test_different_urls_produce_different_qr(self):
        a = generate_qr_bytes("https://chat.whatsapp.com/aaa")
        b = generate_qr_bytes("https://chat.whatsapp.com/bbb")
        assert a != b