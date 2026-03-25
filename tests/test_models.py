from eduqr.models import ClassEntry, TicketTemplate, GenerationConfig


class TestClassEntry:
    def test_display_name_with_suffix(self):
        e = ClassEntry(code="101", link="https://x.com", suffix="AI")
        assert e.display_name == "101 AI"

    def test_display_name_without_suffix(self):
        e = ClassEntry(code="101", link="https://x.com", suffix="")
        assert e.display_name == "101"

    def test_display_name_strips_whitespace(self):
        e = ClassEntry(code="101", link="https://x.com", suffix="  ")
        assert e.display_name == "101"

    def test_default_quantity(self):
        e = ClassEntry(code="101", link="https://x.com")
        assert e.quantity == 10


class TestGenerationConfig:
    def test_tickets_per_page(self):
        cfg = GenerationConfig(cols=2, rows_per_page=5, output_path="out.docx")
        assert cfg.tickets_per_page == 10

    def test_tickets_per_page_1x1(self):
        cfg = GenerationConfig(cols=1, rows_per_page=1, output_path="out.docx")
        assert cfg.tickets_per_page == 1


class TestTicketTemplate:
    def test_defaults(self):
        t = TicketTemplate()
        assert t.title == "Grupo do WhatsApp para os pais"
        assert t.show_class_name is True