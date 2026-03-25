import json
import os
import tempfile
import pytest

from eduqr.models import SavedSession, TicketTemplate
from eduqr import storage


@pytest.fixture
def tmp_saves_path(monkeypatch, tmp_path):
    path = str(tmp_path / "test_sessions.json")
    monkeypatch.setattr(storage, "SAVES_PATH", path)
    return path


def _make_session(name="Turma AI") -> SavedSession:
    return SavedSession(
        name=name,
        raw_text="101\nhttps://chat.whatsapp.com/abc",
        suffix="AI",
        layout_key="2 × 5  —  10 por página",
        template=TicketTemplate(
            title="Grupo do WhatsApp",
            subtitle="Não se preocupe!",
            footer_prefix="Turma:",
            show_class_name=True,
        ),
        quantities={"101 AI": 15},
    )


class TestStorage:
    def test_load_returns_empty_when_no_file(self, tmp_saves_path):
        result = storage.load_sessions()
        assert result == {}

    def test_save_and_load_roundtrip(self, tmp_saves_path):
        session = _make_session()
        sessions = {session.name: session}
        storage.save_sessions(sessions)

        loaded = storage.load_sessions()
        assert "Turma AI" in loaded
        s = loaded["Turma AI"]
        assert s.raw_text == session.raw_text
        assert s.suffix == session.suffix
        assert s.layout_key == session.layout_key
        assert s.quantities == session.quantities

    def test_template_roundtrip(self, tmp_saves_path):
        session = _make_session()
        storage.save_sessions({session.name: session})
        loaded = storage.load_sessions()
        t = loaded[session.name].template
        assert t.title == "Grupo do WhatsApp"
        assert t.subtitle == "Não se preocupe!"
        assert t.footer_prefix == "Turma:"
        assert t.show_class_name is True

    def test_delete_session(self, tmp_saves_path):
        session = _make_session()
        sessions = {session.name: session}
        storage.save_sessions(sessions)
        storage.delete_session(sessions, session.name)
        loaded = storage.load_sessions()
        assert session.name not in loaded

    def test_multiple_sessions(self, tmp_saves_path):
        s1 = _make_session("Sessao 1")
        s2 = _make_session("Sessao 2")
        sessions = {s1.name: s1, s2.name: s2}
        storage.save_sessions(sessions)
        loaded = storage.load_sessions()
        assert len(loaded) == 2

    def test_corrupted_file_returns_empty(self, tmp_saves_path):
        with open(tmp_saves_path, "w") as f:
            f.write("{ invalid json }")
        result = storage.load_sessions()
        assert result == {}

    def test_missing_keys_fall_back_to_defaults(self, tmp_saves_path):
        with open(tmp_saves_path, "w", encoding="utf-8") as f:
            json.dump({"Velha": {"raw_text": "101\nhttps://x.com/a"}}, f)
        loaded = storage.load_sessions()
        s = loaded["Velha"]
        assert s.suffix == ""
        assert s.template.title == "Grupo do WhatsApp para os pais"
        assert s.quantities == {}