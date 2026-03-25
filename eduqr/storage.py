import json
import os
from typing import Optional
from .models import SavedSession, TicketTemplate

SAVES_PATH = os.path.join(os.path.expanduser("~"), ".eduqr_sessions.json")


def load_sessions() -> dict[str, SavedSession]:
    if not os.path.exists(SAVES_PATH):
        return {}
    try:
        with open(SAVES_PATH, "r", encoding="utf-8") as f:
            raw = json.load(f)
        sessions = {}
        for name, d in raw.items():
            t = d.get("template", {})
            sessions[name] = SavedSession(
                name=name,
                raw_text=d.get("raw_text", ""),
                suffix=d.get("suffix", ""),
                layout_key=d.get("layout_key", "2 × 5  (10/pág)"),
                template=TicketTemplate(
                    title=t.get("title", "Grupo do WhatsApp para os pais"),
                    subtitle=t.get("subtitle", "Caso já estiver no grupo, não se preocupe!"),
                    footer_prefix=t.get("footer_prefix", "Turma:"),
                    show_class_name=t.get("show_class_name", True),
                ),
                quantities=d.get("quantities", {}),
            )
        return sessions
    except Exception:
        return {}


def save_sessions(sessions: dict[str, SavedSession]) -> None:
    try:
        data = {}
        for name, s in sessions.items():
            data[name] = {
                "raw_text": s.raw_text,
                "suffix": s.suffix,
                "layout_key": s.layout_key,
                "template": {
                    "title": s.template.title,
                    "subtitle": s.template.subtitle,
                    "footer_prefix": s.template.footer_prefix,
                    "show_class_name": s.template.show_class_name,
                },
                "quantities": s.quantities,
            }
        with open(SAVES_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def delete_session(sessions: dict[str, SavedSession], name: str) -> None:
    sessions.pop(name, None)
    save_sessions(sessions)
