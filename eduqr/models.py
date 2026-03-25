from dataclasses import dataclass, field
from typing import Optional


@dataclass
class TicketTemplate:
    title: str = "Grupo do WhatsApp para os pais"
    subtitle: str = "Caso já estiver no grupo, não se preocupe!"
    footer_prefix: str = "Turma:"
    show_class_name: bool = True


@dataclass
class ClassEntry:
    code: str
    link: str
    quantity: int = 10
    suffix: str = ""

    @property
    def display_name(self) -> str:
        if self.suffix.strip():
            return f"{self.code} {self.suffix.strip()}"
        return self.code


@dataclass
class GenerationConfig:
    cols: int = 2
    rows_per_page: int = 5
    logo_bytes: Optional[bytes] = None
    logo_filename: str = ""
    output_path: str = ""
    template: TicketTemplate = field(default_factory=TicketTemplate)

    @property
    def tickets_per_page(self) -> int:
        return self.cols * self.rows_per_page


@dataclass
class SavedSession:
    name: str
    raw_text: str
    suffix: str
    layout_key: str
    template: TicketTemplate
    quantities: dict[str, int] = field(default_factory=dict)
