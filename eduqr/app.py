import os
import io
import tempfile
import threading
from tkinter import filedialog, messagebox
import tkinter as tk
import tkinter.simpledialog

import customtkinter as ctk
from PIL import Image, ImageTk

from .models import ClassEntry, GenerationConfig, TicketTemplate
from .utils import parse_classes, generate_qr_pil
from .generator import generate_docx
from .storage import load_sessions, save_sessions, delete_session, SavedSession

LAYOUTS: dict[str, tuple[int, int]] = {
    "2 × 5  —  10 por página":  (2, 5),
    "2 × 4  —  8 por página":   (2, 4),
    "2 × 3  —  6 por página":   (2, 3),
    "2 × 2  —  4 por página":   (2, 2),
    "1 × 4  —  4 por página":   (1, 4),
    "1 × 2  —  2 por página":   (1, 2),
    "1 × 1  —  1 por página":   (1, 1),
}

GREEN       = "#1B8A5A"
GREEN_HOVER = "#16724A"
GREEN_LIGHT = "#E8F5EE"
GREEN_MID   = "#D0EDDE"
SURFACE     = "#F7F9F8"
CARD        = "#FFFFFF"
BORDER      = "#DDE5E0"
TEXT_DARK   = "#111D17"
TEXT_MED    = "#3D5247"
TEXT_MUTED  = "#7A9688"
TEXT_FAINT  = "#B0C4BB"
ERROR       = "#C0392B"


class ClassCard(ctk.CTkFrame):
    def __init__(self, master, entry: ClassEntry, on_qty_change, **kwargs):
        super().__init__(master, fg_color=CARD, corner_radius=8,
                         border_width=1, border_color=BORDER, **kwargs)
        self.entry = entry
        self._on_qty_change = on_qty_change
        self._build()

    def _build(self):
        self.grid_columnconfigure(1, weight=1)

        indicator = ctk.CTkFrame(self, width=4, fg_color=GREEN, corner_radius=2)
        indicator.grid(row=0, column=0, rowspan=2, sticky="ns", padx=(10, 8), pady=8)

        ctk.CTkLabel(
            self, text=self.entry.display_name,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=TEXT_DARK, anchor="w",
        ).grid(row=0, column=1, sticky="w", pady=(8, 1))

        short = self.entry.link
        if len(short) > 44:
            short = short[:44] + "..."
        ctk.CTkLabel(
            self, text=short,
            font=ctk.CTkFont(family="Courier New", size=9),
            text_color=TEXT_MUTED, anchor="w",
        ).grid(row=1, column=1, sticky="w", pady=(0, 8))

        qty_frame = ctk.CTkFrame(self, fg_color="transparent")
        qty_frame.grid(row=0, column=2, rowspan=2, padx=(4, 10), pady=8)

        ctk.CTkLabel(
            qty_frame, text="Bilhetes",
            font=ctk.CTkFont(family="Segoe UI", size=9),
            text_color=TEXT_MUTED,
        ).pack(anchor="center")

        spin_frame = ctk.CTkFrame(qty_frame, fg_color=SURFACE, corner_radius=6,
                                   border_width=1, border_color=BORDER)
        spin_frame.pack()

        self._qty_var = tk.StringVar(value=str(self.entry.quantity))
        self._qty_var.trace_add("write", self._on_change)

        btn_cfg = dict(
            master=spin_frame, width=22, height=22, corner_radius=4,
            fg_color="transparent", hover_color=GREEN_LIGHT,
            text_color=TEXT_MED, font=ctk.CTkFont(size=13, weight="bold"),
        )
        ctk.CTkButton(**btn_cfg, text="−", command=self._decrement).pack(side="left", padx=2, pady=2)

        self._entry = ctk.CTkEntry(
            spin_frame, textvariable=self._qty_var,
            width=42, height=22, justify="center",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color="transparent", border_width=0,
            text_color=TEXT_DARK,
        )
        self._entry.pack(side="left")

        ctk.CTkButton(**btn_cfg, text="+", command=self._increment).pack(side="left", padx=2, pady=2)

    def _increment(self):
        try:
            self._qty_var.set(str(int(self._qty_var.get()) + 1))
        except ValueError:
            self._qty_var.set("1")

    def _decrement(self):
        try:
            val = max(1, int(self._qty_var.get()) - 1)
            self._qty_var.set(str(val))
        except ValueError:
            self._qty_var.set("1")

    def _on_change(self, *_):
        try:
            val = int(self._qty_var.get())
            if val >= 1:
                self.entry.quantity = val
                self._on_qty_change()
        except ValueError:
            pass


class TemplateEditor(ctk.CTkToplevel):
    def __init__(self, master, template: TicketTemplate, on_save):
        super().__init__(master)
        self.title("Personalizar texto do bilhete")
        self.configure(fg_color=SURFACE)
        self.grab_set()
        self.resizable(False, False)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"500x420+{(sw-500)//2}+{(sh-420)//2}")
        self._template = template
        self._on_save = on_save
        self._build()

    def _build(self):
        ctk.CTkFrame(self, height=3, fg_color=GREEN, corner_radius=0).pack(fill="x")

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(18, 8))
        ctk.CTkLabel(
            header, text="Texto do Bilhete",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            text_color=TEXT_DARK,
        ).pack(anchor="w")
        ctk.CTkLabel(
            header, text="Personalize o conteúdo que aparece em cada bilhete impresso",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=TEXT_MUTED,
        ).pack(anchor="w")

        form = ctk.CTkFrame(self, fg_color=CARD, corner_radius=10,
                             border_width=1, border_color=BORDER)
        form.pack(fill="x", padx=24, pady=8)
        form.grid_columnconfigure(1, weight=1)

        fields = [
            ("Título principal", self._template.title),
            ("Subtítulo / mensagem", self._template.subtitle),
            ("Prefixo do rodapé", self._template.footer_prefix),
        ]
        self._vars = []
        for i, (label, val) in enumerate(fields):
            ctk.CTkLabel(
                form, text=label,
                font=ctk.CTkFont(family="Segoe UI", size=11),
                text_color=TEXT_MED, anchor="w",
            ).grid(row=i, column=0, sticky="w", padx=(16, 12), pady=(12 if i == 0 else 4, 4))
            var = tk.StringVar(value=val)
            self._vars.append(var)
            ctk.CTkEntry(
                form, textvariable=var,
                font=ctk.CTkFont(family="Segoe UI", size=11),
                fg_color=SURFACE, border_color=BORDER,
                text_color=TEXT_DARK, height=34,
            ).grid(row=i, column=1, sticky="ew", padx=(0, 16), pady=(12 if i == 0 else 4, 4))

        self._show_name_var = tk.BooleanVar(value=self._template.show_class_name)
        check_row = ctk.CTkFrame(form, fg_color="transparent")
        check_row.grid(row=3, column=0, columnspan=2, sticky="w", padx=16, pady=(4, 14))
        ctk.CTkCheckBox(
            check_row, text="Mostrar nome da turma no rodapé",
            variable=self._show_name_var,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=TEXT_MED,
            fg_color=GREEN, hover_color=GREEN_HOVER,
            border_color=BORDER, checkmark_color=CARD,
        ).pack(anchor="w")

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(fill="x", padx=24, pady=(8, 20))
        ctk.CTkButton(
            btns, text="Cancelar",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=SURFACE, hover_color=GREEN_LIGHT,
            text_color=TEXT_MED, border_width=1, border_color=BORDER,
            height=38, width=110, command=self.destroy,
        ).pack(side="right", padx=(8, 0))
        ctk.CTkButton(
            btns, text="Salvar",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color=GREEN, hover_color=GREEN_HOVER,
            text_color=CARD, height=38, width=110, command=self._save,
        ).pack(side="right")

    def _save(self):
        self._template.title = self._vars[0].get().strip() or self._template.title
        self._template.subtitle = self._vars[1].get().strip()
        self._template.footer_prefix = self._vars[2].get().strip()
        self._template.show_class_name = self._show_name_var.get()
        self._on_save(self._template)
        self.destroy()


class PreviewWindow(ctk.CTkToplevel):
    def __init__(self, master, entry: ClassEntry, template: TicketTemplate, logo_bytes):
        super().__init__(master)
        self.title(f"Prévia — {entry.display_name}")
        self.configure(fg_color=SURFACE)
        self.grab_set()
        self.resizable(False, False)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"400x560+{(sw-400)//2}+{(sh-560)//2}")
        self._entry = entry
        self._template = template
        self._logo_bytes = logo_bytes
        self._img_ref = None
        self._build()

    def _build(self):
        ctk.CTkFrame(self, height=3, fg_color=GREEN, corner_radius=0).pack(fill="x")

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(16, 8))
        ctk.CTkLabel(
            header, text="Prévia do Bilhete",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color=TEXT_DARK,
        ).pack(anchor="w")
        ctk.CTkLabel(
            header, text=f"Turma: {self._entry.display_name}  •  QR Code gerado a partir do link real",
            font=ctk.CTkFont(family="Segoe UI", size=10),
            text_color=TEXT_MUTED,
        ).pack(anchor="w")

        ticket = ctk.CTkFrame(self, fg_color=CARD, corner_radius=10,
                               border_width=2, border_color=GREEN)
        ticket.pack(padx=32, pady=10, fill="x")

        ctk.CTkFrame(ticket, height=3, fg_color=GREEN, corner_radius=0).pack(fill="x")

        ctk.CTkLabel(
            ticket, text=self._template.title,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=GREEN, wraplength=300,
        ).pack(pady=(14, 2))

        if self._template.subtitle.strip():
            ctk.CTkLabel(
                ticket, text=self._template.subtitle,
                font=ctk.CTkFont(family="Segoe UI", size=9),
                text_color=TEXT_MUTED, wraplength=300,
            ).pack(pady=(0, 10))

        self._qr_label = ctk.CTkLabel(ticket, text="Gerando QR Code...",
                                       text_color=TEXT_MUTED,
                                       font=ctk.CTkFont(size=11))
        self._qr_label.pack(pady=6)

        if self._template.show_class_name:
            footer_text = self._template.footer_prefix + " " + self._entry.display_name
            ctk.CTkLabel(
                ticket, text=footer_text.strip(),
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                text_color=TEXT_DARK,
            ).pack(pady=(4, 16))

        ctk.CTkButton(
            self, text="Fechar",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=GREEN, hover_color=GREEN_HOVER,
            text_color=CARD, height=38, command=self.destroy,
        ).pack(padx=32, pady=12, fill="x")

        threading.Thread(target=self._load_qr, daemon=True).start()

    def _load_qr(self):
        try:
            pil_img = generate_qr_pil(self._entry.link, 180, self._logo_bytes)
            ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(180, 180))
            self._img_ref = ctk_img
            self.after(0, lambda: self._qr_label.configure(image=ctk_img, text=""))
        except Exception as ex:
            self.after(0, lambda: self._qr_label.configure(text=f"Erro: {ex}"))


class SessionDialog(ctk.CTkToplevel):
    def __init__(self, master, sessions: dict, on_load, on_delete):
        super().__init__(master)
        self.title("Sessões salvas")
        self.configure(fg_color=SURFACE)
        self.grab_set()
        self.resizable(False, False)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"380x380+{(sw-380)//2}+{(sh-380)//2}")
        self._sessions = sessions
        self._on_load = on_load
        self._on_delete = on_delete
        self._build()

    def _build(self):
        ctk.CTkFrame(self, height=3, fg_color=GREEN, corner_radius=0).pack(fill="x")

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=20, pady=(16, 8))
        ctk.CTkLabel(
            header, text="Sessões salvas",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color=TEXT_DARK,
        ).pack(anchor="w")

        list_frame = ctk.CTkScrollableFrame(self, fg_color=CARD, corner_radius=8,
                                             border_width=1, border_color=BORDER)
        list_frame.pack(fill="both", expand=True, padx=20, pady=4)

        self._selected = tk.StringVar()
        for name in self._sessions:
            row = ctk.CTkFrame(list_frame, fg_color="transparent")
            row.pack(fill="x", pady=1)
            rb = ctk.CTkRadioButton(
                row, text=name, variable=self._selected, value=name,
                font=ctk.CTkFont(family="Segoe UI", size=12),
                text_color=TEXT_DARK, fg_color=GREEN, hover_color=GREEN_HOVER,
                border_color=BORDER,
            )
            rb.pack(anchor="w", padx=12, pady=6)

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(fill="x", padx=20, pady=(8, 16))

        ctk.CTkButton(
            btns, text="Deletar",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=SURFACE, hover_color="#FFE8E8",
            text_color=ERROR, border_width=1, border_color="#F0C0C0",
            height=36, width=90, command=self._delete,
        ).pack(side="left")

        ctk.CTkButton(
            btns, text="Cancelar",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=SURFACE, hover_color=GREEN_LIGHT,
            text_color=TEXT_MED, border_width=1, border_color=BORDER,
            height=36, width=90, command=self.destroy,
        ).pack(side="right", padx=(8, 0))
        ctk.CTkButton(
            btns, text="Carregar",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color=GREEN, hover_color=GREEN_HOVER,
            text_color=CARD, height=36, width=100, command=self._load,
        ).pack(side="right")

    def _load(self):
        name = self._selected.get()
        if not name:
            return
        self._on_load(self._sessions[name])
        self.destroy()

    def _delete(self):
        name = self._selected.get()
        if not name:
            return
        if messagebox.askyesno("Deletar sessão", f"Deletar '{name}'?", parent=self):
            self._on_delete(name)
            self.destroy()


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("green")

        self.title("EduQR")
        self.configure(fg_color=SURFACE)

        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = 1080, 720
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        self.minsize(860, 580)

        self._classes: list[ClassEntry] = []
        self._logo_bytes: bytes | None = None
        self._logo_filename: str = ""
        self._template = TicketTemplate()
        self._sessions = load_sessions()
        self._class_cards: list[ClassCard] = []
        self._generating = False

        self._build()

    def _build(self):
        self._build_header()
        self._build_body()
        self._build_footer()

    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color=CARD, corner_radius=0,
                               border_width=0)
        header.pack(fill="x")
        ctk.CTkFrame(header, height=3, fg_color=GREEN, corner_radius=0).pack(fill="x")

        inner = ctk.CTkFrame(header, fg_color="transparent")
        inner.pack(fill="x", padx=24, pady=12)
        inner.grid_columnconfigure(1, weight=1)

        brand = ctk.CTkFrame(inner, fg_color="transparent")
        brand.grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(
            brand, text="EduQR",
            font=ctk.CTkFont(family="Georgia", size=22, weight="bold"),
            text_color=GREEN,
        ).pack(side="left")
        ctk.CTkLabel(
            brand, text="  Gerador de Bilhetes WhatsApp",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color=TEXT_MUTED,
        ).pack(side="left")

        actions = ctk.CTkFrame(inner, fg_color="transparent")
        actions.grid(row=0, column=2, sticky="e")

        ctk.CTkButton(
            actions, text="Sessoes",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=SURFACE, hover_color=GREEN_LIGHT,
            text_color=TEXT_MED, border_width=1, border_color=BORDER,
            height=32, width=90, command=self._open_sessions,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            actions, text="Salvar sessao",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=SURFACE, hover_color=GREEN_LIGHT,
            text_color=TEXT_MED, border_width=1, border_color=BORDER,
            height=32, width=110, command=self._save_session,
        ).pack(side="left")

    def _build_body(self):
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=(12, 0))
        body.grid_columnconfigure(0, weight=5)
        body.grid_columnconfigure(1, weight=4)
        body.grid_rowconfigure(0, weight=1)

        self._build_left_panel(body)
        self._build_right_panel(body)

    def _build_left_panel(self, parent):
        left = ctk.CTkFrame(parent, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.grid_rowconfigure(1, weight=1)
        left.grid_columnconfigure(0, weight=1)

        top_bar = ctk.CTkFrame(left, fg_color="transparent")
        top_bar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        top_bar.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            top_bar, text="TURMAS E LINKS",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            text_color=TEXT_MUTED, anchor="w",
        ).grid(row=0, column=0, sticky="w")

        hint = ctk.CTkLabel(
            top_bar, text="código da turma na linha acima do link",
            font=ctk.CTkFont(family="Segoe UI", size=9),
            text_color=TEXT_FAINT, anchor="w",
        )
        hint.grid(row=0, column=1, sticky="w", padx=(8, 0))

        suffix_frame = ctk.CTkFrame(top_bar, fg_color="transparent")
        suffix_frame.grid(row=0, column=2, sticky="e")

        ctk.CTkLabel(
            suffix_frame, text="Sufixo:",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=TEXT_MED,
        ).pack(side="left", padx=(0, 6))

        self._suffix_var = tk.StringVar(value="")
        ctk.CTkEntry(
            suffix_frame, textvariable=self._suffix_var,
            width=80, height=30,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=CARD, border_color=BORDER,
            text_color=TEXT_DARK, placeholder_text="ex: AI",
        ).pack(side="left")
        self._suffix_var.trace_add("write", lambda *_: self._refresh())

        text_frame = ctk.CTkFrame(left, fg_color=CARD, corner_radius=8,
                                   border_width=1, border_color=BORDER)
        text_frame.grid(row=1, column=0, sticky="nsew")
        text_frame.grid_rowconfigure(0, weight=1)
        text_frame.grid_columnconfigure(0, weight=1)

        self._text_input = ctk.CTkTextbox(
            text_frame,
            font=ctk.CTkFont(family="Courier New", size=11),
            fg_color=CARD, text_color=TEXT_DARK,
            border_width=0, corner_radius=8,
            scrollbar_button_color=GREEN_LIGHT,
            wrap="word",
        )
        self._text_input.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        self._text_input.insert("1.0",
            "101\nhttps://chat.whatsapp.com/...\n"
            "102\nhttps://chat.whatsapp.com/...\n"
            "201\nhttps://chat.whatsapp.com/..."
        )
        self._text_input.bind("<KeyRelease>", lambda _: self._refresh())

    def _build_right_panel(self, parent):
        right = ctk.CTkFrame(parent, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_rowconfigure(2, weight=1)
        right.grid_columnconfigure(0, weight=1)

        top_bar = ctk.CTkFrame(right, fg_color="transparent")
        top_bar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        top_bar.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            top_bar, text="TURMAS DETECTADAS",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            text_color=TEXT_MUTED, anchor="w",
        ).grid(row=0, column=0, sticky="w")

        layout_frame = ctk.CTkFrame(top_bar, fg_color="transparent")
        layout_frame.grid(row=0, column=1, sticky="e")

        ctk.CTkLabel(
            layout_frame, text="Layout:",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=TEXT_MED,
        ).pack(side="left", padx=(0, 6))

        self._layout_var = tk.StringVar(value="2 × 5  —  10 por página")
        ctk.CTkComboBox(
            layout_frame, variable=self._layout_var,
            values=list(LAYOUTS.keys()),
            font=ctk.CTkFont(family="Segoe UI", size=10),
            fg_color=CARD, border_color=BORDER,
            button_color=GREEN, button_hover_color=GREEN_HOVER,
            dropdown_fg_color=CARD, dropdown_text_color=TEXT_DARK,
            text_color=TEXT_DARK, width=188, height=30,
            state="readonly",
        ).pack(side="left")

        self._summary_label = ctk.CTkLabel(
            right, text="",
            font=ctk.CTkFont(family="Segoe UI", size=10),
            text_color=TEXT_MUTED, anchor="w",
        )
        self._summary_label.grid(row=1, column=0, sticky="w", pady=(0, 4))

        self._cards_frame = ctk.CTkScrollableFrame(
            right, fg_color=SURFACE, corner_radius=8,
            border_width=1, border_color=BORDER,
            scrollbar_button_color=GREEN_LIGHT,
        )
        self._cards_frame.grid(row=2, column=0, sticky="nsew")
        self._cards_frame.grid_columnconfigure(0, weight=1)

        preview_btn = ctk.CTkButton(
            right, text="Ver prévia do bilhete",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=SURFACE, hover_color=GREEN_LIGHT,
            text_color=TEXT_MED, border_width=1, border_color=BORDER,
            height=34, command=self._open_preview,
        )
        preview_btn.grid(row=3, column=0, sticky="ew", pady=(6, 0))

        self._refresh()

    def _build_footer(self):
        footer = ctk.CTkFrame(self, fg_color=CARD, corner_radius=0,
                               border_width=0)
        footer.pack(fill="x", pady=(12, 0))
        ctk.CTkFrame(footer, height=1, fg_color=BORDER, corner_radius=0).pack(fill="x")

        inner = ctk.CTkFrame(footer, fg_color="transparent")
        inner.pack(fill="x", padx=20, pady=10)
        inner.grid_columnconfigure(1, weight=1)

        logo_group = ctk.CTkFrame(inner, fg_color="transparent")
        logo_group.grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            logo_group, text="Logo no QR:",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=TEXT_MED,
        ).pack(side="left", padx=(0, 6))

        self._logo_name_var = tk.StringVar(value="Nenhum")
        ctk.CTkLabel(
            logo_group, textvariable=self._logo_name_var,
            font=ctk.CTkFont(family="Courier New", size=9),
            text_color=TEXT_MUTED,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            logo_group, text="Escolher",
            font=ctk.CTkFont(family="Segoe UI", size=10),
            fg_color=SURFACE, hover_color=GREEN_LIGHT,
            text_color=TEXT_MED, border_width=1, border_color=BORDER,
            height=28, width=80, command=self._pick_logo,
        ).pack(side="left", padx=(0, 4))

        ctk.CTkButton(
            logo_group, text="Remover",
            font=ctk.CTkFont(family="Segoe UI", size=10),
            fg_color=SURFACE, hover_color="#FFE8E8",
            text_color=TEXT_MUTED, border_width=1, border_color=BORDER,
            height=28, width=72, command=self._remove_logo,
        ).pack(side="left")

        ctk.CTkButton(
            logo_group, text="Editar texto",
            font=ctk.CTkFont(family="Segoe UI", size=10),
            fg_color=GREEN_LIGHT, hover_color=GREEN_MID,
            text_color=GREEN, border_width=1, border_color=GREEN,
            height=28, width=90, command=self._edit_template,
        ).pack(side="left", padx=(16, 0))

        status_group = ctk.CTkFrame(inner, fg_color="transparent")
        status_group.grid(row=0, column=1, padx=16, sticky="ew")
        status_group.grid_columnconfigure(0, weight=1)

        self._status_var = tk.StringVar(value="Pronto.")
        ctk.CTkLabel(
            status_group, textvariable=self._status_var,
            font=ctk.CTkFont(family="Segoe UI", size=10),
            text_color=TEXT_MUTED, anchor="w",
        ).grid(row=0, column=0, sticky="ew")

        self._progress = ctk.CTkProgressBar(
            status_group, fg_color=SURFACE, progress_color=GREEN,
            height=4, corner_radius=2,
        )
        self._progress.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self._progress.set(0)

        btn_group = ctk.CTkFrame(inner, fg_color="transparent")
        btn_group.grid(row=0, column=2, sticky="e")

        ctk.CTkButton(
            btn_group, text="Imprimir",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=SURFACE, hover_color=GREEN_LIGHT,
            text_color=TEXT_MED, border_width=1, border_color=BORDER,
            height=40, width=100, command=self._print,
        ).pack(side="left", padx=(0, 8))

        self._gen_btn = ctk.CTkButton(
            btn_group, text="Gerar DOCX",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color=GREEN, hover_color=GREEN_HOVER,
            text_color=CARD, height=40, width=130,
            command=self._generate,
        )
        self._gen_btn.pack(side="left")

    def _refresh(self):
        raw = self._text_input.get("1.0", "end-1c").strip()
        suffix = self._suffix_var.get()
        self._classes = parse_classes(raw, suffix)

        for w in self._cards_frame.winfo_children():
            w.destroy()
        self._class_cards.clear()

        if not self._classes:
            ctk.CTkLabel(
                self._cards_frame, text="Nenhuma turma detectada.\nCole os links à esquerda.",
                font=ctk.CTkFont(family="Segoe UI", size=11),
                text_color=TEXT_FAINT,
            ).pack(pady=24)
            self._summary_label.configure(text="")
            return

        for entry in self._classes:
            card = ClassCard(self._cards_frame, entry, self._update_summary)
            card.pack(fill="x", padx=4, pady=3)
            self._class_cards.append(card)

        self._update_summary()

    def _update_summary(self):
        total_tickets = sum(e.quantity for e in self._classes)
        n = len(self._classes)
        self._summary_label.configure(
            text=f"{n} turma{'s' if n != 1 else ''}  •  {total_tickets} bilhete{'s' if total_tickets != 1 else ''} no total"
        )

    def _pick_logo(self):
        path = filedialog.askopenfilename(
            title="Selecionar logo para o QR Code",
            filetypes=[("Imagens", "*.png *.jpg *.jpeg *.bmp *.gif *.webp"), ("Todos", "*.*")],
        )
        if not path:
            return
        with open(path, "rb") as f:
            self._logo_bytes = f.read()
        self._logo_filename = os.path.basename(path)
        name = self._logo_filename
        if len(name) > 22:
            name = name[:20] + "..."
        self._logo_name_var.set(name)

    def _remove_logo(self):
        self._logo_bytes = None
        self._logo_filename = ""
        self._logo_name_var.set("Nenhum")

    def _edit_template(self):
        TemplateEditor(self, self._template, self._on_template_saved)

    def _on_template_saved(self, template: TicketTemplate):
        self._template = template

    def _open_preview(self):
        if not self._classes:
            messagebox.showinfo("Prévia", "Nenhuma turma detectada.")
            return
        entry = self._classes[0]
        PreviewWindow(self, entry, self._template, self._logo_bytes)

    def _open_sessions(self):
        if not self._sessions:
            messagebox.showinfo("Sessoes", "Nenhuma sessao salva.")
            return
        SessionDialog(self, self._sessions, self._load_session, self._delete_session)

    def _save_session(self):
        if not self._classes:
            messagebox.showinfo("Salvar", "Sem turmas para salvar.")
            return
        name = tkinter.simpledialog.askstring("Salvar sessao", "Nome da sessao:", parent=self)
        if not name or not name.strip():
            return
        name = name.strip()
        raw = self._text_input.get("1.0", "end-1c").strip()
        self._sessions[name] = SavedSession(
            name=name,
            raw_text=raw,
            suffix=self._suffix_var.get(),
            layout_key=self._layout_var.get(),
            template=TicketTemplate(
                title=self._template.title,
                subtitle=self._template.subtitle,
                footer_prefix=self._template.footer_prefix,
                show_class_name=self._template.show_class_name,
            ),
            quantities={e.display_name: e.quantity for e in self._classes},
        )
        save_sessions(self._sessions)
        self._status_var.set(f"Sessao '{name}' salva.")

    def _load_session(self, session: SavedSession):
        self._text_input.delete("1.0", "end")
        self._text_input.insert("1.0", session.raw_text)
        self._suffix_var.set(session.suffix)
        if session.layout_key in LAYOUTS:
            self._layout_var.set(session.layout_key)
        self._template = session.template
        self._refresh()
        for entry in self._classes:
            if entry.display_name in session.quantities:
                entry.quantity = session.quantities[entry.display_name]
        for card in self._class_cards:
            if card.entry.display_name in session.quantities:
                card._qty_var.set(str(session.quantities[card.entry.display_name]))
        self._update_summary()

    def _delete_session(self, name: str):
        delete_session(self._sessions, name)

    def _build_config(self, output_path: str) -> GenerationConfig:
        cols, rows = LAYOUTS[self._layout_var.get()]
        return GenerationConfig(
            cols=cols,
            rows_per_page=rows,
            logo_bytes=self._logo_bytes,
            logo_filename=self._logo_filename,
            output_path=output_path,
            template=TicketTemplate(
                title=self._template.title,
                subtitle=self._template.subtitle,
                footer_prefix=self._template.footer_prefix,
                show_class_name=self._template.show_class_name,
            ),
        )

    def _generate(self, for_print: bool = False):
        if not self._classes:
            messagebox.showerror("Erro", "Nenhuma turma encontrada.")
            return
        if self._generating:
            return

        if for_print:
            tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
            output_path = tmp.name
            tmp.close()
        else:
            output_path = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word Document", "*.docx")],
                initialfile="bilhetes.docx",
                title="Salvar como...",
            )
            if not output_path:
                return

        config = self._build_config(output_path)
        classes = list(self._classes)

        self._generating = True
        self._gen_btn.configure(state="disabled", text="Gerando...")
        self._progress.set(0)

        def on_progress(current: int, total: int, name: str):
            pct = current / total if total > 0 else 0
            msg = f"Processando: {name}  ({current}/{total})" if current < total else "Salvando..."
            self.after(0, lambda: self._progress.set(pct))
            self.after(0, lambda: self._status_var.set(msg))

        def run():
            try:
                generate_docx(classes, config, on_progress)
                if for_print:
                    self.after(0, lambda: self._send_to_printer(output_path))
                else:
                    self.after(0, lambda: self._on_success(output_path))
            except Exception as ex:
                self.after(0, lambda: self._on_error(str(ex)))

        threading.Thread(target=run, daemon=True).start()

    def _print(self):
        if not self._classes:
            messagebox.showerror("Erro", "Nenhuma turma encontrada.")
            return
        if messagebox.askyesno("Imprimir", "Gerar e enviar para a impressora padrão?"):
            self._generate(for_print=True)

    def _send_to_printer(self, path: str):
        self._progress.set(1.0)
        self._status_var.set("Enviando para impressora...")
        self._reset_gen_btn()
        try:
            os.startfile(path, "print")
            self._status_var.set("Enviado para impressora.")
        except Exception as ex:
            messagebox.showerror("Erro ao imprimir", str(ex))
            self._status_var.set("Erro ao imprimir.")

    def _on_success(self, path: str):
        self._progress.set(1.0)
        self._status_var.set(f"Salvo: {os.path.basename(path)}")
        self._reset_gen_btn()

        win = ctk.CTkToplevel(self)
        win.title("Concluido")
        win.configure(fg_color=SURFACE)
        win.grab_set()
        win.resizable(False, False)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        win.geometry(f"340x200+{(sw-340)//2}+{(sh-200)//2}")

        ctk.CTkFrame(win, height=3, fg_color=GREEN, corner_radius=0).pack(fill="x")
        ctk.CTkLabel(
            win, text="Bilhetes gerados com sucesso",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color=TEXT_DARK,
        ).pack(pady=(22, 4))
        ctk.CTkLabel(
            win, text=os.path.basename(path),
            font=ctk.CTkFont(family="Courier New", size=9),
            text_color=TEXT_MUTED,
        ).pack()
        ctk.CTkButton(
            win, text="OK",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color=GREEN, hover_color=GREEN_HOVER,
            text_color=CARD, height=38, width=120, command=win.destroy,
        ).pack(pady=18)

    def _on_error(self, msg: str):
        self._reset_gen_btn()
        self._status_var.set("Erro na geracao.")
        self._progress.set(0)
        messagebox.showerror("Erro", f"Ocorreu um erro:\n\n{msg}")

    def _reset_gen_btn(self):
        self._generating = False
        self._gen_btn.configure(state="normal", text="Gerar DOCX")
