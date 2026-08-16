"""Open XML compiler — Deck IR → PPTX (16:9).

Contract-first: the compiler is an *implementation* of the deck contract.
Semantic layouts are a registry (open for extension — adding a layout means
adding one entry, never editing dispatch logic). It renders structure only:
final layout fidelity, PDF and previews come from the COM renderer.
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from .deck_ir import DeckEnvelope, Slide

SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

# Brand palette (default theme until brand:// templates land in the registry).
COLOR_BG = RGBColor(0x0F, 0x14, 0x1E)
COLOR_BG_SOFT = RGBColor(0x17, 0x1F, 0x2E)
COLOR_PANEL = RGBColor(0x1E, 0x2A, 0x3D)
COLOR_ACCENT = RGBColor(0x4D, 0x9D, 0xE0)
COLOR_ACCENT_2 = RGBColor(0x7F, 0xC8, 0xA9)
COLOR_TEXT = RGBColor(0xE8, 0xEC, 0xF2)
COLOR_MUTED = RGBColor(0x9A, 0xA7, 0xBC)
COLOR_DARK_TEXT = RGBColor(0x1A, 0x22, 0x33)

FONT_LATIN = "Segoe UI"
FONT_CJK = "Microsoft YaHei"


class DeckCompiler:
    """Compiles a DeckEnvelope into a PPTX file via python-pptx (Open XML)."""

    def __init__(self, envelope: DeckEnvelope) -> None:
        self.envelope = envelope
        self.prs = Presentation()
        self.prs.slide_width = SLIDE_WIDTH
        self.prs.slide_height = SLIDE_HEIGHT
        self.blank = self.prs.slide_layouts[6]

    # -- helpers -----------------------------------------------------------

    def _set_run_font(self, run, name: str, size_pt: float, color: RGBColor, bold: bool = False) -> None:
        font = run.font
        font.name = name
        font.size = Pt(size_pt)
        font.bold = bold
        font.color.rgb = color
        rpr = run._r.get_or_add_rPr()
        for tag in ("a:latin", "a:ea"):
            existing = rpr.find(qn(tag))
            if existing is None:
                existing = rpr.makeelement(qn(tag), {})
                rpr.append(existing)
            existing.set("typeface", name)

    def _add_textbox(self, slide, left, top, width, height, text: str, *, size: float = 18, color: RGBColor = COLOR_TEXT, bold: bool = False, align=PP_ALIGN.LEFT) -> None:
        box = slide.shapes.add_textbox(left, top, width, height)
        box.shadow.inherit = False
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Pt(4)
        lines = text.split("\n")
        for i, line in enumerate(lines):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.alignment = align
            run = para.add_run()
            run.text = line
            self._set_run_font(run, FONT_LATIN, size, color, bold)
        return box

    def _add_shape(self, slide, shape_type, left, top, width, height, *, fill: RGBColor = COLOR_PANEL, line: RGBColor | None = None, radius: float | None = None):
        shape = slide.shapes.add_shape(shape_type, left, top, width, height)
        shape.shadow.inherit = False
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
        if line is None:
            shape.line.fill.background()
        else:
            shape.line.color.rgb = line
        if radius is not None and shape.adjustments:
            shape.adjustments[0] = radius
        tf = shape.text_frame
        tf.word_wrap = True
        return shape

    def _fill_shape(self, shape, text: str, *, size: float = 16, color: RGBColor = COLOR_TEXT, bold: bool = False, align=PP_ALIGN.CENTER) -> None:
        tf = shape.text_frame
        tf.margin_left = tf.margin_right = Pt(8)
        tf.margin_top = tf.margin_bottom = Pt(6)
        for i, line in enumerate(text.split("\n")):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.alignment = align
            run = para.add_run()
            run.text = line
            self._set_run_font(run, FONT_LATIN, size, color, bold)

    def _background(self, slide, color: RGBColor = COLOR_BG) -> None:
        rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT)
        rect.shadow.inherit = False
        rect.fill.solid()
        rect.fill.fore_color.rgb = color
        rect.line.fill.background()

    def _notes(self, slide, notes: str | None) -> None:
        if notes:
            slide.notes_slide.notes_text_frame.text = notes

    # -- slide dispatch (OCP: one registry, one entry per layout) -----------

    def compile(self, out_path: str | Path) -> Path:
        document = self.envelope.document
        for index, slide_ir in enumerate(document.slides):
            handler = LAYOUTS.get(slide_ir.semantic_layout)
            if handler is None:
                raise KeyError(
                    f"unknown semantic_layout '{slide_ir.semantic_layout}' (slide {index + 1}); "
                    f"known: {', '.join(sorted(LAYOUTS))}"
                )
            slide = self.prs.slides.add_slide(self.blank)
            self._background(slide)
            handler(self, slide, slide_ir)
            self._notes(slide, slide_ir.speaker_notes)
        out = Path(out_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(str(out))
        return out


# -- layouts ---------------------------------------------------------------

def _title_cover(c: DeckCompiler, slide, ir: Slide) -> None:
    c._background(slide, COLOR_BG)
    c._add_shape(slide, MSO_SHAPE.RECTANGLE, Inches(0.9), Inches(2.6), Inches(2.2), Pt(4)).fill.fore_color.rgb = COLOR_ACCENT
    c._add_textbox(slide, Inches(0.9), Inches(2.75), Inches(11.5), Inches(1.6), ir.title or c.envelope.metadata.title, size=54, bold=True)
    subtitle = next((b for b in ir.content_blocks if b["type"] == "text"), None)
    if subtitle:
        text = "\n".join(subtitle.get("paragraphs", []))
        c._add_textbox(slide, Inches(0.9), Inches(4.4), Inches(11.0), Inches(1.2), text, size=20, color=COLOR_MUTED)
    c._add_textbox(slide, Inches(0.9), Inches(6.7), Inches(11.0), Inches(0.4), f"{c.envelope.metadata.title} · {c.envelope.document_id}", size=12, color=COLOR_MUTED)


def _section_cover(c: DeckCompiler, slide, ir: Slide) -> None:
    c._background(slide, COLOR_BG_SOFT)
    c._add_shape(slide, MSO_SHAPE.RECTANGLE, Inches(0.9), Inches(3.15), Inches(0.18), Inches(1.2)).fill.fore_color.rgb = COLOR_ACCENT_2
    c._add_textbox(slide, Inches(1.3), Inches(3.1), Inches(11.0), Inches(1.4), ir.title or "", size=44, bold=True)


def _bullets(c: DeckCompiler, slide, ir: Slide) -> None:
    if ir.title:
        c._add_textbox(slide, Inches(0.9), Inches(0.5), Inches(11.5), Inches(0.8), ir.title, size=30, bold=True)
    y = Inches(1.5)
    for block in ir.content_blocks:
        if block["type"] != "bullets":
            continue
        for item in block.get("items", []):
            c._add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), y, Inches(11.5), Inches(0.85), fill=COLOR_PANEL, radius=0.18)
            shape = slide.shapes[-1]
            c._fill_shape(shape, "●  " + item, size=17, align=PP_ALIGN.LEFT)
            y = y + Inches(1.0)


def _two_columns(c: DeckCompiler, slide, ir: Slide) -> None:
    if ir.title:
        c._add_textbox(slide, Inches(0.9), Inches(0.5), Inches(11.5), Inches(0.8), ir.title, size=30, bold=True)
    texts = [b for b in ir.content_blocks if b["type"] == "bullets"]
    for col, block in enumerate(texts[:2]):
        left = Inches(0.9 + col * 6.0)
        panel = c._add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(1.5), Inches(5.6), Inches(5.2), fill=COLOR_PANEL, radius=0.06)
        c._fill_shape(panel, "\n".join("●  " + i for i in block.get("items", [])), size=16, align=PP_ALIGN.LEFT)


def _comparison(c: DeckCompiler, slide, ir: Slide) -> None:
    if ir.title:
        c._add_textbox(slide, Inches(0.9), Inches(0.5), Inches(11.5), Inches(0.8), ir.title, size=30, bold=True)
    texts = [b for b in ir.content_blocks if b["type"] == "bullets"]
    for col, block in enumerate(texts[:2]):
        left = Inches(0.9 + col * 6.0)
        color = COLOR_ACCENT if col == 0 else COLOR_ACCENT_2
        head = c._add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(1.5), Inches(5.6), Inches(0.7), fill=color, radius=0.12)
        c._fill_shape(head, block.get("items", [""])[0] if block.get("items") else "", size=18, bold=True)
        body = c._add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(2.35), Inches(5.6), Inches(4.35), fill=COLOR_PANEL, radius=0.06)
        c._fill_shape(body, "\n".join("●  " + i for i in block.get("items", [])[1:]), size=15, align=PP_ALIGN.LEFT)


def _timeline(c: DeckCompiler, slide, ir: Slide) -> None:
    if ir.title:
        c._add_textbox(slide, Inches(0.9), Inches(0.5), Inches(11.5), Inches(0.8), ir.title, size=30, bold=True)
    items = next((b["items"] for b in ir.content_blocks if b["type"] == "bullets"), [])
    n = max(len(items), 1)
    step = 11.5 / n
    y_line = Inches(3.6)
    c._add_shape(slide, MSO_SHAPE.RECTANGLE, Inches(0.9), y_line, Inches(11.5), Pt(3)).fill.fore_color.rgb = COLOR_ACCENT
    for i, item in enumerate(items):
        x = Inches(0.9 + step * i + step / 2)
        c._add_shape(slide, MSO_SHAPE.OVAL, x - Pt(9), y_line - Pt(4), Pt(18), Pt(18), fill=COLOR_ACCENT_2)
        c._add_textbox(slide, x - Inches(1.8), y_line + Inches(0.25), Inches(3.6), Inches(1.2), f"{i + 1:02d}  {item}", size=15, align=PP_ALIGN.CENTER)


def _kpi_dashboard(c: DeckCompiler, slide, ir: Slide) -> None:
    if ir.title:
        c._add_textbox(slide, Inches(0.9), Inches(0.5), Inches(11.5), Inches(0.8), ir.title, size=30, bold=True)
    items = next((b["items"] for b in ir.content_blocks if b["type"] == "bullets"), [])
    cols = 3
    card_w = Inches(3.6)
    card_h = Inches(1.9)
    for i, item in enumerate(items[:9]):
        row, col = divmod(i, cols)
        left = Inches(0.9 + col * 3.9)
        top = Inches(1.5 + row * 2.15)
        card = c._add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left, top, card_w, card_h, fill=COLOR_PANEL, radius=0.08)
        parts = item.split("|", 1)
        value = parts[0].strip()
        label = parts[1].strip() if len(parts) > 1 else ""
        c._fill_shape(card, value + "\n" + label, size=26, bold=(len(parts) == 1), color=COLOR_ACCENT_2 if len(parts) > 1 else COLOR_TEXT)


def _technical_architecture(c: DeckCompiler, slide, ir: Slide) -> None:
    if ir.title:
        c._add_textbox(slide, Inches(0.9), Inches(0.5), Inches(11.5), Inches(0.8), ir.title, size=30, bold=True)
    items = next((b["items"] for b in ir.content_blocks if b["type"] == "bullets"), [])
    top_ins = 1.6
    centers: list[float] = []
    for i, item in enumerate(items):
        top = Inches(top_ins + i * 0.95)
        box = c._add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2.8), top, Inches(7.7), Inches(0.85), fill=COLOR_PANEL, radius=0.10)
        c._fill_shape(box, item, size=15)
        centers.append(top_ins + i * 0.95 + 0.425)
    x = Inches(6.65)
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x, Inches(centers[0]), x, Inches(centers[-1]))
    conn.line.color.rgb = COLOR_ACCENT
    conn.shadow.inherit = False


def _image_left_text_right(c: DeckCompiler, slide, ir: Slide) -> None:
    if ir.title:
        c._add_textbox(slide, Inches(0.9), Inches(0.5), Inches(11.5), Inches(0.8), ir.title, size=30, bold=True)
    image = next((b for b in ir.content_blocks if b["type"] == "image"), None)
    if image and Path(image.get("resource", "")).is_file():
        slide.shapes.add_picture(image["resource"], Inches(0.9), Inches(1.5), height=Inches(5.2))
    texts = [b for b in ir.content_blocks if b["type"] == "bullets"]
    if texts:
        panel = c._add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.4), Inches(1.5), Inches(5.0), Inches(5.2), fill=COLOR_PANEL, radius=0.06)
        c._fill_shape(panel, "\n".join("●  " + i for i in texts[0].get("items", [])), size=15, align=PP_ALIGN.LEFT)


def _closing(c: DeckCompiler, slide, ir: Slide) -> None:
    c._background(slide, COLOR_BG)
    c._add_textbox(slide, Inches(1.5), Inches(2.8), Inches(10.3), Inches(1.5), ir.title or "Thanks", size=48, bold=True, align=PP_ALIGN.CENTER)
    text = next((b for b in ir.content_blocks if b["type"] == "text"), None)
    if text:
        c._add_textbox(slide, Inches(1.5), Inches(4.4), Inches(10.3), Inches(0.8), "\n".join(text.get("paragraphs", [])), size=18, color=COLOR_MUTED, align=PP_ALIGN.CENTER)


LAYOUTS: dict[str, Callable] = {
    "title_cover": _title_cover,
    "section_cover": _section_cover,
    "bullets": _bullets,
    "two_columns": _two_columns,
    "comparison": _comparison,
    "timeline": _timeline,
    "kpi_dashboard": _kpi_dashboard,
    "technical_architecture": _technical_architecture,
    "image_left_text_right": _image_left_text_right,
    "closing": _closing,
}


def compile_deck(envelope: DeckEnvelope, out_path: str | Path) -> Path:
    """Compile a DeckEnvelope to PPTX via the Open XML backend."""
    return DeckCompiler(envelope).compile(out_path)
