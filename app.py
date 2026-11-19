from io import BytesIO
import base64
import re
from html import escape

from flask import Flask, render_template, request, send_file, abort
from bs4 import BeautifulSoup, NavigableString, Tag
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, HRFlowable
from PIL import Image as PILImage, ImageOps
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024

PAGE_WIDTH, PAGE_HEIGHT = A4
PDF_CONTENT_WIDTH = PAGE_WIDTH - 100


def image_bytes_from_data_url(src):
    """Decode an image inserted by the browser's image picker."""
    if not src or not src.startswith("data:image/") or "," not in src:
        return None
    try:
        header, encoded = src.split(",", 1)
        raw = base64.b64decode(encoded)
        img = PILImage.open(BytesIO(raw))
        img = ImageOps.exif_transpose(img)
        img.load()
        if img.mode not in ("RGB", "RGBA"):
            img = img.convert("RGB")
        out = BytesIO()
        if img.mode == "RGBA":
            img.save(out, format="PNG")
        else:
            img.save(out, format="JPEG", quality=88)
        out.seek(0)
        return out
    except Exception:
        return None


def image_width_px(tag, default=300):
    style = tag.get("style", "") or ""
    match = re.search(r"width\s*:\s*(\d+(?:\.\d+)?)\s*px", style, re.I)
    if match:
        return float(match.group(1))
    width_attr = tag.get("width")
    if width_attr and str(width_attr).isdigit():
        return float(width_attr)
    return float(default)


def image_alignment(tag):
    style = (tag.get("style", "") or "").lower()
    if "margin-left: auto" in style and "margin-right: auto" in style:
        return "center"
    if "margin-left: auto" in style:
        return "right"
    return "left"


def inline_markup(node):
    """Convert safe, simple editor markup to ReportLab Paragraph markup."""
    if isinstance(node, NavigableString):
        return escape(str(node)).replace("\n", "<br/>")
    if not isinstance(node, Tag):
        return ""
    name = node.name.lower()
    if name == "img":
        return ""
    inner = "".join(inline_markup(child) for child in node.children)
    if name in ("b", "strong"):
        return f"<b>{inner}</b>"
    if name in ("i", "em"):
        return f"<i>{inner}</i>"
    if name == "u":
        return f"<u>{inner}</u>"
    if name == "br":
        return "<br/>"
    if name in ("span", "font"):
        return inner
    return inner


def add_pdf_image(story, tag):
    stream = image_bytes_from_data_url(tag.get("src", ""))
    if not stream:
        return
    px_width = image_width_px(tag)
    # Browser CSS pixels -> PDF points; never exceed printable page width.
    width = min(px_width * 0.75, PDF_CONTENT_WIDTH)
    try:
        pil = PILImage.open(stream)
        w, h = pil.size
        stream.seek(0)
        height = width * (h / max(w, 1))
        flow = RLImage(stream, width=width, height=height)
        align = image_alignment(tag)
        flow.hAlign = {"left": "LEFT", "center": "CENTER", "right": "RIGHT"}.get(align, "LEFT")
        story.append(flow)
        story.append(Spacer(1, 7))
    except Exception:
        pass


def build_pdf(title, html_content):
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        rightMargin=50, leftMargin=50, topMargin=45, bottomMargin=45
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "CustomTitle", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=13, leading=16, textColor=colors.HexColor("#111111"),
        spaceAfter=7
    )
    body_style = ParagraphStyle(
        "CustomBody", parent=styles["Normal"], fontName="Helvetica",
        fontSize=9, leading=12, textColor=colors.HexColor("#333333"),
        spaceAfter=6, wordWrap="CJK"
    )

    story = [Paragraph(escape(title or "Document Title"), title_style),
             HRFlowable(width="100%", thickness=0.7,
                        color=colors.HexColor("#cccccc"),
                        spaceBefore=2, spaceAfter=10)]

    soup = BeautifulSoup(html_content or "", "html.parser")

    def process_node(node):
        if isinstance(node, NavigableString):
            txt = str(node).strip()
            if txt:
                story.append(Paragraph(escape(txt), body_style))
            return
        if not isinstance(node, Tag):
            return
        if node.name.lower() == "img":
            add_pdf_image(story, node)
            return

        # Each paragraph/heading/div is kept in reading order. Images inside it
        # are emitted at their insertion point in the PDF flow.
        if node.name.lower() in ("p", "div", "h1", "h2", "h3", "li", "blockquote"):
            chunks = []
            for child in node.children:
                if isinstance(child, Tag) and child.name.lower() == "img":
                    markup = "".join(chunks).strip()
                    if markup:
                        story.append(Paragraph(markup, body_style))
                    chunks = []
                    add_pdf_image(story, child)
                elif isinstance(child, Tag) and child.name.lower() in ("div", "p", "ul", "ol"):
                    markup = "".join(chunks).strip()
                    if markup:
                        story.append(Paragraph(markup, body_style))
                    chunks = []
                    process_node(child)
                else:
                    chunks.append(inline_markup(child))
            markup = "".join(chunks).strip()
            if markup:
                story.append(Paragraph(markup, body_style))
            elif not node.find("img"):
                story.append(Spacer(1, 3))
        else:
            for child in node.children:
                process_node(child)

    root = soup.body if soup.body else soup
    for child in root.children:
        process_node(child)

    doc.build(story)
    buffer.seek(0)
    return buffer


def add_docx_image(paragraph, tag):
    stream = image_bytes_from_data_url(tag.get("src", ""))
    if not stream:
        return
    width_inches = min(image_width_px(tag) / 96.0, 6.3)
    try:
        run = paragraph.add_run()
        run.add_picture(stream, width=Inches(width_inches))
        align = image_alignment(tag)
        if align == "center":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "right":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        else:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    except Exception:
        pass


def add_docx_inline(paragraph, node):
    if isinstance(node, NavigableString):
        text = str(node)
        if text:
            paragraph.add_run(text)
        return
    if not isinstance(node, Tag):
        return
    if node.name.lower() == "img":
        add_docx_image(paragraph, node)
        return
    if node.name.lower() == "br":
        paragraph.add_run().add_break()
        return

    run = paragraph.add_run()
    if node.name.lower() in ("b", "strong"):
        run.bold = True
    elif node.name.lower() in ("i", "em"):
        run.italic = True
    elif node.name.lower() == "u":
        run.underline = True

    # Put direct text/children into their own runs so basic formatting is retained.
    for child in node.children:
        add_docx_inline(paragraph, child)


def build_docx(title, html_content):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    heading = doc.add_paragraph()
    run = heading.add_run(title or "Document Title")
    run.bold = True
    run.font.size = Pt(13)
    run.font.color.rgb = RGBColor(25, 25, 25)

    soup = BeautifulSoup(html_content or "", "html.parser")
    root = soup.body if soup.body else soup

    def process_docx_node(node):
        if isinstance(node, NavigableString):
            if str(node).strip():
                p = doc.add_paragraph()
                p.add_run(str(node))
            return
        if not isinstance(node, Tag):
            return
        name = node.name.lower()
        if name == "img":
            p = doc.add_paragraph()
            add_docx_image(p, node)
            return
        if name in ("p", "div", "h1", "h2", "h3", "li", "blockquote"):
            p = doc.add_paragraph()
            if name in ("h1", "h2", "h3"):
                for child in node.children:
                    add_docx_inline(p, child)
                for r in p.runs:
                    r.bold = True
                    r.font.size = Pt(12 if name == "h1" else 10)
            else:
                for child in node.children:
                    add_docx_inline(p, child)
            return
        for child in node.children:
            process_docx_node(child)

    for child in root.children:
        process_docx_node(child)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/download", methods=["POST"])
def download_document():
    title = request.form.get("title", "Document Title").strip() or "Document Title"
    html_content = request.form.get("content_html", "")
    file_format = request.form.get("format", "pdf").lower()

    if file_format == "pdf":
        buffer = build_pdf(title, html_content)
        return send_file(
            buffer, as_attachment=True,
            download_name="Smart_Professional_Document.pdf",
            mimetype="application/pdf"
        )
    if file_format in ("word", "docx"):
        buffer = build_docx(title, html_content)
        return send_file(
            buffer, as_attachment=True,
            download_name="Smart_Professional_Document.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    abort(400, "Unsupported download format")


if __name__ == "__main__":
    app.run(debug=True)
