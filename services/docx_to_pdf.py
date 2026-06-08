"""
Converte DOCX → PDF usando mammoth (DOCX→HTML) + fpdf2 (HTML→PDF).
100% Python puro — sem dependências de sistema (LibreOffice, Cairo, etc).
"""
import io
import mammoth
from fpdf import FPDF
from bs4 import BeautifulSoup, NavigableString


class _PDF(FPDF):
    def header(self):
        pass

    def footer(self):
        self.set_y(-12)
        self.set_font("Helvetica", size=8)
        self.set_text_color(120)
        self.cell(0, 6, f"Página {self.page_no()}", align="C")
        self.set_text_color(0)


def _render_elem(pdf: FPDF, elem, base_size: int = 11):
    """Renderiza recursivamente um elemento BeautifulSoup no PDF."""
    if isinstance(elem, NavigableString):
        return

    tag = elem.name

    if tag in ("h1", "h2", "h3", "h4"):
        level = int(tag[1])
        size = max(14 - (level - 1) * 2, 10)
        pdf.ln(3)
        pdf.set_font("Helvetica", style="B", size=size)
        pdf.multi_cell(0, 7, elem.get_text().strip())
        pdf.ln(1)
        pdf.set_font("Helvetica", size=base_size)

    elif tag == "p":
        text = elem.get_text().strip()
        if not text:
            pdf.ln(3)
            return
        # Detecta negrito dominante no parágrafo
        bold_chars = sum(len(b.get_text()) for b in elem.find_all("strong"))
        is_bold = bold_chars > len(text) * 0.5
        pdf.set_font("Helvetica", style="B" if is_bold else "", size=base_size)
        pdf.multi_cell(0, 6, text)
        pdf.ln(1)

    elif tag in ("ul", "ol"):
        for i, li in enumerate(elem.find_all("li", recursive=False)):
            prefix = f"{i + 1}." if tag == "ol" else "•"
            pdf.set_font("Helvetica", size=base_size)
            pdf.multi_cell(0, 6, f"   {prefix}  {li.get_text().strip()}")
        pdf.ln(1)

    elif tag == "table":
        rows = elem.find_all("tr")
        if not rows:
            return
        pdf.ln(2)
        usable_w = pdf.w - pdf.l_margin - pdf.r_margin
        n_cols = max(len(r.find_all(["td", "th"])) for r in rows)
        if n_cols == 0:
            return
        col_w = usable_w / n_cols

        for r_idx, row in enumerate(rows):
            cells = row.find_all(["td", "th"])
            is_header = r_idx == 0 or any(c.name == "th" for c in cells)
            line_h = 6
            # Calcula altura máxima da linha
            row_texts = [c.get_text().strip() for c in cells]

            x0 = pdf.l_margin
            y0 = pdf.get_y()

            # Verifica se cabe na página
            if y0 + line_h > pdf.h - pdf.b_margin:
                pdf.add_page()
                y0 = pdf.get_y()

            for c_idx, (cell, text) in enumerate(zip(cells, row_texts)):
                pdf.set_xy(x0 + c_idx * col_w, y0)
                pdf.set_font("Helvetica", style="B" if is_header else "", size=9)
                pdf.multi_cell(col_w, line_h, text, border=1)

            # Avança para a linha mais alta entre as células
            pdf.set_xy(x0, y0 + line_h)

        pdf.ln(3)

    elif tag in ("div", "section", "article", "body"):
        for child in elem.children:
            _render_elem(pdf, child, base_size)

    elif tag == "br":
        pdf.ln(3)


def docx_bytes_to_pdf(docx_bytes: bytes) -> bytes:
    """Recebe bytes de um DOCX e retorna bytes de PDF."""
    # 1. DOCX → HTML via mammoth
    result = mammoth.convert_to_html(io.BytesIO(docx_bytes))
    soup = BeautifulSoup(result.value, "html.parser")

    # 2. HTML → PDF via fpdf2
    pdf = _PDF(orientation="P", unit="mm", format="A4")
    pdf.set_margins(left=25, top=20, right=25)
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()
    pdf.set_font("Helvetica", size=11)

    for elem in soup.children:
        _render_elem(pdf, elem)

    return bytes(pdf.output())
