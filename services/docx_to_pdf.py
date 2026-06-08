"""
Converte DOCX → PDF usando mammoth (DOCX→HTML) + xhtml2pdf (HTML→PDF).
Funciona em qualquer ambiente Python sem dependências de sistema (LibreOffice, Word).
"""
import io
import mammoth
from xhtml2pdf import pisa


_CSS = """
@page {
    margin: 2cm 2.5cm;
}
body {
    font-family: Arial, Helvetica, sans-serif;
    font-size: 11pt;
    line-height: 1.5;
    color: #000;
}
p {
    margin: 0 0 6pt 0;
}
h1, h2, h3, h4 {
    margin: 12pt 0 4pt 0;
    font-weight: bold;
}
table {
    border-collapse: collapse;
    width: 100%;
    margin-bottom: 8pt;
}
td, th {
    border: 1px solid #999;
    padding: 4px 8px;
    vertical-align: top;
}
strong { font-weight: bold; }
em     { font-style: italic; }
"""


def docx_bytes_to_pdf(docx_bytes: bytes) -> bytes:
    """Recebe bytes de um DOCX e retorna bytes de PDF."""
    # 1. DOCX → HTML
    result = mammoth.convert_to_html(io.BytesIO(docx_bytes))
    html_body = result.value

    full_html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>{_CSS}</style>
</head>
<body>{html_body}</body>
</html>"""

    # 2. HTML → PDF
    buf = io.BytesIO()
    status = pisa.CreatePDF(full_html, dest=buf, encoding="utf-8")
    if status.err:
        raise RuntimeError(f"xhtml2pdf: erro ao gerar PDF (código {status.err})")
    return buf.getvalue()
