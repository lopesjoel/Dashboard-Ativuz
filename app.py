from flask import (
    Flask, render_template, request, redirect,
    url_for, send_file, flash, jsonify, abort, session
)
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
import os as _os
from dotenv import load_dotenv
load_dotenv()
from pathlib import Path
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo

_BRT = ZoneInfo("America/Sao_Paulo")
from io import BytesIO
import json
import platform
import re
import subprocess
import unicodedata
import uuid

from gerar_contrato import gerar_docx, gerar_termo_quitacao, gerar_notificacao_avalista, gerar_notificacao_inadimplente, gerar_vistoria_entrega, gerar_vistoria_nova, nome_arquivo_saida

app = Flask(__name__)
app.secret_key = _os.environ.get("SECRET_KEY", "ativuz-secret-dev-2026")


@app.errorhandler(Exception)
def handle_any_error(e):
    import traceback; traceback.print_exc()
    return jsonify({"error": str(e)}), 500


# ── Template filters ──────────────────────────────────────────────────────────

@app.template_filter('brl')
def _fmt_brl(v):
    if v is None:
        return '—'
    neg = v < 0
    s = f"{abs(v):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    return f"({s})" if neg else s


@app.template_filter('pct_fmt')
def _fmt_pct(v):
    if v is None:
        return '—'
    return f"{v * 100:.1f}%"


# ── Autenticação ──────────────────────────────────────────────────────────────

_ROTAS_PUBLICAS = {"login", "static", "admin_novo_usuario"}

@app.before_request
def verificar_login():
    if request.endpoint in _ROTAS_PUBLICAS:
        return
    if not session.get("usuario"):
        return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("usuario"):
        return redirect(url_for("dashboard"))
    erro = None
    if request.method == "POST":
        nome  = request.form.get("nome", "").strip()
        senha = request.form.get("senha", "")
        sb = _supabase()
        if not sb:
            erro = "Serviço indisponível. Tente novamente."
        else:
            try:
                from supabase import create_client
                url = _os.environ.get("SUPABASE_URL", "")
                key = _os.environ.get("SUPABASE_KEY", "")
                email = (nome if "@" in nome else f"{nome}@ativuz.com").lower()
                auth_client = create_client(url, key)
                res = auth_client.auth.sign_in_with_password({"email": email, "password": senha})
                session["usuario"] = nome
                return redirect(url_for("dashboard"))
            except Exception:
                erro = "Nome ou senha incorretos."
    return render_template("login.html", erro=erro)


@app.route("/logout")
def logout():
    sb = _supabase()
    if sb:
        try:
            sb.auth.sign_out()
        except Exception:
            pass
    session.clear()
    return redirect(url_for("login"))


@app.route("/api/clientes")
def api_clientes():
    import openpyxl
    path = Path(__file__).parent / "dados_clientes.xlsx"
    if not path.exists():
        return jsonify([])
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if len(rows) < 2:
        return jsonify([])
    # Mapeamento dinâmico por nome de coluna (case-insensitive, ignora acentos)
    import unicodedata
    def _norm(s):
        s = unicodedata.normalize("NFD", str(s or "").lower())
        return "".join(c for c in s if unicodedata.category(c) != "Mn")
    headers = [_norm(h) for h in rows[0]]
    def _col(name):
        n = _norm(name)
        return next((i for i, h in enumerate(headers) if n in h), None)
    i_nome    = _col("cliente")
    i_tel     = _col("telefone")
    i_ano     = _col("ano")
    i_chassi  = _col("chassi")
    i_cor     = _col("cor")
    i_marca   = _col("marca")
    i_modelo  = _col("modelo")
    i_placa   = _col("placa")
    i_end     = _col("endereco")
    i_motor   = _col("motor")
    def _v(row, i): return str(row[i] or "") if i is not None and i < len(row) else ""
    q = request.args.get("q", "").lower().strip()
    clientes = []
    for row in rows[1:]:
        if not (i_nome is not None and i_nome < len(row) and row[i_nome]):
            continue
        nome = str(row[i_nome])
        if q and q not in nome.lower():
            continue
        marca  = _v(row, i_marca)
        modelo = _v(row, i_modelo)
        clientes.append({
            "nome":         nome,
            "telefone":     _v(row, i_tel),
            "endereco":     _v(row, i_end),
            "veiculo":      f"{marca} {modelo}".strip(),
            "placa":        _v(row, i_placa),
            "cor":          _v(row, i_cor),
            "ano":          str(int(row[i_ano])) if i_ano is not None and i_ano < len(row) and row[i_ano] else "",
            "chassi":       _v(row, i_chassi),
            "numero_motor": _v(row, i_motor),
        })
    return jsonify(clientes[:30])


@app.route("/admin/novo-usuario", methods=["GET", "POST"])
def admin_novo_usuario():
    token_correto = _os.environ.get("ADMIN_TOKEN", "")
    token = request.args.get("token", "")
    if not token_correto or token != token_correto:
        abort(403)
    mensagem = None
    erro = None
    if request.method == "POST":
        nome  = request.form.get("nome", "").strip()
        senha = request.form.get("senha", "")
        if not nome or not senha:
            erro = "Nome e senha são obrigatórios."
        else:
            sb = _supabase()
            if not sb:
                erro = "Supabase não configurado."
            else:
                try:
                    sb.table("usuarios").insert({
                        "nome": nome,
                        "senha_hash": generate_password_hash(senha),
                        "ativo": True,
                    }).execute()
                    mensagem = f"Usuário '{nome}' criado com sucesso!"
                except Exception as exc:
                    erro = f"Erro ao criar usuário: {exc}"
    return f"""
    <!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">
    <title>Novo Usuário — Admin</title>
    <style>body{{font-family:Inter,sans-serif;background:#f0f2f7;display:flex;
    align-items:center;justify-content:center;min-height:100vh;}}
    .card{{background:#fff;border-radius:14px;padding:2rem;width:360px;
    box-shadow:0 4px 20px rgba(0,0,0,.1);}}
    h1{{font-size:1.1rem;margin-bottom:1.5rem;}}
    label{{font-size:.75rem;font-weight:600;text-transform:uppercase;
    letter-spacing:.05em;color:#475569;display:block;margin-bottom:.3rem;}}
    input{{width:100%;padding:.6rem .8rem;border:1.5px solid #e2e8f0;border-radius:8px;
    font-family:inherit;font-size:.9rem;margin-bottom:1rem;outline:none;}}
    input:focus{{border-color:#4361ee;}}
    button{{width:100%;padding:.7rem;background:#4361ee;color:#fff;border:none;
    border-radius:8px;font-weight:700;font-size:.9rem;cursor:pointer;}}
    .ok{{color:#166534;background:#f0fdf4;border:1px solid #bbf7d0;
    border-radius:8px;padding:.6rem .9rem;font-size:.85rem;margin-bottom:1rem;}}
    .err{{color:#991b1b;background:#fef2f2;border:1px solid #fecaca;
    border-radius:8px;padding:.6rem .9rem;font-size:.85rem;margin-bottom:1rem;}}
    </style></head><body><div class="card">
    <h1>Criar novo usuário</h1>
    {"<div class='ok'>"+mensagem+"</div>" if mensagem else ""}
    {"<div class='err'>"+erro+"</div>" if erro else ""}
    <form method="POST" action="?token={token}">
    <label>Nome</label><input name="nome" required>
    <label>Senha</label><input type="password" name="senha" required>
    <button>Criar usuário</button></form></div></body></html>
    """

# ── Supabase (opcional — só ativa se as env vars estiverem definidas) ─────────

_sb = None

def _supabase():
    global _sb
    if _sb is None:
        url = _os.environ.get("SUPABASE_URL", "")
        key = _os.environ.get("SUPABASE_KEY", "")
        if url and key:
            from supabase import create_client
            _sb = create_client(url, key)
    return _sb

UPLOAD_FOLDER = Path("uploads")
CONTRATOS_FOLDER = Path("contratos")
TEMP_FOLDER = Path("temp_preview")
DOCX_TEMPLATES = Path("docx_templates")

UPLOAD_FOLDER.mkdir(exist_ok=True)
CONTRATOS_FOLDER.mkdir(exist_ok=True)
TEMP_FOLDER.mkdir(exist_ok=True)
DOCX_TEMPLATES.mkdir(exist_ok=True)


# ── helpers ───────────────────────────────────────────────

def _converter_pdf(caminho_docx: str, caminho_pdf: str):
    """Converte .docx para PDF: Word no Windows, LibreOffice no Linux."""
    if platform.system() == "Windows":
        import pythoncom
        from docx2pdf import convert
        pythoncom.CoInitialize()
        try:
            convert(caminho_docx, caminho_pdf)
        finally:
            pythoncom.CoUninitialize()
    else:
        import tempfile, os, shutil
        docx_abs  = str(Path(caminho_docx).resolve())
        pdf_abs   = str(Path(caminho_pdf).resolve())
        if not Path(docx_abs).exists():
            raise FileNotFoundError(f"DOCX não encontrado: {docx_abs!r}")
        with tempfile.TemporaryDirectory() as work_dir:
            # copia para /tmp isolado — evita problemas de permissão/path no LO
            tmp_docx = Path(work_dir) / Path(docx_abs).name
            shutil.copy2(docx_abs, tmp_docx)
            env = {**os.environ, "HOME": work_dir}
            result = subprocess.run(
                [
                    "libreoffice",
                    "--headless", "--norestore", "--nofirststartwizard",
                    "--convert-to", "pdf",
                    "--outdir", work_dir,
                    str(tmp_docx),
                ],
                capture_output=True,
                env=env,
            )
            gerado = Path(work_dir) / (tmp_docx.stem + ".pdf")
            if not gerado.exists():
                stderr = result.stderr.decode(errors="replace") if result.stderr else ""
                stdout = result.stdout.decode(errors="replace") if result.stdout else ""
                raise RuntimeError(
                    f"LibreOffice (exit {result.returncode}) não gerou PDF. "
                    f"docx={str(tmp_docx)!r} stderr={stderr!r} stdout={stdout!r}"
                )
            shutil.copy2(gerado, pdf_abs)


def _slugify(texto: str) -> str:
    """Maiúsculas, sem acentos, espaços → underscore, sem caracteres especiais."""
    norm = unicodedata.normalize('NFD', texto)
    norm = ''.join(c for c in norm if unicodedata.category(c) != 'Mn')
    norm = norm.upper().strip()
    norm = re.sub(r'[^A-Z0-9\s]', '', norm)
    norm = re.sub(r'\s+', '_', norm)
    return norm or "SEM_NOME"


def detectar_tipo(filename: str):
    """Retorna o tipo do template com base no nome do arquivo."""
    norm = unicodedata.normalize('NFD', filename.lower())
    norm = ''.join(c for c in norm if unicodedata.category(c) != 'Mn')
    if 'quitacao' in norm:
        return 'quitacao'
    if 'locacao' in norm:
        return 'locacao'
    if 'notificacao' in norm and 'inadimplente' in norm:
        return 'inadimplente'
    if 'notificacao' in norm:
        return 'notificacao'
    return None


def get_templates():
    result = []
    files = sorted(
        list(UPLOAD_FOLDER.glob("*.docx")) + list(UPLOAD_FOLDER.glob("*.xlsx")),
        key=lambda f: f.name,
    )
    for f in files:
        meta_path = UPLOAD_FOLDER / f"{f.stem}.json"
        display_name = f.stem
        if meta_path.exists():
            display_name = json.loads(meta_path.read_text(encoding="utf-8")).get("nome", f.stem)
        result.append({
            "filename": f.name,
            "nome": display_name,
            "tamanho_kb": round(f.stat().st_size / 1024, 1),
            "data": datetime.fromtimestamp(f.stat().st_mtime).strftime("%d/%m/%Y %H:%M"),
        })
    return result


def _historico_append(locatario_nome: str, template: str, arquivo: str):
    sb = _supabase()
    if not sb:
        return
    try:
        sb.table("historico_docs").insert({
            "locatario_nome": locatario_nome,
            "template": template,
            "arquivo": arquivo,
            "data_hora": datetime.now(_BRT).strftime("%d/%m/%Y %H:%M"),
        }).execute()
    except Exception:
        import traceback; traceback.print_exc()


def _gerar_para_caminho(form, tipo, template_path_str, caminho_saida):
    """Gera o documento para caminho_saida. Retorna nome_pessoa."""
    if tipo == "locacao":
        campos = [
            "locatario_nome", "locatario_rg", "locatario_cpf",
            "locatario_endereco", "locatario_cep", "locatario_telefone",
            "avalista_nome", "avalista_cpf", "avalista_endereco", "avalista_telefone",
            "veiculo_descricao", "veiculo_marca", "veiculo_modelo", "veiculo_ano",
            "veiculo_motor", "veiculo_chassi", "veiculo_cor", "veiculo_placa",
            "contrato_inicio", "contrato_duracao", "valor_semanal",
            "data_dia", "data_mes", "data_ano",
            "testemunha1_nome", "testemunha1_rg", "testemunha1_cpf",
            "testemunha2_nome", "testemunha2_rg", "testemunha2_cpf",
        ]
        dados = {c: form.get(c, "") for c in campos}
        gerar_docx(dados, caminho_saida, template_path=template_path_str)
        return dados["locatario_nome"]

    elif tipo == "notificacao":
        avalista_nome = form.get("avalista_nome_notif", "")
        gerar_notificacao_avalista(
            avalista_nome  = avalista_nome,
            data_contrato  = form.get("data_contrato", ""),
            locatario_nome = form.get("locatario_nome_notif", ""),
            valor_debito   = float(form.get("valor_debito") or 0),
            caminho_saida  = caminho_saida,
            template_path  = template_path_str,
        )
        return avalista_nome

    elif tipo == "inadimplente":
        locatario_nome_inad = form.get("locatario_nome_inad", "")
        gerar_notificacao_inadimplente(
            locatario_nome = locatario_nome_inad,
            data_contrato  = form.get("data_contrato_inad", ""),
            valor_debito   = float(form.get("valor_debito_inad") or 0),
            caminho_saida  = caminho_saida,
            template_path  = template_path_str,
        )
        return locatario_nome_inad

    else:  # quitacao
        def _f(campo): return float(form.get(campo) or 0)
        def _i(campo): return int(float(form.get(campo) or 0))
        devedor_nome = form.get("devedor_nome", "")
        gerar_termo_quitacao(
            devedor_nome          = devedor_nome,
            devedor_cpf           = form.get("devedor_cpf", ""),
            placa                 = form.get("placa", ""),
            mes_referencia_fipe   = form.get("mes_referencia_fipe", ""),
            valor_fipe            = _f("valor_fipe"),
            percentual_fipe       = _f("percentual_fipe"),
            meias_diarias         = _f("meias_diarias"),
            entrada               = _f("entrada"),
            num_parcelas_pagas    = _i("num_parcelas_pagas"),
            valor_parcela_paga    = _f("valor_parcela_paga"),
            num_parcelas_semanais = _i("num_parcelas_semanais"),
            valor_parcela_semanal = _f("valor_parcela_semanal"),
            data_primeira_parcela = form.get("data_primeira_parcela", ""),
            data_assinatura       = form.get("data_assinatura", ""),
            caminho_saida         = caminho_saida,
            template_path         = template_path_str,
        )
        return devedor_nome


# ── página 1 — Templates ──────────────────────────────────

@app.route("/")
def dashboard():
    sb = _supabase()
    total_contratos = 0
    total_vistorias = 0
    total_docs = 0
    valor_mensal = "—"
    contratos = []
    if sb:
        try:
            res = sb.table("contratos_locacao").select(
                "id, locatario_nome, veiculo_placa, veiculo_marca, veiculo_modelo, contrato_inicio, valor_semanal",
                count="exact"
            ).order("criado_em", desc=True).limit(5).execute()
            contratos = res.data or []
            total_contratos = res.count or len(contratos)
        except Exception:
            pass
        try:
            rv = sb.table("vistorias").select("id", count="exact").execute()
            total_vistorias = rv.count or 0
        except Exception:
            pass
    if sb:
        try:
            rd = sb.table("historico_docs").select("id", count="exact").eq("deletado", False).execute()
            total_docs = rd.count or 0
        except Exception:
            pass
    inad = _inad_summary()
    return render_template(
        "dashboard.html",
        active="dashboard",
        total_contratos=total_contratos,
        total_vistorias=total_vistorias,
        total_docs=total_docs,
        valor_mensal=valor_mensal,
        contratos=contratos,
        inad=inad,
    )


@app.route("/templates")
def pagina_templates():
    return render_template("templates.html", templates=get_templates(), active="templates")


@app.route("/upload", methods=["POST"])
def upload_template():
    nome = request.form.get("nome", "").strip()
    arquivo = request.files.get("arquivo")

    if not nome:
        flash("Informe um nome para o template.", "erro")
        return redirect(url_for("pagina_templates"))

    if not arquivo or arquivo.filename == "":
        flash("Selecione um arquivo .docx.", "erro")
        return redirect(url_for("pagina_templates"))

    if not arquivo.filename.lower().endswith((".docx", ".xlsx")):
        flash("Apenas arquivos .docx ou .xlsx são aceitos.", "erro")
        return redirect(url_for("pagina_templates"))

    uid = uuid.uuid4().hex[:8]
    safe_stem = secure_filename(f"{nome}_{uid}")
    dest = UPLOAD_FOLDER / f"{safe_stem}.docx"
    arquivo.save(str(dest))

    meta = UPLOAD_FOLDER / f"{safe_stem}.json"
    meta.write_text(json.dumps({"nome": nome}, ensure_ascii=False), encoding="utf-8")

    flash(f'Template "{nome}" enviado com sucesso!', "ok")
    return redirect(url_for("pagina_templates"))


@app.route("/templates/excluir/<filename>", methods=["POST"])
def excluir_template(filename):
    safe = secure_filename(filename)
    docx = UPLOAD_FOLDER / safe
    meta = UPLOAD_FOLDER / f"{Path(safe).stem}.json"
    if docx.exists():
        docx.unlink()
    if meta.exists():
        meta.unlink()
    flash("Template excluído.", "ok")
    return redirect(url_for("pagina_templates"))


# ── página 2 — Gerar Contrato ─────────────────────────────

@app.route("/gerar")
def pagina_gerar():
    return render_template("gerar.html", templates=get_templates(), active="gerar")


@app.route("/gerar-contrato", methods=["POST"])
def gerar_contrato_route():
    template_filename = request.form.get("template", "")
    if not template_filename:
        flash("Selecione um template.", "erro")
        return redirect(url_for("pagina_gerar"))

    template_path = UPLOAD_FOLDER / secure_filename(template_filename)
    if not template_path.exists():
        flash("Template não encontrado.", "erro")
        return redirect(url_for("pagina_gerar"))

    tipo = detectar_tipo(template_filename)
    if tipo is None:
        return jsonify({
            "error": "Template não reconhecido. Renomeie o arquivo com 'locacao', 'quitacao', 'notificacao' ou 'inadimplente' no nome."
        }), 400

    formato = request.form.get("formato", "docx")

    # ── Nome do arquivo de saída ──────────────────────────
    if tipo == "locacao":
        ano        = datetime.now().strftime("%Y")
        nome_saida = f"{ano}_{_slugify(request.form.get('veiculo_placa', ''))}_{_slugify(request.form.get('locatario_nome', ''))}.docx"
    elif tipo == "notificacao":
        data_slug  = datetime.now().strftime("%d.%m.%Y")
        nome_saida = f"NOTIFICACAO_AVALISTA_{_slugify(request.form.get('avalista_nome_notif', ''))}_{data_slug}.docx"
    elif tipo == "inadimplente":
        data_slug  = datetime.now().strftime("%d.%m.%Y")
        nome_saida = f"NOTIFICACAO_INADIMPLENTE_{_slugify(request.form.get('locatario_nome_inad', ''))}_{data_slug}.docx"
    else:  # quitacao
        data_slug  = datetime.now().strftime("%d.%m.%Y")
        nome_saida = f"QUITACAO_DIVIDA_{_slugify(request.form.get('devedor_nome', ''))}_{data_slug}.docx"

    caminho_saida = str(CONTRATOS_FOLDER / nome_saida)

    # ── Gerar documento ───────────────────────────────────
    try:
        nome_pessoa = _gerar_para_caminho(request.form, tipo, str(template_path), caminho_saida)
    except Exception as e:
        return jsonify({"error": f"Erro ao gerar contrato: {e}"}), 500

    # ── Histórico local ───────────────────────────────────
    meta_path = UPLOAD_FOLDER / f"{template_path.stem}.json"
    nome_template = template_filename
    if meta_path.exists():
        try:
            nome_template = json.loads(meta_path.read_text(encoding="utf-8")).get("nome", template_filename)
        except Exception:
            pass
    try:
        _historico_append(nome_pessoa, nome_template, nome_saida)
    except Exception:
        pass

    # ── Download direto ────────────────────────────────────
    if formato == "pdf":
        nome_pdf    = nome_saida.replace(".docx", ".pdf")
        caminho_pdf = str(CONTRATOS_FOLDER / nome_pdf)
        try:
            _converter_pdf(caminho_saida, caminho_pdf)
            pdf_bytes = BytesIO(Path(caminho_pdf).read_bytes())
            return send_file(
                pdf_bytes,
                as_attachment=True,
                download_name=nome_pdf,
                mimetype="application/pdf",
            )
        except Exception as e:
            docx_url = url_for("download_contrato", filename=nome_saida)
            return jsonify({
                "error": f"Erro ao gerar PDF: {e}",
                "docx_url": docx_url,
            }), 422

    return send_file(
        caminho_saida,
        as_attachment=True,
        download_name=nome_saida,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.route("/preview-contrato", methods=["POST"])
def preview_contrato():
    import mammoth

    template_filename = request.form.get("template", "")
    if not template_filename:
        return jsonify({"error": "Selecione um template."}), 400

    template_path = UPLOAD_FOLDER / secure_filename(template_filename)
    if not template_path.exists():
        return jsonify({"error": "Template não encontrado."}), 400

    tipo = detectar_tipo(template_filename)
    if tipo is None:
        return jsonify({"error": "Template não reconhecido."}), 400
    if tipo == "vistoria":
        return jsonify({"error": "Pré-visualização não disponível para vistoria (formato .xlsx)."}), 400

    temp_id = uuid.uuid4().hex
    caminho_temp = str(TEMP_FOLDER / f"{temp_id}.docx")

    try:
        _gerar_para_caminho(request.form, tipo, str(template_path), caminho_temp)
    except Exception as e:
        return jsonify({"error": f"Erro ao gerar pré-visualização: {e}"}), 500

    try:
        with open(caminho_temp, "rb") as f:
            result = mammoth.convert_to_html(f)
        html = result.value
    except Exception as e:
        Path(caminho_temp).unlink(missing_ok=True)
        return jsonify({"error": f"Erro ao converter para HTML: {e}"}), 500

    return jsonify({"html": html, "temp_id": temp_id})


@app.route("/cleanup-temp/<temp_id>", methods=["POST"])
def cleanup_temp(temp_id):
    if not re.match(r'^[0-9a-f]{32}$', temp_id):
        abort(400)
    caminho = TEMP_FOLDER / f"{temp_id}.docx"
    if caminho.exists():
        caminho.unlink()
    return jsonify({"ok": True})


# ── página 3 — Histórico ──────────────────────────────────

@app.route("/historico")
def pagina_historico():
    sb = _supabase()
    historico = []
    if sb:
        try:
            res = sb.table("historico_docs").select("*") \
                .eq("deletado", False) \
                .order("criado_em", desc=True).execute()
            historico = res.data or []
        except Exception:
            pass
    return render_template("historico.html", historico=historico, active="historico")


@app.route("/historico/download/<path:filename>")
def download_contrato(filename):
    caminho = (CONTRATOS_FOLDER / filename).resolve()
    if not str(caminho).startswith(str(CONTRATOS_FOLDER.resolve())):
        abort(400)
    if not caminho.exists():
        flash("Arquivo não encontrado.", "erro")
        return redirect(url_for("pagina_historico"))
    ext = Path(filename).suffix.lower()
    mime = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        if ext == ".xlsx"
        else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    return send_file(
        str(caminho),
        as_attachment=True,
        download_name=Path(filename).name,
        mimetype=mime,
    )


@app.route("/historico/download-pdf/<path:filename>")
def download_contrato_pdf(filename):
    caminho_docx = (CONTRATOS_FOLDER / filename).resolve()
    if not str(caminho_docx).startswith(str(CONTRATOS_FOLDER.resolve())):
        abort(400)
    if not caminho_docx.exists():
        return jsonify({"error": "Arquivo não encontrado."}), 404

    nome_pdf    = Path(filename).stem + ".pdf"
    caminho_pdf = CONTRATOS_FOLDER / nome_pdf
    try:
        _converter_pdf(str(caminho_docx), str(caminho_pdf))
        pdf_bytes = BytesIO(Path(caminho_pdf).read_bytes())
        return send_file(
            pdf_bytes,
            as_attachment=True,
            download_name=nome_pdf,
            mimetype="application/pdf",
        )
    except Exception as e:
        return jsonify({"error": f"Erro ao gerar PDF: {e}"}), 422


@app.route("/historico/excluir/<entry_id>", methods=["POST"])
def excluir_contrato(entry_id):
    sb = _supabase()
    if sb:
        try:
            sb.table("historico_docs").update({"deletado": True}).eq("id", entry_id).execute()
        except Exception:
            import traceback; traceback.print_exc()
    return jsonify({"ok": True})


@app.route("/historico/exportar-excel")
def exportar_historico_excel():
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment

    sb = _supabase()
    historico = []
    if sb:
        try:
            res = sb.table("historico_docs").select("*") \
                .eq("deletado", False) \
                .order("criado_em", desc=True).execute()
            historico = res.data or []
        except Exception:
            pass
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Histórico"

    cabecalho = ["Locatário", "Template", "Data / Hora", "Nome do Arquivo"]
    ws.append(cabecalho)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1E3A5F")
        cell.alignment = Alignment(horizontal="center")

    for item in historico:
        ws.append([
            item.get("locatario_nome", ""),
            item.get("template", ""),
            item.get("data_hora", ""),
            item.get("arquivo", ""),
        ])

    for col in ws.columns:
        max_len = max((len(str(cell.value or "")) for cell in col), default=0)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    data_hoje = datetime.now().strftime("%d-%m-%Y")
    nome_arquivo = f"HISTORICO_ATIVUZ_{data_hoje}.xlsx"

    return send_file(
        buf,
        as_attachment=True,
        download_name=nome_arquivo,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ── Contrato de Locação — helpers e rotas ────────────────────────────────────

CONTRATO_LOCACAO_TEMPLATE = DOCX_TEMPLATES / "CONTRATO DE LOCAÇÃO EDITADO.docx"


def _salvar_contrato_locacao(insert: dict, caminho_docx: str, storage_path: str, edit_id: str = None):
    """INSERT no Supabase (main thread) + upload do arquivo (background)."""
    import threading, traceback as _tb
    sb = _supabase()
    if not sb:
        return
    try:
        sb.table("contratos_locacao").insert(insert).execute()
    except Exception:
        _tb.print_exc()
    if edit_id:
        try:
            old = sb.table("contratos_locacao").select("arquivo_path").eq("id", edit_id).single().execute()
            _old_path = (old.data or {}).get("arquivo_path")
            sb.table("contratos_locacao").delete().eq("id", edit_id).execute()
        except Exception:
            _old_path = None
            _tb.print_exc()
    else:
        _old_path = None

    _docx_bytes = Path(caminho_docx).read_bytes()
    _sp = storage_path
    _op = _old_path

    def _bg():
        try:
            sb2 = _supabase()
            if not sb2:
                return
            try:
                sb2.storage.from_("documentos").upload(
                    _sp, _docx_bytes,
                    {"content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                     "upsert": "true"},
                )
            except Exception:
                _tb.print_exc()
            if _op and _op != _sp:
                try:
                    sb2.storage.from_("documentos").remove([_op])
                except Exception:
                    pass
        except Exception:
            _tb.print_exc()

    threading.Thread(target=_bg, daemon=True).start()

_MESES_PT = ["janeiro","fevereiro","março","abril","maio","junho",
             "julho","agosto","setembro","outubro","novembro","dezembro"]


@app.route("/contrato-locacao")
def pagina_contrato_locacao():
    agora = datetime.now(_BRT)
    defaults = {
        "data_dia": agora.strftime("%d"),
        "data_mes": _MESES_PT[agora.month - 1],
        "data_ano": agora.strftime("%Y"),
    }
    return render_template("contrato_locacao.html", defaults=defaults, active="contrato_locacao")


@app.route("/contrato-locacao/gerar", methods=["POST"])
def gerar_contrato_locacao_route():
    """Usado pelo fluxo de edição a partir de historico_contratos."""
    if not CONTRATO_LOCACAO_TEMPLATE.exists():
        return jsonify({"error": "Template não encontrado em docx_templates/."}), 404

    campos = [
        "locatario_nome", "locatario_rg", "locatario_cpf",
        "locatario_endereco", "locatario_cep", "locatario_telefone",
        "avalista_nome", "avalista_cpf", "avalista_endereco", "avalista_telefone",
        "veiculo_descricao", "veiculo_marca", "veiculo_modelo", "veiculo_ano",
        "veiculo_motor", "veiculo_chassi", "veiculo_cor", "veiculo_placa",
        "contrato_inicio", "contrato_duracao", "valor_semanal",
        "data_dia", "data_mes", "data_ano",
        "testemunha1_nome", "testemunha1_rg", "testemunha1_cpf",
        "testemunha2_nome", "testemunha2_rg", "testemunha2_cpf",
    ]
    dados   = {c: request.form.get(c, "") for c in campos}
    edit_id = request.form.get("edit_id", "").strip()

    placa_slug    = _slugify(dados.get("veiculo_placa") or "PLACA")
    nome_slug     = _slugify((dados.get("locatario_nome") or "LOCATARIO").split()[0])
    data_slug     = datetime.now(_BRT).strftime("%d.%m.%Y")
    nome_docx     = f"CONTRATO_LOCACAO_{placa_slug}_{nome_slug}_{data_slug}.docx"
    caminho_saida = str(CONTRATOS_FOLDER / nome_docx)

    try:
        gerar_docx(dados, caminho_saida, template_path=str(CONTRATO_LOCACAO_TEMPLATE))
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": f"Erro ao gerar contrato: {e}"}), 500

    _storage_path = f"contratos/{nome_docx}"
    _insert = {**dados, "arquivo_path": _storage_path}
    _salvar_contrato_locacao(_insert, caminho_saida, _storage_path, edit_id=edit_id or None)

    return jsonify({"redirect_url": url_for("historico_contratos")})


@app.route("/historico/contratos")
def historico_contratos():
    sb = _supabase()
    contratos = []
    erro = None
    if sb:
        try:
            res = sb.table("contratos_locacao").select(
                "id, locatario_nome, locatario_cpf, veiculo_placa, veiculo_marca, "
                "veiculo_modelo, contrato_inicio, valor_semanal, arquivo_path, criado_em"
            ).neq("deletado", True).order("criado_em", desc=True).execute()
            contratos = res.data or []
        except Exception as e:
            erro = str(e)
    else:
        erro = "Supabase não configurado."
    return render_template("historico_contratos.html", contratos=contratos, erro=erro,
                           active="hist_contratos")


@app.route("/historico/contratos/<contrato_id>/excluir", methods=["POST"])
def excluir_contrato_locacao(contrato_id):
    sb = _supabase()
    if sb:
        try:
            sb.table("contratos_locacao").update({"deletado": True}).eq("id", contrato_id).execute()
        except Exception:
            import traceback; traceback.print_exc()
    return jsonify({"ok": True})


@app.route("/historico/contratos/download/<contrato_id>")
def download_contrato_locacao_docx(contrato_id):
    sb = _supabase()
    if not sb:
        abort(503)
    try:
        res = sb.table("contratos_locacao").select("arquivo_path").eq("id", contrato_id).single().execute()
        path = res.data["arquivo_path"]
        signed = sb.storage.from_("documentos").create_signed_url(path, 60)
        return redirect(signed["signedURL"])
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/historico/contratos/download/<contrato_id>/pdf")
def download_contrato_locacao_pdf(contrato_id):
    sb = _supabase()
    if not sb:
        abort(503)
    try:
        res = sb.table("contratos_locacao").select("arquivo_path").eq("id", contrato_id).single().execute()
        docx_path = res.data["arquivo_path"]
        docx_bytes = sb.storage.from_("documentos").download(docx_path)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    if not isinstance(docx_bytes, (bytes, bytearray)):
        docx_bytes = getattr(docx_bytes, 'content', None) or bytes(docx_bytes)

    tmp_docx = TEMP_FOLDER / f"{uuid.uuid4().hex}.docx"
    tmp_pdf  = tmp_docx.with_suffix(".pdf")
    TEMP_FOLDER.mkdir(exist_ok=True)
    try:
        tmp_docx.write_bytes(docx_bytes)
        _converter_pdf(str(tmp_docx), str(tmp_pdf))
        if not tmp_pdf.exists():
            raise FileNotFoundError(f"LibreOffice não gerou o PDF")
        pdf_bytes = tmp_pdf.read_bytes()
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": f"Erro ao gerar PDF: {e}"}), 422
    finally:
        tmp_docx.unlink(missing_ok=True)
        tmp_pdf.unlink(missing_ok=True)

    nome_pdf = Path(docx_path).stem + ".pdf"
    return send_file(BytesIO(pdf_bytes), as_attachment=True,
                     download_name=nome_pdf, mimetype="application/pdf")


@app.route("/historico/contratos/<contrato_id>/editar")
def editar_contrato_locacao(contrato_id):
    sb = _supabase()
    if not sb:
        flash("Supabase não configurado.", "erro")
        return redirect(url_for("historico_contratos"))
    try:
        res = sb.table("contratos_locacao").select("*").eq("id", contrato_id).single().execute()
        contrato = res.data
    except Exception as e:
        flash(f"Erro ao buscar contrato: {e}", "erro")
        return redirect(url_for("historico_contratos"))
    agora = datetime.now(_BRT)
    defaults = {
        "data_dia": contrato.get("data_dia") or agora.strftime("%d"),
        "data_mes": contrato.get("data_mes") or _MESES_PT[agora.month - 1],
        "data_ano": contrato.get("data_ano") or agora.strftime("%Y"),
    }
    return render_template("contrato_locacao.html", contrato=contrato,
                           edit_id=contrato_id, defaults=defaults,
                           active="hist_contratos")


# ── Vistoria de Entrega ───────────────────────────────────────────────────────

VISTORIA_TEMPLATE = DOCX_TEMPLATES / "VISTORIA_TESTE_1.docx"


@app.route("/vistoria", methods=["GET"])
def pagina_vistoria():
    return render_template("vistoria.html", active="vistoria", vistoria=None, edit_id=None, acessorios={}, usuario=session.get("usuario", ""))


@app.route("/vistoria", methods=["POST"])
def processar_vistoria():
    if not VISTORIA_TEMPLATE.exists():
        flash("Template de vistoria não encontrado em docx_templates/.", "erro")
        return redirect(url_for("pagina_vistoria"))

    dados = {
        "cliente":                  request.form.get("cliente", ""),
        "tel":                      request.form.get("tel", ""),
        "preenchido_por":           request.form.get("preenchido_por", ""),
        "endereco":                 request.form.get("endereco", ""),
        "chassi":                   request.form.get("chassi", ""),
        "motor":                    request.form.get("motor", ""),
        "veiculo":                  request.form.get("veiculo", ""),
        "placa":                    request.form.get("placa", ""),
        "ano":                      request.form.get("ano", ""),
        "cor":                      request.form.get("cor", ""),
        "hodometro_entrega":        request.form.get("hodometro_entrega", ""),
        "hodometro_retorno":        request.form.get("hodometro_retorno", ""),
        "combustivel":              request.form.get("combustivel", ""),
        "data":                     datetime.now().strftime("%d/%m/%Y"),
        "danos":                    request.form.get("danos", ""),
        "observacoes":              request.form.get("observacoes", ""),
        "sintomas":                 request.form.get("sintomas", ""),
        "assinatura_cliente":       request.form.get("assinatura_cliente", ""),
        "assinatura_responsavel":   request.form.get("assinatura_responsavel", ""),
        # Acessórios (form fields sem prefixo, mapeados para chaves que gerar_contrato usa)
        "acc_calotas":          request.form.get("calotas", ""),
        "acc_buzina":           request.form.get("buzina", ""),
        "acc_doc_crlv":         request.form.get("doc_crlv", ""),
        "acc_triangulo":        request.form.get("triangulo", ""),
        "acc_antena":           request.form.get("antena", ""),
        "acc_sensor_re":        request.form.get("sensor_re", ""),
        "acc_som":              request.form.get("som", ""),
        "acc_tapetes":          request.form.get("tapetes", ""),
        "acc_limpadores":       request.form.get("limpadores", ""),
        "acc_chave_roda":       request.form.get("chave_roda", ""),
        "acc_vidros":           request.form.get("vidros", ""),
        "acc_oleo_motor":       request.form.get("oleo_motor", ""),
        "acc_alarme":           request.form.get("alarme", ""),
        "acc_lampadas":         request.form.get("lampadas", ""),
        "acc_macaco":           request.form.get("macaco", ""),
        "acc_estepe":           request.form.get("estepe", ""),
        "acc_gnv":              request.form.get("gnv", ""),
        "acc_agua":             request.form.get("agua", ""),
        "acc_borracha_psg_d":   request.form.get("borracha_psg_d", ""),
        "acc_borr_mtr":         request.form.get("borracha_mtr_d", ""),
        "acc_asa_urubu_dd":     request.form.get("asa_urubu_dd", ""),
        "acc_asa_urub_td":      request.form.get("asa_urubu_td", ""),
        "acc_tapete_mala":      request.form.get("tapete_mala", ""),
        "acc_tampa_prx":        request.form.get("tampa_paraxq", ""),
        "acc_borracha_psg_t":   request.form.get("borracha_psg_t", ""),
        "acc_borr_mtr_t":       request.form.get("borracha_mtr_t", ""),
        "acc_asa_urubu_de":     request.form.get("asa_urubu_de", ""),
        "acc_asa_urub_te":      request.form.get("asa_urubu_te", ""),
        "acc_bagagito":         request.form.get("bagagito", ""),
        "acc_linguet":          request.form.get("lingueta", ""),
    }

    # ── Salvar fotos temporariamente ──────────────────────────────────────────
    fotos_paths = []
    for foto in request.files.getlist("fotos"):
        if foto and foto.filename:
            safe = secure_filename(foto.filename)
            ext = Path(safe).suffix.lower()
            if ext in (".jpg", ".jpeg", ".png"):
                p = TEMP_FOLDER / f"{uuid.uuid4().hex}{ext}"
                foto.save(str(p))
                fotos_paths.append(p)

    formato = request.form.get("formato", "pdf").lower()

    # ── Gerar arquivo ─────────────────────────────────────────────────────────
    placa = _slugify(dados.get("placa", "PLACA"))
    data_slug = datetime.now().strftime("%d.%m.%Y")
    nome_docx = f"VISTORIA_{placa}_{data_slug}.docx"
    nome_pdf  = f"VISTORIA_{placa}_{data_slug}.pdf"
    caminho_docx = str(CONTRATOS_FOLDER / nome_docx)
    caminho_pdf  = str(CONTRATOS_FOLDER / nome_pdf)

    try:
        gerar_vistoria_entrega(dados, fotos_paths, caminho_docx, str(VISTORIA_TEMPLATE))
    except Exception as e:
        import traceback; traceback.print_exc()
        for p in fotos_paths:
            p.unlink(missing_ok=True)
        return jsonify({"error": f"Erro ao gerar vistoria: {e}"}), 500

    for p in fotos_paths:
        p.unlink(missing_ok=True)

    # ── Histórico + resposta ──────────────────────────────────────────────────
    if formato == "docx":
        _historico_append(dados.get("cliente", ""), "VISTORIA", nome_docx)
        return jsonify({"download_url": url_for("baixar_vistoria", nome=nome_docx)})

    try:
        _converter_pdf(caminho_docx, caminho_pdf)
    except Exception as e:
        flash(f"Erro ao converter para PDF: {e}", "erro")
        return redirect(url_for("pagina_vistoria"))

    _historico_append(dados.get("cliente", ""), "VISTORIA", nome_pdf)
    return jsonify({"download_url": url_for("baixar_vistoria", nome=nome_pdf)})


@app.route("/vistoria/gerar", methods=["POST"])
def gerar_vistoria_route():
    foto_path = None
    try:
        return _gerar_vistoria_impl()
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": f"Erro interno: {e}"}), 500
    finally:
        if foto_path:
            Path(foto_path).unlink(missing_ok=True)


def _gerar_vistoria_impl():
    agora = datetime.now(_BRT)
    dados = {
        "cliente_nome":      request.form.get("cliente_nome", ""),
        "cliente_telefone":  request.form.get("cliente_telefone", ""),
        "cliente_endereco":  request.form.get("cliente_endereco", ""),
        "preenchido_por":    request.form.get("preenchido_por", ""),
        "data_vistoria":     agora.strftime("%d/%m/%Y"),
        "veiculo":           request.form.get("veiculo", ""),
        "placa":             request.form.get("placa", "").upper(),
        "cor":               request.form.get("cor", ""),
        "ano":               request.form.get("ano", ""),
        "chassi":            request.form.get("chassi", ""),
        "numero_motor":      request.form.get("numero_motor", ""),
        "data_hora":         agora.strftime("%d/%m/%Y %H:%M"),
        "hodometro_entrega": request.form.get("hodometro_entrega", ""),
        "hodometro_retorno": request.form.get("hodometro_retorno", ""),
        "combustivel":       request.form.get("combustivel", ""),
        "acc_calotas":          request.form.get("acc_calotas", ""),
        "acc_buzina":           request.form.get("acc_buzina", ""),
        "acc_doc_crlv":         request.form.get("acc_doc_crlv", ""),
        "acc_triangulo":        request.form.get("acc_triangulo", ""),
        "acc_antena":           request.form.get("acc_antena", ""),
        "acc_sensor_re":        request.form.get("acc_sensor_re", ""),
        "acc_som":              request.form.get("acc_som", ""),
        "acc_tapetes":          request.form.get("acc_tapetes", ""),
        "acc_limpadores":       request.form.get("acc_limpadores", ""),
        "acc_chave_roda":       request.form.get("acc_chave_roda", ""),
        "acc_vidros_eletricos": request.form.get("acc_vidros_eletricos", ""),
        "acc_oleo_motor":       request.form.get("acc_oleo_motor", ""),
        "acc_alarme":           request.form.get("acc_alarme", ""),
        "acc_lampadas":         request.form.get("acc_lampadas", ""),
        "acc_macaco":           request.form.get("acc_macaco", ""),
        "acc_estepe":           request.form.get("acc_estepe", ""),
        "acc_gnv":              request.form.get("acc_gnv", ""),
        "acc_agua":             request.form.get("acc_agua", ""),
        "acc_borr_psg_dir":     request.form.get("acc_borr_psg_dir", ""),
        "acc_borr_mtr_dir":     request.form.get("acc_borr_mtr_dir", ""),
        "acc_asa_dd":           request.form.get("acc_asa_dd", ""),
        "acc_asa_td":           request.form.get("acc_asa_td", ""),
        "acc_tapete_mala":      request.form.get("acc_tapete_mala", ""),
        "acc_tampa_parachoque": request.form.get("acc_tampa_parachoque", ""),
        "acc_borr_psg_tras":    request.form.get("acc_borr_psg_tras", ""),
        "acc_borr_mtr_tras":    request.form.get("acc_borr_mtr_tras", ""),
        "acc_asa_de":           request.form.get("acc_asa_de", ""),
        "acc_asa_te":           request.form.get("acc_asa_te", ""),
        "acc_bagagito":         request.form.get("acc_bagagito", ""),
        "acc_lingueta":         request.form.get("acc_lingueta", ""),
        "obs_gerais":        request.form.get("obs_gerais", ""),
        "desc_sintomas":     request.form.get("desc_sintomas", ""),
    }

    fotos_paths = []

    def _salvar_foto(file_storage):
        if not file_storage or not file_storage.filename:
            return None
        ext = Path(secure_filename(file_storage.filename)).suffix.lower()
        if ext not in ('.jpg', '.jpeg', '.png'):
            return None
        p = TEMP_FOLDER / f"{uuid.uuid4().hex}{ext}"
        file_storage.save(str(p))
        return str(p)

    foto_painel = _salvar_foto(request.files.get("foto_painel"))
    if foto_painel:
        fotos_paths.append(foto_painel)

    for f in request.files.getlist("fotos_veiculo"):
        p = _salvar_foto(f)
        if p:
            fotos_paths.append(p)

    edit_id    = request.form.get("edit_id", "").strip()
    placa_slug = _slugify(dados["placa"] or "PLACA")
    data_slug  = agora.strftime("%d.%m.%Y")
    nome_docx    = f"VISTORIA_{placa_slug}_{data_slug}.docx"
    caminho_docx = str(CONTRATOS_FOLDER / nome_docx)

    try:
        gerar_vistoria_nova(dados, fotos_paths, caminho_docx)
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": f"Erro ao gerar vistoria: {e}"}), 500
    finally:
        for p in fotos_paths:
            Path(p).unlink(missing_ok=True)

    # ── Supabase: INSERT na thread principal (rápido), upload no background ────
    if _os.environ.get("SUPABASE_URL") and _os.environ.get("SUPABASE_KEY"):
        import threading, traceback as _tb
        _storage_path = f"vistorias/{nome_docx}"
        _docx_bytes   = Path(caminho_docx).read_bytes()
        _insert       = {
            "cliente":           dados["cliente_nome"],
            "telefone":          dados["cliente_telefone"],
            "endereco":          dados["cliente_endereco"],
            "preenchido_por":    dados["preenchido_por"],
            "veiculo":           dados["veiculo"],
            "placa":             dados["placa"],
            "cor":               dados["cor"],
            "ano":               dados["ano"],
            "chassi":            dados["chassi"],
            "numero_motor":      dados["numero_motor"],
            "data_hora":         dados["data_hora"],
            "hodometro_entrega": dados["hodometro_entrega"],
            "hodometro_retorno": dados["hodometro_retorno"],
            "combustivel":       dados["combustivel"],
            "obs_gerais":        dados["obs_gerais"],
            "desc_sintomas":     dados["desc_sintomas"],
            "arquivo_path":      _storage_path,
            "acessorios":        {k: v for k, v in dados.items() if k.startswith('acc_')},
        }

        _old_storage_path = None
        sb = _supabase()
        if sb:
            try:
                sb.table("vistorias").insert(_insert).execute()
            except Exception:
                _tb.print_exc()
            if edit_id:
                try:
                    old = sb.table("vistorias").select("arquivo_path").eq("id", edit_id).single().execute()
                    _old_storage_path = (old.data or {}).get("arquivo_path")
                    sb.table("vistorias").delete().eq("id", edit_id).execute()
                except Exception:
                    _tb.print_exc()

        _ost = _old_storage_path

        def _bg():
            try:
                sb2 = _supabase()
                if not sb2:
                    return
                try:
                    sb2.storage.from_("documentos").upload(
                        _storage_path, _docx_bytes,
                        {"content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                         "upsert": "true"},
                    )
                except Exception:
                    _tb.print_exc()
                if _ost and _ost != _storage_path:
                    try:
                        sb2.storage.from_("documentos").remove([_ost])
                    except Exception:
                        pass
            except Exception:
                _tb.print_exc()

        threading.Thread(target=_bg, daemon=True).start()

    try:
        _historico_append(dados["cliente_nome"], "VISTORIA", nome_docx)
    except Exception:
        import traceback; traceback.print_exc()

    return jsonify({"redirect_url": url_for("historico_vistorias")})


# ── Histórico de Vistorias (Supabase) ─────────────────────────────────────────

@app.route("/historico/vistorias")
def historico_vistorias():
    sb = _supabase()
    vistorias = []
    erro = None
    if sb:
        try:
            res = sb.table("vistorias").select(
                "id, cliente, placa, veiculo, preenchido_por, data_hora, criado_em, arquivo_path"
            ).neq("deletado", True).order("criado_em", desc=True).execute()
            vistorias = res.data or []
        except Exception as e:
            erro = str(e)
    else:
        erro = "Supabase não configurado (SUPABASE_URL / SUPABASE_KEY ausentes)."
    return render_template("historico_vistorias.html", vistorias=vistorias, erro=erro, active="hist_vistorias")


@app.route("/historico/vistorias/<vistoria_id>/excluir", methods=["POST"])
def excluir_vistoria(vistoria_id):
    sb = _supabase()
    if sb:
        try:
            sb.table("vistorias").update({"deletado": True}).eq("id", vistoria_id).execute()
        except Exception:
            import traceback; traceback.print_exc()
    return jsonify({"ok": True})


@app.route("/historico/vistorias/download/<vistoria_id>")
def download_vistoria_supabase(vistoria_id):
    sb = _supabase()
    if not sb:
        abort(503)
    try:
        res = sb.table("vistorias").select("arquivo_path").eq("id", vistoria_id).single().execute()
        path = res.data["arquivo_path"]
        signed = sb.storage.from_("documentos").create_signed_url(path, 60)
        return redirect(signed["signedURL"])
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/historico/vistorias/<vistoria_id>/editar")
def editar_vistoria(vistoria_id):
    sb = _supabase()
    if not sb:
        flash("Supabase não configurado.", "erro")
        return redirect(url_for("historico_vistorias"))
    try:
        res = sb.table("vistorias").select("*").eq("id", vistoria_id).single().execute()
        vistoria = res.data
    except Exception as e:
        flash(f"Erro ao buscar vistoria: {e}", "erro")
        return redirect(url_for("historico_vistorias"))
    return render_template("vistoria.html", active="vistoria", vistoria=vistoria,
                           edit_id=vistoria_id, acessorios=vistoria.get("acessorios") or {})


@app.route("/historico/vistorias/download/<vistoria_id>/pdf")
def download_vistoria_pdf(vistoria_id):
    sb = _supabase()
    if not sb:
        abort(503)
    try:
        res = sb.table("vistorias").select("arquivo_path").eq("id", vistoria_id).single().execute()
        docx_storage_path = res.data["arquivo_path"]
        docx_bytes = sb.storage.from_("documentos").download(docx_storage_path)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    if not isinstance(docx_bytes, (bytes, bytearray)):
        docx_bytes = getattr(docx_bytes, 'content', None) or bytes(docx_bytes)
    if len(docx_bytes) == 0:
        return jsonify({"error": f"Download vazio para {docx_storage_path!r}"}), 500
    if docx_bytes[:4] != b'PK\x03\x04':
        return jsonify({"error": f"Arquivo baixado não é DOCX válido. Início: {docx_bytes[:20]!r}"}), 500

    TEMP_FOLDER.mkdir(exist_ok=True)
    tmp_docx = TEMP_FOLDER / f"{uuid.uuid4().hex}.docx"
    tmp_pdf  = tmp_docx.with_suffix(".pdf")
    try:
        tmp_docx.write_bytes(docx_bytes)
        _converter_pdf(str(tmp_docx), str(tmp_pdf))
        pdf_bytes = tmp_pdf.read_bytes()
    except Exception as e:
        return jsonify({"error": f"Erro ao gerar PDF: {e}"}), 422
    finally:
        tmp_docx.unlink(missing_ok=True)
        tmp_pdf.unlink(missing_ok=True)

    nome_pdf = Path(docx_storage_path).stem + ".pdf"
    return send_file(BytesIO(pdf_bytes), as_attachment=True, download_name=nome_pdf, mimetype="application/pdf")


@app.route("/vistoria/download/<nome>")
def baixar_vistoria(nome):
    caminho = CONTRATOS_FOLDER / nome
    if not caminho.exists() or caminho.parent.resolve() != CONTRATOS_FOLDER.resolve():
        abort(404)
    ext = caminho.suffix.lower()
    mime = "application/pdf" if ext == ".pdf" else \
           "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    return send_file(str(caminho), as_attachment=True, download_name=nome, mimetype=mime)


@app.route("/debug/libreoffice")
def debug_libreoffice():
    """Diagnóstico: testa conversão de um DOCX mínimo a PDF."""
    import tempfile, os, shutil, platform as _pf
    lines = [f"platform={_pf.system()}"]
    # versão do LO
    try:
        r = subprocess.run(["libreoffice", "--version"], capture_output=True, timeout=10)
        lines.append(f"version={r.stdout.decode(errors='replace').strip()}")
    except Exception as e:
        lines.append(f"version_error={e}")
    # criação de DOCX mínimo
    try:
        from docx import Document
        with tempfile.TemporaryDirectory() as wd:
            p_docx = Path(wd) / "test.docx"
            p_pdf  = Path(wd) / "test.pdf"
            doc = Document(); doc.add_paragraph("Teste LibreOffice"); doc.save(str(p_docx))
            _converter_pdf(str(p_docx), str(p_pdf))
            lines.append(f"pdf_ok={p_pdf.exists()} size={p_pdf.stat().st_size if p_pdf.exists() else 0}")
    except Exception as e:
        lines.append(f"convert_error={e}")
    return "<br>".join(lines)


# ── Controle de Inadimplência ─────────────────────────────────────────────────

def _brl(v):
    """Format float as Brazilian currency (R$ 1.234,56)."""
    try:
        v = float(v)
    except (TypeError, ValueError):
        v = 0.0
    parts = f"{v:,.2f}".split(".")
    return "R$ " + parts[0].replace(",", ".") + "," + parts[1]


def _parse_valor_excel(raw):
    if isinstance(raw, (int, float)):
        return float(raw)
    if not raw:
        return 0.0
    s = str(raw).replace("R$", "").replace(" ", "").strip()
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def _inad_summary():
    """Read CONTAS-A-RECEBER.xlsx and return aggregated data for the dashboard."""
    import openpyxl

    def _nh(s):
        s = unicodedata.normalize("NFD", str(s or "").lower())
        return "".join(c for c in s if unicodedata.category(c) != "Mn")

    xlsx_path = Path(__file__).parent / "docx_templates" / "CONTAS-A-RECEBER.xlsx"
    if not xlsx_path.exists():
        return None

    try:
        wb = openpyxl.load_workbook(str(xlsx_path), read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()

        header_idx = 0
        _htargets = ["receber de", "vencimento", "valor"]
        for ri, row in enumerate(rows[:10]):
            nh_row = [_nh(str(c or "")) for c in row]
            if sum(1 for t in _htargets if any(t in n for n in nh_row)) >= 2:
                header_idx = ri
                break

        header    = rows[header_idx]
        data_rows = rows[header_idx + 1:]

        def _ci(keyword):
            nk = _nh(keyword)
            return next((i for i, h in enumerate(header)
                         if h is not None and nk in _nh(str(h))), None)

        i_nome  = _ci("receber de (fantasia)") or _ci("receber de")
        i_valor = _ci("valor previsto") or _ci("valor")
        i_venc  = _ci("data de vencimento") or _ci("vencimento")
        i_sit   = _ci("situacao (data de vencimento)") or _ci("situacao")
        i_doc   = _ci("numero do documento") or _ci("documento")
        i_unid  = _ci("unidade")

        hoje = date.today()
        _MULTA_VALS = {600, 630, 650, 680, 700, 800, 1200}

        def _get(row, idx):
            return row[idx] if idx is not None and idx < len(row) else None

        registros = []
        for row in data_rows:
            nome_raw = _get(row, i_nome)
            if not nome_raw:
                continue
            nome = str(nome_raw).strip()
            if not nome:
                continue

            valor = _parse_valor_excel(_get(row, i_valor))
            if valor <= 0:
                continue

            venc_raw = _get(row, i_venc)
            sit_raw  = str(_get(row, i_sit) or "").lower()

            venc_date = None
            if venc_raw:
                if isinstance(venc_raw, datetime):
                    venc_date = venc_raw.date()
                elif isinstance(venc_raw, date):
                    venc_date = venc_raw
                else:
                    for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]:
                        try:
                            venc_date = datetime.strptime(str(venc_raw).strip(), fmt).date()
                            break
                        except (ValueError, TypeError):
                            pass
            if venc_date is None:
                continue

            if "a vencer" in sit_raw and venc_date > hoje:
                continue

            if venc_date == hoje:
                dias = 0
            else:
                dias = (hoje - venc_date).days
                if dias < 0:
                    continue

            qualifica = int(round(valor)) in _MULTA_VALS
            multa = valor * 0.05 if (dias >= 1 and qualifica) else 0.0
            juros = valor * 0.005 * dias if dias >= 1 else 0.0
            total = valor + multa + juros

            if dias == 0:    etapa = "D+0"
            elif dias == 1:  etapa = "D+1"
            elif dias == 2:  etapa = "D+2"
            elif dias == 3:  etapa = "D+3"
            elif dias == 4:  etapa = "D+4"
            elif dias <= 6:  etapa = "D+5"
            elif dias <= 9:  etapa = "D+7"
            elif dias <= 14: etapa = "D+10"
            else:            etapa = "D+15"

            doc  = str(_get(row, i_doc) or "").strip()
            unid = str(_get(row, i_unid) or "").strip()

            registros.append({
                "nome":    nome,
                "placa":   doc or unid or "—",
                "venc":    venc_date.strftime("%d/%m/%Y"),
                "dias":    dias,
                "total_s": _brl(total),
                "etapa":   etapa,
            })

        if not registros:
            return {"total_s": _brl(0), "casos": 0, "hoje": 0,
                    "por_etapa": {}, "recentes": [], "total_raw": 0}

        total_aberto = sum(_parse_valor_excel(r["total_s"]) for r in registros)
        # Re-compute from original values stored above via closure isn't possible;
        # sum total_s back from formatted strings by re-reading raw registros
        total_aberto_raw = 0.0
        for row in data_rows:
            nome_raw = _get(row, i_nome)
            if not nome_raw or not str(nome_raw).strip():
                continue
            valor = _parse_valor_excel(_get(row, i_valor))
            if valor <= 0:
                continue
            venc_raw = _get(row, i_venc)
            sit_raw  = str(_get(row, i_sit) or "").lower()
            venc_date = None
            if venc_raw:
                if isinstance(venc_raw, datetime):
                    venc_date = venc_raw.date()
                elif isinstance(venc_raw, date):
                    venc_date = venc_raw
                else:
                    for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]:
                        try:
                            venc_date = datetime.strptime(str(venc_raw).strip(), fmt).date()
                            break
                        except (ValueError, TypeError):
                            pass
            if venc_date is None:
                continue
            if "a vencer" in sit_raw and venc_date > hoje:
                continue
            if venc_date == hoje:
                dias = 0
            else:
                dias = (hoje - venc_date).days
                if dias < 0:
                    continue
            qualifica = int(round(valor)) in _MULTA_VALS
            multa = valor * 0.05 if (dias >= 1 and qualifica) else 0.0
            juros = valor * 0.005 * dias if dias >= 1 else 0.0
            total_aberto_raw += valor + multa + juros

        nomes_unicos = set(r["nome"] for r in registros)
        hoje_count   = sum(1 for r in registros if r["dias"] == 0)

        etapas = ["D+0","D+1","D+2","D+3","D+4","D+5","D+7","D+10","D+15"]
        por_etapa = {e: 0 for e in etapas}
        for r in registros:
            if r["etapa"] in por_etapa:
                por_etapa[r["etapa"]] += 1

        recentes = sorted(registros, key=lambda r: r["dias"], reverse=True)[:8]

        return {
            "total_s":   _brl(total_aberto_raw),
            "total_raw": total_aberto_raw,
            "casos":     len(nomes_unicos),
            "hoje":      hoje_count,
            "por_etapa": por_etapa,
            "recentes":  recentes,
        }

    except Exception:
        import traceback; traceback.print_exc()
        return None


@app.route("/inadimplencia")
def pagina_inadimplencia():
    from urllib.parse import quote as _url_quote
    from collections import Counter
    import openpyxl

    def _nh(s):
        s = unicodedata.normalize("NFD", str(s or "").lower())
        return "".join(c for c in s if unicodedata.category(c) != "Mn")

    _base = Path(__file__).parent / "docx_templates"
    xlsx_path = _base / "CONTAS-A-RECEBER.xlsx"

    hoje = date.today()
    registros_vencidos = []   # VENCIDO rows + A VENCER rows whose date == hoje (D+0)
    registros_a_vencer = []   # A VENCER rows whose date > hoje
    erro_leitura = None

    if xlsx_path.exists():
        try:
            wb = openpyxl.load_workbook(str(xlsx_path), read_only=True, data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            wb.close()

            # Detect header row (needs "receber de" + "vencimento" + "valor")
            header_idx = 0
            _htargets = ["receber de", "vencimento", "valor"]
            for ri, row in enumerate(rows[:10]):
                nh_row = [_nh(str(c or "")) for c in row]
                if sum(1 for t in _htargets if any(t in n for n in nh_row)) >= 2:
                    header_idx = ri
                    break

            header    = rows[header_idx]
            data_rows = rows[header_idx + 1:]

            def _ci(keyword):
                nk = _nh(keyword)
                return next((i for i, h in enumerate(header)
                             if h is not None and nk in _nh(str(h))), None)

            i_nome  = _ci("receber de (fantasia)") or _ci("receber de")
            i_valor = _ci("valor previsto") or _ci("valor")
            i_venc  = _ci("data de vencimento") or _ci("vencimento")
            i_sit   = _ci("situacao (data de vencimento)") or _ci("situacao")
            i_tipo  = _ci("tipo de fatura") or _ci("tipo")
            i_doc   = _ci("numero do documento") or _ci("documento")
            i_unid  = _ci("unidade")

            # Count all rows per name (vencido + a vencer) for reincidence
            name_counts = Counter()
            for row in data_rows:
                if i_nome is not None and i_nome < len(row) and row[i_nome]:
                    n = str(row[i_nome]).strip()
                    if n:
                        name_counts[n] += 1

            def _get(row, idx):
                return row[idx] if idx is not None and idx < len(row) else None

            for row in data_rows:
                nome_raw = _get(row, i_nome)
                if not nome_raw:
                    continue
                nome = str(nome_raw).strip()
                if not nome:
                    continue

                valor    = _parse_valor_excel(_get(row, i_valor))
                venc_raw = _get(row, i_venc)
                sit_raw  = _get(row, i_sit)
                tipo_raw = _get(row, i_tipo)
                doc_raw  = _get(row, i_doc)
                unid_raw = _get(row, i_unid)

                # Parse vencimento date (openpyxl returns datetime objects directly)
                venc_date = None
                if venc_raw:
                    if isinstance(venc_raw, datetime):
                        venc_date = venc_raw.date()
                    elif isinstance(venc_raw, date):
                        venc_date = venc_raw
                    else:
                        for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]:
                            try:
                                venc_date = datetime.strptime(str(venc_raw).strip(), fmt).date()
                                break
                            except (ValueError, TypeError):
                                pass
                if venc_date is None:
                    continue

                situacao    = _nh(str(sit_raw  or ""))
                tipo_fatura = str(tipo_raw or "").strip()
                num_doc     = str(doc_raw  or "").strip()
                unidade     = str(unid_raw or "").strip()
                for s in ("None", "nan", ""):
                    if num_doc == s: num_doc = ""
                    if unidade == s: unidade = ""

                reincidente = name_counts[nome] > 1
                is_fatura   = _nh(tipo_fatura) == "fatura"
                data_fmt    = venc_date.strftime("%d/%m/%Y")

                # ── A VENCER (future dates) ──────────────────────────────────
                if "a vencer" in situacao and venc_date > hoje:
                    dias_ate = (venc_date - hoje).days
                    registros_a_vencer.append({
                        "nome":            nome,
                        "num_doc":         num_doc,
                        "unidade":         unidade,
                        "data_vencimento": data_fmt,
                        "dias_ate":        dias_ate,
                        "reincidente":     reincidente,
                        "tipo_fatura":     tipo_fatura,
                        "valor_s":         _brl(valor),
                        "_valor":          valor,
                    })
                    continue

                # ── VENCIDO or vence hoje ────────────────────────────────────
                if "a vencer" in situacao and venc_date == hoje:
                    dias = 0
                else:
                    dias = (hoje - venc_date).days
                    if dias < 0:
                        continue

                if dias == 0:
                    etapa, etapa_cls = "D+0",  "stage-d0"
                    proxima = "Enviar lembrete de vencimento"
                elif dias == 1:
                    etapa, etapa_cls = "D+1",  "stage-d1"
                    proxima = "Cobrança formal + aplicar multa e juros"
                elif dias == 2:
                    etapa, etapa_cls = "D+2",  "stage-d2"
                    proxima = "Cobrança formal + aplicar multa e juros"
                elif dias == 3:
                    etapa, etapa_cls = "D+3",  "stage-d3"
                    proxima = "Pressão + avisar suspensão em 48h"
                elif dias == 4:
                    etapa, etapa_cls = "D+4",  "stage-d4"
                    proxima = "Suspensão iminente — último aviso"
                elif dias <= 6:
                    etapa, etapa_cls = "D+5",  "stage-d5"
                    proxima = "Bloqueio do veículo"
                elif dias <= 9:
                    etapa, etapa_cls = "D+7",  "stage-d7"
                    proxima = "Notificação formal — prazo 48h para negativação"
                elif dias <= 14:
                    etapa, etapa_cls = "D+10", "stage-d10"
                    proxima = "Negativação + encaminhamento jurídico"
                else:
                    etapa, etapa_cls = "D+15", "stage-d15"
                    proxima = "Recolhimento + protesto em cartório + execução contratual"

                # Multa applies only to specific contract values (locação fixa)
                _MULTA_VALS = {600, 630, 650, 680, 700, 800, 1200}
                valor_qualifica_multa = int(round(valor)) in _MULTA_VALS
                tem_multa = dias >= 1 and valor_qualifica_multa
                multa  = valor * 0.05 if tem_multa else 0.0
                juros  = valor * 0.005 * dias if dias >= 1 else 0.0
                total  = valor + multa + juros
                pausar = total * 0.5

                dias_label = (f"{dias} dia{'s' if dias != 1 else ''} de atraso"
                              if dias > 0 else "Vence hoje")
                dias_s = f"{dias} dia{'s' if dias != 1 else ''}"

                # WhatsApp messages per stage
                if dias == 0:
                    msg = f"Oi, {nome}! 😊 Passando para avisar que sua parcela de *{_brl(valor)}* vence *hoje*. Qualquer dúvida, é só chamar!\n\n\n*Ativuz Veículos*"
                elif dias == 1:
                    msg = f"{nome}, sua parcela de *{_brl(valor)}* venceu ontem (vencimento: {data_fmt}). O valor atualizado é *{_brl(total)}*. Assim que puder, regularize para evitar encargos adicionais.\n\n\n*Ativuz Veículos*"
                elif dias == 2:
                    msg = f"{nome}, seu pagamento de *{_brl(total)}* ainda está em aberto (vencimento: {data_fmt}). Caso tenha alguma dúvida ou dificuldade, entre em contato antes que os encargos aumentem.\n\n\n*Ativuz Veículos*"
                elif dias == 3:
                    msg = f"{nome}, seu pagamento está em aberto há *{dias_s}* (vencimento: {data_fmt}). Valor atualizado: *{_brl(total)}*. Caso haja algum imprevisto, entre em contato — mas precisamos regularizar em breve para evitar a suspensão do serviço.\n\n\n*Ativuz Veículos*"
                elif dias == 4:
                    msg = f"{nome}, sua parcela de *{_brl(valor)}* segue em aberto há *{dias_s}* (vencimento: {data_fmt}). Valor atualizado: *{_brl(total)}*. Regularize o quanto antes para evitar a suspensão do veículo.\n\n\n*Ativuz Veículos*"
                elif dias <= 6:   # D+5
                    msg = f"{nome}, infelizmente precisamos suspender o serviço por inadimplência, conforme contrato. Valor atualizado: *{_brl(total)}*. Para reativação, basta regularizar o pagamento. Estamos à disposição.\n\n\n*Ativuz Veículos*"
                elif dias <= 9:   # D+7
                    msg = f"{nome}, seu débito de *{_brl(total)}* está em aberto há {dias_s}. Esta é uma notificação formal com prazo de *48 horas* para regularização antes de tomarmos as próximas medidas previstas em contrato.\n\n\n*Ativuz Veículos*"
                elif dias <= 14:  # D+10
                    msg = f"{nome}, informamos que seu débito foi encaminhado para negativação e assessoria jurídica. Valor atualizado: *{_brl(total)}*.\n\n\n*Ativuz Veículos*"
                else:             # D+15
                    msg = f"{nome}, comunicamos que serão iniciados os procedimentos de protesto em cartório e execução contratual. Valor atualizado: *{_brl(total)}*.\n\n\n*Ativuz Veículos*"

                # Pausar cobrança: FATURA only, D+1 or D+2
                mostrar_pausar = is_fatura and 1 <= dias <= 2
                if mostrar_pausar:
                    if dias == 1:
                        msg_pausar = f"{nome}, sua parcela de *{_brl(valor)}* está em aberto (vencimento: {data_fmt}). O valor atualizado é *{_brl(total)}*. 📌 Pague *{_brl(pausar)}* hoje e quite *{_brl(pausar)}* até a sexta-feira desta semana. ⚠️ Juros de 0,5% ao dia continuam correndo sobre o saldo restante. Sem pagamento até sexta, a cobrança retoma no sábado. ⚠️ Não se trata de desconto. O valor total do débito permanece integral.\n\n\n*Ativuz Veículos*"
                    else:  # dias == 2
                        msg_pausar = f"{nome}, sua parcela de *{_brl(valor)}* está em aberto há *2 dias* (vencimento: {data_fmt}). O valor atualizado é *{_brl(total)}*. 📌 Pague *{_brl(pausar)}* hoje e quite *{_brl(pausar)}* até a sexta-feira desta semana. ⚠️ Juros de 0,5% ao dia continuam correndo sobre o saldo restante. Sem pagamento até sexta, a cobrança retoma no sábado. ⚠️ Não se trata de desconto. O valor total do débito permanece integral.\n\n\n*Ativuz Veículos*"
                else:
                    msg_pausar = None

                wa_cobranca = "https://wa.me/?text=" + _url_quote(msg)
                wa_pausar   = (("https://wa.me/?text=" + _url_quote(msg_pausar))
                               if mostrar_pausar else None)

                sit_key = "vence-hoje" if dias == 0 else "vencido"

                registros_vencidos.append({
                    "nome":             nome,
                    "num_doc":          num_doc,
                    "unidade":          unidade,
                    "tipo_fatura":      tipo_fatura,
                    "data_vencimento":  data_fmt,
                    "dias_atraso":      dias,
                    "dias_label":       dias_label,
                    "reincidente":      reincidente,
                    "is_fatura":        is_fatura,
                    "tem_multa":        tem_multa,
                    "etapa":            etapa,
                    "etapa_cls":        etapa_cls,
                    "proxima_acao":     proxima,
                    "situacao_key":     sit_key,
                    "wa_cobranca":      wa_cobranca,
                    "wa_pausar":        wa_pausar,
                    "msg_cobranca_txt": msg,
                    "msg_pausar_txt":   msg_pausar,
                    "valor_s":          _brl(valor),
                    "multa_s":          _brl(multa),
                    "juros_s":          _brl(juros),
                    "total_s":          _brl(total),
                    "pausar_s":         _brl(pausar),
                    "_valor":           valor,
                    "_multa":           multa,
                    "_juros":           juros,
                    "_total":           total,
                })

        except Exception as e:
            import traceback; traceback.print_exc()
            erro_leitura = str(e)
    else:
        erro_leitura = (
            "Planilha não encontrada em docx_templates/. "
            "Salve o arquivo como CONTAS-A-RECEBER.xlsx nessa pasta."
        )

    registros_vencidos.sort(key=lambda r: r["dias_atraso"], reverse=True)
    registros_a_vencer.sort(key=lambda r: r["dias_ate"])

    total_vencidos        = len(registros_vencidos)
    total_a_vencer_cnt    = len(registros_a_vencer)
    total_valor_orig      = _brl(sum(r["_valor"] for r in registros_vencidos))
    total_valor_atual     = _brl(sum(r["_total"] for r in registros_vencidos))
    total_a_vencer_val    = _brl(sum(r["_valor"] for r in registros_a_vencer))
    criticos              = sum(1 for r in registros_vencidos if r["dias_atraso"] >= 7)
    reincidentes_criticos = sum(1 for r in registros_vencidos
                                if r["dias_atraso"] >= 7 and r["reincidente"])

    return render_template(
        "inadimplencia.html",
        registros=registros_vencidos,
        registros_a_vencer=registros_a_vencer,
        total_registros=total_vencidos,
        total_a_vencer=total_a_vencer_cnt,
        total_valor_orig=total_valor_orig,
        total_valor_atual=total_valor_atual,
        total_a_vencer_val=total_a_vencer_val,
        criticos=criticos,
        reincidentes=reincidentes_criticos,
        erro_leitura=erro_leitura,
        hoje=hoje.strftime("%d/%m/%Y"),
        active="inadimplencia",
    )


@app.route("/inadimplencia/upload", methods=["POST"])
def inadimplencia_upload():
    f = request.files.get("planilha")
    if not f or not f.filename:
        flash("Nenhum arquivo selecionado.", "error")
        return redirect(url_for("pagina_inadimplencia"))
    if not f.filename.lower().endswith(".xlsx"):
        flash("Apenas arquivos .xlsx são aceitos.", "error")
        return redirect(url_for("pagina_inadimplencia"))
    dest = Path(__file__).parent / "docx_templates" / "CONTAS-A-RECEBER.xlsx"
    f.save(str(dest))
    flash("Planilha atualizada! Os dados abaixo refletem o arquivo enviado.", "success")
    return redirect(url_for("pagina_inadimplencia"))


_MESES_PT = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
             "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
_MESES_PT_CURTO = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
                   "Jul", "Ago", "Set", "Out", "Nov", "Dez"]


def _nome_mes_label(mes, ano, acumulado=False):
    if acumulado:
        return f"Jan–{_MESES_PT_CURTO[mes - 1]} {ano}"
    return f"{_MESES_PT[mes - 1]} {ano}"


# ── Configurações DRE ────────────────────────────────────────────────────────
# IR/CSLL: manter 0.0 para Simples Nacional; alterar aqui para mudar regime.
_DRE_IR_CSLL = 0.0

# Códigos cujo Tipo no sistema (ENTRADA/SAÍDA) é oposto ao tratamento no DRE.
# Ao adicionar novos, revisar o grupo de destino no _DRE_LAYOUT abaixo.
_DRE_CODIGOS_SINAL_INVERTIDO = frozenset([
    "02.02.06.006",  # SAÍDA → receita operacional (Combustível Reembolsável)
    "01.01.02.008",  # ENTRADA → dedução (Desconto Concedido a Clientes)
    "02.04.06.004",  # SAÍDA → receita financeira (Desconto Pgto Boletos)
])

# ── DRE Layout ────────────────────────────────────────────────────────────────
# sign: +1 = receita, -1 = despesa.  item tuple: (codigo, label).
# Códigos tratados como string — nunca converter para número.

_DRE_LAYOUT = [
    {"id": "rb", "label": "(+) Receita Operacional Bruta", "grupos": [
        {"id": "rb-loc",  "label": "Locação", "sign": +1, "itens": [
            ("01.01.01.001", "Locação"),
            ("01.01.01.002", "KM Excedente"),
            ("01.01.01.003", "Multa Atraso Pagamento"),
            ("01.01.01.004", "Multa Quebra de Contrato"),
            ("01.01.01.005", "Acordo/Renegociação"),
            ("01.01.01.006", "Taxa de Adm. de Veículos"),
        ]},
        {"id": "rb-reimb", "label": "Reembolsos", "sign": +1, "itens": [
            ("01.01.02.001", "Manutenção Reembolsável"),
            ("01.01.02.002", "Entrada de Multa de Trânsito"),
            ("01.01.02.004", "Multa Dev. Antecipada"),
            ("01.01.02.005", "Reembolso de Sinistro"),
            ("01.01.02.006", "Outros Reembolsos"),
            ("02.02.06.006", "Combustível Reembolsável"),  # SAÍDA no sistema → sinal invertido
        ]},
        {"id": "rb-fi", "label": "Reembolsos — Frota Investidores", "sign": +1, "itens": [
            ("01.01.02.009", "Reembolso de Manutenções — FI"),
            ("01.01.02.010", "Reembolso de Multas — FI"),
            ("01.01.02.011", "Reembolso Desp. Operacionais — FI"),
            ("01.02.01.005", "Recebimentos — Frota Investidores"),
        ]},
    ]},
    {"id": "ded", "label": "(-) Deduções", "grupos": [
        {"id": "ded-imp",  "label": "Impostos", "sign": -1,
         "note": "incide sobre o faturamento — Simples Nacional", "itens": [
            ("02.01.01.001", "PIS"),
            ("02.01.01.005", "Simples Nacional"),
            ("02.01.01.006", "Outros Impostos"),
        ]},
        {"id": "ded-desc", "label": "Descontos", "sign": -1, "itens": [
            ("01.01.02.008", "Desconto Concedido a Clientes"),  # ENTRADA → sinal invertido
        ]},
    ]},
    # subtotal: receita_liquida
    {"id": "custos", "label": "(-) Custos — Custo Direto da Operação/Frota", "grupos": [
        {"id": "c-lic",   "label": "Licenciamento", "sign": -1, "itens": [
            ("02.02.01.001", "Emplacamento"),
            ("02.02.01.002", "Transferência Veicular"),
            ("02.02.01.003", "IPVA"),
            ("02.02.01.005", "Taxa de Licenciamento"),
            ("02.02.01.006", "Despesas com Cartório"),
            ("02.02.01.007", "Transferência Veicular"),
            ("02.02.01.008", "Taxa Bombeiros"),
            ("02.02.01.009", "Vistoria"),
            ("02.02.01.010", "Documentação Veicular"),
        ]},
        {"id": "c-desp",  "label": "Honorários Despachante", "sign": -1, "itens": [
            ("02.02.02.001", "Honorários Despachante"),
        ]},
        {"id": "c-seg",   "label": "Seguro/Assistência 24h", "sign": -1, "itens": [
            ("02.02.03.001", "Seguro Total"),
            ("02.02.03.003", "Reserva Operacional p/ Sinistros"),
            ("02.02.03.004", "Guincho"),
            ("02.02.03.005", "Rastreador Veicular"),
        ]},
        {"id": "c-sub",   "label": "Sublocação/Transporte", "sign": -1, "itens": [
            ("02.02.04.002", "Frete Não Reembolsável"),
            ("02.02.04.003", "Multas Não Reembolsáveis"),
            ("02.02.04.004", "Taxa de Frete"),
            ("02.02.04.005", "Gasolina"),
            ("02.02.04.006", "Uber ou App de Transporte"),
        ]},
        {"id": "c-man",   "label": "Manutenção", "sign": -1, "itens": [
            ("02.02.05.001", "Manutenção Preventiva"),
            ("02.02.05.002", "Manutenção Corretiva"),
            ("02.02.05.006", "Compra Equipamento GNV"),
            ("02.02.05.007", "Equipamentos (Sensor, SIM, etc.)"),
            ("02.02.05.008", "Lavagem Veicular Não Reembolsável"),
            ("02.02.05.009", "Compra de Peças"),
            ("02.02.05.010", "Compra de Pneus"),
        ]},
        {"id": "c-reimb", "label": "Despesas Reembolsáveis", "sign": -1, "itens": [
            ("02.02.06.001", "Saída de Multa de Trânsito"),
            ("02.02.06.002", "Saída de Sinistro"),
            ("02.02.06.010", "Reembolso de Clientes"),
        ]},
        {"id": "c-fi",    "label": "Despesas c/ Frota de Investidores", "sign": -1, "itens": [
            ("02.02.07.01", "Manutenções — Frota Investidores"),
            ("02.02.07.02", "Sinistros — Frota Investidores"),
            ("02.02.07.03", "Desp. Operacionais — FI"),
            ("02.02.07.04", "Desp. Operacionais — FI"),
        ]},
    ]},
    # subtotal: margem
    {"id": "sga", "label": "(-) SG&A — Despesas Gerais e Administrativas", "grupos": [
        {"id": "s-sal", "label": "Salários", "sign": -1, "itens": [
            ("02.03.01.001", "Salário"),
            ("02.03.01.002", "Adiantamento Salarial"),
            ("02.03.01.003", "Férias"),
            ("02.03.01.004", "13° Salário"),
            ("02.03.01.005", "Rescisão"),
            ("02.03.01.006", "Prêmios"),
            ("02.03.01.007", "Comissão"),
            ("02.03.01.10",  "ASO e Saúde e Seg. do Trabalho"),
        ]},
        {"id": "s-ben", "label": "Benefícios", "sign": -1, "itens": [
            ("02.03.02.001", "VT"),
            ("02.03.02.003", "VR"),
            ("02.03.02.005", "Assistência Médica"),
            ("02.03.02.006", "P.C.M.S.O"),
            ("02.03.02.007", "Treinamento"),
            ("02.03.02.008", "Outros Benefícios"),
        ]},
        {"id": "s-imp", "label": "Impostos Folha", "sign": -1, "itens": [
            ("02.03.03.001", "INSS"),
            ("02.03.03.002", "FGTS"),
        ]},
        {"id": "s-pro", "label": "Pró-labore", "sign": -1,
         "highlight": True, "note": "remuneração dos sócios", "itens": [
            ("02.03.04.001", "Pró-labore Folha"),
        ]},
        {"id": "s-com", "label": "Despesas Comerciais", "sign": -1, "itens": [
            ("02.04.01.001", "Marketing"),
        ]},
        {"id": "s-ocu", "label": "Ocupação", "sign": -1, "itens": [
            ("02.04.02.001", "Aluguel de Imóveis"),
            ("02.04.02.003", "IPTU"),
            ("02.04.02.004", "Água"),
            ("02.04.02.005", "Luz"),
            ("02.04.02.006", "Manutenção Predial"),
        ]},
        {"id": "s-sup", "label": "Suprimentos", "sign": -1, "itens": [
            ("02.04.03.001", "Telefone"),
            ("02.04.03.007", "Bens de Pequeno Valor"),
        ]},
        {"id": "s-adm", "label": "Despesas Administrativas", "sign": -1, "itens": [
            ("02.04.04.001", "Desp. Administrativas e de Escritório"),
            ("02.04.04.004", "Taxas e Despesas Legais"),
            ("02.04.04.005", "Outras Despesas"),
            ("02.04.04.007", "Despesas Jurídicas"),
        ]},
        {"id": "s-svc", "label": "Serviços Prestados", "sign": -1, "itens": [
            ("02.04.05.001", "Honorários Advocatícios"),
            ("02.04.05.003", "Softwares"),
            ("02.04.05.004", "Órgãos de Proteção ao Crédito"),
            ("02.04.05.005", "Assessoria Administrativa"),
            ("02.04.05.006", "Honorários de Consultoria"),
            ("02.04.05.007", "Serviços Contábeis"),
            ("02.04.05.008", "Serviços de Limpeza"),
            ("02.04.05.009", "Serviços Manut. Máquinas e Equip."),
        ]},
        {"id": "s-out", "label": "Outras Saídas", "sign": -1, "itens": [
            ("02.04.07.002", "Outras Saídas"),
        ]},
    ]},
    # subtotal: ebitda; depreciação=0; subtotal: ebit
    {"id": "rfin", "label": "(-) Resultado Financeiro", "grupos": [
        {"id": "rf-desp", "label": "Despesas Bancárias e Financeiras", "sign": -1, "itens": [
            ("02.04.06.001", "Tarifa Bancária"),
            ("02.04.06.003", "Juros e Multas Bancárias Pagos"),
            ("02.04.06.005", "Taxa Maquineta"),
            ("03.01.03.002", "Consórcio Contemplado Juros"),
            ("03.01.03.004", "Financiamento Juros"),
            ("04.01.02.002", "Pgto Juros sobre Mútuos"),
        ]},
        {"id": "rf-rec",  "label": "Receitas Financeiras", "sign": +1, "itens": [
            ("03.03.01.003", "Rendimento de Aplicações"),
            ("02.04.06.004", "Desconto Pgto Boletos"),  # SAÍDA no sistema → receita financeira
        ]},
    ]},
    {"id": "rnop", "label": "(-) Resultados Não Operacionais", "grupos": [
        {"id": "rnop-out", "label": "Outras Entradas", "sign": +1, "itens": [
            ("01.02.01.002", "Depósitos Não Identificados"),
            ("01.02.01.003", "Outras Entradas"),
        ]},
        {"id": "rnop-inv", "label": "Outros Investimentos", "sign": +1, "itens": [
            ("03.03.02.001", "Outros Investimentos"),
        ]},
    ]},
    # subtotal: lucro_liquido ── abaixo: fluxo de caixa
    {"id": "inv", "label": "(-) Investimentos", "grupos": [
        {"id": "inv-venda", "label": "Venda de Veículos", "sign": +1, "itens": [
            ("01.02.01.004", "Venda de Veículo"),
        ]},
        {"id": "inv-comp",  "label": "Compra de Veículos", "sign": -1, "itens": [
            ("03.01.02.001", "Compra de Veículos à Vista"),
            ("03.01.02.002", "Entrada Compra Veículo"),
            ("03.01.02.003", "Adiantamento de Consórcio"),
        ]},
        {"id": "inv-fin",   "label": "Financiamentos Veiculares", "sign": -1, "itens": [
            ("03.01.03.001", "Consórcio Contemplado"),
            ("03.01.03.003", "Pagamento de Financiamento"),
            ("03.01.03.005", "Consórcio Parcela Não Contemplada"),
            ("03.01.03.006", "Quitação Antecipada de Parcelas"),
        ]},
        {"id": "inv-imob",  "label": "Outros Imobilizados", "sign": -1, "itens": [
            ("03.02.01.001", "Instalações"),
            ("03.02.01.002", "Computadores e Periféricos"),
            ("03.02.01.003", "Móveis e Utensílios"),
            ("03.02.01.004", "Sistemas e Softwares"),
            ("03.02.01.005", "Outras Imobilizações"),
        ]},
        {"id": "inv-obra",  "label": "Construção da Oficina", "sign": -1, "itens": [
            ("02.04.08.01", "Compra de Material para Oficina"),
            ("02.04.08.02", "Pagamento da Mão de Obra"),
            ("02.04.08.03", "Aluguel de Equipamentos"),
        ]},
        {"id": "inv-aplic", "label": "Aplicações Financeiras", "sign": -1, "itens": [
            ("03.03.01.001", "Aplicações Financeiras"),
        ]},
        {"id": "inv-resg",  "label": "Resgate de Aplicação", "sign": +1, "itens": [
            ("03.03.01.002", "Resgate de Aplicação Financeira"),
        ]},
    ]},
    {"id": "financ", "label": "(-) Financiamentos", "grupos": [
        {"id": "fin-ent",      "label": "Entradas de Mútuos", "sign": +1, "itens": [
            ("04.01.01.001", "Entrada de Mútuos"),
            ("04.01.01.002", "Entrada Caução"),
        ]},
        {"id": "fin-pgto",     "label": "Pgto de Mútuos", "sign": -1, "itens": [
            ("04.01.02.001", "Saída de Mútuos"),
            ("04.01.02.003", "Saída Caução"),
        ]},
        {"id": "fin-reimb-in", "label": "Entrada de Reembolso", "sign": +1, "itens": [
            ("04.04.02.01", "Entrada de Reembolso"),
        ]},
        {"id": "fin-reimb-out","label": "Saída de Reembolso", "sign": -1, "itens": [
            ("04.04.02.02", "Saída de Reembolso"),
        ]},
    ]},
    {"id": "aporte", "label": "(+) Aporte de Capital", "grupos": [
        {"id": "aporte-g", "label": "Aporte de Capital", "sign": +1, "itens": [
            ("04.04.01.003", "Aporte de Capital"),
        ]},
    ]},
    # subtotal: fluxo_acionista
    {"id": "distrib", "label": "(-) Distribuição de Resultado", "grupos": [
        {"id": "distrib-g",     "label": "Distribuição de Resultado", "sign": -1, "itens": [
            ("04.04.01.002", "Distribuição de Resultado"),
        ]},
        {"id": "distrib-lucro", "label": "Distribuição de Lucros Mensal", "sign": -1, "itens": [
            ("02.03.04.002", "Distribuição de Lucros Mensal"),
        ]},
    ]},
    # subtotal: fluxo_livre
]


def _dre_ler_lancamentos():
    import openpyxl

    base = Path(__file__).resolve().parent
    pasta = base / "planilhas" / "dre"

    arquivos = sorted(pasta.glob("*.xlsx")) if pasta.is_dir() else []
    if not arquivos:
        return []

    def _split(s):
        if ' - ' in s:
            a, b = s.split(' - ', 1)
            return a.strip(), b.strip()
        return s.strip(), s.strip()

    def _ler_arquivo(path):
        registros = []
        try:
            wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            wb.close()
            for row in rows[5:]:
                if not row[0]:
                    continue
                dt    = row[3]
                valor = row[9]
                if not isinstance(dt, datetime) or not isinstance(valor, (int, float)):
                    continue
                cod, _ = _split(str(row[2]) if row[2] else "")
                registros.append({"codigo": cod, "dt": dt, "valor": float(valor)})
        except Exception:
            pass
        return registros

    # Deduplicação por nome de arquivo: ignora se o mesmo nome aparecer duas vezes
    nomes_vistos = set()
    result = []
    for arq in arquivos:
        if arq.name in nomes_vistos:
            continue
        nomes_vistos.add(arq.name)
        result.extend(_ler_arquivo(arq))
    return result


def _dre_calcular(lancamentos):
    from collections import defaultdict
    code_val = defaultdict(float)
    for l in lancamentos:
        if l["codigo"]:
            code_val[l["codigo"]] += l["valor"]

    sections = []
    for sec_def in _DRE_LAYOUT:
        grupos = []
        sec_total = 0.0
        for grp_def in sec_def["grupos"]:
            sign = grp_def["sign"]
            itens = []
            grp_abs = 0.0
            for codigo, label in grp_def["itens"]:
                v = code_val.get(codigo, 0.0)
                grp_abs += v
                itens.append({"codigo": codigo, "label": label, "val": sign * v})
            grp_total = sign * grp_abs
            sec_total += grp_total
            grupos.append({
                "id": grp_def["id"], "label": grp_def["label"],
                "sign": sign, "total": grp_total, "itens": itens,
                "highlight": grp_def.get("highlight", False),
                "note": grp_def.get("note", ""),
            })
        sections.append({"id": sec_def["id"], "label": sec_def["label"],
                         "total": sec_total, "grupos": grupos})

    def _s(sid):
        for s in sections:
            if s["id"] == sid:
                return s["total"]
        return 0.0

    rb     = _s("rb")
    ded    = _s("ded")
    rl     = rb + ded
    custos = _s("custos")
    margem = rl + custos
    sga    = _s("sga")
    ebitda = margem + sga
    ebit   = ebitda        # depreciação = 0
    rfin   = _s("rfin")
    rnop   = _s("rnop")
    lucro_antes_ir = ebit + rfin + rnop
    ir_csll = _DRE_IR_CSLL
    ll     = lucro_antes_ir - ir_csll
    inv    = _s("inv")
    financ = _s("financ")
    aporte = _s("aporte")
    fluxo_ac  = ll + inv + financ + aporte
    distrib   = _s("distrib")
    fluxo_liv = fluxo_ac + distrib

    # Add %RL to each section (informational; template may display for L1 rows)
    for sec in sections:
        sec["pct"] = sec["total"] / rl if rl else 0.0

    return {
        "sections": sections,
        "receita_bruta": rb,  "deducoes": ded,
        "receita_liquida": rl,
        "custos": custos,
        "margem": margem,       "pct_margem": margem / rl if rl else 0,
        "sga": sga,
        "ebitda": ebitda,       "pct_ebitda": ebitda / rl if rl else 0,
        "depreciacao": 0.0,
        "ebit": ebit,           "pct_ebit": ebit / rl if rl else 0,
        "rfin": rfin,           "rnop": rnop,
        "lucro_antes_ir": lucro_antes_ir,
        "pct_lajir": lucro_antes_ir / rl if rl else 0,
        "ir_csll": ir_csll,
        "lucro_liquido": ll,    "pct_ll": ll / rl if rl else 0,
        "inv": inv,  "financ": financ,  "aporte": aporte,
        "fluxo_acionista": fluxo_ac,
        "distrib": distrib,
        "fluxo_livre": fluxo_liv,
    }


@app.route("/dre")
def pagina_dre():
    from calendar import monthrange
    hoje = datetime.now(_BRT)
    try:
        mes = int(request.args.get("mes", hoje.month))
        ano = int(request.args.get("ano", hoje.year))
        if not (1 <= mes <= 12):
            mes = hoje.month
    except (ValueError, TypeError):
        mes, ano = hoje.month, hoje.year

    acumulado = request.args.get("acumulado", "0") == "1"

    # Build current period date range
    if acumulado:
        d_ini = datetime(ano, 1, 1)
        d_fim = datetime(ano, mes, monthrange(ano, mes)[1], 23, 59, 59)
    else:
        d_ini = datetime(ano, mes, 1)
        d_fim = datetime(ano, mes, monthrange(ano, mes)[1], 23, 59, 59)

    # Build previous period date range (previous month, or same period -1 year for acumulado)
    if acumulado:
        d_ini_prev = datetime(ano - 1, 1, 1)
        d_fim_prev = datetime(ano - 1, mes, monthrange(ano - 1, mes)[1], 23, 59, 59)
    else:
        prev_mes = mes - 1 if mes > 1 else 12
        prev_ano = ano if mes > 1 else ano - 1
        d_ini_prev = datetime(prev_ano, prev_mes, 1)
        d_fim_prev = datetime(prev_ano, prev_mes, monthrange(prev_ano, prev_mes)[1], 23, 59, 59)

    todos = _dre_ler_lancamentos()

    def _filtrar(d0, d1):
        return [l for l in todos if d0 <= l["dt"] <= d1]

    dre_atual = _dre_calcular(_filtrar(d_ini, d_fim))
    dre_prev  = _dre_calcular(_filtrar(d_ini_prev, d_fim_prev))

    meses_disponiveis = sorted({(l["dt"].year, l["dt"].month) for l in todos})

    return render_template("dre.html",
        active="dre",
        dre=dre_atual,
        dre_prev=dre_prev,
        mes=mes, ano=ano,
        acumulado=acumulado,
        meses_disponiveis=meses_disponiveis,
        label_atual=_nome_mes_label(mes, ano, acumulado),
        label_prev=_nome_mes_label(d_ini_prev.month, d_ini_prev.year, acumulado),
    )


# ── Financiamentos & Consórcios ───────────────────────────────────────────────

@app.route("/financiamentos")
def pagina_financiamentos():
    from math import ceil
    hoje = datetime.now(_BRT).date()

    sb   = _supabase()
    rows = sb.table("financiamentos_contratos").select("*").order("created_at").execute().data or []

    def _restante(vcto_str):
        if not vcto_str:
            return 0
        try:
            vcto = date.fromisoformat(str(vcto_str)[:10])
        except Exception:
            return 0
        dias = (vcto - hoje).days
        return ceil(dias / 30.44) if dias > 0 else 0

    contratos = []
    for r in rows:
        parcelas = int(r["parcelas_total"])
        parcela  = float(r["valor_parcela"])
        restante = _restante(r.get("data_vencimento"))
        pagas    = parcelas - restante

        contratos.append({
            "operacao":        r["operacao"],
            "contrato":        r.get("contrato") or "",
            "placa":           r.get("placa") or "",
            "data_vencimento": str(r.get("data_vencimento") or "")[:10] or None,
            "restante":        restante,
            "parcelas_total":  parcelas,
            "valor_parcela":   parcela,
            "total_pago":      pagas * parcela,
            "total":           parcelas * parcela,
            "devedor":         restante * parcela,
            "c_prazo":         min(restante, 12) * parcela,
            "l_prazo":         max(restante - 12, 0) * parcela,
            "pct_quitado":     pagas / parcelas if parcelas else 0,
            "quitado":         restante == 0,
        })

    contratos.sort(key=lambda x: x["pct_quitado"], reverse=True)

    ativos   = [c for c in contratos if not c["quitado"]]
    quitados = [c for c in contratos if     c["quitado"]]

    soma_devedor    = sum(c["devedor"]       for c in ativos)
    soma_total_pago = sum(c["total_pago"]    for c in contratos)
    soma_c_prazo    = sum(c["c_prazo"]       for c in ativos)
    soma_l_prazo    = sum(c["l_prazo"]       for c in ativos)
    soma_mensal     = sum(c["valor_parcela"] for c in ativos)
    tempo_medio     = sum(c["restante"]      for c in ativos) / len(ativos) if ativos else 0

    ativos_vcto = [c for c in ativos if c["data_vencimento"]]
    mais_perto  = min(ativos_vcto, key=lambda x: x["data_vencimento"])["operacao"] if ativos_vcto else "—"

    cards = {
        "saldo_devedor": soma_devedor,
        "total_pago":    soma_total_pago,
        "curto_prazo":   soma_c_prazo,
        "longo_prazo":   soma_l_prazo,
        "valor_mensal":  soma_mensal,
        "tempo_medio":   round(tempo_medio, 1),
        "mais_perto":    mais_perto,
        "r_quitados":    sum(c["total_pago"] for c in quitados),
        "pct_cp":        soma_c_prazo / soma_devedor if soma_devedor else 0,
        "pct_lp":        soma_l_prazo / soma_devedor if soma_devedor else 0,
    }

    return render_template("financiamentos.html",
        active="financiamentos",
        contratos=contratos,
        cards=cards,
    )


# ── Checklist ─────────────────────────────────────────────────────────────────

def _veiculos_xlsx_path():
    base = Path(__file__).resolve().parent
    return base / "data" / "veiculos.xlsx"


_IMAGEM_MAP = [
    ("GOL",     "gol_sf.png"),
    ("VOYAGE",  "voyage.png"),
    ("POLO",    "polo.png"),
    ("DOLPHIN", "byd.png"),
    ("SANDERO", "sandero.png"),
    ("CG",      "CG.png"),
    ("NXR",     "nxr.png"),
]
_BLEND_MULTIPLY = {"gol_sf.png", "voyage.png", "polo.png", "byd.png", "CG.png"}


def _imagem_veiculo(modelo):
    m = (modelo or "").upper()
    for keyword, fname in _IMAGEM_MAP:
        if keyword in m:
            return fname
    return None


def _ler_veiculos():
    import openpyxl
    xlsx_path = _veiculos_xlsx_path()
    if not xlsx_path.exists():
        return [], "Planilha não encontrada em data/veiculos.xlsx."

    try:
        wb = openpyxl.load_workbook(str(xlsx_path), read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()

        if len(rows) <= 4:
            return [], "Planilha sem dados suficientes (cabeçalho esperado na linha 5)."

        header_row = rows[4]

        def _norm(s):
            s = unicodedata.normalize("NFD", str(s or "").lower())
            return "".join(c for c in s if unicodedata.category(c) != "Mn")

        headers_norm = [_norm(h) for h in header_row]

        def _ci(keyword):
            kn = _norm(keyword)
            return next((i for i, h in enumerate(headers_norm) if kn in h), None)

        i_placa    = _ci("placa")
        i_modelo   = _ci("modelo")
        i_cliente  = _ci("razao social cliente") or _ci("razao social") or _ci("cliente")
        i_contrato = _ci("contrato de locacao") or _ci("contrato")
        i_unidade  = _ci("unidade do veiculo") or _ci("unidade")
        i_inicio   = _ci("inicio de contrato") or _ci("inicio")
        i_termino  = _ci("termino de contrato") or _ci("termino")

        def _v(row, i):
            if i is None or i >= len(row):
                return ""
            v = row[i]
            if v is None:
                return ""
            if hasattr(v, "strftime"):
                return v.strftime("%d/%m/%Y")
            return str(v).strip()

        veiculos = []
        for row in rows[5:]:
            placa = _v(row, i_placa)
            if not placa:
                continue
            modelo = _v(row, i_modelo)
            img    = _imagem_veiculo(modelo)
            veiculos.append({
                "placa":    placa,
                "modelo":   modelo,
                "cliente":  _v(row, i_cliente),
                "contrato": _v(row, i_contrato),
                "unidade":  _v(row, i_unidade),
                "inicio":   _v(row, i_inicio),
                "termino":  _v(row, i_termino),
                "imagem":   img,
                "blend":    img in _BLEND_MULTIPLY if img else False,
            })
        veiculos.sort(key=lambda v: v["cliente"].lower())
        return veiculos, None

    except Exception as e:
        import traceback; traceback.print_exc()
        return [], str(e)


@app.route("/checklist")
def pagina_checklist():
    veiculos, erro_leitura = _ler_veiculos()

    badge_data = {}
    sb = _supabase()
    if sb:
        try:
            contratos_res = sb.table("checklist_contratos").select("id, contrato").execute()
            if contratos_res.data:
                ids_map = {r["id"]: r["contrato"] for r in contratos_res.data}
                itens_res = sb.table("checklist_itens").select("contrato_id, marcado").execute()
                for item in (itens_res.data or []):
                    cid  = item["contrato_id"]
                    cnum = ids_map.get(cid)
                    if cnum:
                        if cnum not in badge_data:
                            badge_data[cnum] = {"total": 0, "marcados": 0}
                        badge_data[cnum]["total"] += 1
                        if item["marcado"]:
                            badge_data[cnum]["marcados"] += 1
        except Exception:
            pass

    return render_template("checklist.html",
                           active="checklist",
                           veiculos=veiculos,
                           badge_data=badge_data,
                           erro_leitura=erro_leitura)


@app.route("/api/checklist/contrato")
def api_checklist_get():
    contrato = request.args.get("contrato", "").strip()
    placa    = request.args.get("placa", "").strip()
    cliente  = request.args.get("cliente", "").strip()
    unidade  = request.args.get("unidade", "").strip()

    if not contrato:
        return jsonify({"error": "Contrato obrigatório"}), 400

    sb = _supabase()
    if not sb:
        return jsonify({"error": "Supabase indisponível"}), 503

    ITENS_PADRAO = ["INDICAÇÃO DE CONDUTOR", "CONTRATO", "NOTA PROMISSÓRIA", "CHAVE RESERVA"]

    res = sb.table("checklist_contratos").select("*").eq("contrato", contrato).execute()

    if res.data:
        rec        = res.data[0]
        contrato_id = rec["id"]
        itens_res  = sb.table("checklist_itens").select("*").eq("contrato_id", contrato_id).order("created_at").execute()
        itens = itens_res.data or []
        if not itens:
            for nome in ITENS_PADRAO:
                sb.table("checklist_itens").insert({"contrato_id": contrato_id, "nome": nome, "marcado": False}).execute()
            itens = sb.table("checklist_itens").select("*").eq("contrato_id", contrato_id).order("created_at").execute().data or []
    else:
        ins = sb.table("checklist_contratos").insert({
            "contrato": contrato, "placa": placa, "cliente": cliente, "unidade": unidade,
        }).execute()
        rec         = ins.data[0]
        contrato_id = rec["id"]
        for nome in ITENS_PADRAO:
            sb.table("checklist_itens").insert({"contrato_id": contrato_id, "nome": nome, "marcado": False}).execute()
        itens = sb.table("checklist_itens").select("*").eq("contrato_id", contrato_id).order("created_at").execute().data or []

    return jsonify({
        "contrato_id": contrato_id,
        "contrato":    contrato,
        "placa":       rec.get("placa", placa),
        "cliente":     rec.get("cliente", cliente),
        "unidade":     rec.get("unidade", unidade),
        "itens": [{"id": i["id"], "nome": i["nome"], "marcado": i["marcado"]} for i in itens],
    })


@app.route("/api/checklist/salvar", methods=["POST"])
def api_checklist_salvar():
    data        = request.get_json(force=True)
    contrato_id = data.get("contrato_id")
    itens       = data.get("itens", [])

    if not contrato_id:
        return jsonify({"error": "contrato_id obrigatório"}), 400

    sb = _supabase()
    if not sb:
        return jsonify({"error": "Supabase indisponível"}), 503

    try:
        for item in itens:
            if item.get("is_new"):
                sb.table("checklist_itens").insert({
                    "contrato_id": contrato_id,
                    "nome":        item["nome"],
                    "marcado":     item.get("marcado", False),
                }).execute()
            else:
                sb.table("checklist_itens").update({
                    "marcado": bool(item.get("marcado", False))
                }).eq("id", item["id"]).execute()
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True)
