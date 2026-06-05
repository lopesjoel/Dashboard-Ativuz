"""
Importa snapshots históricos de inadimplência a partir do histórico git.

Para cada dia que o arquivo CONTAS-A-RECEBER.xlsx foi alterado, extrai a versão
do arquivo naquele commit, calcula os totais de inadimplência usando aquela data
como "hoje" (para os atrasos ficarem corretos), e salva no Supabase.

Uso:
    python importar_historico_inad.py [--dry-run]
"""

import os
import sys
import subprocess
import tempfile
import unicodedata
from datetime import date, datetime
from pathlib import Path

# ── Config ────────────────────────────────────────────────────────────────────
SUPABASE_URL = os.environ.get("SUPABASE_URL", "https://btcwgwrajgsqndnpjman.supabase.co")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImJ0Y3dnd3JhamdzcW5kbnBqbWFuIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc3NjU0MTc3NiwiZXhwIjoyMDkyMTE3Nzc2fQ.F_Ymbg9_U1f6DDVExuM2OXy7adxenwte0qT1Zzrn2hU")
REPO_ROOT    = Path(__file__).parent
XLSX_PATH    = "planilhas/CONTAS-A-RECEBER.xlsx"
DRY_RUN      = "--dry-run" in sys.argv

DIAS_PT = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
_NOMES_EXCLUIDOS = {"MARCELO BENTO DE ARAUJO"}


def _nh(s):
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.lower().strip()


def _parse_valor(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("R$", "").replace(" ", "")
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def calcular_snapshot(xlsx_bytes: bytes, hoje: date) -> dict:
    """Processa um XLSX em memória usando 'hoje' como data de referência."""
    import openpyxl
    import io
    from collections import Counter

    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    # detecta header
    header_idx = 0
    for ri, row in enumerate(rows[:10]):
        nh_row = [_nh(str(c or "")) for c in row]
        if sum(1 for t in ["receber de", "vencimento", "valor"]
               if any(t in n for n in nh_row)) >= 2:
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

    def _get(row, idx):
        return row[idx] if idx is not None and idx < len(row) else None

    name_counts = Counter()
    for row in data_rows:
        n = str(_get(row, i_nome) or "").strip()
        if n and n.upper() not in _NOMES_EXCLUIDOS:
            name_counts[n] += 1

    nomes_vencidos = set()
    total_valor    = 0.0
    criticos       = 0
    etapas = ["Hoje", "Terça-feira", "Quarta-feira", "Quinta-feira",
              "Sexta-feira", "D+5", "D+7", "D+10", "D+15"]
    por_etapa = {e: 0 for e in etapas}

    for row in data_rows:
        nome_raw = _get(row, i_nome)
        if not nome_raw:
            continue
        nome = str(nome_raw).strip()
        if not nome or nome.upper() in _NOMES_EXCLUIDOS:
            continue

        valor    = _parse_valor(_get(row, i_valor))
        venc_raw = _get(row, i_venc)
        sit_raw  = _get(row, i_sit)

        venc_date = None
        if venc_raw:
            if isinstance(venc_raw, datetime):
                venc_date = venc_raw.date()
            elif isinstance(venc_raw, date):
                venc_date = venc_raw
            else:
                venc_str = str(venc_raw).strip().upper()
                if venc_str == "HOJE":
                    venc_date = hoje
                else:
                    for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]:
                        try:
                            venc_date = datetime.strptime(venc_str, fmt).date()
                            break
                        except (ValueError, TypeError):
                            pass
        if venc_date is None:
            continue

        situacao = _nh(str(sit_raw or ""))
        if "a vencer" in situacao and venc_date > hoje:
            continue

        dias = (hoje - venc_date).days
        if dias < 0:
            continue

        # multa + juros
        multa      = valor * 0.10 if dias >= 2 else 0.0
        juros_mora = valor * 0.005 * dias if dias >= 3 else 0.0
        total      = valor + multa + juros_mora

        nomes_vencidos.add(nome)
        total_valor += total
        if dias >= 7:
            criticos += 1

        if dias == 0:    etapa = "Hoje"
        elif dias == 1:  etapa = "Terça-feira"
        elif dias == 2:  etapa = "Quarta-feira"
        elif dias == 3:  etapa = "Quinta-feira"
        elif dias == 4:  etapa = "Sexta-feira"
        elif dias <= 6:  etapa = "D+5"
        elif dias <= 9:  etapa = "D+7"
        elif dias <= 14: etapa = "D+10"
        else:            etapa = "D+15"

        if etapa in por_etapa:
            por_etapa[etapa] += 1

    return {
        "total_casos": len(nomes_vencidos),
        "total_valor": round(total_valor, 2),
        "criticos":    criticos,
        "por_etapa":   por_etapa,
    }


def get_commits_por_data():
    """Retorna dict {date: commit_hash} — último commit de cada dia."""
    result = subprocess.run(
        ["git", "log", "--format=%h %ad", "--date=short", "--", XLSX_PATH],
        capture_output=True, text=True, cwd=REPO_ROOT
    )
    commits = {}
    for line in result.stdout.strip().splitlines():
        parts = line.strip().split()
        if len(parts) >= 2:
            h, d = parts[0], parts[1]
            if d not in commits:       # primeiro aparecido = mais recente do dia
                commits[d] = h
    return commits


def extract_xlsx(commit_hash: str) -> bytes:
    result = subprocess.run(
        ["git", "show", f"{commit_hash}:{XLSX_PATH}"],
        capture_output=True, cwd=REPO_ROOT
    )
    if result.returncode != 0:
        raise RuntimeError(f"git show falhou para {commit_hash}: {result.stderr.decode()}")
    return result.stdout


def main():
    from supabase import create_client
    sb = create_client(SUPABASE_URL, SUPABASE_KEY)

    commits = get_commits_por_data()
    datas   = sorted(commits.keys())

    print(f"{'[DRY-RUN] ' if DRY_RUN else ''}Encontrados {len(datas)} dias com dados:\n")

    for data_str in datas:
        hash_ = commits[data_str]
        d     = date.fromisoformat(data_str)
        dia   = DIAS_PT[d.weekday()]

        try:
            xlsx_bytes = extract_xlsx(hash_)
            snap       = calcular_snapshot(xlsx_bytes, d)

            print(f"  {dia} {data_str}  |  {snap['total_casos']} casos  "
                  f"|  R$ {snap['total_valor']:,.2f}  |  {snap['criticos']} críticos", end="")

            if DRY_RUN:
                print("  [não salvo]")
                continue

            payload = {
                "semana":      data_str,
                "total_casos": snap["total_casos"],
                "total_valor": snap["total_valor"],
                "criticos":    snap["criticos"],
                "por_etapa":   snap["por_etapa"],
            }
            existing = sb.table("inad_snapshots").select("id").eq("semana", data_str).execute()
            if existing.data:
                sb.table("inad_snapshots").update(payload).eq("semana", data_str).execute()
                print("  [atualizado]")
            else:
                sb.table("inad_snapshots").insert(payload).execute()
                print("  [salvo]")

        except Exception as e:
            print(f"\n  ERRO em {data_str}: {e}")

    print("\nConcluído.")


if __name__ == "__main__":
    main()
