"""
Baixa e substitui planilhas do sistema Ativuz automaticamente.

Uso:
  python atualizar_planilha.py --setup     # primeira vez: abre o browser para login
  python atualizar_planilha.py             # execução normal (headless), usada pelo agendador
  python atualizar_planilha.py --visivel   # roda visível (para depurar sem salvar sessão)

Requisitos:
  pip install playwright
  playwright install chromium
"""

import asyncio
import argparse
import calendar
import shutil
import sys
from datetime import date, datetime
from pathlib import Path

from playwright.async_api import async_playwright, Page

# ─── CAMINHOS ────────────────────────────────────────────────────────────────

ROOT        = Path(__file__).parent.parent
PERFIL_DIR  = Path.home() / ".playwright_ativuz"

DEST_CONTAS = ROOT / "planilhas" / "CONTAS-A-RECEBER.xlsx"

# ─────────────────────────────────────────────────────────────────────────────


def log(msg: str):
    print(f"[{datetime.now().strftime('%d/%m %H:%M:%S')}] {msg}", flush=True)


def ultimo_dia_mes() -> str:
    """Retorna o último dia do mês atual no formato dd/mm/aaaa."""
    hoje  = date.today()
    ultimo = calendar.monthrange(hoje.year, hoje.month)[1]
    return f"{ultimo:02d}/{hoje.month:02d}/{hoje.year}"


async def preencher_campo(page: Page, rotulo: str, valor: str, nth: int = 0):
    """Preenche um campo por label, placeholder ou posição."""
    campo = None

    # 1) por label associado (for/id)
    try:
        loc = page.get_by_label(rotulo, exact=False)
        if await loc.count() > 0:
            campo = loc.first
    except Exception:
        pass

    # 2) por texto do label como elemento próximo ao input
    if campo is None:
        try:
            loc = page.locator(f"label:has-text('{rotulo}') + input, label:has-text('{rotulo}') ~ input")
            if await loc.count() > 0:
                campo = loc.first
        except Exception:
            pass

    # 3) por placeholder
    if campo is None:
        try:
            loc = page.locator(f"input[placeholder*='{rotulo}']")
            if await loc.count() > 0:
                campo = loc.first
        except Exception:
            pass

    # 4) fallback: nth input visível da página
    if campo is None:
        campo = page.locator("input:visible").nth(nth)

    await campo.click()
    await campo.fill(valor)
    await campo.press("Tab")


async def baixar_contas_a_receber(page: Page) -> bool:
    """
    Relatório: https://app.bluefleet.com.br/analytics/report/72
    Campos:
      - Data Inicial: 01/08/2025
      - Data Final:   último dia do mês atual
      - Conta de Recebimento: 1.0 [OFICIAL] SICREDI | BANCO GESTOR
      - Demais campos: em branco
    """
    nome = "CONTAS-A-RECEBER.xlsx"
    url  = "https://app.bluefleet.com.br/analytics/report/72?showingAllReports=True"

    log(f"[{nome}] Acessando relatório diretamente...")
    await page.goto(url, wait_until="networkidle", timeout=40_000)
    await page.wait_for_timeout(2_000)

    # Verifica se caiu na tela de login
    if "login" in page.url.lower() or await page.locator("input[type='password']").count() > 0:
        log(f"[{nome}] ERRO: sessão expirada. Rode --setup novamente para fazer login.")
        return False

    # ── Data Inicial ──────────────────────────────────────────────────────────
    log(f"[{nome}] Preenchendo Data Inicial: 01/08/2025")
    campo_ini = page.locator("#txtParameterStartDate")
    await campo_ini.wait_for(state="visible", timeout=10_000)
    await campo_ini.fill("01/08/2025")
    await campo_ini.press("Tab")
    await page.wait_for_timeout(400)

    # ── Data Final ────────────────────────────────────────────────────────────
    data_final = ultimo_dia_mes()
    log(f"[{nome}] Preenchendo Data Final: {data_final}")
    campo_fim = page.locator("#txtParameterEndDate")
    await campo_fim.fill(data_final)
    await campo_fim.press("Tab")
    await page.wait_for_timeout(400)

    # ── Conta de Recebimento (campo autocomplete type=search) ─────────────────
    log(f"[{nome}] Preenchendo Conta de Recebimento...")
    conta_texto = "1.0 [OFICIAL] SICREDI | BANCO GESTOR"
    campo_conta = page.locator("input[type='search']").first
    await campo_conta.click()
    await campo_conta.fill("SICREDI")
    await page.wait_for_timeout(1_200)
    # Seleciona o primeiro item da lista que aparecer
    opcao = page.locator("li, .autocomplete-item, [role='option']").filter(has_text="SICREDI").first
    if await opcao.count() > 0:
        await opcao.click()
    else:
        # Tenta digitar o texto completo e pressionar Enter
        await campo_conta.fill(conta_texto)
        await campo_conta.press("Enter")
    await page.wait_for_timeout(600)

    # ── Gerar Relatório ───────────────────────────────────────────────────────
    log(f"[{nome}] Clicando em Gerar Relatório...")
    await page.get_by_role("button", name="Gerar Relatório", exact=False).first.click()
    await page.wait_for_load_state("networkidle", timeout=30_000)
    await page.wait_for_timeout(2_000)

    # ── Exportar ──────────────────────────────────────────────────────────────
    log(f"[{nome}] Clicando em Exportar...")
    await page.get_by_role("button", name="Exportar", exact=False).first.click()
    await page.wait_for_timeout(1_000)

    # ── Exportar Excel ────────────────────────────────────────────────────────
    log(f"[{nome}] Clicando em Exportar Excel...")
    backup = DEST_CONTAS.with_suffix(".xlsx.bak")
    if DEST_CONTAS.exists():
        shutil.copy2(DEST_CONTAS, backup)

    try:
        async with page.expect_download(timeout=60_000) as dl_info:
            await page.get_by_text("Exportar Excel", exact=False).first.click()

        download = await dl_info.value
        if download.failure():
            log(f"[{nome}] ERRO no download: {download.failure()}")
            if backup.exists():
                shutil.copy2(backup, DEST_CONTAS)
            return False

        await download.save_as(str(DEST_CONTAS))
        if backup.exists():
            backup.unlink()
        log(f"[{nome}] Salvo em: {DEST_CONTAS}")
        return True

    except Exception as ex:
        log(f"[{nome}] ERRO ao capturar download: {ex}")
        if backup.exists():
            shutil.copy2(backup, DEST_CONTAS)
        return False


# ─── LISTA DE ROTINAS ─────────────────────────────────────────────────────────
# Adicione mais funções acima e liste-as aqui.
ROTINAS = [
    baixar_contas_a_receber,
    # baixar_veiculos,  # adicionar depois
]


# ─── MAIN ─────────────────────────────────────────────────────────────────────

STATE_FILE = PERFIL_DIR / "state.json"


async def main(setup: bool, visivel: bool):
    PERFIL_DIR.mkdir(parents=True, exist_ok=True)

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=(not setup and not visivel),
            args=["--start-maximized"],
        )

        if setup:
            ctx  = await browser.new_context(viewport={"width": 1440, "height": 900}, accept_downloads=True)
            page = await ctx.new_page()
            log("Modo setup — abrindo Bluefleet para login...")
            await page.goto("https://app.bluefleet.com.br", wait_until="networkidle", timeout=30_000)
            log("Faça login no navegador que abriu.")
            log("Quando a página inicial carregar (menu Gestão visível), pressione ENTER.")
            await asyncio.get_event_loop().run_in_executor(None, input)
            await ctx.storage_state(path=str(STATE_FILE))
            log(f"Sessão salva em {STATE_FILE}. Rode sem --setup para testar.")
            await browser.close()
            return

        if not STATE_FILE.exists():
            log("ERRO: sessão não encontrada. Rode --setup primeiro.")
            await browser.close()
            sys.exit(1)

        ctx  = await browser.new_context(
            storage_state=str(STATE_FILE),
            viewport={"width": 1440, "height": 900},
            accept_downloads=True,
        )
        page = await ctx.new_page()

        erros = []
        for rotina in ROTINAS:
            try:
                ok = await rotina(page)
                if not ok:
                    erros.append(rotina.__name__)
            except Exception as ex:
                log(f"ERRO inesperado em {rotina.__name__}: {ex}")
                erros.append(rotina.__name__)

        # Atualiza o state após execução (renova tokens se o site os atualizou)
        await ctx.storage_state(path=str(STATE_FILE))
        await browser.close()

        if erros:
            log(f"Concluído com erros: {', '.join(erros)}")
            sys.exit(1)
        else:
            log("Todas as planilhas atualizadas com sucesso.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--setup",   action="store_true", help="Abre browser para login")
    parser.add_argument("--visivel", action="store_true", help="Roda com browser visível (debug)")
    args = parser.parse_args()
    asyncio.run(main(args.setup, args.visivel))
