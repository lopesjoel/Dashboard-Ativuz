"""
Baixa e substitui planilhas do sistema Ativuz automaticamente.

Uso:
  python atualizar_planilha.py --setup     # primeira vez: abre o browser para você fazer login
  python atualizar_planilha.py             # execução normal (headless), usada pelo agendador

Requisitos:
  pip install playwright
  playwright install chromium
"""

import asyncio
import argparse
import shutil
import sys
from datetime import datetime
from pathlib import Path

from playwright.async_api import async_playwright

# ─── CONFIGURAÇÃO ────────────────────────────────────────────────────────────

# Pasta onde o Playwright salva cookies/sessão (não apagar entre execuções)
PERFIL_DIR = Path.home() / ".playwright_ativuz"

# Lista de planilhas para baixar.
# Cada item: (url_da_pagina, seletor_do_botao_download, caminho_destino)
PLANILHAS = [
    {
        "nome": "veiculos.xlsx",
        "url": "PREENCHER_URL_AQUI",
        "seletor": "PREENCHER_SELETOR_AQUI",  # ex: "button:has-text('Exportar')"
        "destino": Path(__file__).parent.parent / "data" / "veiculos.xlsx",
    },
    # Descomente e preencha se precisar baixar CONTAS-A-RECEBER também:
    # {
    #     "nome": "CONTAS-A-RECEBER.xlsx",
    #     "url": "PREENCHER_URL_AQUI",
    #     "seletor": "PREENCHER_SELETOR_AQUI",
    #     "destino": Path(__file__).parent.parent / "docx_templates" / "CONTAS-A-RECEBER.xlsx",
    # },
]

# ─────────────────────────────────────────────────────────────────────────────


def log(msg: str):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


async def baixar(planilha: dict, page) -> bool:
    nome    = planilha["nome"]
    url     = planilha["url"]
    seletor = planilha["seletor"]
    destino = Path(planilha["destino"])

    log(f"Acessando página para {nome}...")
    await page.goto(url, wait_until="networkidle", timeout=30_000)

    log(f"Aguardando botão de download ({seletor})...")
    await page.wait_for_selector(seletor, timeout=15_000)

    backup = destino.with_suffix(".xlsx.bak")
    if destino.exists():
        shutil.copy2(destino, backup)
        log(f"Backup criado: {backup.name}")

    log(f"Iniciando download de {nome}...")
    async with page.expect_download(timeout=60_000) as dl_info:
        await page.click(seletor)

    download = await dl_info.value
    if download.failure():
        log(f"ERRO no download de {nome}: {download.failure()}")
        if backup.exists():
            shutil.copy2(backup, destino)
        return False

    await download.save_as(str(destino))
    if backup.exists():
        backup.unlink()

    log(f"OK — {nome} atualizado em {destino}")
    return True


async def main(setup: bool):
    PERFIL_DIR.mkdir(parents=True, exist_ok=True)

    async with async_playwright() as pw:
        ctx = await pw.chromium.launch_persistent_context(
            str(PERFIL_DIR),
            headless=not setup,          # visível no --setup, invisível no agendador
            accept_downloads=True,
            viewport={"width": 1280, "height": 900},
        )

        page = ctx.pages[0] if ctx.pages else await ctx.new_page()

        if setup:
            log("Modo setup — faça login no navegador que abriu.")
            log("Quando terminar o login, pressione ENTER aqui para salvar a sessão.")
            await asyncio.get_event_loop().run_in_executor(None, input)
            await ctx.storage_state(path=str(PERFIL_DIR / "state.json"))
            log("Sessão salva. Feche o browser e rode sem --setup para testar.")
            await ctx.close()
            return

        erros = []
        for planilha in PLANILHAS:
            ok = await baixar(planilha, page)
            if not ok:
                erros.append(planilha["nome"])

        await ctx.close()

        if erros:
            log(f"Concluído com erros: {', '.join(erros)}")
            sys.exit(1)
        else:
            log("Todas as planilhas atualizadas com sucesso.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--setup", action="store_true",
                        help="Abre o browser para fazer login e salvar a sessão")
    args = parser.parse_args()
    asyncio.run(main(args.setup))
