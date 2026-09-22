import os
import re
import glob
import time
import asyncio
import logging
from datetime import datetime, timezone
from dotenv import load_dotenv
from playwright.async_api import async_playwright
import pandas as pd

from gmail_helper import obter_codigo_verificacao

# Carrega as variáveis de ambiente (.env)
load_dotenv()

GMAIL_USER = os.getenv("GMAIL_USER")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD")
SHAREPOINT_URL = os.getenv("SHAREPOINT_URL")
STATE_FILE = "state.json"

# Configuração dinâmica da pasta de destino (Lê do .env ou usa 'downloads' por padrão)
ENV_OUTPUT_DIR = os.getenv("OUTPUT_DIR", "").strip()
if ENV_OUTPUT_DIR:
    DOWNLOAD_DIR = os.path.abspath(ENV_OUTPUT_DIR)
else:
    DOWNLOAD_DIR = os.path.abspath("downloads")

# Garantia da existência da pasta de destino
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# -------------------------------------------------------------
# CONFIGURAÇÃO DE LOGGING (ARQUIVO + TERMINAL)
# -------------------------------------------------------------
LOG_FILE = os.path.join(DOWNLOAD_DIR, "execucao.log")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler()
    ]
)

# Remove o arquivo de sessão antigo caso exista
if os.path.exists(STATE_FILE):
    try:
        os.remove(STATE_FILE)
        logging.info(f"🧹 Arquivo de sessão antigo '{STATE_FILE}' removido para garantir fluxo do zero.")
    except Exception as e:
        logging.warning(f"⚠️ Não foi possível remover {STATE_FILE}: {e}")


async def clicar_elemento_em_frames(page, textos):
    """
    Procura e clica num elemento pelo texto, buscando no documento principal
    e em todos os IFrames ativos na página.
    """
    for frame in page.frames:
        for texto in textos:
            try:
                locator = frame.get_by_text(texto, exact=True).first
                if await locator.is_visible(timeout=1000):
                    await locator.click()
                    return True
            except Exception:
                continue
    return False


def processar_e_salvar_planilha(caminho_arquivo):
    """
    Realiza o pós-processamento da planilha:
    1. Lê a planilha baixada removendo as 4 primeiras linhas (cola como valor).
    2. Exclui a 8ª coluna (índice 7) e depois a 1ª coluna (índice 0).
    3. Salva a planilha com a data atual no formato DD-MM-YYYY.xlsx na pasta de saída.
    """
    data_hoje = datetime.now().strftime("%d-%m-%Y")
    logging.info(f"📂 Processando arquivo baixado: {os.path.basename(caminho_arquivo)}")

    # 1. Lê a planilha pulando as 4 primeiras linhas (linhas 1, 2, 3 e 4 do Excel)
    df = pd.read_excel(caminho_arquivo, skiprows=4)

    total_colunas = df.shape[1]
    logging.info(f"📊 Quantidade inicial de colunas tratadas: {total_colunas}")

    # 2. Excluir a 8ª coluna (índice 7)
    if total_colunas >= 8:
        coluna_8_nome = df.columns[7]
        df.drop(df.columns[7], axis=1, inplace=True)
        logging.info(f"🗑️ 8ª coluna removida: '{coluna_8_nome}'")
    else:
        logging.warning("⚠️ A planilha possui menos de 8 colunas. Remoção da 8ª coluna ignorada.")

    # 3. Excluir a 1ª coluna (índice 0)
    if df.shape[1] >= 1:
        coluna_1_nome = df.columns[0]
        df.drop(df.columns[0], axis=1, inplace=True)
        logging.info(f"🗑️ 1ª coluna removida: '{coluna_1_nome}'")

    # 4. Salvar novo arquivo com a data atual na pasta configurada
    nome_novo_arquivo = f"{data_hoje}.xlsx"
    caminho_novo_arquivo = os.path.join(DOWNLOAD_DIR, nome_novo_arquivo)

    df.to_excel(caminho_novo_arquivo, index=False)

    logging.info(f"🎉 PÓS-PROCESSAMENTO CONCLUÍDO COM SUCESSO! Planilha salva em: {caminho_novo_arquivo}")
    return caminho_novo_arquivo


async def main():
    logging.info("=" * 60)
    logging.info("🚀 INICIANDO EXECUÇÃO DA AUTOMAÇÃO DE DOWNLOAD")
    logging.info(f"📂 Pasta de destino e logs: {DOWNLOAD_DIR}")
    logging.info("=" * 60)

    if not GMAIL_USER or not GMAIL_APP_PASSWORD or not SHAREPOINT_URL:
        logging.error("❌ ERRO CRÍTICO: GMAIL_USER, GMAIL_APP_PASSWORD ou SHAREPOINT_URL ausentes no .env!")
        return

    try:
        async with async_playwright() as p:
            browser = await p.chromium.launch(
                headless=True,
                args=["--no-sandbox", "--disable-setuid-sandbox"]
            )

            context = await browser.new_context(
                accept_downloads=True,
                locale="pt-BR",
                viewport={"width": 1280, "height": 720}
            )

            page = await context.new_page()

            logging.info(f"🌐 Acessando SharePoint: {SHAREPOINT_URL}")
            await page.goto(SHAREPOINT_URL, wait_until="domcontentloaded")
            await page.wait_for_timeout(2000)
            await page.screenshot(path=os.path.join(DOWNLOAD_DIR, "01_pagina_carregada.png"))

            # -------------------------------------------------------------
            # 1. AUTENTICAÇÃO
            # -------------------------------------------------------------
            logging.info("🟡 Preenchendo e-mail de acesso...")
            email_locator = page.locator("input[type='email'], input[name='i0116'], input[id='i0116']").first
            await email_locator.wait_for(state="visible", timeout=15000)
            
            await email_locator.click()
            await email_locator.fill(GMAIL_USER)
            await page.screenshot(path=os.path.join(DOWNLOAD_DIR, "02_email_preenchido.png"))

            logging.info("🚀 Enviando e-mail de verificação...")
            timestamp_solicitacao = datetime.now(timezone.utc)

            btn_enviar = page.locator("#idSIButton9, input[type='submit']").first
            if await btn_enviar.is_visible():
                await btn_enviar.click()
            else:
                await email_locator.press("Enter")

            await page.wait_for_timeout(1500)
            await page.screenshot(path=os.path.join(DOWNLOAD_DIR, "03_logo_apos_clicar_next.png"))

            logging.info("⏳ Localizando campo OTP (#txtTOAACode)...")
            otp_locator = page.locator("#txtTOAACode, input[name='txtTOAACode']").first
            await otp_locator.wait_for(state="visible", timeout=20000)
            await page.screenshot(path=os.path.join(DOWNLOAD_DIR, "04_tela_otp_visivel.png"))

            logging.info("📩 Consultando código no Gmail via IMAP...")
            codigo_otp = await asyncio.to_thread(
                obter_codigo_verificacao, 
                GMAIL_USER, 
                GMAIL_APP_PASSWORD, 
                timestamp_solicitacao
            )

            logging.info(f"🔑 Preenchendo código OTP: {codigo_otp}")
            await otp_locator.click()
            await otp_locator.fill(codigo_otp)
            await page.screenshot(path=os.path.join(DOWNLOAD_DIR, "05_otp_digitado.png"))

            logging.info("🚀 Confirmando código OTP...")
            btn_verificar = page.locator("#btnSubmitCode, input[name='btnSubmitCode']").first
            if await btn_verificar.is_visible():
                await btn_verificar.click()
            else:
                await otp_locator.press("Enter")

            await page.wait_for_timeout(5000)
            await page.screenshot(path=os.path.join(DOWNLOAD_DIR, "06_pos_submissao_otp.png"))

            # -------------------------------------------------------------
            # 2. DOWNLOAD VIA FLUENT UI (EXCEL ONLINE)
            # -------------------------------------------------------------
            logging.info("📊 Aguardando carregamento da interface da planilha (12s)...")
            await page.wait_for_timeout(12000)
            await page.screenshot(path=os.path.join(DOWNLOAD_DIR, "07_excel_carregado.png"))

            logging.info("📥 1. Abrindo menu 'Arquivo' / 'Filer'...")
            clicou = await clicar_elemento_em_frames(page, ["Arquivo", "Filer", "File"])
            
            if not clicou:
                logging.info("⚠️ Tentando atalho de teclado Alt+A / Alt+F...")
                await page.keyboard.press("Alt+a")
                await page.wait_for_timeout(1000)
                await page.screenshot(path=os.path.join(DOWNLOAD_DIR, "07b_tentativa_atalho.png"))
                if not await page.get_by_text("Criar uma Cópia", exact=True).is_visible():
                    await page.keyboard.press("Alt+f")

            await page.wait_for_timeout(2000)
            await page.screenshot(path=os.path.join(DOWNLOAD_DIR, "08_menu_arquivo_aberto.png"))

            logging.info("📥 2. Clicando em 'Criar uma Cópia'...")
            clicou_copia = await clicar_elemento_em_frames(page, ["Criar uma Cópia", "Opret en kopi", "Save As", "Salvar como"])
            if not clicou_copia:
                await page.locator("text=/Criar uma Cópia/i, text=/Opret en kopi/i").first.click()

            await page.wait_for_timeout(2000)
            await page.screenshot(path=os.path.join(DOWNLOAD_DIR, "09_sub_menu_copia_aberto.png"))

            logging.info("📥 3. Clicando em 'Baixar uma Cópia'...")
            async with page.expect_download(timeout=60000) as download_info:
                clicou_baixar = await clicar_elemento_em_frames(page, ["Baixar uma Cópia", "Download en kopi", "Download a Copy"])
                if not clicou_baixar:
                    await page.locator("text=/Baixar uma Cópia/i, text=/Download en kopi/i").first.click()

                await page.screenshot(path=os.path.join(DOWNLOAD_DIR, "10_clique_baixar_copia.png"))
                
                await page.wait_for_timeout(2000)
                for frame in page.frames:
                    try:
                        btn_confirmar = frame.locator("button:has-text('Baixar uma Cópia'), button:has-text('Download'), button:has-text('Baixar')").last
                        if await btn_confirmar.is_visible(timeout=1000):
                            logging.info("🔘 Clicando na confirmação do modal de download...")
                            await page.screenshot(path=os.path.join(DOWNLOAD_DIR, "11_modal_confirmacao.png"))
                            await btn_confirmar.click()
                            break
                    except Exception:
                        continue

            download = await download_info.value
            caminho_final = os.path.join(DOWNLOAD_DIR, download.suggested_filename)
            await download.save_as(caminho_final)

            logging.info(f"🎉 DOWNLOAD CONCLUÍDO COM SUCESSO: {caminho_final}")

            # -------------------------------------------------------------
            # 3. PÓS-PROCESSAMENTO DA PLANILHA
            # -------------------------------------------------------------
            logging.info("⚙️ Iniciando pós-processamento com Pandas...")
            processar_e_salvar_planilha(caminho_final)

            await page.screenshot(path=os.path.join(DOWNLOAD_DIR, "12_download_concluido.png"))
            await context.close()
            await browser.close()
            
            logging.info("✅ EXECUÇÃO FINALIZADA COM SUCESSO!\n")

    except Exception as e:
        logging.error(f"❌ OCORREU UM ERRO DURANTE A EXECUÇÃO DA AUTOMAÇÃO:", exc_info=True)
        logging.error("❌ EXECUÇÃO FINALIZADA COM ERROS.\n")

if __name__ == "__main__":
    asyncio.run(main()) 