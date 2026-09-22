import os
import sys
import glob
import time
import shutil
import asyncio
import logging
from datetime import datetime, timezone
from dotenv import load_dotenv

# -------------------------------------------------------------
# 1. CONFIGURAÇÃO DE DIRETÓRIOS E LOGGING
# -------------------------------------------------------------
load_dotenv()

# Pasta base do executável (onde o .exe está rodando)
if getattr(sys, 'frozen', False):
    EXE_DIR = os.path.dirname(sys.executable)
else:
    EXE_DIR = os.path.dirname(os.path.abspath(__file__))

# Logs e Screenshots sempre salvos na mesma pasta do executável
LOG_DIR = EXE_DIR
LOG_FILE = os.path.join(LOG_DIR, "execucao.log")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler()
    ]
)

# Captura de erros fatais
def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    logging.critical("💥 ERRO CRÍTICO NÃO TRATADO NA EXECUÇÃO:", exc_info=(exc_type, exc_value, exc_traceback))

sys.excepthook = handle_exception

logging.info("=" * 60)
logging.info("🚀 INICIANDO APLICAÇÃO")
logging.info(f"📂 Pasta local do executável (Logs/Screenshots): {EXE_DIR}")
logging.info("=" * 60)

# Importações após logger pronto
try:
    from playwright.async_api import async_playwright
    import pandas as pd
    from gmail_helper import obter_codigo_verificacao
except Exception as err:
    logging.critical("💥 ERRO AO IMPORTAR BIBLIOTECAS NECESSÁRIAS:", exc_info=True)
    sys.exit(1)

# Leitura e tratamento dos caminhos de saída
ENV_OUTPUT_DIR = os.getenv("OUTPUT_DIR", "").strip().strip('"').strip("'")
ENV_OUTPUT_DIR_2 = os.getenv("OUTPUT_DIR_2", "").strip().strip('"').strip("'")

DOWNLOAD_DIR = os.path.normpath(ENV_OUTPUT_DIR) if ENV_OUTPUT_DIR else os.path.join(EXE_DIR, "downloads")
OUTPUT_DIR_2 = os.path.normpath(ENV_OUTPUT_DIR_2) if ENV_OUTPUT_DIR_2 else None

# Garante criação das pastas de saída
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
if OUTPUT_DIR_2:
    os.makedirs(OUTPUT_DIR_2, exist_ok=True)

GMAIL_USER = os.getenv("GMAIL_USER")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD")
SHAREPOINT_URL = os.getenv("SHAREPOINT_URL")
STATE_FILE = os.path.join(EXE_DIR, "state.json")

# Limpeza de sessão antiga
if os.path.exists(STATE_FILE):
    try:
        os.remove(STATE_FILE)
        logging.info(f"🧹 Arquivo de sessão antigo '{STATE_FILE}' removido.")
    except Exception as e:
        logging.warning(f"⚠️ Não foi possível remover {STATE_FILE}: {e}")


async def clicar_elemento_em_frames(page, textos):
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
    Trata a planilha baixada e salva nas duas pastas configuradas.
    """
    data_hoje = datetime.now().strftime("%d-%m-%Y")
    logging.info(f"📂 Processando arquivo baixado: {os.path.basename(caminho_arquivo)}")

    # 1. Leitura e exclusão de colunas
    df = pd.read_excel(caminho_arquivo, skiprows=4)
    total_colunas = df.shape[1]
    logging.info(f"📊 Quantidade inicial de colunas tratadas: {total_colunas}")

    if total_colunas >= 8:
        coluna_8_nome = df.columns[7]
        df.drop(df.columns[7], axis=1, inplace=True)
        logging.info(f"🗑️ 8ª coluna removida: '{coluna_8_nome}'")
    else:
        logging.warning("⚠️ A planilha possui menos de 8 colunas. Remoção da 8ª coluna ignorada.")

    if df.shape[1] >= 1:
        coluna_1_nome = df.columns[0]
        df.drop(df.columns[0], axis=1, inplace=True)
        logging.info(f"🗑️ 1ª coluna removida: '{coluna_1_nome}'")

    # 2. Salva no 1º destino (Com data: DD-MM-YYYY.xlsx)
    nome_novo_arquivo = f"{data_hoje}.xlsx"
    caminho_destino_1 = os.path.join(DOWNLOAD_DIR, nome_novo_arquivo)
    df.to_excel(caminho_destino_1, index=False)
    logging.info(f"🎉 Planilha salva com sucesso em (Destino 1): {caminho_destino_1}")

    # 3. Copia para o 2º destino (Nome fixo: Disponibilidade Marcelo.xlsx)
    if OUTPUT_DIR_2:
        caminho_destino_2 = os.path.join(OUTPUT_DIR_2, "Disponibilidade Marcelo.xlsx")
        try:
            shutil.copy2(caminho_destino_1, caminho_destino_2)
            logging.info(f"📋 Cópia criada com sucesso em (Destino 2): {caminho_destino_2}")
        except Exception as e:
            logging.error(f"❌ Erro ao copiar arquivo para o Destino 2 ({caminho_destino_2}): {e}")

    return caminho_destino_1


async def main():
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
            await page.screenshot(path=os.path.join(EXE_DIR, "01_pagina_carregada.png"))

            # 1. AUTENTICAÇÃO
            logging.info("🟡 Preenchendo e-mail de acesso...")
            email_locator = page.locator("input[type='email'], input[name='i0116'], input[id='i0116']").first
            await email_locator.wait_for(state="visible", timeout=15000)
            
            await email_locator.click()
            await email_locator.fill(GMAIL_USER)
            await page.screenshot(path=os.path.join(EXE_DIR, "02_email_preenchido.png"))

            logging.info("🚀 Enviando e-mail de verificação...")
            timestamp_solicitacao = datetime.now(timezone.utc)

            btn_enviar = page.locator("#idSIButton9, input[type='submit']").first
            if await btn_enviar.is_visible():
                await btn_enviar.click()
            else:
                await email_locator.press("Enter")

            await page.wait_for_timeout(1500)
            await page.screenshot(path=os.path.join(EXE_DIR, "03_logo_apos_clicar_next.png"))

            logging.info("⏳ Localizando campo OTP (#txtTOAACode)...")
            otp_locator = page.locator("#txtTOAACode, input[name='txtTOAACode']").first
            await otp_locator.wait_for(state="visible", timeout=20000)
            await page.screenshot(path=os.path.join(EXE_DIR, "04_tela_otp_visivel.png"))

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
            await page.screenshot(path=os.path.join(EXE_DIR, "05_otp_digitado.png"))

            logging.info("🚀 Confirmando código OTP...")
            btn_verificar = page.locator("#btnSubmitCode, input[name='btnSubmitCode']").first
            if await btn_verificar.is_visible():
                await btn_verificar.click()
            else:
                await otp_locator.press("Enter")

            await page.wait_for_timeout(5000)
            await page.screenshot(path=os.path.join(EXE_DIR, "06_pos_submissao_otp.png"))

            # 2. DOWNLOAD VIA FLUENT UI
            logging.info("📊 Aguardando carregamento da interface da planilha (12s)...")
            await page.wait_for_timeout(12000)
            await page.screenshot(path=os.path.join(EXE_DIR, "07_excel_carregado.png"))

            logging.info("📥 1. Abrindo menu 'Arquivo' / 'Filer'...")
            clicou = await clicar_elemento_em_frames(page, ["Arquivo", "Filer", "File"])
            
            if not clicou:
                logging.info("⚠️ Tentando atalho de teclado Alt+A / Alt+F...")
                await page.keyboard.press("Alt+a")
                await page.wait_for_timeout(1000)
                await page.screenshot(path=os.path.join(EXE_DIR, "07b_tentativa_atalho.png"))
                if not await page.get_by_text("Criar uma Cópia", exact=True).is_visible():
                    await page.keyboard.press("Alt+f")

            await page.wait_for_timeout(2000)
            await page.screenshot(path=os.path.join(EXE_DIR, "08_menu_arquivo_aberto.png"))

            logging.info("📥 2. Clicando em 'Criar uma Cópia'...")
            clicou_copia = await clicar_elemento_em_frames(page, ["Criar uma Cópia", "Opret en kopi", "Save As", "Salvar como"])
            if not clicou_copia:
                await page.locator("text=/Criar uma Cópia/i, text=/Opret en kopi/i").first.click()

            await page.wait_for_timeout(2000)
            await page.screenshot(path=os.path.join(EXE_DIR, "09_sub_menu_copia_aberto.png"))

            logging.info("📥 3. Clicando em 'Baixar uma Cópia'...")
            async with page.expect_download(timeout=60000) as download_info:
                clicou_baixar = await clicar_elemento_em_frames(page, ["Baixar uma Cópia", "Download en kopi", "Download a Copy"])
                if not clicou_baixar:
                    await page.locator("text=/Baixar uma Cópia/i, text=/Download en kopi/i").first.click()

                await page.screenshot(path=os.path.join(EXE_DIR, "10_clique_baixar_copia.png"))
                
                await page.wait_for_timeout(2000)
                for frame in page.frames:
                    try:
                        btn_confirmar = frame.locator("button:has-text('Baixar uma Cópia'), button:has-text('Download'), button:has-text('Baixar')").last
                        if await btn_confirmar.is_visible(timeout=1000):
                            logging.info("🔘 Clicando na confirmação do modal de download...")
                            await page.screenshot(path=os.path.join(EXE_DIR, "11_modal_confirmacao.png"))
                            await btn_confirmar.click()
                            break
                    except Exception:
                        continue

            download = await download_info.value
            caminho_temp = os.path.join(EXE_DIR, download.suggested_filename)
            await download.save_as(caminho_temp)

            logging.info(f"🎉 DOWNLOAD TEMPORÁRIO CONCLUÍDO: {caminho_temp}")

            # 3. PÓS-PROCESSAMENTO E ENVIO
            logging.info("⚙️ Iniciando pós-processamento com Pandas...")
            processar_e_salvar_planilha(caminho_temp)

            # Limpa o download bruto original temporário
            if os.path.exists(caminho_temp):
                try:
                    os.remove(caminho_temp)
                except Exception:
                    pass

            await page.screenshot(path=os.path.join(EXE_DIR, "12_download_concluido.png"))
            await context.close()
            await browser.close()
            
            logging.info("✅ EXECUÇÃO FINALIZADA COM SUCESSO!\n")

    except Exception as e:
        logging.error("❌ OCORREU UM ERRO DURANTE A EXECUÇÃO DA AUTOMAÇÃO:", exc_info=True)
        logging.error("❌ EXECUÇÃO FINALIZADA COM ERROS.\n")

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except Exception as e:
        logging.critical("💥 FALHA CRÍTICA AO EXECUTAR A APLICAÇÃO:", exc_info=True)