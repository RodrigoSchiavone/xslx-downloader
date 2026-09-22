import os
import time
from playwright.sync_api import sync_playwright
from gmail_helper import obter_codigo_verificacao

def baixar_planilha(url_sharepoint: str, email_user: str, app_password: str, pasta_destino: str = "./downloads"):
    """Executa a automação de login via OTP e faz o download do arquivo."""
    os.makedirs(pasta_destino, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        print("Acessando a página do SharePoint...")
        page.goto(url_sharepoint)

        # 1. Preenchimento do e-mail
        print("Preenchendo e-mail de acesso...")
        page.wait_for_selector('input[type="email"]')
        page.fill('input[type="email"]', email_user)
        page.click('input[type="submit"], button[type="submit"]')

        # 2. Obtenção do código via IMAP
        codigo_otp = obter_codigo_verificacao(email_user, app_password)

        # 3. Inserção do código OTP
        print("Preenchendo código OTP no SharePoint...")
        page.wait_for_selector('input[name="otc"], input[type="text"]')
        page.fill('input[name="otc"], input[type="text"]', codigo_otp)
        page.click('input[type="submit"], button[type="submit"]')

        # 4. Aguardar carregamento da planilha
        print("Aguardando carregamento da interface do Excel Online...")
        page.wait_for_load_state("networkidle")
        time.sleep(5)

        # 5. Execução do Download
        print("Iniciando o download do arquivo...")
        try:
            page.click('button:has-text("Arquivo"), button:has-text("File")')
            page.wait_for_timeout(1000)
            page.click('text="Salvar como", text="Save As"')
            page.wait_for_timeout(1000)

            with page.expect_download() as download_info:
                page.click('text="Baixar uma cópia", text="Download a Copy"')
            
            download = download_info.value
            caminho_final = os.path.join(pasta_destino, download.suggested_filename)
            download.save_as(caminho_final)
            print(f"Download concluído com sucesso: {caminho_final}")

        except Exception as e:
            print(f"Erro no menu visual: {e}. Tentando download via URL estendida...")
            url_direta = url_sharepoint.replace("action=default", "action=download")
            with page.expect_download() as download_info:
                page.goto(url_direta)
            download = download_info.value
            caminho_final = os.path.join(pasta_destino, download.suggested_filename)
            download.save_as(caminho_final)
            print(f"Download via URL direta concluído: {caminho_final}")

        browser.close()