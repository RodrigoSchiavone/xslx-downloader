import re
import time
import imaplib
import email
from datetime import datetime, timezone
from email.header import decode_header
from email.utils import parsedate_to_datetime


def decode_mime_header(header_value: str) -> str:
    """Decodifica cabeçalhos de e-mail (Assunto/Remetente)."""
    if not header_value:
        return ""
    decoded_fragments = decode_header(header_value)
    subject_str = ""
    for fragment, encoding in decoded_fragments:
        if isinstance(fragment, bytes):
            subject_str += fragment.decode(encoding or "utf-8", errors="ignore")
        else:
            subject_str += fragment
    return subject_str


def obter_codigo_verificacao(email_user: str, app_password: str, timestamp_inicio: datetime = None, timeout: int = 90) -> str:
    """
    Conecta ao Gmail via IMAP e busca apenas e-mails recebidos APÓS o timestamp_inicio.
    """
    if timestamp_inicio is None:
        timestamp_inicio = datetime.now(timezone.utc)
    elif timestamp_inicio.tzinfo is None:
        # Se vier sem fuso horário, assume UTC
        timestamp_inicio = timestamp_inicio.replace(tzinfo=timezone.utc)

    print(f"📩 Aguardando e-mail da Microsoft enviado após: {timestamp_inicio.strftime('%H:%M:%S UTC')}...")
    tempo_inicial = time.time()

    while time.time() - tempo_inicial < timeout:
        try:
            mail = imaplib.IMAP4_SSL("imap.gmail.com")
            mail.login(email_user, app_password)
            mail.select("inbox")

            status, messages = mail.search(None, 'ALL')
            email_ids = messages[0].split()

            if email_ids:
                # Olha as últimas 5 mensagens recebidas
                ultimos_ids = email_ids[-5:]
                ultimos_ids.reverse()

                for e_id in ultimos_ids:
                    # Busca os cabeçalhos Subject, From e Date
                    status, data = mail.fetch(e_id, '(BODY.PEEK[HEADER.FIELDS (SUBJECT FROM DATE)])')

                    for response_part in data:
                        if isinstance(response_part, tuple):
                            msg = email.message_from_bytes(response_part[1])
                            assunto = decode_mime_header(msg.get("Subject"))
                            remetente = decode_mime_header(msg.get("From"))
                            data_str = msg.get("Date")

                            # Converte a data do cabeçalho do e-mail para datetime
                            try:
                                data_email = parsedate_to_datetime(data_str)
                                if data_email.tzinfo is None:
                                    data_email = data_email.replace(tzinfo=timezone.utc)
                            except Exception:
                                data_email = datetime.now(timezone.utc)

                            # 1. Checa se o e-mail é de um remetente/assunto válido
                            if "no-reply@notify.microsoft.com" in remetente.lower() or "onedrive" in assunto.lower():
                                # 2. Valida se o e-mail foi enviado APÓS o momento da requisição
                                if data_email >= timestamp_inicio:
                                    print(f"🔍 E-mail novo recebido às {data_email.strftime('%H:%M:%S UTC')} | Assunto: '{assunto}'")
                                    
                                    codigo_match = re.search(r'^\s*(\d+)', assunto)
                                    if codigo_match:
                                        codigo = codigo_match.group(1)
                                        print(f"✅ Código OTP NOVO capturado: {codigo}")
                                        mail.logout()
                                        return codigo
                                else:
                                    print(f"⏳ E-mail antigo ignorado (recebido às {data_email.strftime('%H:%M:%S UTC')}). Aguardando o novo...")

            mail.logout()
        except Exception as e:
            print(f"⏳ Tentando consultar o Gmail novamente... ({e})")

        time.sleep(4)  # Aguarda 4 segundos entre as tentativas

    raise TimeoutError("Tempo limite esgotado sem receber um novo e-mail de verificação.")