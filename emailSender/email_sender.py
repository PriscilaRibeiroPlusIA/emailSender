# email_sender.py
import os
import smtplib
import ssl
import socket
import certifi  # Importa a biblioteca certifi
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from typing import List
import config


def enviar_email(destinatario: str, numero_processo: str, caminhos_anexos: List[str]) -> bool:
    """Monta e envia o email com a mensagem e assinatura atualizadas."""
    assunto = f"Encaminhamento Comprovante - Processo nº {numero_processo}"

    # --- CORPO DO EMAIL ATUALIZADO COM A NOVA ASSINATURA ---
    corpo_email = (
        f"Prezados(as),\n\n"
        f"Encaminho em anexo comprovante(s) de cumprimento referente ao processo número {numero_processo}.\n\n"
        f"Atenciosamente,\n\n"
        f"Priscila Ribeiro\n"
        f"Técnica do Seguro Social\n"
        f"Matr. 1446118\n"
        f"SADJ-INSS"
    )
    # --- FIM DA ATUALIZAÇÃO DO CORPO DO EMAIL ---

    mensagem = MIMEMultipart()
    mensagem["From"] = config.EMAIL_REMETENTE
    mensagem["To"] = destinatario
    mensagem["Subject"] = assunto
    mensagem.attach(MIMEText(corpo_email, "plain", "utf-8"))

    print(f"  [Email Sender] Preparando email para: {destinatario}, Assunto: {assunto}")

    if not caminhos_anexos:
        print("    - Nenhum anexo a ser enviado.")
    else:
        print(f"    - Anexando {len(caminhos_anexos)} arquivo(s)...")

    for caminho_anexo in caminhos_anexos:
        if not os.path.exists(caminho_anexo):
            print(f"    - Alerta: Arquivo de anexo não encontrado e será ignorado: {caminho_anexo}")
            continue
        try:
            nome_arquivo_anexo = os.path.basename(caminho_anexo)
            with open(caminho_anexo, "rb") as anexo_file:
                parte = MIMEApplication(anexo_file.read(), Name=nome_arquivo_anexo)
            parte['Content-Disposition'] = f'attachment; filename="{nome_arquivo_anexo}"'
            mensagem.attach(parte)
            print(f"      -> Anexado: {nome_arquivo_anexo}")
        except Exception as e:
            print(f"    - Erro ao anexar o arquivo {caminho_anexo}: {e}")

    try:
        cafile_path = certifi.where()
        context = ssl.create_default_context(cafile=cafile_path)
        print(f"  [Email Sender] Usando contexto SSL padrão com CAFile explícito de certifi: {cafile_path}")

        with smtplib.SMTP(config.SERVIDOR_SMTP, config.PORTA_SMTP) as server:
            print(f"  [Email Sender] Conectando ao servidor SMTP: {config.SERVIDOR_SMTP}:{config.PORTA_SMTP}...")
            server.ehlo()
            server.starttls(context=context)
            server.ehlo()
            print(f"  [Email Sender] Fazendo login com o usuário: {config.EMAIL_REMETENTE}...")
            server.login(config.EMAIL_REMETENTE, config.SENHA_REMETENTE)
            print(f"  [Email Sender] Enviando email para {destinatario}...")
            server.sendmail(config.EMAIL_REMETENTE, destinatario, mensagem.as_string())
        print(
            f"  [Email Sender] Email enviado com sucesso para {destinatario} referente ao processo {numero_processo}!")
        return True
    except smtplib.SMTPAuthenticationError:
        print(f"  [Email Sender] Erro de AUTENTICAÇÃO SMTP para {config.EMAIL_REMETENTE}.")
        print(f"  Verifique email/senha e configurações de segurança da conta (ex: 'app password').")
        return False
    except smtplib.SMTPException as e:
        print(f"  [Email Sender] Erro SMTP ao enviar email para {destinatario}: {e}")
        return False
    except socket.gaierror:
        print(
            f"  [Email Sender] Erro de rede: Não foi possível encontrar o servidor SMTP '{config.SERVIDOR_SMTP}'. Verifique o nome e sua conexão.")
        return False
    except Exception as e:
        print(f"  [Email Sender] Erro geral e inesperado ao enviar email para {destinatario}: {e}")
        print(f"  Tipo de erro: {type(e).__name__}")
        return False