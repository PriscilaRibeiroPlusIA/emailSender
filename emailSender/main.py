# main.py
import os
import time
import shutil

try:
    import config
    import pdf_processor
    import excel_reader
    import email_sender
except ImportError as e:
    print(f"ERRO CRÍTICO em main.py: Falha ao importar um dos módulos do projeto: {e}")
    print(
        "Verifique se todos os arquivos .py (config, pdf_processor, excel_reader, email_sender) estão na mesma pasta.")
    exit("Módulo essencial ausente.")


def carregar_pdfs_processados() -> set:
    """Carrega a lista de nomes de arquivos PDF que já foram processados do log."""
    processados = set()
    if os.path.exists(config.ARQUIVO_PROCESSADOS_LOG):
        try:
            with open(config.ARQUIVO_PROCESSADOS_LOG, "r", encoding="utf-8") as f:
                for linha in f:
                    processados.add(linha.strip())
        except Exception as e:
            print(f"Erro ao ler o arquivo de log de processados ({config.ARQUIVO_PROCESSADOS_LOG}): {e}")
    return processados


def marcar_como_processado_e_mover(nome_arquivo_pdf: str, sucesso_envio: bool):
    """
    Adiciona o nome do arquivo PDF à lista de processados no log
    e move o arquivo para a pasta de sucesso ou erro.
    """
    try:
        with open(config.ARQUIVO_PROCESSADOS_LOG, "a", encoding="utf-8") as f:
            f.write(nome_arquivo_pdf + "\n")
    except Exception as e:
        print(f"Erro ao escrever no arquivo de log de processados ({config.ARQUIVO_PROCESSADOS_LOG}): {e}")

    pasta_destino = config.PASTA_PROCESSADOS_SUCESSO if sucesso_envio else config.PASTA_PROCESSADOS_ERRO
    os.makedirs(pasta_destino, exist_ok=True)

    caminho_origem_pdf = os.path.join(config.PASTA_PROCESSOS_PDF, nome_arquivo_pdf)
    caminho_destino_pdf = os.path.join(pasta_destino, nome_arquivo_pdf)

    try:
        if os.path.exists(caminho_destino_pdf):
            base, ext = os.path.splitext(nome_arquivo_pdf)
            timestamp = int(time.time())
            novo_nome_pdf_destino = f"{base}_{timestamp}{ext}"
            caminho_destino_pdf = os.path.join(pasta_destino, novo_nome_pdf_destino)
            print(
                f"  Aviso: PDF {nome_arquivo_pdf} já existe em {os.path.basename(pasta_destino)}. Renomeando no destino para {novo_nome_pdf_destino}")

        shutil.move(caminho_origem_pdf, caminho_destino_pdf)
        status_movido = "sucesso" if sucesso_envio else "erro"
        print(
            f"  PDF {nome_arquivo_pdf} movido para a pasta de processados com {status_movido}: {os.path.basename(pasta_destino)}")
    except Exception as e:
        print(f"  Erro ao mover PDF {nome_arquivo_pdf} para {pasta_destino}: {e}")


def processar_um_pdf(caminho_pdf: str, nome_pdf: str) -> bool:
    """
    Processa um único arquivo PDF: extrai dados, busca email, monta e envia.
    O número do processo é obtido do nome do arquivo PDF.
    Os comprovantes são unificados em um único PDF.
    """
    print(f"\n>>> Iniciando processamento do PDF: {nome_pdf} <<<")

    numero_processo, _ = os.path.splitext(nome_pdf)
    print(
        f"  [Main Process] Número do Processo obtido do nome do arquivo: '{numero_processo}' (Tipo: {type(numero_processo).__name__})")

    if not numero_processo:
        print(
            f"  [Main Process] Não foi possível obter um número de processo válido do nome do arquivo: {nome_pdf}. Pulando.")
        return False

    texto_pdf = pdf_processor.extrair_texto_do_pdf(caminho_pdf)
    if not texto_pdf:
        print(f"Falha ao extrair texto do PDF {nome_pdf}. PDF não será processado.")
        return False

    dados_vara_comarca = pdf_processor.extrair_informacoes_processo(texto_pdf, nome_pdf)

    vara_civel = None
    comarca = None

    if dados_vara_comarca:
        vara_civel = dados_vara_comarca.get("vara_civel")
        comarca = dados_vara_comarca.get("comarca")

    print(f"  [Main Process] Vara Cível extraída: '{vara_civel}' (Tipo: {type(vara_civel).__name__})")
    print(f"  [Main Process] Comarca extraída: '{comarca}' (Tipo: {type(comarca).__name__})")

    dados_completos = True
    if not numero_processo:
        print(f"  [Main Process] ERRO INTERNO: numero_processo (do nome do arquivo) está vazio.")
        dados_completos = False
    if not vara_civel:
        print(f"  [Main Process] ERRO: vara_civel não foi extraída ou está vazia.")
        dados_completos = False
    if not comarca:
        print(f"  [Main Process] ERRO: comarca não foi extraída ou está vazia.")
        dados_completos = False

    if not dados_completos:
        print(
            f"Dados essenciais (número do processo do arquivo, vara, comarca) incompletos para o PDF {nome_pdf}. Pulando.")
        return False

    print(f"  Dados para busca no Excel -> Vara: '{vara_civel}', Comarca: '{comarca}'")

    comprovantes_originais = pdf_processor.identificar_comprovantes(numero_processo)

    lista_final_de_anexos_para_email = []

    if comprovantes_originais:
        print(
            f"  [Main Process] {len(comprovantes_originais)} comprovante(s) original(is) encontrado(s). Tentando unificar em um único PDF.")
        caminho_pdf_unificado = pdf_processor.criar_pdf_unificado(
            comprovantes_originais,
            numero_processo,
            config.PASTA_COMPROVANTES
        )
        if caminho_pdf_unificado and os.path.exists(caminho_pdf_unificado):
            lista_final_de_anexos_para_email.append(caminho_pdf_unificado)
            print(f"  [Main Process] PDF unificado pronto para anexo: {os.path.basename(caminho_pdf_unificado)}")
        else:
            print(
                f"  [Main Process] ATENÇÃO: Falha ao criar PDF unificado. O email será enviado sem comprovantes unificados.")
    else:
        print("  [Main Process] Nenhum comprovante original encontrado para o processo.")

    email_da_vara = excel_reader.buscar_email_vara(vara_civel, comarca)
    if not email_da_vara:
        print(
            f"Email da vara '{vara_civel}' na comarca '{comarca}' não encontrado para o PDF {nome_pdf}. Email não será enviado.")
        if lista_final_de_anexos_para_email and os.path.exists(lista_final_de_anexos_para_email[0]):
            if config.PASTA_COMPROVANTES in os.path.abspath(
                    lista_final_de_anexos_para_email[0]) and "Comprovantes_Unificados_Processo_" in \
                    lista_final_de_anexos_para_email[0]:
                try:
                    os.remove(lista_final_de_anexos_para_email[0])
                    print(
                        f"  [Main Process] PDF unificado '{os.path.basename(lista_final_de_anexos_para_email[0])}' removido pois o email não será enviado.")
                except Exception as e_del:
                    print(f"  [Main Process] Erro ao remover PDF unificado: {e_del}")
        return False

    email_destinatario_final = email_da_vara

    print(f"  [Main Process] Email será enviado para: {email_destinatario_final}")
    if email_destinatario_final == config.EMAIL_REMETENTE and email_da_vara != config.EMAIL_REMETENTE:
        print(
            f"  ALERTA DE TESTE: O email está configurado para ser enviado para o remetente ({config.EMAIL_REMETENTE}), mas o email da vara encontrado foi {email_da_vara}.")

    sucesso_ao_enviar = email_sender.enviar_email(
        destinatario=email_destinatario_final,
        numero_processo=numero_processo,
        caminhos_anexos=lista_final_de_anexos_para_email
    )

    if sucesso_ao_enviar:
        print(f"Processamento do PDF {nome_pdf} concluído com sucesso (email enviado).")
        return True
    else:
        print(f"Falha ao enviar email para o PDF {nome_pdf}.")
        return False


def executar_uma_vez():  # Nome da função alterado para refletir a nova funcionalidade
    """Função principal para verificar a pasta de PDFs e processá-los UMA VEZ."""
    print("====================================================")
    print("Iniciando Sistema de Envio de Emails Automatizado (Execução Única)")  # Mensagem ajustada
    print(f"Data e Hora Início: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    print("====================================================")
    print(f"Verificando pasta de PDFs: {config.PASTA_PROCESSOS_PDF}")  # Mensagem ajustada
    print(f"Pasta de comprovantes (base para subpastas de processo): {config.PASTA_COMPROVANTES}")
    print(f"Planilha de emails: {config.CAMINHO_PLANILHA_EMAILS}")
    print(f"Log de processados: {config.ARQUIVO_PROCESSADOS_LOG}")
    print("----------------------------------------------------")

    os.makedirs(config.PASTA_PROCESSADOS_SUCESSO, exist_ok=True)
    os.makedirs(config.PASTA_PROCESSADOS_ERRO, exist_ok=True)
    os.makedirs(config.PASTA_COMPROVANTES, exist_ok=True)
    if config.ARQUIVO_PROCESSADOS_LOG and not os.path.exists(os.path.dirname(config.ARQUIVO_PROCESSADOS_LOG)):
        if os.path.dirname(config.ARQUIVO_PROCESSADOS_LOG):
            os.makedirs(os.path.dirname(config.ARQUIVO_PROCESSADOS_LOG), exist_ok=True)

    pdfs_ja_processados_nesta_sessao = carregar_pdfs_processados()
    print(f"Carregados {len(pdfs_ja_processados_nesta_sessao)} PDFs do log de já processados.")

    # --- INÍCIO DA LÓGICA QUE ESTAVA DENTRO DO 'while True:' ---
    # Agora executa apenas uma vez
    print(f"\n[{time.strftime('%Y-%m-%d %H:%M:%S')}] Verificando por novos PDFs...")
    novos_pdfs_foram_detectados = False
    arquivos_na_pasta_monitorada = []
    try:
        if not os.path.isdir(config.PASTA_PROCESSOS_PDF):
            print(
                f"ERRO CRÍTICO: A pasta de processos PDF '{config.PASTA_PROCESSOS_PDF}' não existe ou não é um diretório.")
            # Em execução única, podemos simplesmente sair se a pasta não existir.
            return

        arquivos_na_pasta_monitorada = os.listdir(config.PASTA_PROCESSOS_PDF)
    except Exception as e_listdir:
        print(f"ERRO ao listar arquivos em '{config.PASTA_PROCESSOS_PDF}': {e_listdir}")
        return  # Sai se não conseguir listar arquivos

    for nome_do_arquivo in arquivos_na_pasta_monitorada:
        caminho_completo_do_pdf = os.path.join(config.PASTA_PROCESSOS_PDF, nome_do_arquivo)

        if nome_do_arquivo.lower().endswith(".pdf") and \
                os.path.isfile(caminho_completo_do_pdf) and \
                nome_do_arquivo not in pdfs_ja_processados_nesta_sessao:
            novos_pdfs_foram_detectados = True
            envio_bem_sucedido = processar_um_pdf(caminho_completo_do_pdf, nome_do_arquivo)

            marcar_como_processado_e_mover(nome_do_arquivo, sucesso_envio=envio_bem_sucedido)
            pdfs_ja_processados_nesta_sessao.add(
                nome_do_arquivo)  # Adiciona mesmo se falhar, para não tentar de novo nesta execução

    if not novos_pdfs_foram_detectados:
        print("Nenhum novo PDF encontrado para processamento nesta execução.")
    # --- FIM DA LÓGICA QUE ESTAVA DENTRO DO 'while True:' ---

    print("\n----------------------------------------------------")
    print(
        f"Verificação e processamento concluídos às {time.strftime('%Y-%m-%d %H:%M:%S')}.")  # Nova mensagem de conclusão
    print("====================================================")


if __name__ == "__main__":
    try:
        executar_uma_vez()  # Chama a função de execução única
    except Exception as e_global:
        print("\n----------------------------------------------------")
        print(f"UM ERRO GLOBAL INESPERADO OCORREU NO SCRIPT: {e_global}")
        print(f"Tipo de erro: {type(e_global).__name__}")
        print("----------------------------------------------------")
    finally:
        print("Script finalizado.")