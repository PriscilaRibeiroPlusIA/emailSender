# pdf_processor.py
import os
import re
from typing import Dict, Optional, List
import shutil

# Importações para unificação e conversão de PDF
try:
    import PyPDF2  # Para unir PDFs
except ImportError:
    print("ERRO: A biblioteca 'PyPDF2' não está instalada. Por favor, instale com: pip install PyPDF2")
    PyPDF2 = None

try:
    from PIL import Image  # Pillow para converter imagens
except ImportError:
    print("ERRO: A biblioteca 'Pillow' não está instalada. Por favor, instale com: pip install Pillow")
    Image = None

try:
    from reportlab.pdfgen import canvas  # Para converter texto para PDF
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import inch
    from reportlab.lib.utils import ImageReader  # Para melhor controle de imagem com reportlab
    from reportlab.pdfbase.ttfonts import TTFont  # Para registrar fontes se necessário
    from reportlab.pdfbase import pdfmetrics
except ImportError:
    print("ERRO: A biblioteca 'reportlab' não está instalada. Por favor, instale com: pip install reportlab")
    canvas = None
    TTFont = None  # Define como None se a importação falhar

# Importações existentes
try:
    import pdfplumber
except ImportError:
    print("ERRO: A biblioteca 'pdfplumber' não está instalada. Por favor, instale com: pip install pdfplumber")
    pdfplumber = None

try:
    import config
except ImportError:
    print("ERRO CRÍTICO em pdf_processor.py: O arquivo config.py não foi encontrado ou não pôde ser importado.")


    class FallbackConfig:
        PASTA_COMPROVANTES = "pasta_comprovantes_nao_configurada"


    config = FallbackConfig()


def extrair_texto_do_pdf(caminho_pdf: str) -> Optional[str]:
    """Extrai todo o texto de um arquivo PDF."""
    if not pdfplumber:
        print("  [PDF Extractor] pdfplumber não está disponível. Não é possível extrair texto.")
        return None
    texto_completo = ""
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for i, pagina in enumerate(pdf.pages):
                texto_pagina = pagina.extract_text(x_tolerance=2, y_tolerance=2)
                if texto_pagina:
                    texto_completo += texto_pagina + "\n"
        return texto_completo
    except Exception as e:
        print(f"  [PDF Extractor] Erro ao ler o PDF {os.path.basename(caminho_pdf)}: {e}")
        return None


def extrair_informacoes_processo(texto_pdf: str, nome_arquivo_pdf_original: str) -> Optional[Dict[str, str]]:
    """
    Extrai SOMENTE vara cível e comarca do texto do PDF.
    O número do processo virá do nome do arquivo no script main.py.
    """
    if not texto_pdf:
        print(
            f"  [PDF Extractor] Texto do PDF está vazio para {nome_arquivo_pdf_original}. Não é possível extrair Vara/Comarca.")
        return None

    dados_pdf = {}
    print(f"--- Analisando conteúdo do PDF: {nome_arquivo_pdf_original} para Vara e Comarca ---")

    padrao_vara = r"(\d+ª\s*vara\s*c[íi]vel)"
    match_vara = re.search(padrao_vara, texto_pdf, re.IGNORECASE)
    if match_vara:
        vara_capturada = match_vara.group(1).strip()
        vara_limpa = vara_capturada.replace("\n", " ").replace("\r", " ")
        dados_pdf["vara_civel"] = re.sub(r'\s+', ' ', vara_limpa).strip()
        print(f"  [PDF Extractor] Vara Cível encontrada (após limpeza): {dados_pdf['vara_civel']}")
    else:
        print(f"  [PDF Extractor] Vara cível não encontrada com o padrão atual no PDF: {nome_arquivo_pdf_original}.")

    padrao_comarca = r"Comarca\s+de\s+([A-Za-zÀ-ú\s-]+(?:SP)?)"
    match_comarca = re.search(padrao_comarca, texto_pdf, re.IGNORECASE)
    if match_comarca:
        comarca_capturada = match_comarca.group(1).strip()
        comarca_limpa = comarca_capturada.replace("\n", " ").replace("\r", " ")
        dados_pdf["comarca"] = re.sub(r'\s+', ' ', comarca_limpa).strip()
        print(f"  [PDF Extractor] Comarca encontrada (após limpeza): {dados_pdf['comarca']}")
    else:
        print(f"  [PDF Extractor] Comarca não encontrada com o padrão atual no PDF: {nome_arquivo_pdf_original}.")

    if dados_pdf:
        return dados_pdf
    else:
        print(
            f"  [PDF Extractor] Nenhuma informação de Vara ou Comarca foi extraída do PDF {nome_arquivo_pdf_original}.")
        return None


def identificar_comprovantes(numero_processo: str) -> List[str]:
    """
    Identifica os arquivos de comprovante na subpasta nomeada com o numero_processo.
    """
    comprovantes_encontrados = []
    if not numero_processo:
        print(
            "  [Attachment Finder] Número do processo (do nome do arquivo PDF) não fornecido. Não é possível buscar comprovantes.")
        return []

    # O nome da subpasta é o próprio numero_processo (nome do arquivo PDF sem extensão)
    caminho_subpasta_comprovantes = os.path.join(config.PASTA_COMPROVANTES, numero_processo)
    print(f"  [Attachment Finder] Procurando comprovantes na subpasta: {caminho_subpasta_comprovantes}")

    if not os.path.isdir(caminho_subpasta_comprovantes):
        print(
            f"  [Attachment Finder] Subpasta de comprovantes '{os.path.basename(caminho_subpasta_comprovantes)}' não encontrada em '{config.PASTA_COMPROVANTES}'.")
        return []
    try:
        for nome_item_na_subpasta in os.listdir(caminho_subpasta_comprovantes):
            # Ignora subpastas temporárias de conversão
            if nome_item_na_subpasta == "_temp_conversion":
                continue
            caminho_completo_item = os.path.join(caminho_subpasta_comprovantes, nome_item_na_subpasta)
            if os.path.isfile(caminho_completo_item):
                comprovantes_encontrados.append(caminho_completo_item)
                print(f"    -> Comprovante encontrado na subpasta: {nome_item_na_subpasta}")

        if not comprovantes_encontrados:
            print(
                f"  [Attachment Finder] Nenhum arquivo (comprovante) encontrado dentro da subpasta '{os.path.basename(caminho_subpasta_comprovantes)}'.")
        else:
            print(
                f"  [Attachment Finder] Total de {len(comprovantes_encontrados)} comprovante(s) encontrado(s) na subpasta.")
    except Exception as e:
        print(
            f"  [Attachment Finder] Erro ao listar arquivos na subpasta '{os.path.basename(caminho_subpasta_comprovantes)}': {e}")
        return []
    return comprovantes_encontrados


def criar_pdf_de_texto(caminho_arquivo_texto: str, caminho_pdf_saida: str):
    """Converte um arquivo de texto simples para PDF usando reportlab."""
    if not canvas:
        print(
            f"  [PDF Unifier] Reportlab não está disponível. Não é possível converter texto de '{os.path.basename(caminho_arquivo_texto)}'.")
        return False  # Indica falha
    try:
        # Tenta registrar uma fonte comum que suporte mais caracteres, como a DejaVu Sans
        # Você pode precisar baixar o arquivo .ttf e colocar em uma pasta acessível
        # ou usar fontes padrão como Helvetica, Times-Roman
        try:
            if TTFont:  # Verifica se TTFont foi importado
                # Use um nome de fonte padrão se DejaVuSans não estiver disponível ou não for necessário
                pdfmetrics.registerFont(TTFont('DejaVuSans', 'DejaVuSans.ttf'))  # Exemplo, requer o .ttf
                FONT_NAME = 'DejaVuSans'
            else:
                FONT_NAME = 'Helvetica'  # Fallback
        except:  # Se o registro da fonte falhar, usa Helvetica como padrão
            FONT_NAME = 'Helvetica'
            print(
                f"  [PDF Unifier] Fonte DejaVuSans não encontrada/registrada, usando {FONT_NAME} para {os.path.basename(caminho_arquivo_texto)}.")

        c = canvas.Canvas(caminho_pdf_saida, pagesize=A4)
        c.setFont(FONT_NAME, 9)
        largura_pagina, altura_pagina = A4
        margem_x = 0.75 * inch
        margem_y = 0.75 * inch
        largura_texto_util = largura_pagina - (2 * margem_x)
        y_pos = altura_pagina - margem_y
        line_height = 11

        with open(caminho_arquivo_texto, 'r', encoding='utf-8', errors='ignore') as f:
            for line in f:
                linha_tratada = line.rstrip('\n\r')

                # Lógica simples de quebra de linha se a linha for muito longa
                # (ReportLab tem ferramentas mais avançadas para isso, como Paragraphs)
                while c.stringWidth(linha_tratada, FONT_NAME, 9) > largura_texto_util:
                    # Encontra o último espaço antes de estourar a largura
                    ponto_quebra = -1
                    for i in range(len(linha_tratada) - 1, 0, -1):
                        if linha_tratada[i] == ' ':
                            sub_linha = linha_tratada[:i]
                            if c.stringWidth(sub_linha, FONT_NAME, 9) <= largura_texto_util:
                                ponto_quebra = i
                                break
                    if ponto_quebra == -1:  # Não achou espaço, quebra na força
                        ponto_quebra = int(largura_texto_util / (c.stringWidth("M", FONT_NAME, 9) + 1e-6))  # Estimativa

                    c.drawString(margem_x, y_pos, linha_tratada[:ponto_quebra])
                    linha_tratada = linha_tratada[ponto_quebra:].lstrip()
                    y_pos -= line_height
                    if y_pos < margem_y:
                        c.showPage()
                        c.setFont(FONT_NAME, 9)
                        y_pos = altura_pagina - margem_y

                c.drawString(margem_x, y_pos, linha_tratada)
                y_pos -= line_height
                if y_pos < margem_y:
                    c.showPage()
                    c.setFont(FONT_NAME, 9)
                    y_pos = altura_pagina - margem_y
        c.save()
        print(
            f"    -> Texto '{os.path.basename(caminho_arquivo_texto)}' convertido para PDF: {os.path.basename(caminho_pdf_saida)}")
        return True
    except Exception as e:
        print(f"  [PDF Unifier] Erro ao converter texto '{os.path.basename(caminho_arquivo_texto)}' para PDF: {e}")
        return False


def criar_pdf_de_imagem(caminho_arquivo_imagem: str, caminho_pdf_saida: str):
    """Converte uma imagem para PDF usando Pillow e Reportlab para melhor posicionamento."""
    if not Image or not canvas or not ImageReader:
        print(
            f"  [PDF Unifier] Pillow ou Reportlab não estão disponíveis. Não é possível converter imagem de '{os.path.basename(caminho_arquivo_imagem)}'.")
        return False
    try:
        img = Image.open(caminho_arquivo_imagem)
        img_convertida = img
        if img.mode == 'RGBA' or img.mode == 'P':  # RGBA e P (paleta) podem ter transparência
            # Cria um fundo branco e cola a imagem com máscara de transparência
            img_convertida = Image.new("RGB", img.size, (255, 255, 255))
            img_convertida.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' or (
                        img.mode == 'P' and 'transparency' in img.info) else None)
        elif img.mode != 'RGB' and img.mode != 'L':  # L é grayscale
            img_convertida = img.convert('RGB')

        largura_img, altura_img = img_convertida.size
        largura_a4, altura_a4 = A4

        margem = 0.5 * inch
        max_largura_conteudo = largura_a4 - (2 * margem)
        max_altura_conteudo = altura_a4 - (2 * margem)

        ratio = min(max_largura_conteudo / largura_img, max_altura_conteudo / altura_img)
        nova_largura = largura_img * ratio
        nova_altura = altura_img * ratio

        c = canvas.Canvas(caminho_pdf_saida, pagesize=A4)
        x_pos = margem + (max_largura_conteudo - nova_largura) / 2
        y_pos = margem + (max_altura_conteudo - nova_altura) / 2
        c.drawImage(ImageReader(img_convertida), x_pos, y_pos, width=nova_largura, height=nova_altura,
                    preserveAspectRatio=True, anchor='c', mask='auto')
        c.save()
        print(
            f"    -> Imagem '{os.path.basename(caminho_arquivo_imagem)}' convertida para PDF: {os.path.basename(caminho_pdf_saida)}")
        return True
    except Exception as e:
        print(f"  [PDF Unifier] Erro ao converter imagem '{os.path.basename(caminho_arquivo_imagem)}' para PDF: {e}")
        return False


def criar_pdf_unificado(lista_arquivos_originais: List[str], numero_processo: str, pasta_base_comprovantes: str) -> \
Optional[str]:
    """
    Converte arquivos (PDF, imagem, texto) para PDF e os une em um único PDF.
    Retorna o caminho do PDF unificado ou None se falhar.
    """
    if not PyPDF2 or not Image or not canvas:
        print(
            "  [PDF Unifier] Bibliotecas necessárias (PyPDF2, Pillow, Reportlab) não estão todas disponíveis. Não é possível unificar.")
        return None

    pdfs_para_unir = []
    # Cria uma subpasta _temp_conversion DENTRO da pasta do processo para os PDFs convertidos
    pasta_processo_especifico = os.path.join(pasta_base_comprovantes, numero_processo)
    pasta_temporaria_conversao = os.path.join(pasta_processo_especifico, "_temp_conversion")

    # Certifique-se que a pasta do processo específico exista (onde o PDF final será salvo)
    os.makedirs(pasta_processo_especifico, exist_ok=True)
    os.makedirs(pasta_temporaria_conversao, exist_ok=True)

    print(f"  [PDF Unifier] Iniciando unificação para processo {numero_processo}.")

    for i, caminho_arquivo in enumerate(lista_arquivos_originais):
        nome_base, extensao = os.path.splitext(os.path.basename(caminho_arquivo))
        extensao = extensao.lower()
        # Salva PDFs convertidos na pasta temporária
        caminho_pdf_convertido = os.path.join(pasta_temporaria_conversao, f"temp_{nome_base.replace('.', '_')}_{i}.pdf")

        if extensao == ".pdf":
            pdfs_para_unir.append(caminho_arquivo)
            print(f"    -> PDF original adicionado à lista de união: {os.path.basename(caminho_arquivo)}")
        elif extensao in [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff"]:
            if criar_pdf_de_imagem(caminho_arquivo, caminho_pdf_convertido):
                pdfs_para_unir.append(caminho_pdf_convertido)
        elif extensao in [".txt", ".rel", ".lst", ".prn"]:  # Adicione outras extensões de texto aqui
            print(f"    -> Tratando arquivo '{os.path.basename(caminho_arquivo)}' (ext: {extensao}) como texto.")
            if criar_pdf_de_texto(caminho_arquivo, caminho_pdf_convertido):
                pdfs_para_unir.append(caminho_pdf_convertido)
        else:
            print(
                f"    -> ATENÇÃO: Arquivo '{os.path.basename(caminho_arquivo)}' com extensão '{extensao}' não suportado para conversão. Será ignorado.")

    if not pdfs_para_unir:
        print("  [PDF Unifier] Nenhum arquivo PDF (original ou convertido) para unir.")
        shutil.rmtree(pasta_temporaria_conversao, ignore_errors=True)  # Limpa pasta temporária
        return None

    nome_pdf_final = f"Comprovantes_Unificados_Processo_{numero_processo}.pdf"
    # Salva o PDF final na subpasta do processo, não na pasta _temp_conversion
    caminho_pdf_final = os.path.join(pasta_processo_especifico, nome_pdf_final)

    try:
        merger = PyPDF2.PdfWriter()
        for pdf_path in pdfs_para_unir:
            try:
                reader = PyPDF2.PdfReader(pdf_path)
                for page in reader.pages:
                    merger.add_page(page)
                print(f"      -> Páginas de '{os.path.basename(pdf_path)}' adicionadas ao PDF final.")
            except Exception as e_read:
                print(
                    f"  [PDF Unifier] Erro ao ler o PDF '{os.path.basename(pdf_path)}' durante a união: {e_read}. Será ignorado.")

        if len(merger.pages) > 0:
            with open(caminho_pdf_final, "wb") as f_out:
                merger.write(f_out)
            print(
                f"  [PDF Unifier] PDF unificado criado com sucesso: {nome_pdf_final} em {os.path.dirname(caminho_pdf_final)}")

            # Limpeza da pasta temporária de conversão após o sucesso
            shutil.rmtree(pasta_temporaria_conversao, ignore_errors=True)
            print(f"      -> Pasta temporária '{os.path.basename(pasta_temporaria_conversao)}' removida.")
            return caminho_pdf_final
        else:
            print("  [PDF Unifier] Nenhuma página foi adicionada ao PDF final. O arquivo unificado não será criado.")
            shutil.rmtree(pasta_temporaria_conversao, ignore_errors=True)
            return None

    except Exception as e:
        print(f"  [PDF Unifier] Erro ao criar o PDF unificado: {e}")
        shutil.rmtree(pasta_temporaria_conversao, ignore_errors=True)  # Limpa pasta temporária em caso de erro
        return None


# Bloco de teste para executar este script isoladamente
if __name__ == "__main__":
    # Teste para extrair_informacoes_processo
    print("--- Iniciando Teste do Módulo pdf_processor.py (Extração) ---")
    caminho_pdf_teste_extracao = r"C:\Users\Priscila\APSDJ\ProcessosBaixadosTemp\00009563620218260404.pdf"  # Use um PDF real
    if not os.path.exists(caminho_pdf_teste_extracao):
        print(f"AVISO: Arquivo para teste de extração não encontrado: {caminho_pdf_teste_extracao}")
    else:
        nome_arquivo_teste_extracao = os.path.basename(caminho_pdf_teste_extracao)
        texto_extraido_teste = extrair_texto_do_pdf(caminho_pdf_teste_extracao)
        if texto_extraido_teste:
            informacoes_teste = extrair_informacoes_processo(texto_extraido_teste, nome_arquivo_teste_extracao)
            print(f"Resultado da extração para {nome_arquivo_teste_extracao}: {informacoes_teste}")

    # Teste para criar_pdf_unificado
    print("\n--- Testando a função criar_pdf_unificado ---")
    if not (PyPDF2 and Image and canvas):
        print("Bibliotecas PyPDF2, Pillow ou Reportlab não estão disponíveis. Teste de unificação pulado.")
    else:
        numero_processo_teste_uniao = "TESTE_001"
        subpasta_processo_teste = os.path.join(config.PASTA_COMPROVANTES, numero_processo_teste_uniao)
        os.makedirs(subpasta_processo_teste, exist_ok=True)  # Garante que a subpasta do processo exista

        # Crie arquivos de teste DENTRO da subpasta_processo_teste
        arquivos_teste_para_unir = []

        # 1. PDF de Teste
        try:
            pdf_teste_path = os.path.join(subpasta_processo_teste, "exemplo.pdf")
            c_test = canvas.Canvas(pdf_teste_path, pagesize=A4)
            c_test.drawString(1 * inch, A4[1] - 1 * inch, "Página 1 do PDF de Teste.")
            c_test.showPage()
            c_test.drawString(1 * inch, A4[1] - 1 * inch, "Página 2 do PDF de Teste.")
            c_test.save()
            arquivos_teste_para_unir.append(pdf_teste_path)
            print(f"Arquivo PDF de teste criado: {pdf_teste_path}")
        except Exception as e:
            print(f"Falha ao criar PDF de teste: {e}")

        # 2. Imagem de Teste
        try:
            img_teste_path = os.path.join(subpasta_processo_teste, "exemplo_imagem.png")
            img_teste = Image.new('RGB', (600, 400), color='skyblue')
            # Adicionar texto à imagem com Pillow (requer Pillow > 9.0.0 para textsize e text)
            # from PIL import ImageDraw
            # draw = ImageDraw.Draw(img_teste)
            # draw.text((10,10), "Imagem de Teste", fill=(0,0,0)) # Requer fonte, ou usa padrão
            img_teste.save(img_teste_path)
            arquivos_teste_para_unir.append(img_teste_path)
            print(f"Arquivo de imagem de teste criado: {img_teste_path}")
        except Exception as e:
            print(f"Falha ao criar imagem de teste: {e}")

        # 3. Arquivo de Texto de Teste
        try:
            txt_teste_path = os.path.join(subpasta_processo_teste, "exemplo_texto.txt")
            with open(txt_teste_path, "w", encoding="utf-8") as f_txt:
                f_txt.write("Olá Priscila,\n\n")
                f_txt.write("Este é um arquivo de texto que será convertido para PDF.\n")
                f_txt.write("Ele contém múltiplas linhas e alguns caracteres acentuados como á, ç, õ.\n\n")
                f_txt.write(
                    "Uma linha bem longa para testar a quebra de linha automática se implementada de forma simples: " + "palavra " * 20 + "\n")
                f_txt.write("Fim do arquivo de texto.")
            arquivos_teste_para_unir.append(txt_teste_path)
            print(f"Arquivo de texto de teste criado: {txt_teste_path}")
        except Exception as e:
            print(f"Falha ao criar arquivo de texto de teste: {e}")

        # 4. Arquivo "tipo MS-DOS" (simulado)
        try:
            msdos_teste_path = os.path.join(subpasta_processo_teste, "simulado_msdos.rel")
            with open(msdos_teste_path, "w", encoding="cp850",
                      errors="ignore") as f_msdos:  # Tenta cp850 comum em DOS BR
                f_msdos.write("Relatório do Sistema Antigo - Processo XYZ\n")
                f_msdos.write("------------------------------------------\n")
                f_msdos.write("Item 1: Valor A\n")
                f_msdos.write("Item 2: Valor B\n")
            arquivos_teste_para_unir.append(msdos_teste_path)
            print(f"Arquivo 'tipo MS-DOS' de teste criado: {msdos_teste_path}")
        except Exception as e:
            print(f"Falha ao criar arquivo MS-DOS de teste: {e}")

        if arquivos_teste_para_unir:  # Verifica se a lista de arquivos de teste não está vazia
            print(f"\nArquivos originais para unificação: {arquivos_teste_para_unir}")
            pdf_unificado_final = criar_pdf_unificado(arquivos_teste_para_unir, numero_processo_teste_uniao,
                                                      config.PASTA_COMPROVANTES)
            if pdf_unificado_final and os.path.exists(pdf_unificado_final):
                print(f"\n--- Teste Unificação: PDF unificado criado em: {pdf_unificado_final} ---")
                print("Por favor, verifique o arquivo.")
            else:
                print("\n--- Teste Unificação: Falha ao criar PDF unificado ou arquivo final não encontrado. ---")
        else:
            print("\n--- Teste Unificação: Nenhum arquivo de teste foi criado/adicionado para unificação. ---")
            print(f"Verifique a pasta {subpasta_processo_teste}.")

    print("\n--- Fim do Teste do Módulo pdf_processor.py ---")