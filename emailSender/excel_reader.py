# excel_reader.py
import os
import pandas as pd
from typing import Optional

# Tente importar config e avise se faltar
try:
    import config
except ImportError:
    print("ERRO CRÍTICO em excel_reader.py: O arquivo config.py não foi encontrado ou não pôde ser importado.")


    class FallbackConfig:  # Definições de fallback
        CAMINHO_PLANILHA_EMAILS = "planilha_nao_configurada.xls"
        COLUNA_VARA_EXCEL = "Vara"
        COLUNA_COMARCA_EXCEL = "Comarca"
        COLUNA_EMAIL_EXCEL = "e-mail"


    config = FallbackConfig()


def buscar_email_vara(vara_civel_pdf: str, comarca_pdf: str) -> Optional[str]:
    """
    Busca o email da vara e comarca na planilha Excel, usando 'xlrd' como engine,
    com depuração detalhada do que o pandas lê.
    """
    if not os.path.exists(config.CAMINHO_PLANILHA_EMAILS):
        print(f"  [Excel Reader] Erro: Planilha de emails não encontrada em '{config.CAMINHO_PLANILHA_EMAILS}'")
        return None
    try:
        # Tenta ler a planilha (por padrão, a primeira aba)
        df = pd.read_excel(config.CAMINHO_PLANILHA_EMAILS, engine='xlrd')

        print(
            f"  [Excel Reader] Planilha '{os.path.basename(config.CAMINHO_PLANILHA_EMAILS)}' (.xls) carregada com engine 'xlrd'.")

        # --- NOVOS PRINTS DE DEPURAÇÃO ---
        print(f"  [Excel DEBUG] Colunas detectadas pela pandas: {df.columns.tolist()}")
        print(f"  [Excel DEBUG] Primeiras 3 linhas lidas pela pandas (como DataFrame):\n{df.head(3)}")
        # --- FIM DOS NOVOS PRINTS DE DEPURAÇÃO ---

        print(f"\n  Procurando por Vara PDF: '{vara_civel_pdf}', Comarca PDF: '{comarca_pdf}'.")
        vara_pdf_norm = str(vara_civel_pdf).strip().lower()
        comarca_pdf_norm = str(comarca_pdf).strip().lower()
        print(f"  Valores normalizados do PDF -> Vara: '{vara_pdf_norm}', Comarca: '{comarca_pdf_norm}'")

        for index, row in df.iterrows():
            print(f"\n    --- [Excel DEBUG] Verificando Linha Excel {index + 2} (índice DataFrame: {index}) ---")

            raw_vara_excel = row.get(config.COLUNA_VARA_EXCEL)
            raw_comarca_excel = row.get(config.COLUNA_COMARCA_EXCEL)
            raw_email_excel = row.get(config.COLUNA_EMAIL_EXCEL)

            vara_excel = str(raw_vara_excel if pd.notna(raw_vara_excel) else "").strip().lower()
            comarca_excel = str(raw_comarca_excel if pd.notna(raw_comarca_excel) else "").strip().lower()
            email_excel = str(raw_email_excel if pd.notna(raw_email_excel) else "").strip()

            print(
                f"      Lido da Coluna '{config.COLUNA_VARA_EXCEL}': '{raw_vara_excel}' (Tipo: {type(raw_vara_excel).__name__}) -> Normalizado: '{vara_excel}'")
            print(
                f"      Lido da Coluna '{config.COLUNA_COMARCA_EXCEL}': '{raw_comarca_excel}' (Tipo: {type(raw_comarca_excel).__name__}) -> Normalizado: '{comarca_excel}'")
            print(
                f"      Lido da Coluna '{config.COLUNA_EMAIL_EXCEL}': '{raw_email_excel}' (Tipo: {type(raw_email_excel).__name__}) -> Normalizado: '{email_excel}'")

            match_vara = False
            if vara_excel:
                match_vara = vara_excel in vara_pdf_norm

            match_comarca = False
            if comarca_excel:
                match_comarca = comarca_excel in comarca_pdf_norm

            print(f"      Comparando PDF Vara ('{vara_pdf_norm}') com Excel Vara ('{vara_excel}'). Match? {match_vara}")
            print(
                f"      Comparando PDF Comarca ('{comarca_pdf_norm}') com Excel Comarca ('{comarca_excel}'). Match? {match_comarca}")

            if vara_excel and comarca_excel:
                if match_vara and match_comarca:
                    print(f"    DEBUG: CONDIÇÃO DE MATCH POSITIVA para Vara E Comarca na linha {index + 2}.")
                    if pd.notna(raw_email_excel) and email_excel and "@" in email_excel:
                        print(
                            f"    -> Email VÁLIDO encontrado para Vara: '{vara_civel_pdf}', Comarca: '{comarca_pdf}' -> {email_excel} (Linha {index + 2} da planilha)")
                        return email_excel
                    else:
                        print(
                            f"    -> Correspondência de Vara/Comarca encontrada na Linha {index + 2}, mas o email ('{email_excel}') parece inválido ou está vazio.")

        print(
            f"\n  [Excel Reader] Email não encontrado para Vara: '{vara_civel_pdf}', Comarca: '{comarca_pdf}' na planilha após verificar todas as linhas.")
        return None

    except pd.errors.ParserError as pe:
        print(f"  [Excel Reader] Erro de Parsing ao ler a planilha Excel: {pe}")
        return None
    except ValueError as ve:
        print(f"  [Excel Reader] Erro de Valor ao processar a planilha Excel: {ve}")
        return None
    except FileNotFoundError:
        print(
            f"  [Excel Reader] Erro Crítico: Planilha de emails não foi encontrada em '{config.CAMINHO_PLANILHA_EMAILS}' durante a tentativa de leitura.")
        return None
    except ImportError:
        print(f"  [Excel Reader] Erro de Importação. Certifique-se que 'pandas' e 'xlrd' estão instalados.")
        return None
    except Exception as e:
        print(f"  [Excel Reader] Erro geral e inesperado ao ler ou processar a planilha Excel: {e}")
        print(f"  Tipo de erro: {type(e).__name__}")
        return None


# Bloco de teste
if __name__ == "__main__":
    print("--- Iniciando Teste do Módulo excel_reader.py ---")
    # Certifique-se que config.py está correto e a planilha de teste existe
    # com os cabeçalhos configurados em config.py e uma linha de dados correspondente.
    vara_teste = "1ª vara cível"
    comarca_teste = "Franca SP"

    # Crie uma linha na sua planilha de teste que corresponda a:
    # Coluna configurada como COLUNA_VARA_EXCEL: "1ª" (ou algo que esteja contido em "1ª vara cível")
    # Coluna configurada como COLUNA_COMARCA_EXCEL: "Franca" (ou algo que esteja contido em "franca sp")
    # Coluna configurada como COLUNA_EMAIL_EXCEL: "seu_email_de_teste@exemplo.com"

    print(f"Tentando buscar email para Vara: '{vara_teste}', Comarca: '{comarca_teste}'")
    email_encontrado = buscar_email_vara(vara_teste, comarca_teste)
    if email_encontrado:
        print(f"\n--- Teste: Email encontrado: {email_encontrado} ---")
    else:
        print(f"\n--- Teste: Email NÃO encontrado ---")
    print("\n--- Fim do Teste do Módulo excel_reader.py ---")