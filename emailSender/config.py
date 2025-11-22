# config.py
import os
from dotenv import load_dotenv # Importa a função para carregar o .env

# Carrega as variáveis do arquivo .env para o ambiente do sistema operacional.
# É importante chamar load_dotenv() no início do script, antes de tentar acessar as variáveis.
load_dotenv()

# --- Configurações Carregadas do Arquivo .env ---
# Acessa as variáveis usando os.getenv().
# O segundo argumento de os.getenv() é um valor padrão opcional se a variável não for encontrada no .env.
EMAIL_REMETENTE = os.getenv("EMAIL_REMETENTE")
SENHA_REMETENTE = os.getenv("SENHA_REMETENTE") # Para senhas, geralmente não se define um valor padrão por segurança.
SERVIDOR_SMTP = os.getenv("SERVIDOR_SMTP", "smtp.gmail.com") # Exemplo de valor padrão
PORTA_SMTP_STR = os.getenv("PORTA_SMTP", "587") # Exemplo de valor padrão, lido como string

# Validação e conversão da porta SMTP
if PORTA_SMTP_STR and PORTA_SMTP_STR.isdigit():
    PORTA_SMTP = int(PORTA_SMTP_STR)
else:
    print(f"AVISO: PORTA_SMTP '{PORTA_SMTP_STR}' não é um número válido ou não foi encontrada no .env. Usando porta padrão 587.")
    PORTA_SMTP = 587 # Porta padrão se a do .env for inválida

# Verificação crítica para variáveis essenciais (especialmente a senha)
if not EMAIL_REMETENTE:
    print("ERRO CRÍTICO: A variável EMAIL_REMETENTE não foi configurada no arquivo .env.")
    print("Por favor, defina EMAIL_REMETENTE no seu arquivo .env e tente novamente.")
    # Em um cenário real, você poderia querer sair do programa aqui:
    # exit("Configuração de email ausente.")
if not SENHA_REMETENTE:
    print("ERRO CRÍTICO: A variável SENHA_REMETENTE não foi configurada no arquivo .env.")
    print("Por favor, defina SENHA_REMETENTE no seu arquivo .env e tente novamente.")
    # exit("Configuração de senha ausente.")

# --- Outras Configurações do Projeto (Caminhos, Nomes de Arquivos, Colunas) ---
PASTA_PROCESSOS_PDF = r"C:\Users\Priscila\APSDJ\ProcessosBaixadosTemp"
PASTA_APSDJ = r"C:\Users\Priscila\APSDJ"
PASTA_COMPROVANTES = r"C:\Users\Priscila\APSDJ\Comprovantes"
# Certifique-se que o nome do arquivo e a extensão .xls estão corretos
NOME_ARQUIVO_EXCEL_EMAILS = "TESTE EMAILS VARAS CIVEIS ESTADO DE SÃO PAULO-FINAL.xls"
CAMINHO_PLANILHA_EMAILS = os.path.join(PASTA_APSDJ, NOME_ARQUIVO_EXCEL_EMAILS)

# Colunas esperadas na planilha Excel
# !!! AJUSTADO COM BASE NA IMAGEM DA SUA PLANILHA !!!
# Verifique se estes nomes correspondem EXATAMENTE aos cabeçalhos da sua planilha.
COLUNA_VARA_EXCEL = "Vara"      # <<< ALTERADO
COLUNA_COMARCA_EXCEL = "Comarca"  # <<< ALTERADO
COLUNA_EMAIL_EXCEL = "e-mail"   # <<< ALTERADO (atenção ao minúsculo e hífen)

# Arquivo de log para PDFs já processados
ARQUIVO_PROCESSADOS_LOG = os.path.join(PASTA_APSDJ, "processos_ja_enviados.txt")

# Pastas para mover os PDFs após processamento (dentro da PASTA_PROCESSOS_PDF)
PASTA_PROCESSADOS_SUCESSO = os.path.join(PASTA_PROCESSOS_PDF, "ProcessadosComSucesso")
PASTA_PROCESSADOS_ERRO = os.path.join(PASTA_PROCESSOS_PDF, "ProcessadosComErro")

# Opcional: Imprimir uma confirmação de que as configurações foram carregadas (para depuração)
# print(f"Configurações carregadas: Email Remetente: {EMAIL_REMETENTE}, Servidor SMTP: {SERVIDOR_SMTP}:{PORTA_SMTP}")
# print(f"Planilha de emails: {CAMINHO_PLANILHA_EMAILS}")
# print(f"Colunas Excel: Vara='{COLUNA_VARA_EXCEL}', Comarca='{COLUNA_COMARCA_EXCEL}', Email='{COLUNA_EMAIL_EXCEL}'")