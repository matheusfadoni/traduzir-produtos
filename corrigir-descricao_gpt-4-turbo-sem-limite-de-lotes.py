import os
from dotenv import load_dotenv # type: ignore
from google.oauth2.service_account import Credentials # type: ignore
from googleapiclient.discovery import build # type: ignore
from openai import OpenAI # type: ignore
import time
from datetime import datetime

# Carrega variáveis de ambiente do arquivo .env (inclui chave da OpenAI)
load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# Autenticação Google Sheets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_file(
    '/app/credentials-automacao-atendimento-zendesk.json',
    scopes=SCOPES
)
service = build('sheets', 'v4', credentials=creds)

# ID e aba da planilha
spreadsheet_id = '1sbQasVfVH2cEhUHL4mNRATKOsveEKtmlgICK9C5Ku2w'
sheet_name = 'Hoja 1'
range_ler = f"{sheet_name}!A2:G"

# Prompt fixo a ser usado no campo 'system'
with open("prompt1-0-2.txt", "r", encoding="utf-8") as f:
    PROMPT_SYSTEM = f.read()

# Função de log
LOG_PATH = "log_execucao.txt"
def registrar_log(mensagem):
    data_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_PATH, "a", encoding="utf-8") as log_file:
        log_file.write(f"[{data_hora}] {mensagem}\n")

# Prepara os dados no formato SKU|NOME|CATEGORIA|SUBCATEGORIA|MARCA
def montar_entrada_lote(lote):
    linhas = []
    for row in lote:
        sku = row[0] if len(row) > 0 else ''
        nome = row[1].replace('|', ' ') if len(row) > 1 else ''
        cat = row[3].replace('|', ' ') if len(row) > 3 else ''
        subcat = row[4].replace('|', ' ') if len(row) > 4 else ''
        marca = row[5].replace('|', ' ') if len(row) > 5 else ''
        linhas.append(f"{sku}|{nome}|{cat}|{subcat}|{marca}")
    return "\n".join(linhas)

# Converte a resposta da API em um dicionário {SKU: NOME_CORRIGIDO}
def tratar_resposta(resposta):
    linhas = resposta.strip().split("\n")
    saida = {}
    for linha in linhas:
        partes = linha.split("|", 1)
        if len(partes) == 2:
            sku, nome_corrigido = partes
            saida[sku.strip()] = nome_corrigido.strip()
    return saida

try:
    registrar_log("\n______________________\nIniciando leitura da planilha.")
    valores = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_ler
    ).execute().get('values', [])

    pendentes = [(i + 2, row) for i, row in enumerate(valores) if len(row) < 7 or row[6].strip().lower() != 'true']
    atualizados = []

    TAMANHO_LOTE = 4 # Variável para definir quantos produtos serao tratados por requisição
    for i in range(0, len(pendentes), TAMANHO_LOTE):
        lote = pendentes[i:i + TAMANHO_LOTE]
        entrada = montar_entrada_lote([x[1] for x in lote])
        registrar_log(f"Enviando lote {i//TAMANHO_LOTE + 1} para a IA")

        atualizados = []  # Aqui zeramos a cada lote, evitando acumular de outros lotes

        try:
            resposta = client.chat.completions.create(
                model="gpt-4-turbo",
                temperature=0,
                max_tokens=1600,
                messages=[
                    {"role": "system", "content": PROMPT_SYSTEM},
                    {"role": "user", "content": entrada}
                ]
            )
            usage = resposta.usage
            registrar_log(f"Tokens usados no lote {i//TAMANHO_LOTE + 1}: prompt={usage.prompt_tokens}, completion={usage.completion_tokens}, total={usage.total_tokens}")
            registrar_log(f"Lote {i//TAMANHO_LOTE + 1} processado com sucesso pela IA")
            correcoes = tratar_resposta(resposta.choices[0].message.content)

            for linha, row in lote:
                sku = row[0] if len(row) > 0 else ''
                if sku in correcoes:
                    row += [''] * (7 - len(row))
                    row[1] = correcoes[sku]
                    row[6] = True
                    dados = row[:7]  
                    atualizados.append((linha, dados))

            # Subir imediatamente após cada lote processado:
            for linha, dados in atualizados:
                try:
                    resp = service.spreadsheets().values().update(
                        spreadsheetId=spreadsheet_id,
                        range=f"{sheet_name}!A{linha}:G{linha}",
                        valueInputOption="RAW",
                        body={"values": [dados]}
                    ).execute()
                except Exception as e:
                    registrar_log(f"Erro ao atualizar linha {linha}: {type(e).__name__}: {e}")
            registrar_log("Atualizações enviadas para a planilha.")
            time.sleep(180)  # Intervalo antes de processar o próximo lote

        except Exception as e:
            registrar_log(f"Erro no lote {i//TAMANHO_LOTE + 1}: {type(e).__name__}: {e}")

    registrar_log("Fim da execução.")

except Exception as e:
    registrar_log(f"Erro geral: {type(e).__name__}: {e}")
