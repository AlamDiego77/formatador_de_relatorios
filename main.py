import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, Border, Side

# --- 1. CONFIGURAÇÃO E ESTILOS ---

# Caminhos
modelo_path = "modelo_incidentes_formatados.xlsx"
pasta_dados = "dados_xlsx"
arquivo_saida = "Incidentes_Formatados_Final.xlsx"

# Estilos base para aplicar nas novas linhas
altura_linha = 75
fonte_padrao = Font(name="Calibri", size=11)
alinhamento_celula = Alignment(horizontal="center", vertical="center", wrap_text=True)
borda_padrao = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# --- CABEÇALHOS DO SEU MODELO (ORDEM EXATA) ---
cabecalhos_modelo = [
    "LOJA", "REGIONAL", "ABERTURA", "FECHAMENTO", "CAUSA", 
    "TIPO", "STATUS", "LINK OPERANDO", "DATA", "INICIO", 
    "FIM", "IMPACTO", "DISPONIBILIDADE", "SOLUÇÃO"
]

# --- 2. CARREGAMENTO DO MODELO ---

try:
    # Carrega o workbook modelo
    wb = load_workbook(modelo_path)
    ws = wb.active
except FileNotFoundError:
    print(f"ERRO: Arquivo modelo \'{modelo_path}\' não encontrado.")
    exit()

# --- 3. PROCESSAMENTO E INSERÇÃO DOS DADOS (Arquivo por Arquivo) ---

# Define a linha inicial para inserção dos dados (linha 2, após o cabeçalho)
linha_atual = 2
print(f"Iniciando inserção de dados na planilha a partir da linha {linha_atual}...")

# Processa cada arquivo de dados na pasta
for nome_arquivo in os.listdir(pasta_dados):
    if nome_arquivo.endswith(".xlsx"):
        caminho_arquivo = os.path.join(pasta_dados, nome_arquivo)
        print(f"Lendo arquivo: {nome_arquivo}")

        # Lê a planilha com pandas
        df = pd.read_excel(caminho_arquivo)

        # --- APLICAÇÃO DA REGRA: REMOVER COLUNA "TITULO" ---
        if "Titulo" in df.columns:
            df = df.drop(columns=["Titulo"])
            print(f"  - Coluna \"Titulo\" removida do arquivo {nome_arquivo}.")

        # --- REORDENAR COLUNAS PARA CORRESPONDER AO MODELO ---
        # Cria um novo DataFrame com as colunas na ordem do modelo.
        # Se uma coluna do modelo não existir no df, ela será preenchida com NaN (vazio).
        df_reordenado = pd.DataFrame(columns=cabecalhos_modelo)
        for col in cabecalhos_modelo:
            if col in df.columns:
                df_reordenado[col] = df[col]
            else:
                df_reordenado[col] = None # Garante que a coluna exista e esteja vazia
        
        # --- APLICAÇÃO DA REGRA DE NEGÓCIO: "MATRIZ TI - Cuiabá" ---
        if not df_reordenado.empty:
            # As colunas agora estão na ordem do modelo, então podemos usar os nomes diretamente
            df_reordenado.loc[df_reordenado["LOJA"] == "MATRIZ TI - Cuiabá", ["ABERTURA", "FECHAMENTO"]] = ["06:00", "22:00"]
            print(f"  - Regra \"MATRIZ TI - Cuiabá\" aplicada no arquivo {nome_arquivo}.")

        # --- INSERÇÃO E FORMATAÇÃO (com Openpyxl) ---
        # Itera sobre as linhas do DataFrame REORDENADO
        for linha_dados in dataframe_to_rows(df_reordenado, index=False, header=False):
            # Adiciona a linha de dados na posição correta
            for col_idx, valor_celula in enumerate(linha_dados, 1):
                ws.cell(row=linha_atual, column=col_idx, value=valor_celula)
            
            # Aplica os estilos de formatação a essa nova linha
            ws.row_dimensions[linha_atual].height = altura_linha
            for celula in ws[linha_atual]:
                celula.font = fonte_padrao
                celula.alignment = alinhamento_celula
                celula.border = borda_padrao
                
            # Incrementa o contador para a próxima linha
            linha_atual += 1

# --- 4. SALVAR O RESULTADO ---

try:
    wb.save(arquivo_saida)
    print(f"\n✅ Planilha final salva com sucesso como \'{arquivo_saida}\'")
except Exception as e:
    print(f"\nERRO ao salvar o arquivo final: {e}")
