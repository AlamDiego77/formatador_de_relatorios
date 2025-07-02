import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side

# Caminhos
modelo_path = "modelo_incidentes_formatados.xlsx"
pasta_dados = "dados_xlsx"
saida_path = "Incidentes_Formatados_Final.xlsx"

# Estilos base
altura_linha = 75
fonte_padrao = Font(name='Calibri', size=11)
alinhamento_celula = Alignment(horizontal='center', vertical='center', wrap_text=True)
borda_padrao = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

# Carrega modelo
wb = load_workbook(modelo_path)
ws = wb.active

# Identifica onde começa a tabela (linha de cabeçalho + 1)
linha_inicio = 2
cabecalho_modelo = [cell.value for cell in ws[linha_inicio - 1] if cell.value]

# Lê e junta os dados de todos os arquivos
todos_dados = pd.DataFrame()
for nome_arq in os.listdir(pasta_dados):
    if nome_arq.endswith(".xlsx"):
        df = pd.read_excel(os.path.join(pasta_dados, nome_arq))
        if "Titulo" in df.columns:
            df = df.drop(columns=["Titulo"])
        todos_dados = pd.concat([todos_dados, df], ignore_index=True)

# Remove linhas completamente vazias
todos_dados = todos_dados.dropna(how='all')

# Reorganiza as colunas para seguir a ordem do modelo
todos_dados = todos_dados[[col for col in cabecalho_modelo if col in todos_dados.columns]]

# Insere os dados um a um com formatação
for i, row in todos_dados.iterrows():
    linha_excel = linha_inicio + i
    ws.insert_rows(linha_excel)
    ws.row_dimensions[linha_excel].height = altura_linha
    for j, col_name in enumerate(cabecalho_modelo):
        valor = row[col_name] if col_name in row else ""
        col_letter = get_column_letter(j + 1)
        celula = ws[f"{col_letter}{linha_excel}"]
        celula.value = valor
        celula.font = fonte_padrao
        celula.alignment = alinhamento_celula
        celula.border = borda_padrao

# Remove linhas em branco após a última inserção
ultima_linha_usada = linha_inicio + len(todos_dados)
total_linhas = ws.max_row

if total_linhas > ultima_linha_usada:
    ws.delete_rows(ultima_linha_usada, total_linhas - ultima_linha_usada)

# Salva o resultado final
wb.save(saida_path)
print("✅ Arquivo final salvo como:", saida_path)
