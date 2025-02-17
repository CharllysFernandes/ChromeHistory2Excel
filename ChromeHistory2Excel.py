import os
import sqlite3
import pandas as pd
from datetime import datetime

# Caminho do banco de dados do Chrome (ajuste conforme necessário)
chrome_history_path = os.path.expanduser("~") + r"\AppData\Local\Google\Chrome\User Data\Default\History"

# Copia o banco de dados para evitar erros de bloqueio
backup_path = "history_backup"
os.system(f'copy "{chrome_history_path}" "{backup_path}"')

# Conecta ao banco de dados SQLite
conn = sqlite3.connect(backup_path)
cursor = conn.cursor()

# Consulta para obter os dados do histórico
query = """
    SELECT datetime(last_visit_time/1000000-11644473600, 'unixepoch', 'localtime') AS visit_time, url 
    FROM urls
    ORDER BY last_visit_time DESC
"""

# Executa a consulta e armazena os resultados em um DataFrame
cursor.execute(query)
rows = cursor.fetchall()
df = pd.DataFrame(rows, columns=["Data e Horário", "URL"])

# Separa a data do horário em duas colunas
df["Data"] = df["Data e Horário"].apply(lambda x: x.split(' ')[0])
df["Horário"] = df["Data e Horário"].apply(lambda x: x.split(' ')[1])
df = df.drop(columns=["Data e Horário"])

# Salva os dados em um arquivo Excel com hiperlinks
with pd.ExcelWriter("historico_chrome.xlsx", engine='xlsxwriter') as writer:
    df[["Data", "Horário", "URL"]].to_excel(writer, index=False, sheet_name='Histórico')
    workbook = writer.book
    worksheet = writer.sheets['Histórico']
    
    # Adiciona hiperlinks
    for row_num, url in enumerate(df["URL"], start=2):  # Excel rows are 1-indexed and header is in row 1
        worksheet.write_url(f'C{row_num}', url, string=url)

# Fecha a conexão e remove o backup
conn.close()
os.remove(backup_path)

print("Histórico exportado com sucesso para historico_chrome.xlsx")