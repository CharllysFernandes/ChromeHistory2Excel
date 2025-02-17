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

# Salva os dados em um arquivo Excel
df.to_excel("historico_chrome.xlsx", index=False)

# Fecha a conexão e remove o backup
conn.close()
os.remove(backup_path)

print("Histórico exportado com sucesso para historico_chrome.xlsx")