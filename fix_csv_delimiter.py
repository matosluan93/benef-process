import re

with open('app.py', encoding='utf-8') as f:
    content = f.read()

# Corrige write_totalpass_csv
content = content.replace(
    "writer = csv.DictWriter(f, fieldnames=['CNPJ_EMPRESA', 'E-MAIL_COLABORADOR', 'MATRÍCULA'])",
    "writer = csv.DictWriter(f, fieldnames=['CNPJ_EMPRESA', 'E-MAIL_COLABORADOR', 'MATRÍCULA'], delimiter=';')"
)

# Corrige write_newvalue_csv
content = content.replace(
    "writer = csv.DictWriter(f, fieldnames=['Nome', 'Nome Empresa', 'CNPJ Empresa', 'Número de CPF'])",
    "writer = csv.DictWriter(f, fieldnames=['Nome', 'Nome Empresa', 'CNPJ Empresa', 'Número de CPF'], delimiter=';')"
)

# Corrige write_wellhub_csv (se existir como função)
content = content.replace(
    "writer = csv.DictWriter(f, fieldnames=COLS)",
    "writer = csv.DictWriter(f, fieldnames=COLS, delimiter=';')"
)

# Corrige o CSV do Wellhub gerado inline no api_process
content = content.replace(
    "_w = _csv.writer(_f)",
    "_w = _csv.writer(_f, delimiter=';')"
)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)

print("✓ Delimitador corrigido para ponto e vírgula em todos os CSVs")
print("  Os arquivos vão abrir no Excel com colunas separadas automaticamente")
