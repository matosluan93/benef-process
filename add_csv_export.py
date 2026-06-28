"""
add_csv_export.py
Adiciona geração de CSV no formato exato de cada cliente ao BenefProcess.
Execute UMA VEZ dentro da pasta do projeto:
    py add_csv_export.py

O que é adicionado (zero mudança na análise):
  - Geração dos CSVs logo após os Excels já existentes em api_process
  - 3 novas rotas de download:
      GET /api/download/<session_id>/totalpass_csv
      GET /api/download/<session_id>/newvalue_csv
      GET /api/download/<session_id>/wellhub_csv_zip
"""
import os, ast

app_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'app.py')

with open(app_path, encoding='utf-8') as f:
    code = f.read()

# ── 1. Funções de geração de CSV ─────────────────────────────────────────────
csv_functions = '''
# ─── Geração de CSV por cliente ──────────────────────────────────────────────
# Colunas e mapeamentos definidos pelos layouts oficiais de cada cliente.
# Nenhuma regra de negócio é alterada aqui — só o formato de saída.

def write_totalpass_csv(tp_rows, filepath):
    """
    Layout TotalPass: CNPJ_EMPRESA | E-MAIL_COLABORADOR | MATRÍCULA
    """
    import csv
    with open(filepath, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=['CNPJ_EMPRESA', 'E-MAIL_COLABORADOR', 'MATRÍCULA'])
        writer.writeheader()
        for row in tp_rows:
            writer.writerow({
                'CNPJ_EMPRESA':        row.get('CNPJ da Empresa', ''),
                'E-MAIL_COLABORADOR':  row.get('Email Funcional', ''),
                'MATRÍCULA':           row.get('Matrícula', ''),
            })

def write_newvalue_csv(nv_rows, filepath):
    """
    Layout New Value: Nome | Nome Empresa | CNPJ Empresa | Número de CPF
    """
    import csv
    with open(filepath, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=['Nome', 'Nome Empresa', 'CNPJ Empresa', 'Número de CPF'])
        writer.writeheader()
        for row in nv_rows:
            writer.writerow({
                'Nome':           row.get('Nome do profissional', ''),
                'Nome Empresa':   row.get('Nome da Empresa', ''),
                'CNPJ Empresa':   row.get('CNPJ da Empresa', ''),
                'Número de CPF':  row.get('CPF do profissional', ''),
            })

def write_wellhub_csv(wh_rows, filepath):
    """
    Layout Wellhub — 10 colunas com nomes exatos do sistema Wellhub.
    """
    import csv
    COLS = [
        'Name (nome) obrigatório',
        'Email obrigatório',
        'National ID (cpf) obrigatório',
        'Employee ID (matrícula) obrigatório',
        'Department',
        'Cost Center',
        'Office Zip Code',
        'Payroll ID',
        'Payroll Enabled (folha de pagamento habilitada?) obrigatório',
        'Employee Segment',
    ]
    with open(filepath, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=COLS)
        writer.writeheader()
        for row in wh_rows:
            writer.writerow({
                'Name (nome) obrigatório':                                   row.get('Name', ''),
                'Email obrigatório':                                         row.get('Email', ''),
                'National ID (cpf) obrigatório':                             row.get('National ID', ''),
                'Employee ID (matrícula) obrigatório':                       row.get('Employee ID', ''),
                'Department':                                                row.get('Department', ''),
                'Cost Center':                                               row.get('Cost Center', ''),
                'Office Zip Code':                                           row.get('Office Zip Code', ''),
                'Payroll ID':                                                row.get('Payroll ID', ''),
                'Payroll Enabled (folha de pagamento habilitada?) obrigatório': row.get('Payroll Enabled', 'YES'),
                'Employee Segment':                                          '',
            })

'''

if 'write_totalpass_csv' not in code:
    # Insere antes das rotas Flask
    code = code.replace("@app.route('/')\ndef index():", csv_functions + "@app.route('/')\ndef index():")
    print("✓ Funções write_totalpass_csv / write_newvalue_csv / write_wellhub_csv adicionadas")
else:
    print("  Funções CSV já existem, pulando")

# ── 2. Geração dos CSVs dentro de api_process ─────────────────────────────────
# Ancora após a linha que grava o auditoria.xlsx — que já existe no código.
trigger_audit = "write_styled_excel(audit_rows, os.path.join(session_dir, 'auditoria.xlsx'), 'Auditoria',  AUDIT_COLS)"
csv_generation = """write_styled_excel(audit_rows, os.path.join(session_dir, 'auditoria.xlsx'), 'Auditoria',  AUDIT_COLS)

        # ── CSVs no formato exato de cada cliente ──────────────────────────
        write_totalpass_csv(tp_rows,  os.path.join(session_dir, 'totalpass.csv'))
        write_newvalue_csv(nv_rows,   os.path.join(session_dir, 'newvalue.csv'))

        # Wellhub: um CSV por empresa (espelha os xlsx já gerados)
        for meta in wh_file_meta:
            safe = re.sub(r'[^\\w]', '_', meta['empresa']).upper().strip('_')
            csv_path = os.path.join(session_dir, f"wh_{re.sub(r'[^\\w]','',meta['cnpj'])}.csv")
            write_wellhub_csv(meta.get('_rows_raw', []), csv_path)
            meta['csv_path'] = csv_path
            meta['csv_name'] = f"{safe}_Wellhub.csv"

        # ZIP de todos os CSVs Wellhub
        wh_csv_zip = os.path.join(session_dir, 'wellhub_csv_todos.zip')
        with zipfile.ZipFile(wh_csv_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
            for meta in wh_file_meta:
                if os.path.exists(meta.get('csv_path', '')):
                    zf.write(meta['csv_path'], meta['csv_name'])"""

if 'write_totalpass_csv(tp_rows' not in code:
    if trigger_audit in code:
        code = code.replace(trigger_audit, csv_generation)
        print("✓ Geração dos CSVs inserida em api_process")
    else:
        print("  AVISO: âncora não encontrada. Insira manualmente após write_styled_excel(audit_rows...)")
else:
    print("  Geração CSV já existe em api_process, pulando")

# ── 3. Armazena rows raw do Wellhub para uso no CSV ──────────────────────────
# O pipeline Wellhub já cria os rows via to_wellhub_row; precisamos guardá-los
# junto com o meta para o CSV. Âncora na linha que monta wh_by_cnpj.
trigger_wh_rows = "wh_by_cnpj[cnpj]['rows'].append(to_wellhub_row(r, internal_map))"
wh_rows_with_raw = """wh_row = to_wellhub_row(r, internal_map)
            wh_by_cnpj[cnpj]['rows'].append(wh_row)"""

if "wh_row = to_wellhub_row" not in code:
    code = code.replace(trigger_wh_rows, wh_rows_with_raw)
    print("✓ Captura de wh_row ajustada")

# Guarda os rows no meta para uso no CSV
trigger_meta_append = "wh_file_meta.append({"
meta_with_rows = """# guarda rows para uso no CSV
            _rows_raw = grp['rows'][:]
            wh_file_meta.append({"""

if "'_rows_raw'" not in code and trigger_meta_append in code:
    code = code.replace(trigger_meta_append, meta_with_rows, 1)
    # Adiciona _rows_raw no dict
    code = code.replace(
        "'filename': fname,\n            })",
        "'filename': fname,\n                '_rows_raw': _rows_raw,\n            })"
    )
    print("✓ _rows_raw adicionado ao wh_file_meta")
else:
    print("  _rows_raw já existe ou âncora não encontrada, pulando")

# ── 4. Rotas de download CSV ──────────────────────────────────────────────────
new_routes = '''
@app.route('/api/download/<session_id>/totalpass_csv')
def download_totalpass_csv(session_id):
    path = os.path.join(TEMP_DIR, session_id, 'totalpass.csv')
    if not os.path.exists(path): return 'Arquivo não encontrado.', 404
    return send_file(path, as_attachment=True, download_name='TotalPass_Colaboradores.csv')

@app.route('/api/download/<session_id>/newvalue_csv')
def download_newvalue_csv(session_id):
    path = os.path.join(TEMP_DIR, session_id, 'newvalue.csv')
    if not os.path.exists(path): return 'Arquivo não encontrado.', 404
    return send_file(path, as_attachment=True, download_name='NewValue_Colaboradores.csv')

@app.route('/api/download/<session_id>/wellhub_csv_zip')
def download_wellhub_csv_zip(session_id):
    path = os.path.join(TEMP_DIR, session_id, 'wellhub_csv_todos.zip')
    if not os.path.exists(path): return 'Arquivo não encontrado.', 404
    return send_file(path, as_attachment=True, download_name='Wellhub_CSVs_Todos.zip')

@app.route('/api/wellhub-csv-zip-filtered/<session_id>', methods=['POST'])
def wellhub_csv_zip_filtered(session_id):
    """Download ZIP com CSVs Wellhub apenas das empresas selecionadas."""
    session_dir    = os.path.join(TEMP_DIR, session_id)
    meta_path      = os.path.join(session_dir, 'wh_meta.json')
    if not os.path.exists(meta_path): return jsonify({'error': 'Sessão expirada.'}), 400
    selected_cnpjs = set(request.json.get('cnpjs', []))
    with open(meta_path, encoding='utf-8') as f: wh_file_meta = json.load(f)
    selected = [m for m in wh_file_meta if m['cnpj'] in selected_cnpjs]
    if not selected: return jsonify({'error': 'Nenhuma empresa selecionada.'}), 400
    zip_path = os.path.join(session_dir, 'wellhub_csv_filtrado.zip')
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for meta in selected:
            csv_p = meta.get('csv_path', '')
            if os.path.exists(csv_p):
                zf.write(csv_p, meta.get('csv_name', os.path.basename(csv_p)))
    return send_file(zip_path, as_attachment=True,
                     download_name=f'Wellhub_CSV_{len(selected)}_Empresas.zip')

'''

if 'download_totalpass_csv' not in code:
    code = code.replace("if __name__ == '__main__':", new_routes + "if __name__ == '__main__':")
    print("✓ 4 rotas CSV adicionadas")
else:
    print("  Rotas CSV já existem, pulando")

# ── 5. Valida e salva ─────────────────────────────────────────────────────────
try:
    ast.parse(code)
    print("✓ Sintaxe Python válida")
except SyntaxError as e:
    print(f"✗ Erro de sintaxe: {e}")
    exit(1)

with open(app_path, 'w', encoding='utf-8') as f:
    f.write(code)

print(f"\n✓ app.py atualizado: {app_path}")
print("\nURLs de download disponíveis após processar:")
print("  /api/download/<session_id>/totalpass_csv    → TotalPass_Colaboradores.csv")
print("  /api/download/<session_id>/newvalue_csv     → NewValue_Colaboradores.csv")
print("  /api/download/<session_id>/wellhub_csv_zip  → Wellhub_CSVs_Todos.zip (todas empresas)")
print("  POST /api/wellhub-csv-zip-filtered/<id>     → ZIP filtrado por CNPJ selecionado")
