import re

with open('app.py', encoding='utf-8') as f:
    lines = f.readlines()

anchor = "        write_styled_excel(audit_rows, os.path.join(session_dir, 'auditoria.xlsx'), 'Auditoria',  AUDIT_COLS)\n"

insert = """
        # ── Gera CSVs no formato de cada cliente ──────────────────────────
        write_totalpass_csv(tp_rows, os.path.join(session_dir, 'totalpass.csv'))
        write_newvalue_csv(nv_rows,  os.path.join(session_dir, 'newvalue.csv'))

        # Wellhub: um CSV por empresa (lê do xlsx já gerado)
        import csv as _csv, pandas as _pd
        wh_csv_zip_path = os.path.join(session_dir, 'wellhub_csv_todos.zip')
        _wh_csvs = []
        for meta in wh_file_meta:
            _df = _pd.read_excel(meta['filepath'], dtype=str)
            _safe = re.sub(r'[^\\w]', '_', meta['empresa']).upper().strip('_')
            _csv_path = os.path.join(session_dir, f"wh_{re.sub(r'[^\\w]','',meta['cnpj'])}.csv")
            _csv_name = f"{_safe}_Wellhub.csv"
            COLS = ['Name (nome) obrigatório','Email obrigatório','National ID (cpf) obrigatório',
                    'Employee ID (matrícula) obrigatório','Department','Cost Center',
                    'Office Zip Code','Payroll ID',
                    'Payroll Enabled (folha de pagamento habilitada?) obrigatório','Employee Segment']
            with open(_csv_path, 'w', newline='', encoding='utf-8-sig') as _f:
                _w = _csv.writer(_f)
                _w.writerow(COLS)
                for _, _row in _df.iterrows():
                    _w.writerow([
                        _row.get('Name',''), _row.get('Email',''),
                        _row.get('National ID',''), _row.get('Employee ID',''),
                        '','','','',
                        _row.get('Payroll Enabled','YES'), ''
                    ])
            meta['csv_path'] = _csv_path
            meta['csv_name'] = _csv_name
            _wh_csvs.append((_csv_path, _csv_name))

        import zipfile as _zf
        with _zf.ZipFile(wh_csv_zip_path, 'w', _zf.ZIP_DEFLATED) as _z:
            for _p, _n in _wh_csvs:
                _z.write(_p, _n)
        # ──────────────────────────────────────────────────────────────────
"""

found = False
new_lines = []
for line in lines:
    new_lines.append(line)
    if line == anchor and not found:
        new_lines.append(insert)
        found = True

if found:
    with open('app.py', 'w', encoding='utf-8') as f:
        f.writelines(new_lines)
    print("✓ Chamadas CSV inseridas com sucesso")
else:
    print("✗ Âncora não encontrada — tente: grep -n 'write_styled_excel.*auditoria' app.py")
