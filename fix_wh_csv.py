import re

with open('app.py', encoding='utf-8') as f:
    content = f.read()

# Remove a parte do Wellhub do bloco inserido anteriormente
old = """        # Wellhub: um CSV por empresa (lê do xlsx já gerado)
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
        # ──────────────────────────────────────────────────────────────────"""

new = """        # ──────────────────────────────────────────────────────────────────"""

if old in content:
    content = content.replace(old, new)
    print("✓ Wellhub CSV removido da posição errada")
else:
    print("  Bloco Wellhub não encontrado na posição antiga, continuando")

# Adiciona Wellhub CSV no lugar certo: após wh_meta.json ser salvo
anchor = """        with open(os.path.join(session_dir, 'wh_meta.json'), 'w', encoding='utf-8') as jf:
            json.dump(wh_file_meta, jf, ensure_ascii=False)"""

wh_csv_block = """        with open(os.path.join(session_dir, 'wh_meta.json'), 'w', encoding='utf-8') as jf:
            json.dump(wh_file_meta, jf, ensure_ascii=False)

        # Wellhub: um CSV por empresa
        import csv as _csv
        _WCOLS = ['Name (nome) obrigatório','Email obrigatório',
                  'National ID (cpf) obrigatório','Employee ID (matrícula) obrigatório',
                  'Department','Cost Center','Office Zip Code','Payroll ID',
                  'Payroll Enabled (folha de pagamento habilitada?) obrigatório','Employee Segment']
        _wh_csvs = []
        for meta in wh_file_meta:
            _safe     = re.sub(r'[^\\w]', '_', meta['empresa']).upper().strip('_')
            _csv_path = os.path.join(session_dir, f"wh_{re.sub(r'[^\\w]','',meta['cnpj'])}.csv")
            _csv_name = f"{_safe}_Wellhub.csv"
            import pandas as _pd
            _df = _pd.read_excel(meta['filepath'], dtype=str)
            with open(_csv_path, 'w', newline='', encoding='utf-8-sig') as _f:
                _w = _csv.writer(_f)
                _w.writerow(_WCOLS)
                for _, _row in _df.iterrows():
                    _w.writerow([
                        _row.get('Name',''), _row.get('Email',''),
                        _row.get('National ID',''), _row.get('Employee ID',''),
                        '','','','', _row.get('Payroll Enabled','YES'), ''
                    ])
            _wh_csvs.append((_csv_path, _csv_name))

        import zipfile as _zf
        _wh_zip = os.path.join(session_dir, 'wellhub_csv_todos.zip')
        with _zf.ZipFile(_wh_zip, 'w', _zf.ZIP_DEFLATED) as _z:
            for _p, _n in _wh_csvs:
                _z.write(_p, _n)"""

if 'wellhub_csv_todos.zip' not in content:
    if anchor in content:
        content = content.replace(anchor, wh_csv_block)
        print("✓ Wellhub CSV inserido no lugar certo")
    else:
        print("✗ Âncora wh_meta.json não encontrada")
else:
    print("  Wellhub CSV já está no lugar certo, pulando")

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)

print("✓ app.py salvo")
