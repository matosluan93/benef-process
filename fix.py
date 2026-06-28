import re

with open('app.py', encoding='utf-8') as f:
    c = f.read()

old_start = c.find('def parse_mixed_dates')
old_end   = c.find('\ndef ', old_start + 1)
old_fn    = c[old_start:old_end]

new_fn = (
    "def parse_mixed_dates(s):\n"
    "    try:\n"
    "        if pd.isna(s) or str(s).strip() == '': return pd.NaT\n"
    "    except Exception: pass\n"
    "    val = str(s).strip()\n"
    "    if val.replace('.', '', 1).isdigit():\n"
    "        try: return pd.to_datetime(float(val), origin='1899-12-30', unit='D')\n"
    "        except Exception: pass\n"
    "    if re.match(r'^\\d{4}[-/]\\d', val):\n"
    "        return pd.to_datetime(val, dayfirst=False, errors='coerce')\n"
    "    return pd.to_datetime(val, dayfirst=True, errors='coerce')"
)

if old_fn in c:
    c = c.replace(old_fn, new_fn)
    with open('app.py', 'w', encoding='utf-8') as f:
        f.write(c)
    print('CORRIGIDO com sucesso')
else:
    print('AVISO: funcao nao encontrada no formato esperado')
