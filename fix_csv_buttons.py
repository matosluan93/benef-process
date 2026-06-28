path = 'templates/index.html'
with open(path, encoding='utf-8') as f:
    html = f.read()

old = """        <button style="flex:1;min-width:160px;padding:10px 0;border:none;border-radius:8px;font-weight:600;font-size:13px;cursor:pointer;background:#0F2A56;color:#fff"
                onclick="dl('totalpass_csv')">⬇ TotalPass · CSV</button>
        <button style="flex:1;min-width:160px;padding:10px 0;border:none;border-radius:8px;font-weight:600;font-size:13px;cursor:pointer;background:#E91E8C;color:#fff"
                onclick="dl('wellhub_csv_zip')">⬇ Wellhub · ZIP CSV</button>
        <button style="flex:1;min-width:160px;padding:10px 0;border:none;border-radius:8px;font-weight:600;font-size:13px;cursor:pointer;background:#F59E0B;color:#fff"
                onclick="dl('newvalue_csv')">⬇ New Value · CSV</button>
        <button style="flex:1;min-width:160px;padding:10px 0;border:none;border-radius:8px;font-weight:600;font-size:13px;cursor:pointer;background:#1A4A8A;color:#fff"
                onclick="dl('auditppt')">📊 Apresentação PPT</button>"""

new = """        <button style="flex:1;min-width:160px;padding:10px 0;border:none;border-radius:8px;font-weight:600;font-size:13px;cursor:pointer;background:#1A4A8A;color:#fff"
                onclick="dl('auditppt')">📊 Apresentação PPT</button>"""

if old in html:
    html = html.replace(old, new)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(html)
    print("✓ Botões CSV removidos — seção Downloads adicionais ficou só com PPT")
else:
    print("✗ Trecho não encontrado")
