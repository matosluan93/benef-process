path = 'templates/index.html'
with open(path, encoding='utf-8') as f:
    html = f.read()

old = """    <div id="downloadAllWrap" class="alert-green" style="margin-bottom:18px">
      <span style="font-weight:500">Baixar os três arquivos tratados de uma vez:</span>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <button class="btn btn-teal btn-sm" onclick="downloadAll()">⬇ Baixar tudo (Excel)</button>
        <button class="btn btn-ghost btn-sm" onclick="downloadAllCSV()">⬇ Baixar tudo (CSV)</button>
      </div>
    </div>"""

if old in html:
    html = html.replace(old, '')
    with open(path, 'w', encoding='utf-8') as f:
        f.write(html)
    print("✓ Seção 'Baixar tudo' removida")
else:
    # tenta variante sem o botão CSV
    old2 = """    <div id="downloadAllWrap" class="alert-green" style="margin-bottom:18px">
      <span style="font-weight:500">Baixar os três arquivos tratados de uma vez:</span>
      <button class="btn btn-teal btn-sm" onclick="downloadAll()">⬇ Baixar tudo</button>
    </div>"""
    if old2 in html:
        html = html.replace(old2, '')
        with open(path, 'w', encoding='utf-8') as f:
            f.write(html)
        print("✓ Seção 'Baixar tudo' removida (variante)")
    else:
        print("✗ Trecho não encontrado")
