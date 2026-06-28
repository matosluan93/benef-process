"""
add_ppt_button.py
Adiciona o botão de PPT e o botão 'Baixar tudo (CSV)' que ficaram faltando.
Execute dentro da pasta do projeto:
    py add_ppt_button.py
"""
import os

html_path = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    'templates', 'index.html'
)

with open(html_path, encoding='utf-8') as f:
    html = f.read()

changed = False

# ── 1. Seção PPT + CSV global (após o card de auditoria) ─────────────────────
PPT_SECTION = """
    <div style="background:#fff;border-radius:12px;box-shadow:0 1px 5px rgba(0,0,0,.08);padding:16px 22px;margin-bottom:14px">
      <div style="font-weight:700;font-size:14px;display:flex;align-items:center;gap:8px;margin-bottom:10px">
        📥 Downloads adicionais
      </div>
      <div style="display:flex;gap:10px;flex-wrap:wrap">
        <button style="flex:1;min-width:160px;padding:10px 0;border:none;border-radius:8px;font-weight:600;font-size:13px;cursor:pointer;background:#0F2A56;color:#fff"
                onclick="dl('totalpass_csv')">⬇ TotalPass · CSV</button>
        <button style="flex:1;min-width:160px;padding:10px 0;border:none;border-radius:8px;font-weight:600;font-size:13px;cursor:pointer;background:#E91E8C;color:#fff"
                onclick="dl('wellhub_csv_zip')">⬇ Wellhub · ZIP CSV</button>
        <button style="flex:1;min-width:160px;padding:10px 0;border:none;border-radius:8px;font-weight:600;font-size:13px;cursor:pointer;background:#F59E0B;color:#fff"
                onclick="dl('newvalue_csv')">⬇ New Value · CSV</button>
        <button style="flex:1;min-width:160px;padding:10px 0;border:none;border-radius:8px;font-weight:600;font-size:13px;cursor:pointer;background:#1A4A8A;color:#fff"
                onclick="dl('auditppt')">📊 Apresentação PPT</button>
      </div>
    </div>"""

ANCHOR_AUDIT = """<button class="btn btn-ghost btn-sm" onclick="dl('auditoria')">⬇ Auditoria</button>
    </div>"""

if 'auditppt' not in html:
    if ANCHOR_AUDIT in html:
        html = html.replace(ANCHOR_AUDIT, ANCHOR_AUDIT + PPT_SECTION)
        print("✓ Seção PPT + Downloads adicionais inserida")
        changed = True
    else:
        print("✗ Âncora do card de auditoria não encontrada — cole o trecho abaixo manualmente")
        print("  (logo após o botão ⬇ Auditoria no HTML)")
        print(PPT_SECTION)
else:
    print("  Botão PPT já existe, pulando")

# ── 2. Botão 'Baixar tudo (CSV)' na faixa verde ───────────────────────────────
OLD_GREEN_BTN = 'onclick="downloadAll()">⬇ Baixar tudo</button>'
NEW_GREEN_BTN = '''onclick="downloadAll()">⬇ Baixar tudo (Excel)</button>
        <button class="btn btn-ghost btn-sm" onclick="downloadAllCSV()">⬇ Baixar tudo (CSV)</button>'''

if 'downloadAllCSV' not in html:
    if OLD_GREEN_BTN in html:
        html = html.replace(OLD_GREEN_BTN, NEW_GREEN_BTN)
        print("✓ Botão 'Baixar tudo (CSV)' adicionado")
        changed = True
    else:
        # tenta variante já parcialmente modificada
        ALT = 'onclick="downloadAll()">⬇ Baixar tudo (Excel)</button>'
        if ALT in html:
            html = html.replace(
                ALT,
                ALT + '\n        <button class="btn btn-ghost btn-sm" onclick="downloadAllCSV()">⬇ Baixar tudo (CSV)</button>'
            )
            print("✓ Botão 'Baixar tudo (CSV)' adicionado (variante)")
            changed = True
        else:
            print("  Âncora do Baixar tudo não encontrada — sem alteração")
else:
    print("  Botão CSV global já existe, pulando")

# ── 3. Função downloadAllCSV no JS ────────────────────────────────────────────
if 'function downloadAllCSV' not in html:
    html = html.replace(
        "function downloadAll() {",
        """function downloadAllCSV() {
  dl('totalpass_csv');
  setTimeout(() => dl('wellhub_csv_zip'),  600);
  setTimeout(() => dl('newvalue_csv'),     1200);
}

function downloadAll() {"""
    )
    print("✓ Função downloadAllCSV adicionada no JS")
    changed = True
else:
    print("  downloadAllCSV já existe no JS, pulando")

# ── Salva ─────────────────────────────────────────────────────────────────────
if changed:
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"\n✓ templates/index.html salvo com sucesso")
else:
    print("\n  Nenhuma alteração necessária")
