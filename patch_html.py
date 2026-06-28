"""
patch_html.py
Adiciona botões de download CSV e PPT na tela de resultados do BenefProcess.
Execute UMA VEZ dentro da pasta do projeto:
    py patch_html.py
"""
import os, re

html_path = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    'templates', 'index.html'
)

with open(html_path, encoding='utf-8') as f:
    html = f.read()

# ── 1. CSS: botão CSV outline ─────────────────────────────────────────────────
css_new = """
    /* ── Botão CSV ── */
    .btn-csv {
      display: block; width: 100%;
      padding: 9px 0; margin-top: 8px;
      background: transparent;
      border: 2px solid var(--c);
      border-radius: 8px;
      color: var(--c);
      font-weight: 700; font-size: 13px;
      cursor: pointer;
      transition: background .15s, color .15s;
    }
    .btn-csv:hover { background: var(--c); color: #fff; }

    /* ── Seção PPT / CSV global ── */
    .extra-downloads {
      background: #fff; border-radius: 12px;
      box-shadow: 0 1px 5px rgba(0,0,0,.08);
      padding: 16px 22px; margin-bottom: 14px;
    }
    .extra-downloads h4 {
      font-weight: 700; font-size: 14px;
      display: flex; align-items: center; gap: 8px;
      margin-bottom: 10px;
    }
    .extra-dl-grid {
      display: flex; gap: 10px; flex-wrap: wrap;
    }
    .extra-dl-btn {
      flex: 1; min-width: 160px;
      padding: 10px 0; border: none; border-radius: 8px;
      font-weight: 600; font-size: 13px; cursor: pointer;
      transition: opacity .15s;
    }
    .extra-dl-btn:hover { opacity: .88; }
"""

anchor_css = "    .reset-btn { text-align: center; margin-top: 4px; }"
if '.btn-csv' not in html:
    html = html.replace(anchor_css, anchor_css + css_new)
    print("✓ CSS adicionado")
else:
    print("  CSS já existe, pulando")

# ── 2. mkCard: adiciona csvKey e botão CSV ────────────────────────────────────
old_mkcard = """function mkCard(label, color, eligible, excluded, reasons, extraHTML, btnLabel, dlKey) {"""
new_mkcard = """function mkCard(label, color, eligible, excluded, reasons, extraHTML, btnLabel, dlKey, csvKey) {"""

old_card_btn = """    <button class="btn-process" style="background:${color}" onclick="dl('${dlKey}')">${btnLabel}</button>`;"""
new_card_btn = """    <button class="btn-process" style="background:${color}" onclick="dl('${dlKey}')">${btnLabel}</button>
    ${csvKey ? `<button class="btn-csv" style="--c:${color}" onclick="dl('${csvKey}')">⬇ Baixar CSV</button>` : ''}`;"""

if 'csvKey' not in html:
    html = html.replace(old_mkcard, new_mkcard)
    html = html.replace(old_card_btn, new_card_btn)
    print("✓ mkCard atualizado com botão CSV")
else:
    print("  mkCard já tem csvKey, pulando")

# ── 3. renderResults: passa csvKey para cada card ─────────────────────────────
old_tp = "grid.appendChild(mkCard('TotalPass', '#5B4FE9', data.tp.eligible, data.tp.excluded,\n    data.tp.reasons, '', '⬇ Baixar TotalPass', 'totalpass'));"
new_tp = "grid.appendChild(mkCard('TotalPass', '#5B4FE9', data.tp.eligible, data.tp.excluded,\n    data.tp.reasons, '', '⬇ Baixar TotalPass (Excel)', 'totalpass', 'totalpass_csv'));"

old_wh = """grid.appendChild(mkCard('Wellhub', '#E91E8C', wh.eligible, wh.excluded,\n    wh.reasons, compHTML,\n    `⬇ Baixar ZIP (${companies.length} arquivo${companies.length!==1?'s':''})`, 'wellhub'));"""
new_wh = """grid.appendChild(mkCard('Wellhub', '#E91E8C', wh.eligible, wh.excluded,\n    wh.reasons, compHTML,\n    `⬇ Baixar ZIP Excel (${companies.length} arquivo${companies.length!==1?'s':''})`, 'wellhub', 'wellhub_csv_zip'));"""

old_nv = "grid.appendChild(mkCard('New Value', '#F59E0B', data.nv.eligible, data.nv.excluded,\n    data.nv.reasons, '', '⬇ Baixar New Value', 'newvalue'));"
new_nv = "grid.appendChild(mkCard('New Value', '#F59E0B', data.nv.eligible, data.nv.excluded,\n    data.nv.reasons, '', '⬇ Baixar New Value (Excel)', 'newvalue', 'newvalue_csv'));"

if 'totalpass_csv' not in html:
    html = html.replace(old_tp, new_tp)
    html = html.replace(old_wh, new_wh)
    html = html.replace(old_nv, new_nv)
    print("✓ renderResults: csvKey adicionado em TotalPass, Wellhub e New Value")
else:
    print("  renderResults já tem csvKey, pulando")

# ── 4. Card PPT na tela de resultados ─────────────────────────────────────────
old_audit = """    <div class="audit-card">
      <div>
        <h4>📋 Planilha de Auditoria</h4>
        <p id="auditSub"></p>
      </div>
      <button class="btn btn-ghost btn-sm" onclick="dl('auditoria')">⬇ Auditoria</button>
    </div>"""

new_audit = """    <div class="audit-card">
      <div>
        <h4>📋 Planilha de Auditoria</h4>
        <p id="auditSub"></p>
      </div>
      <button class="btn btn-ghost btn-sm" onclick="dl('auditoria')">⬇ Auditoria</button>
    </div>

    <div class="extra-downloads">
      <h4>📥 Downloads adicionais</h4>
      <div class="extra-dl-grid">
        <button class="extra-dl-btn" style="background:#0F2A56;color:#fff"
                onclick="dl('totalpass_csv')">⬇ TotalPass · CSV</button>
        <button class="extra-dl-btn" style="background:#E91E8C;color:#fff"
                onclick="dl('wellhub_csv_zip')">⬇ Wellhub · ZIP CSV</button>
        <button class="extra-dl-btn" style="background:#F59E0B;color:#fff"
                onclick="dl('newvalue_csv')">⬇ New Value · CSV</button>
        <button class="extra-dl-btn" style="background:#1A4A8A;color:#fff"
                onclick="dl('auditppt')">📊 Apresentação PPT</button>
      </div>
    </div>"""

if 'extra-downloads' not in html:
    html = html.replace(old_audit, new_audit)
    print("✓ Seção 'Downloads adicionais' com CSVs e PPT adicionada")
else:
    print("  Seção extra-downloads já existe, pulando")

# ── 5. downloadAll: mantém Excel + adiciona CSV ───────────────────────────────
old_dl_all = """function downloadAll() {
  dl('totalpass');
  setTimeout(() => dl('wellhub'),  600);
  setTimeout(() => dl('newvalue'), 1200);
}"""

new_dl_all = """function downloadAll() {
  dl('totalpass');
  setTimeout(() => dl('wellhub'),   600);
  setTimeout(() => dl('newvalue'), 1200);
}

function downloadAllCSV() {
  dl('totalpass_csv');
  setTimeout(() => dl('wellhub_csv_zip'),  600);
  setTimeout(() => dl('newvalue_csv'),     1200);
}"""

if 'downloadAllCSV' not in html:
    html = html.replace(old_dl_all, new_dl_all)
    print("✓ downloadAllCSV adicionado")
else:
    print("  downloadAllCSV já existe, pulando")

# ── 6. Botão "Baixar CSVs" na faixa verde ─────────────────────────────────────
old_green = """    <div id="downloadAllWrap" class="alert-green" style="margin-bottom:18px">
      <span style="font-weight:500">Baixar os três arquivos tratados de uma vez:</span>
      <button class="btn btn-teal btn-sm" onclick="downloadAll()">⬇ Baixar tudo</button>
    </div>"""

new_green = """    <div id="downloadAllWrap" class="alert-green" style="margin-bottom:18px">
      <span style="font-weight:500">Baixar os três arquivos tratados de uma vez:</span>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <button class="btn btn-teal btn-sm" onclick="downloadAll()">⬇ Baixar tudo (Excel)</button>
        <button class="btn btn-ghost btn-sm" onclick="downloadAllCSV()">⬇ Baixar tudo (CSV)</button>
      </div>
    </div>"""

if 'downloadAllCSV()' not in html:
    html = html.replace(old_green, new_green)
    print("✓ Botão 'Baixar tudo CSV' adicionado na faixa verde")
else:
    print("  Botão CSV já existe na faixa verde, pulando")

# ── Salva ─────────────────────────────────────────────────────────────────────
with open(html_path, 'w', encoding='utf-8') as f:
    f.write(html)

print(f"\n✓ templates/index.html atualizado")
print("\nO que aparece agora na tela de resultados:")
print("  - Cada card tem botão Excel + botão CSV abaixo")
print("  - Faixa verde: 'Baixar tudo (Excel)' e 'Baixar tudo (CSV)'")
print("  - Seção 'Downloads adicionais': botões CSV individuais + botão PPT")
