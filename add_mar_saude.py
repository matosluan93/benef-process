"""
add_mar_saude.py
Adiciona o fluxo completo do Mar Saúde ao BenefProcess.
Execute UMA VEZ dentro da pasta do projeto:
    py add_mar_saude.py

O que é adicionado:
  - Segunda zona de upload na tela inicial (Mar Saúde)
  - Modal com mini-fluxo independente (aba → colunas → resultado)
  - Backend: 5 novas rotas isoladas, sem tocar nas existentes
  - Regra: apenas exclusão de PJ (sem checagem de domínio de e-mail)
  - Saída: Excel + CSV com as 20 colunas exatas do layout Mar Saúde
"""
import os, ast

# ─── Constantes ───────────────────────────────────────────────────────────────
BACKEND_CODE = '''
# ══════════════════════════════════════════════════════════════════════════════
# MAR SAÚDE — fluxo independente, sem alterar nenhuma regra existente
# ══════════════════════════════════════════════════════════════════════════════

MAR_SAUDE_COLS = [
    'Empresa', 'Nome Empresa', 'CNPJ Empresa', 'Célula',
    'Desc. Atividade (Serviço)', 'Nome', 'Nome Social', 'Nome Cargo',
    'Data de Admissão', 'Sexo', 'Data de Nascimento', 'Logradouro',
    'Endereço', 'Complemento', 'Bairro', 'Cidade', 'Complemento CEP',
    'E-mail Funcional', 'Número de CPF', 'Desc. Unidade Adm. (Cliente)',
]

MAR_SAUDE_COL_ALIASES = {
    'empresa':       ['empresa'],
    'nome_empresa':  ['nome empresa'],
    'cnpj':         ['cnpj empresa', 'cnpj da empresa', 'cnpj'],
    'celula':       ['célula', 'celula', 'centro de custo'],
    'atividade':    ['desc. atividade (serviço)', 'desc. atividade', 'atividade', 'servico'],
    'nome':         ['nome profissional', 'nome do profissional', 'nome'],
    'nome_social':  ['nome social'],
    'cargo':        ['nome cargo', 'nome função', 'nome funcao', 'cargo', 'função'],
    'data_admissao':['admissão', 'admissao', 'data de admissão', 'data de admissao', 'dt admissão'],
    'sexo':         ['sexo', 'gênero', 'genero'],
    'dt_nascimento':['data de nascimento', 'nascimento', 'dt nascimento', 'data nascimento'],
    'logradouro':   ['logradouro', 'tipo logradouro'],
    'endereco':     ['endereço', 'endereco', 'rua', 'end'],
    'complemento':  ['complemento'],
    'bairro':       ['bairro'],
    'cidade':       ['cidade', 'município', 'municipio'],
    'cep':          ['complemento cep', 'cep', 'cod postal'],
    'email':        ['e-mail funcional', 'email funcional', 'e-mail', 'email'],
    'cpf':          ['número de cpf', 'cpf do profissional', 'cpf', 'nr cpf'],
    'unidade_adm':  ['desc. unidade adm. (cliente)', 'desc. unidade adm', 'unidade adm'],
    'tipo_vinculo':   ['desc. tipo de vínculo', 'desc. tipo de vinculo',
                       'desc. vínculo', 'desc. vinculo', 'vínculo', 'vinculo', 'regime'],
    'nome_sindicato': ['nome sindicato', 'sindicato', 'nome_sindicato',
                       'desc. sindicato', 'desc sindicato'],
}

def _auto_map_ms(headers):
    """Mapeamento automático usando aliases do Mar Saúde."""
    mapping = {}
    headers_lower = [(h, norm(h)) for h in headers]
    for key, aliases in MAR_SAUDE_COL_ALIASES.items():
        for alias in aliases:
            for h, l in headers_lower:
                if l == alias or l == alias.replace(' ', '_'):
                    mapping[key] = h; break
                if len(alias) >= 5 and alias in l:
                    if key not in mapping: mapping[key] = h
            if key in mapping: break
    return mapping

def _to_ms_row(row_dict, col_map):
    """Converte linha para as 20 colunas do Mar Saúde."""
    get = lambda k: clean_val(row_dict.get(col_map.get(k, '___ms___'), ''))
    def _date(k):
        col = col_map.get(k, '')
        val = row_dict.get(col, '') if col else ''
        try:
            return format_date_val(parse_mixed_dates(val)) if val else ''
        except Exception:
            return clean_val(val)
    return {
        'Empresa':                      get('empresa'),
        'Nome Empresa':                 get('nome_empresa'),
        'CNPJ Empresa':                 get('cnpj'),
        'Célula':                       get('celula'),
        'Desc. Atividade (Serviço)':    get('atividade'),
        'Nome':                         get('nome'),
        'Nome Social':                  get('nome_social'),
        'Nome Cargo':                   get('cargo'),
        'Data de Admissão':             _date('data_admissao'),
        'Sexo':                         get('sexo'),
        'Data de Nascimento':           _date('dt_nascimento'),
        'Logradouro':                   get('logradouro'),
        'Endereço':                     get('endereco'),
        'Complemento':                  get('complemento'),
        'Bairro':                       get('bairro'),
        'Cidade':                       get('cidade'),
        'Complemento CEP':              get('cep'),
        'E-mail Funcional':             get('email'),
        'Número de CPF':                get('cpf'),
        'Desc. Unidade Adm. (Cliente)': get('unidade_adm'),
    }

def _write_ms_csv(rows, filepath):
    import csv
    with open(filepath, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=MAR_SAUDE_COLS, delimiter=';')
        w.writeheader()
        w.writerows(rows)

@app.route('/api/upload-mar-saude', methods=['POST'])
def api_upload_mar_saude():
    if 'file' not in request.files:
        return jsonify({'error': 'Nenhum arquivo enviado.'}), 400
    f = request.files['file']
    if not f.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Envie um arquivo .xlsx ou .xls.'}), 400
    sid         = str(uuid.uuid4())
    session_dir = os.path.join(TEMP_DIR, sid)
    os.makedirs(session_dir, exist_ok=True)
    upload_path = os.path.join(session_dir, 'ms_source.xlsx')
    f.save(upload_path)
    try:
        xl = pd.ExcelFile(upload_path)
        return jsonify({'session_id': sid, 'sheet_names': xl.sheet_names, 'filename': f.filename})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/get-columns-mar-saude', methods=['POST'])
def api_get_columns_mar_saude():
    data        = request.json
    sid         = data.get('session_id')
    sheet_name  = data.get('sheet_name')
    upload_path = os.path.join(TEMP_DIR, sid, 'ms_source.xlsx')
    if not os.path.exists(upload_path):
        return jsonify({'error': 'Sessão expirada.'}), 400
    try:
        df = pd.ExcelFile(upload_path).parse(sheet_name, dtype=str, na_filter=False)
        df.columns = [str(c).strip().lstrip('\\ufeff') for c in df.columns]
        cols = list(df.columns)
        return jsonify({'headers': cols, 'auto_map': _auto_map_ms(cols), 'row_count': len(df)})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/process-mar-saude', methods=['POST'])
def api_process_mar_saude():
    data        = request.json
    sid         = data.get('session_id')
    sheet_name  = data.get('sheet_name')
    col_map_usr = data.get('col_map', {})
    upload_path = os.path.join(TEMP_DIR, sid, 'ms_source.xlsx')
    if not os.path.exists(upload_path):
        return jsonify({'error': 'Sessão expirada.'}), 400
    try:
        df = pd.ExcelFile(upload_path).parse(sheet_name, dtype=str, na_filter=False)
        df.columns = [str(c).strip().lstrip('\\ufeff') for c in df.columns]
        mask = df.apply(lambda row: row.str.strip().ne('').any(), axis=1)
        df   = df[mask].reset_index(drop=True)

        auto = _auto_map_ms(list(df.columns))
        # col_map_usr sobrescreve o auto
        final_map = {**auto, **{k: v for k, v in col_map_usr.items() if v}}

        # Converte datas
        for dk in ('data_admissao', 'dt_nascimento'):
            col = final_map.get(dk)
            if col and col in df.columns:
                df[col] = df[col].apply(parse_mixed_dates)

        sindicato_col = final_map.get('nome_sindicato')
        vinculo_col   = final_map.get('tipo_vinculo')
        eligible, excluded = [], []
        for _, row in df.iterrows():
            is_pj = False

            # Critério primário: Nome Sindicato = "Somente Pj"
            if sindicato_col and sindicato_col in row.index:
                sind_val = norm(clean_val(row[sindicato_col]))
                if 'somente pj' in sind_val:
                    is_pj = True

            # Critério fallback: tipo_vinculo com keywords PJ
            if not is_pj and vinculo_col and vinculo_col in row.index:
                vinculo_norm = norm(clean_val(row[vinculo_col]))
                if vinculo_norm and any(k in vinculo_norm for k in PJ_KW):
                    is_pj = True

            if is_pj:
                excluded.append({'_motivo': 'Pessoa Jurídica (PJ)', **row.to_dict()})
            else:
                eligible.append(row.to_dict())

        ms_rows  = [_to_ms_row(r, final_map) for r in eligible]
        ex_rows  = [_to_ms_row(r, final_map) for r in excluded]

        session_dir = os.path.join(TEMP_DIR, sid)
        write_styled_excel(ms_rows, os.path.join(session_dir, 'ms_output.xlsx'),
                           'Mar Saúde', MAR_SAUDE_COLS)
        _write_ms_csv(ms_rows, os.path.join(session_dir, 'ms_output.csv'))
        write_styled_excel(ex_rows, os.path.join(session_dir, 'ms_audit.xlsx'),
                           'Excluídos', MAR_SAUDE_COLS)

        return jsonify({
            'total':    len(df),
            'eligible': len(ms_rows),
            'excluded': len(ex_rows),
        })
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'detail': traceback.format_exc()}), 500

@app.route('/api/download/<session_id>/ms_excel')
def download_ms_excel(session_id):
    path = os.path.join(TEMP_DIR, session_id, 'ms_output.xlsx')
    if os.path.exists(path):
        return send_file(path, as_attachment=True, download_name='MarSaude_Elegiveis.xlsx')
    return 'Arquivo não encontrado.', 404

@app.route('/api/download/<session_id>/ms_csv')
def download_ms_csv(session_id):
    path = os.path.join(TEMP_DIR, session_id, 'ms_output.csv')
    if os.path.exists(path):
        return send_file(path, as_attachment=True, download_name='MarSaude_Elegiveis.csv')
    return 'Arquivo não encontrado.', 404

@app.route('/api/download/<session_id>/ms_audit')
def download_ms_audit(session_id):
    path = os.path.join(TEMP_DIR, session_id, 'ms_audit.xlsx')
    if os.path.exists(path):
        return send_file(path, as_attachment=True, download_name='MarSaude_Excluidos.xlsx')
    return 'Arquivo não encontrado.', 404
# ══════════════════════════════════════════════════════════════════════════════
'''

MODAL_HTML = '''
<!-- ══ MAR SAÚDE — Upload zone + Modal independente ══════════════════════ -->
<div style="max-width:860px;margin:0 auto;padding:0 18px 28px">
  <div style="background:#fff;border-radius:12px;box-shadow:0 1px 5px rgba(0,0,0,.08);
              padding:24px;border-top:3px solid #0891B2">
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:6px">
      <div style="width:32px;height:32px;background:#0891B2;border-radius:8px;
                  display:flex;align-items:center;justify-content:center;font-size:16px">🏥</div>
      <div>
        <div style="font-weight:800;font-size:15px;color:#0F2A56">Mar Saúde</div>
        <div style="font-size:11px;color:#64748B">Somente Stefanini · Remove apenas PJ · A cada 15 dias</div>
      </div>
    </div>
    <div id="ms-dropzone"
         style="border:2px dashed #BAE6FD;border-radius:10px;padding:28px 20px;
                text-align:center;cursor:pointer;background:#F0F9FF;transition:all .18s;margin-top:12px"
         onclick="document.getElementById('ms-file-input').click()"
         ondragover="event.preventDefault();this.style.borderColor='#0891B2'"
         ondragleave="this.style.borderColor='#BAE6FD'"
         ondrop="event.preventDefault();this.style.borderColor='#BAE6FD';msUploadFile(event.dataTransfer.files[0])">
      <div style="font-size:32px;margin-bottom:6px">📋</div>
      <div id="ms-drop-label" style="font-weight:700;font-size:14px;color:#0F2A56">
        Clique ou arraste a base da Stefanini aqui
      </div>
      <div style="font-size:12px;color:#64748B;margin-top:4px">Formatos aceitos: .xlsx · .xls</div>
    </div>
    <input type="file" id="ms-file-input" accept=".xlsx,.xls" style="display:none"
           onchange="if(this.files[0]) msUploadFile(this.files[0])"/>
  </div>
</div>

<!-- Modal Mar Saúde -->
<div id="ms-modal" style="display:none;position:fixed;inset:0;z-index:1000;
     background:rgba(10,30,61,.6);align-items:center;justify-content:center">
  <div style="background:#fff;border-radius:16px;width:min(680px,96vw);
              max-height:90vh;overflow-y:auto;padding:28px;position:relative">
    <!-- Fechar -->
    <button onclick="msCloseModal()"
            style="position:absolute;top:16px;right:16px;border:none;background:#F1F5F9;
                   border-radius:50%;width:32px;height:32px;font-size:16px;cursor:pointer">✕</button>

    <div style="display:flex;align-items:center;gap:10px;margin-bottom:20px">
      <div style="width:36px;height:36px;background:#0891B2;border-radius:10px;
                  display:flex;align-items:center;justify-content:center;font-size:18px">🏥</div>
      <div>
        <div style="font-weight:800;font-size:17px;color:#0F2A56">Mar Saúde</div>
        <div id="ms-modal-filename" style="font-size:12px;color:#64748B"></div>
      </div>
    </div>

    <!-- Passo A: Selecionar aba -->
    <div id="ms-step-sheet">
      <div style="font-weight:700;font-size:14px;color:#0F2A56;margin-bottom:10px">
        Selecionar aba da planilha
      </div>
      <select id="ms-sheet-select" style="width:100%;padding:10px;border-radius:8px;
              border:1px solid #CBD5E1;font-size:13px;margin-bottom:16px"></select>
      <button onclick="msConfirmSheet()"
              style="width:100%;padding:11px;background:#0891B2;color:#fff;border:none;
                     border-radius:8px;font-weight:700;font-size:14px;cursor:pointer">
        Confirmar aba →
      </button>
    </div>

    <!-- Passo B: Mapeamento (só vínculo) -->
    <div id="ms-step-cols" style="display:none">
      <div style="font-weight:700;font-size:14px;color:#0F2A56;margin-bottom:4px">
        Mapeamento de colunas
      </div>
      <div style="font-size:12px;color:#64748B;margin-bottom:14px">
        Campos detectados automaticamente. Ajuste se necessário.
      </div>
      <div id="ms-col-rows" style="max-height:300px;overflow-y:auto"></div>
      <div style="margin-top:16px;display:flex;gap:10px">
        <button onclick="document.getElementById('ms-step-cols').style.display='none';
                         document.getElementById('ms-step-sheet').style.display='block'"
                style="flex:1;padding:10px;background:#F1F5F9;border:none;border-radius:8px;
                       font-weight:600;font-size:13px;cursor:pointer">← Voltar</button>
        <button onclick="msProcess()"
                style="flex:2;padding:10px;background:#0F2A56;color:#fff;border:none;
                       border-radius:8px;font-weight:700;font-size:13px;cursor:pointer">
                🚀 Processar Mar Saúde</button>
      </div>
    </div>

    <!-- Passo C: Resultado -->
    <div id="ms-step-result" style="display:none">
      <div style="background:#F0FDF9;border:1px solid #A7F3D0;border-radius:10px;
                  padding:16px;margin-bottom:16px;text-align:center">
        <div style="font-size:28px;margin-bottom:4px">✅</div>
        <div style="font-weight:800;font-size:16px;color:#065F46">Processamento concluído!</div>
        <div id="ms-result-sub" style="font-size:13px;color:#047857;margin-top:4px"></div>
      </div>
      <div style="display:flex;flex-direction:column;gap:10px">
        <button onclick="msDl('ms_excel')"
                style="padding:12px;background:#0F2A56;color:#fff;border:none;border-radius:8px;
                       font-weight:700;font-size:14px;cursor:pointer">
                ⬇ Baixar Excel — Elegíveis (Mar Saúde)</button>
        <button onclick="msDl('ms_csv')"
                style="padding:12px;background:transparent;border:2px solid #0891B2;
                       color:#0891B2;border-radius:8px;font-weight:700;font-size:14px;cursor:pointer">
                ⬇ Baixar CSV — Elegíveis (Mar Saúde)</button>
        <button onclick="msDl('ms_audit')"
                style="padding:12px;background:#F1F5F9;color:#475569;border:none;
                       border-radius:8px;font-weight:600;font-size:13px;cursor:pointer">
                ⬇ Planilha de Excluídos (PJ)</button>
      </div>
      <button onclick="msCloseModal()"
              style="width:100%;margin-top:12px;padding:10px;background:transparent;
                     border:1px solid #CBD5E1;border-radius:8px;font-size:13px;
                     color:#64748B;cursor:pointer">Fechar</button>
    </div>

    <!-- Loading -->
    <div id="ms-loading" style="display:none;text-align:center;padding:30px">
      <div style="width:38px;height:38px;border:4px solid #E2E8F0;border-top-color:#0891B2;
                  border-radius:50%;animation:spin .75s linear infinite;margin:0 auto 12px"></div>
      <div style="color:#64748B;font-size:14px" id="ms-loading-text">Processando...</div>
    </div>
  </div>
</div>

<script>
// ─── Estado Mar Saúde ─────────────────────────────────────────────────────────
const MS = { sessionId: null, filename: '', sheetNames: [], selectedSheet: '', colMap: {} };

const MS_FIELDS = [
  { key: 'nome_sindicato', label: 'Nome Sindicato (identifica PJ = "Somente Pj")', req: false },
  { key: 'tipo_vinculo', label: 'Tipo de Vínculo (fallback PJ)', req: false },
  { key: 'nome',         label: 'Nome do profissional',              req: false },
  { key: 'cpf',          label: 'Número de CPF',                     req: false },
  { key: 'email',        label: 'E-mail Funcional',                  req: false },
  { key: 'data_admissao',label: 'Data de Admissão',                  req: false },
  { key: 'cnpj',         label: 'CNPJ Empresa',                      req: false },
  { key: 'nome_empresa', label: 'Nome Empresa',                      req: false },
  { key: 'cargo',        label: 'Nome Cargo',                        req: false },
  { key: 'celula',       label: 'Célula',                            req: false },
  { key: 'atividade',    label: 'Desc. Atividade (Serviço)',          req: false },
  { key: 'sexo',         label: 'Sexo',                              req: false },
  { key: 'dt_nascimento',label: 'Data de Nascimento',                req: false },
  { key: 'cidade',       label: 'Cidade',                            req: false },
  { key: 'unidade_adm',  label: 'Desc. Unidade Adm. (Cliente)',      req: false },
];

function msShowLoading(txt) {
  document.getElementById('ms-loading-text').textContent = txt || 'Processando...';
  ['ms-step-sheet','ms-step-cols','ms-step-result'].forEach(id =>
    document.getElementById(id).style.display = 'none');
  document.getElementById('ms-loading').style.display = 'block';
}
function msHideLoading() { document.getElementById('ms-loading').style.display = 'none'; }

function msOpenModal() {
  const m = document.getElementById('ms-modal');
  m.style.display = 'flex';
  document.getElementById('ms-step-sheet').style.display = 'block';
  document.getElementById('ms-step-cols').style.display = 'none';
  document.getElementById('ms-step-result').style.display = 'none';
}
function msCloseModal() {
  document.getElementById('ms-modal').style.display = 'none';
}

async function msUploadFile(file) {
  if (!file) return;
  if (!file.name.match(/\\.(xlsx|xls)$/i)) {
    alert('Envie um arquivo .xlsx ou .xls.');
    return;
  }
  document.getElementById('ms-drop-label').textContent = file.name;
  msShowLoading('Lendo arquivo...');
  msOpenModal();
  const fd = new FormData();
  fd.append('file', file);
  try {
    const res  = await fetch('/api/upload-mar-saude', { method: 'POST', body: fd });
    const data = await res.json();
    if (data.error) { msHideLoading(); alert('Erro: ' + data.error); return; }
    MS.sessionId  = data.session_id;
    MS.filename   = data.filename;
    MS.sheetNames = data.sheet_names;
    document.getElementById('ms-modal-filename').textContent = data.filename;
    const sel = document.getElementById('ms-sheet-select');
    sel.innerHTML = MS.sheetNames.map(s => `<option value="${s}">${s}</option>`).join('');
    msHideLoading();
    document.getElementById('ms-step-sheet').style.display = 'block';
  } catch(e) { msHideLoading(); alert('Erro: ' + e.message); }
}

async function msConfirmSheet() {
  MS.selectedSheet = document.getElementById('ms-sheet-select').value;
  msShowLoading('Detectando colunas...');
  try {
    const res  = await fetch('/api/get-columns-mar-saude', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({ session_id: MS.sessionId, sheet_name: MS.selectedSheet })
    });
    const data = await res.json();
    if (data.error) { msHideLoading(); alert('Erro: ' + data.error); return; }
    MS.colMap = data.auto_map;
    const headers = data.headers;
    // Renderiza campos de mapeamento
    const container = document.getElementById('ms-col-rows');
    container.innerHTML = '';
    MS_FIELDS.forEach(f => {
      const div = document.createElement('div');
      div.style.cssText = 'display:flex;align-items:center;gap:10px;padding:8px 0;border-bottom:1px solid #F1F5F9';
      const lbl = document.createElement('div');
      lbl.style.cssText = 'width:220px;font-size:12px;font-weight:500;flex-shrink:0;color:#374151';
      lbl.textContent = f.label;
      const sel = document.createElement('select');
      sel.style.cssText = 'flex:1;padding:7px;border-radius:6px;border:1px solid #CBD5E1;font-size:12px';
      sel.innerHTML = '<option value="">— Não mapeado —</option>' +
        headers.map(h => `<option value="${h}" ${MS.colMap[f.key]===h?'selected':''}>${h}</option>`).join('');
      sel.addEventListener('change', () => { MS.colMap[f.key] = sel.value; });
      div.append(lbl, sel);
      container.appendChild(div);
    });
    msHideLoading();
    document.getElementById('ms-step-cols').style.display = 'block';
  } catch(e) { msHideLoading(); alert('Erro: ' + e.message); }
}

async function msProcess() {
  msShowLoading('Aplicando regras de exclusão...');
  try {
    const res  = await fetch('/api/process-mar-saude', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({
        session_id:  MS.sessionId,
        sheet_name:  MS.selectedSheet,
        col_map:     MS.colMap,
      })
    });
    const data = await res.json();
    if (data.error) { msHideLoading(); alert('Erro: ' + data.error); return; }
    msHideLoading();
    document.getElementById('ms-result-sub').textContent =
      `${data.total} registros · ${data.eligible} elegíveis · ${data.excluded} PJs excluídos`;
    document.getElementById('ms-step-result').style.display = 'block';
  } catch(e) { msHideLoading(); alert('Erro: ' + e.message); }
}

function msDl(type) {
  if (!MS.sessionId) return;
  window.location.href = `/api/download/${MS.sessionId}/${type}`;
}
</script>
'''

# ─── Aplica no app.py ─────────────────────────────────────────────────────────
app_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'app.py')
with open(app_path, encoding='utf-8') as f:
    code = f.read()

if 'api_upload_mar_saude' not in code:
    code = code.replace("if __name__ == '__main__':", BACKEND_CODE + "if __name__ == '__main__':")
    print("✓ Rotas Mar Saúde adicionadas ao app.py")
else:
    print("  Rotas Mar Saúde já existem, pulando")

try:
    ast.parse(code)
    print("✓ Sintaxe Python válida")
except SyntaxError as e:
    print(f"✗ Erro de sintaxe: {e}"); exit(1)

with open(app_path, 'w', encoding='utf-8') as f:
    f.write(code)

# ─── Aplica no index.html ─────────────────────────────────────────────────────
html_path = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), 'templates', 'index.html')
with open(html_path, encoding='utf-8') as f:
    html = f.read()

if 'ms-modal' not in html:
    # Insere antes do </body>
    html = html.replace('</body>', MODAL_HTML + '\n</body>')
    # Adiciona título de seção entre o step1 e o conteúdo do step1
    html = html.replace(
        '<div id="step1">',
        '<div id="step1">'
        '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px">'
        '<div><h2 class="page-title">Envie a planilha de profissionais</h2>'
        '<p style="font-size:12px;color:#64748B;margin-bottom:12px">'
        'TotalPass · Wellhub · New Value — base com as 3 abas</p></div></div>'
    )
    html = html.replace(
        '<h2 class="page-title">Envie a planilha de profissionais</h2>\n    <p class="page-sub">',
        '<!-- titulo movido acima -->\n    <p class="page-sub" style="display:none">'
    )
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print("✓ Zona de upload Mar Saúde adicionada ao index.html")
else:
    print("  Modal Mar Saúde já existe no HTML, pulando")

print("\n✓ Instalação concluída.")
print("  Rode: py app.py")
print("\n  Na tela inicial aparecerá:")
print("  1. Upload normal (TotalPass, Wellhub, New Value)")
print("  2. Upload Mar Saúde com modal independente")
