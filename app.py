import os
import uuid
import re
import json
import zipfile
from datetime import date
from ppt_generator import generate as generate_ppt
from flask import Flask, request, jsonify, render_template, send_file
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

TEMP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp')
os.makedirs(TEMP_DIR, exist_ok=True)

# ─── Configuração de colunas ──────────────────────────────────────────────────
# ORDEM IMPORTA: auto_map_cols para no primeiro alias que bater.
# Aliases mais específicos (descritivos) vêm antes de aliases curtos/ambíguos.

COL_ALIASES = {
    'empresa': [
        'nome da empresa', 'nome empresa',
        'empresa', 'company', 'razao social', 'unidade',
    ],
    'cnpj': [
        'cnpj da empresa', 'cnpj empresa', 'cnpj filial',
        'cnpj', 'cnpj/cpf',
    ],
    'nome': [
        'nome profissional', 'nome do profissional',
        'nome', 'colaborador', 'funcionario',
    ],
    'matricula': [
        'matrícula', 'matricula',
        'matrícula do profissional', 'matricula do profissional',
        'chapa', 'registro',
    ],
    'email': [
        # 'email funcional' (sem hífen) cobre IHM cujo cabeçalho é 'EMAIL FUNCIONAL'
        'email funcional', 'e-mail funcional',
        'email', 'e-mail',
        'email profissional', 'e-mail profissional',
    ],
    'tipo_vinculo': [
        # FIX: Stefanini tem 'Desc. Tipo de Vínculo' = 'CONTRATADO'/'ESTAGIÁRIO'.
        # A coluna 'Vínculo' contém apenas o código de letra ('C'), que nunca
        # bateria nas keywords de ESTAGIO_KW / PJ_KW.
        'desc. tipo de vínculo', 'desc. tipo de vinculo',
        'desc. vínculo', 'desc. vinculo',
        'vínculo', 'vinculo',
        'tipo de vínculo', 'tipo de vinculo',
        'regime', 'modalidade',
    ],
    'data_admissao': [
        'admissão', 'admissao',
        'data de admissão', 'data de admissao',
        'dt admissão', 'dt. admissão', 'dt admissao', 'dt. admissao',
        'data admissão', 'data admissao',
        'data_admissao', 'data_admissão',
    ],
    'desconto_folha': [
        'desconto em folha', 'desconto folha',
        'aceite wellhub', 'wellhub folha',
    ],
    'cargo': [
        # FIX: 'nome cargo' e 'nome função' primeiro para evitar que 'função'
        # (alias curto) bata na coluna de CÓDIGO NUMÉRICO 'Função' da Stefanini
        # ('8449') em vez do campo descritivo 'Nome Cargo' / 'Nome Função'.
        'nome cargo', 'nome função', 'nome funcao',
        'função do profissional', 'funcao do profissional',
        'função', 'funcao', 'cargo',
    ],
    'cpf': [
        'cpf do profissional', 'número de cpf', 'numero de cpf',
        'cpf', 'nr cpf',
    ],
}

# ─── Regex ────────────────────────────────────────────────────────────────────

EMAIL_RE  = re.compile(r'^[^\s@]+@[^\s@]+\.[^\s@]+$')
ISO_DATE  = re.compile(r'^\d{4}[-/]\d{1,2}[-/]\d{1,2}')   # YYYY-MM-DD ou YYYY/MM/DD
CPF_CLEAN = re.compile(r'\D')

# ─── Funções de normalização ──────────────────────────────────────────────────

def norm(v):
    """Lowercase + strip + remove BOM invisível (\\ufeff) de cabeçalhos ERP."""
    return str(v or '').strip().lstrip('\ufeff').lower()

def clean_val(v):
    """
    Retorna string limpa. Trata isna() com try/except para não quebrar
    em tipos ambíguos (arrays, objetos customizados) que surgem com na_filter=False.
    """
    try:
        if pd.isna(v):
            return ''
    except (TypeError, ValueError):
        pass
    v_str = str(v).replace('\xa0', ' ').strip()
    if v_str.lower() in ('nan', 'none', 'nat', 'null', '-', ''):
        return ''
    return v_str

def format_cpf(v):
    """
    FIX: CPF da Topaz chega sem formatação ('61213829020').
    Formata para '612.138.290-20' quando há exatamente 11 dígitos.
    """
    raw = CPF_CLEAN.sub('', str(v))
    if len(raw) == 11:
        return f'{raw[:3]}.{raw[3:6]}.{raw[6:9]}-{raw[9:]}'
    return clean_val(v)

# ─── Parsing de datas ─────────────────────────────────────────────────────────

def parse_mixed_dates(s):
    try:
        if pd.isna(s) or str(s).strip() == '': return pd.NaT
    except Exception: pass
    val = str(s).strip()
    if val.replace('.', '', 1).isdigit():
        try: return pd.to_datetime(float(val), origin='1899-12-30', unit='D')
        except Exception: pass
    if re.match(r'^\d{4}[-/]\d', val):
        return pd.to_datetime(val, dayfirst=False, errors='coerce')
    return pd.to_datetime(val, dayfirst=True, errors='coerce')
def format_date_val(v):
    """
    Converte qualquer objeto data/timestamp/string para 'dd/mm/yyyy'.

    FIX: usa pd.Timestamp(v) como conversor universal para cobrir pd.Timestamp,
    numpy.datetime64 e datetime.datetime — todos os tipos que df.iterrows() pode
    retornar para colunas datetime, dependendo da versão do pandas.
    """
    try:
        if pd.isna(v):
            return ''
    except (TypeError, ValueError):
        pass

    # Converte qualquer tipo date-like para Timestamp de forma segura
    try:
        ts = pd.Timestamp(v)
        if pd.isna(ts):
            return ''
        return ts.strftime('%d/%m/%Y')
    except Exception:
        pass

    # Fallback para strings
    v_str = str(v).strip()
    if v_str.lower() in ('nat', 'nan', 'none', ''):
        return ''
    try:
        if ISO_DATE.match(v_str):
            return pd.to_datetime(v_str, dayfirst=False, errors='coerce').strftime('%d/%m/%Y')
        return pd.to_datetime(v_str, dayfirst=True, errors='coerce').strftime('%d/%m/%Y')
    except Exception:
        return v_str

# ─── Regras de negócio ────────────────────────────────────────────────────────

ESTAGIO_KW   = ['estagiário', 'estagiario', 'estágio', 'estagio', 'estag']
APRENDIZ_KW  = ['aprendiz', 'menor aprendiz', 'jovem aprendiz']
TERCEIRO_KW  = ['terceiro', 'terceirizado', 'prestador']
PJ_KW        = ['pj', 'pessoa jurídica', 'pessoa juridica', 'p.j.', 'autonomo', 'autônomo']
DESCONTO_POS = {'sim', 's', 'yes', 'y', 'ativo', 'habilitado', 'enabled', 'true', '1', 'ativado', 'ok'}

def is_valid_email(v):
    return bool(EMAIL_RE.match(norm(v)))

def is_future_date(v):
    if v is None or (hasattr(v, '__class__') and v.__class__.__name__ == 'NaTType'):
        return False
    if isinstance(v, pd.Timestamp):
        return not pd.isna(v) and v.date() > date.today()
    try:
        ts = pd.Timestamp(v)
        return not pd.isna(ts) and ts.date() > date.today()
    except Exception:
        return False

def detect_vinculo_label(v):
    v = norm(v)
    if any(k in v for k in ESTAGIO_KW):  return 'Estagiário'
    if any(k in v for k in APRENDIZ_KW): return 'Aprendiz'
    if any(k in v for k in TERCEIRO_KW): return 'Terceiro/Terceirizado'
    if any(k in v for k in PJ_KW):       return 'Pessoa Jurídica (PJ)'
    return None

def auto_map_cols(headers):
    """
    Mapeia automaticamente os cabeçalhos de uma aba para os campos internos.
    Para no primeiro alias que bater — exact match tem prioridade sobre partial.
    """
    mapping = {}
    headers_lower = [(h, norm(h)) for h in headers]
    for key, aliases in COL_ALIASES.items():
        for alias in aliases:
            for h, l in headers_lower:
                if l == alias or l == alias.replace(' ', '_'):
                    mapping[key] = h
                    break
                if len(alias) >= 5 and alias in l:
                    if key not in mapping:
                        mapping[key] = h
            if key in mapping:
                break
    return mapping

# ─── Carregamento centralizado ────────────────────────────────────────────────

def _load_and_rename(xl, sheet_map, per_sheet_map):
    """
    Fonte única de verdade para leitura, limpeza e renomeação de colunas.
    Elimina a dessincronização que existia quando o mesmo bloco era duplicado
    em api_process, api_preview e api_get_columns.

    Correções aplicadas aqui:
    - Remove BOM (\\ufeff) dos cabeçalhos antes de qualquer comparação.
    - Elimina linhas completamente vazias com filtro compatível com na_filter=False.
    - Deduplica por CPF após concat para remover colaboradores que aparecem em
      mais de uma aba (609 matrículas duplicadas identificadas nos outputs).
    """
    frames = []
    for grupo, sheet_name in sheet_map.items():
        df = xl.parse(sheet_name, dtype=str, na_filter=False)

        # Remove BOM e espaços dos nomes de coluna
        df.columns = [str(c).strip().lstrip('\ufeff') for c in df.columns]

        # Remove linhas onde TODAS as células são vazias (compatível com na_filter=False)
        mask = df.apply(lambda row: row.str.strip().ne('').any(), axis=1)
        df = df[mask].reset_index(drop=True)

        sheet_cols    = list(df.columns)
        sheet_auto    = auto_map_cols(sheet_cols)
        sheet_col_map = per_sheet_map.get(grupo, {})

        rename = {}
        for key in COL_ALIASES.keys():
            user_col = sheet_col_map.get(key, '')
            if user_col and user_col in sheet_cols:
                rename[user_col] = f'_f_{key}'
            elif sheet_auto.get(key) and sheet_auto[key] in sheet_cols:
                rename[sheet_auto[key]] = f'_f_{key}'

        df = df.rename(columns=rename)
        df['_grupo'] = grupo
        frames.append(df)

    if not frames:
        return pd.DataFrame()

    consolidated = pd.concat(frames, ignore_index=True)

    # Deduplica por CPF: colaboradores que aparecem em mais de uma aba
    # (ex: Stefanini + IHM) entravam duplicados no output. CPF é usado como
    # chave porque é único por pessoa; matrícula pode repetir entre empresas.
    cpf_col = '_f_cpf'
    if cpf_col in consolidated.columns:
        # Normaliza o CPF (remove pontuação) para comparação antes de deduplicar
        cpf_norm = consolidated[cpf_col].apply(
            lambda v: CPF_CLEAN.sub('', str(v)) if clean_val(v) else ''
        )
        consolidated = consolidated[cpf_norm.ne('') & ~cpf_norm.duplicated(keep='first')]
        consolidated = consolidated.reset_index(drop=True)

    return consolidated


def _apply_date_conversion(df, internal_map):
    """Converte a coluna de admissão para Timestamp de forma centralizada."""
    adm_col = internal_map.get('data_admissao')
    if adm_col and adm_col in df.columns:
        df[adm_col] = df[adm_col].apply(parse_mixed_dates)
    return df, adm_col

# ─── Processamento por processo ───────────────────────────────────────────────

def apply_rules(df, col_map, process_name, check_fn):
    eligible_rows, excluded_rows = [], []
    email_col    = col_map.get('email')
    admissao_col = col_map.get('data_admissao')
    vinculo_col  = col_map.get('tipo_vinculo')

    for _, row in df.iterrows():
        email   = clean_val(row[email_col]) if email_col and email_col in row.index else ''
        adm     = row[admissao_col] if admissao_col and admissao_col in row.index else None
        vinculo = clean_val(row[vinculo_col]) if vinculo_col and vinculo_col in row.index else ''

        if not email or not is_valid_email(email):
            excluded_rows.append({**row.to_dict(), '_motivo': 'E-mail inválido ou ausente', '_processo': process_name})
            continue

        if adm is not None and is_future_date(adm):
            excluded_rows.append({**row.to_dict(), '_motivo': 'Data de admissão futura', '_processo': process_name})
            continue

        reason = check_fn(row, col_map, vinculo)
        if reason:
            excluded_rows.append({**row.to_dict(), '_motivo': reason, '_processo': process_name})
            continue

        # Regras de domínio de e-mail por grupo
        if row.get('_grupo') == 'IHM' and '@' in email:
            if 'ihm' not in email.split('@')[-1]:
                excluded_rows.append({**row.to_dict(), '_motivo': 'E-mail sem domínio IHM', '_processo': process_name})
                continue

        TOPAZ_KW = ['topaz', 'tpz', 'grupotopaz', 'grupo-topaz', 'newm', 'top-systems', 'topsystems']
        if row.get('_grupo') == 'Topaz' and '@' in email:
            if not any(kw in email.split('@')[-1] for kw in TOPAZ_KW):
                excluded_rows.append({**row.to_dict(), '_motivo': 'E-mail sem domínio Topaz', '_processo': process_name})
                continue

        STEFANINI_KW = ['stefanini']
        if row.get('_grupo') == 'Stefanini' and '@' in email:
            if not any(kw in email.split('@')[-1] for kw in STEFANINI_KW):
                excluded_rows.append({**row.to_dict(), '_motivo': 'E-mail sem domínio Stefanini', '_processo': process_name})
                continue

        eligible_rows.append(row.to_dict())
    return eligible_rows, excluded_rows

def check_tp(row, col_map, vinculo):
    return detect_vinculo_label(vinculo) if vinculo else None

def check_wh(row, col_map, vinculo):
    """
    FIX: vinculo era comparado em case original; keywords são lowercase.
    norm() garante comparação case-insensitive independente da planilha.
    """
    vinculo_norm = norm(vinculo)
    if vinculo_norm:
        bad = ESTAGIO_KW + APRENDIZ_KW + TERCEIRO_KW + PJ_KW
        if any(k in vinculo_norm for k in bad):
            return detect_vinculo_label(vinculo) or 'Vínculo inelegível para Wellhub'
    desconto_col = col_map.get('desconto_folha')
    if desconto_col and desconto_col in row.index:
        val = clean_val(row[desconto_col])
        if val:
            if norm(val) not in DESCONTO_POS:
                return 'Sem desconto em folha habilitado'
    return None

def check_nv(row, col_map, vinculo):
    vinculo_norm = norm(vinculo)
    if vinculo_norm and any(k in vinculo_norm for k in PJ_KW):
        return 'Pessoa Jurídica (PJ)'
    return None

WELLHUB_COLS = ['Name', 'Email', 'National ID', 'Employee ID',
                'Department', 'Cost Center', 'Office Zip Code',
                'Payroll ID', 'Payroll Enabled']

WELLHUB_EMPRESAS_KW = [
    ('gauge',), ('woopi',), ('stefanini', 'data', 'analytics'),
    ('tpz',), ('w3haus',), ('cyber', 'seguranca'),
    ('orbitall', 'atendimento'), ('orbitall', 'servicos'),
    ('orbitall', 'processamentos'), ('stefanini', 'consultoria', 'assessoria'),
    ('intelligenti',), ('n 1 software',), ('n1 software',),
]

def is_wellhub_empresa(nome):
    n = norm(nome)
    return any(all(kw in n for kw in kws) for kws in WELLHUB_EMPRESAS_KW)

# ─── Geração de Excel ─────────────────────────────────────────────────────────

FINAL_COLS = [
    'Nome da Empresa',
    'CNPJ da Empresa',
    'Matrícula',
    'Nome do profissional',
    'Função do profissional',
    'Data de Admissão',
    'Email Funcional',
    'CPF do profissional',
]
AUDIT_COLS = FINAL_COLS + ['Processo', 'Motivo da exclusão', 'Data de processamento']

COL_WIDTHS = {
    'Nome da Empresa': 35, 'CNPJ da Empresa': 22, 'Matrícula': 22,
    'Nome do profissional': 42, 'Função do profissional': 30,
    'Data de Admissão': 18, 'Email Funcional': 38, 'CPF do profissional': 22,
    'Processo': 14, 'Motivo da exclusão': 38, 'Data de processamento': 22,
}

def to_final_row(row_dict, col_map):
    get = lambda k: clean_val(row_dict.get(col_map.get(k, '___absent___'), ''))
    adm_col = col_map.get('data_admissao', '')
    adm_val = row_dict.get(adm_col, '') if adm_col else ''
    return {
        'Nome da Empresa':        get('empresa'),
        'CNPJ da Empresa':        get('cnpj'),
        'Matrícula':              get('matricula'),
        'Nome do profissional':   get('nome'),
        'Função do profissional': get('cargo'),
        'Data de Admissão':       format_date_val(adm_val),
        'Email Funcional':        get('email'),
        # FIX: format_cpf garante formatação padronizada para CPFs sem pontuação (Topaz).
        'CPF do profissional':    format_cpf(row_dict.get(col_map.get('cpf', '___absent___'), '')),
    }

def to_wellhub_row(row_dict, col_map):
    get = lambda k: clean_val(row_dict.get(col_map.get(k, '___absent___'), ''))
    return {
        'Name':            get('nome'),
        'Email':           get('email'),
        'National ID':     format_cpf(row_dict.get(col_map.get('cpf', '___absent___'), '')),
        'Employee ID':     get('matricula'),
        'Department':      '',
        'Cost Center':     '',
        'Office Zip Code': '',
        'Payroll ID':      '',
        'Payroll Enabled': 'YES',
    }

def write_styled_excel(rows, filepath, sheet_name, columns):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]
    header_fill   = PatternFill('solid', fgColor='0F2A56')
    header_font   = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
    data_font     = Font(name='Calibri', size=10)
    alt_fill      = PatternFill('solid', fgColor='F0F4F8')
    bottom_border = Border(bottom=Side(style='thin', color='DBEAFE'))

    for col_idx, col_name in enumerate(columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = Alignment(horizontal='left', vertical='center')
        ws.column_dimensions[cell.column_letter].width = COL_WIDTHS.get(col_name, 20)

    for row_idx, row_data in enumerate(rows, 2):
        for col_idx, col_name in enumerate(columns, 1):
            cell        = ws.cell(row=row_idx, column=col_idx, value=str(row_data.get(col_name, '') or ''))
            cell.font   = data_font
            cell.border = bottom_border
            if row_idx % 2 == 0:
                cell.fill = alt_fill

    ws.row_dimensions[1].height = 22
    ws.freeze_panes = 'A2'
    wb.save(filepath)

# ─── Rotas Flask ──────────────────────────────────────────────────────────────


# ─── Geração de CSV por cliente ──────────────────────────────────────────────
# Colunas e mapeamentos definidos pelos layouts oficiais de cada cliente.
# Nenhuma regra de negócio é alterada aqui — só o formato de saída.

def write_totalpass_csv(tp_rows, filepath):
    """
    Layout TotalPass: CNPJ_EMPRESA | E-MAIL_COLABORADOR | MATRÍCULA
    """
    import csv
    with open(filepath, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=['CNPJ_EMPRESA', 'E-MAIL_COLABORADOR', 'MATRÍCULA'], delimiter=';')
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
        writer = csv.DictWriter(f, fieldnames=['Nome', 'Nome Empresa', 'CNPJ Empresa', 'Número de CPF'], delimiter=';')
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
        writer = csv.DictWriter(f, fieldnames=COLS, delimiter=';')
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

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/upload', methods=['POST'])
def api_upload():
    if 'file' not in request.files:
        return jsonify({'error': 'Nenhum arquivo enviado.'}), 400
    f = request.files['file']
    if not f.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Formato inválido. Envie um arquivo .xlsx ou .xls.'}), 400

    session_id  = str(uuid.uuid4())
    session_dir = os.path.join(TEMP_DIR, session_id)
    os.makedirs(session_dir, exist_ok=True)
    upload_path = os.path.join(session_dir, 'source.xlsx')
    f.save(upload_path)

    try:
        xl       = pd.ExcelFile(upload_path)
        auto_map = {}
        for grupo in ['Stefanini', 'Topaz', 'IHM']:
            for s in xl.sheet_names:
                if grupo.lower() in s.lower() or s.lower() in grupo.lower():
                    auto_map[grupo] = s
                    break
        return jsonify({
            'session_id':     session_id,
            'sheet_names':    xl.sheet_names,
            'auto_sheet_map': auto_map,
            'filename':       f.filename,
        })
    except Exception as e:
        return jsonify({'error': f'Erro ao ler o arquivo: {str(e)}'}), 500

@app.route('/api/get-columns', methods=['POST'])
def api_get_columns():
    data        = request.json
    sid         = data.get('session_id')
    sheet_map   = data.get('sheet_map', {})
    upload_path = os.path.join(TEMP_DIR, sid, 'source.xlsx')
    if not os.path.exists(upload_path):
        return jsonify({'error': 'Sessão expirada.'}), 400

    try:
        xl         = pd.ExcelFile(upload_path)
        sheet_data = {}
        for grupo, sheet_name in sheet_map.items():
            df = xl.parse(sheet_name, dtype=str, na_filter=False)
            df.columns = [str(c).strip().lstrip('\ufeff') for c in df.columns]
            mask = df.apply(lambda row: row.str.strip().ne('').any(), axis=1)
            df   = df[mask].reset_index(drop=True)
            cols = list(df.columns)
            sheet_data[grupo] = {
                'headers':   cols,
                'auto_map':  auto_map_cols(cols),
                'row_count': len(df),
            }
        return jsonify({'sheet_data': sheet_data})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/process', methods=['POST'])
def api_process():
    data          = request.json
    sid           = data.get('session_id')
    sheet_map     = data.get('sheet_map', {})
    per_sheet_map = data.get('per_sheet_col_map', {})
    upload_path   = os.path.join(TEMP_DIR, sid, 'source.xlsx')
    if not os.path.exists(upload_path):
        return jsonify({'error': 'Sessão expirada.'}), 400

    try:
        xl           = pd.ExcelFile(upload_path)
        consolidated = _load_and_rename(xl, sheet_map, per_sheet_map)
        internal_map = {key: f'_f_{key}' for key in COL_ALIASES.keys()}

        consolidated, _ = _apply_date_conversion(consolidated, internal_map)

        today_str   = date.today().strftime('%d/%m/%Y')
        session_dir = os.path.join(TEMP_DIR, sid)

        tp_elig, tp_excl = apply_rules(consolidated, internal_map, 'TotalPass', check_tp)
        wh_elig, wh_excl = apply_rules(consolidated, internal_map, 'Wellhub',   check_wh)
        nv_elig, nv_excl = apply_rules(consolidated, internal_map, 'New Value', check_nv)

        tp_rows = [to_final_row(r, internal_map) for r in tp_elig]
        wh_rows = [to_final_row(r, internal_map) for r in wh_elig]
        nv_rows = [to_final_row(r, internal_map) for r in nv_elig]

        audit_rows = []
        for excl_list in (tp_excl, wh_excl, nv_excl):
            for r in excl_list:
                row = to_final_row(r, internal_map)
                row['Processo']              = r.get('_processo', '')
                row['Motivo da exclusão']    = r.get('_motivo', '')
                row['Data de processamento'] = today_str
                audit_rows.append(row)

        write_styled_excel(tp_rows,    os.path.join(session_dir, 'totalpass.xlsx'), 'TotalPass', FINAL_COLS)
        write_styled_excel(nv_rows,    os.path.join(session_dir, 'newvalue.xlsx'),  'New Value',  FINAL_COLS)
        write_styled_excel(audit_rows, os.path.join(session_dir, 'auditoria.xlsx'), 'Auditoria',  AUDIT_COLS)

        # ── Gera CSVs no formato de cada cliente ──────────────────────────
        write_totalpass_csv(tp_rows, os.path.join(session_dir, 'totalpass.csv'))
        write_newvalue_csv(nv_rows,  os.path.join(session_dir, 'newvalue.csv'))

        # ──────────────────────────────────────────────────────────────────

        # Salva total da base para o PPT de auditoria
        with open(os.path.join(session_dir, 'total_base.txt'), 'w') as _f:
            _f.write(str(len(consolidated)))

        wh_by_cnpj = {}
        cnpj_col   = internal_map.get('cnpj', '')
        emp_col    = internal_map.get('empresa', '')
        for r in wh_elig:
            cnpj_raw = clean_val(r.get(cnpj_col, ''))
            emp_raw  = clean_val(r.get(emp_col, ''))
            cnpj     = cnpj_raw if cnpj_raw else f"GRUPO_{r.get('_grupo', '')}"
            empresa  = emp_raw  if emp_raw  else r.get('_grupo', '')
            if cnpj not in wh_by_cnpj:
                wh_by_cnpj[cnpj] = {'empresa': empresa, 'rows': []}
            wh_row = to_wellhub_row(r, internal_map)
            wh_by_cnpj[cnpj]['rows'].append(wh_row)

        wh_by_cnpj   = {c: g for c, g in wh_by_cnpj.items() if is_wellhub_empresa(g['empresa'])}
        wh_file_meta = []
        for cnpj, grp in wh_by_cnpj.items():
            safe_name = re.sub(r'[^\w]', '_', grp['empresa']).upper().strip('_')
            fname     = f"{safe_name}_Wellhub.xlsx"
            fpath     = os.path.join(session_dir, f"wh_{re.sub(r'[^\w]', '', cnpj)}.xlsx")
            write_styled_excel(grp['rows'], fpath, 'Wellhub', WELLHUB_COLS)
            # guarda rows para uso no CSV
            _rows_raw = grp['rows'][:]
            wh_file_meta.append({
                'cnpj':     cnpj,
                'empresa':  grp['empresa'],
                'count':    len(grp['rows']),
                'filepath': fpath,
                'filename': fname,
                '_rows_raw': _rows_raw,
            })

        zip_path = os.path.join(session_dir, 'wellhub_todos.zip')
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for meta in wh_file_meta:
                zf.write(meta['filepath'], meta['filename'])

        with open(os.path.join(session_dir, 'wh_meta.json'), 'w', encoding='utf-8') as jf:
            json.dump(wh_file_meta, jf, ensure_ascii=False)

        def count_reasons(rows):
            counts = {}
            for r in rows:
                key = r.get('_motivo', 'Desconhecido')
                counts[key] = counts.get(key, 0) + 1
            return counts

        return jsonify({
            'total': len(consolidated),
            'tp': {'eligible': len(tp_rows),  'excluded': len(tp_excl),  'reasons': count_reasons(tp_excl)},
            'wh': {
                'eligible':  sum(m['count'] for m in wh_file_meta),
                'excluded':  len(wh_excl),
                'reasons':   count_reasons(wh_excl),
                'companies': wh_file_meta,
            },
            'nv': {'eligible': len(nv_rows),  'excluded': len(nv_excl),  'reasons': count_reasons(nv_excl)},
            'audit_count': len(audit_rows),
        })
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'detail': traceback.format_exc()}), 500

@app.route('/api/preview', methods=['POST'])
def api_preview():
    data          = request.json
    sid           = data.get('session_id')
    sheet_map     = data.get('sheet_map', {})
    per_sheet_map = data.get('per_sheet_col_map', {})
    upload_path   = os.path.join(TEMP_DIR, sid, 'source.xlsx')
    if not os.path.exists(upload_path):
        return jsonify({'error': 'Sessão expirada.'}), 400

    try:
        xl           = pd.ExcelFile(upload_path)
        consolidated = _load_and_rename(xl, sheet_map, per_sheet_map)
        internal_map = {key: f'_f_{key}' for key in COL_ALIASES.keys()}
        consolidated, _ = _apply_date_conversion(consolidated, internal_map)

        def cell(row, key):
            return clean_val(row.get(internal_map.get(key, '___absent___'), ''))

        rows = []
        for _, row in consolidated.head(15).iterrows():
            rows.append({
                'empresa':      cell(row, 'empresa'),
                'nome':         cell(row, 'nome'),
                'matricula':    cell(row, 'matricula'),
                'email':        cell(row, 'email'),
                'tipo_vinculo': cell(row, 'tipo_vinculo'),
                'cargo':        cell(row, 'cargo'),
                '_grupo':       row.get('_grupo', ''),
            })

        return jsonify({'rows': rows, 'total': len(consolidated)})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/wellhub-zip/<session_id>', methods=['POST'])
def wellhub_zip_filtered(session_id):
    session_dir = os.path.join(TEMP_DIR, session_id)
    meta_path   = os.path.join(session_dir, 'wh_meta.json')
    if not os.path.exists(meta_path):
        return jsonify({'error': 'Sessão expirada.'}), 400

    selected_cnpjs = set(request.json.get('cnpjs', []))
    with open(meta_path, encoding='utf-8') as f:
        wh_file_meta = json.load(f)

    selected = [m for m in wh_file_meta if m['cnpj'] in selected_cnpjs]
    if not selected:
        return jsonify({'error': 'Nenhuma empresa selecionada.'}), 400

    zip_path = os.path.join(session_dir, 'wellhub_filtrado.zip')
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for meta in selected:
            if os.path.exists(meta['filepath']):
                zf.write(meta['filepath'], meta['filename'])

    return send_file(zip_path, as_attachment=True, download_name=f'Wellhub_{len(selected)}_Empresas.zip')

@app.route('/api/download/<session_id>/<file_type>')
def api_download(session_id, file_type):
    session_dir = os.path.join(TEMP_DIR, session_id)
    if file_type == 'wellhub':
        path = os.path.join(session_dir, 'wellhub_todos.zip')
        if os.path.exists(path):
            return send_file(path, as_attachment=True, download_name='Wellhub_Todos_CNPJs.zip')
        return 'Arquivo não encontrado.', 404

    file_map = {
        'totalpass': ('totalpass.xlsx', 'Base_TotalPass_Tratada.xlsx'),
        'newvalue':  ('newvalue.xlsx',  'Base_NewValue_Tratada.xlsx'),
        'auditoria': ('auditoria.xlsx', 'Base_Auditoria_Exclusoes.xlsx'),
    }
    if file_type not in file_map:
        return 'Tipo inválido.', 404
    internal, display = file_map[file_type]
    path = os.path.join(session_dir, internal)
    if os.path.exists(path):
        return send_file(path, as_attachment=True, download_name=display)
    return 'Arquivo não encontrado.', 404



@app.route('/api/download/<session_id>/auditppt')
def download_audit_ppt(session_id):
    session_dir = os.path.join(TEMP_DIR, session_id)
    audit_path  = os.path.join(session_dir, 'auditoria.xlsx')
    ppt_path    = os.path.join(session_dir, 'auditoria_ppt.pptx')
    if not os.path.exists(audit_path):
        return 'Sessão expirada ou processamento não concluído.', 404
    try:
        total_base_path = os.path.join(session_dir, 'total_base.txt')
        total_base = int(open(total_base_path).read().strip()) if os.path.exists(total_base_path) else 0
        generate_ppt(audit_path, ppt_path, total_base)
        return send_file(ppt_path, as_attachment=True,
                         download_name='Auditoria_Exclusoes_BenefProcess.pptx')
    except Exception as e:
        import traceback
        return f'Erro ao gerar PPT: {str(e)}\n{traceback.format_exc()}', 500


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

if __name__ == '__main__':
    print('=' * 50)
    print('  BenefProcess — Tratamento de Benefícios')
    print('  Stefanini · Gente e Cultura')
    print('=' * 50)
    print('  Acesse no navegador: http://localhost:5050')
    print('=' * 50)
    app.run(debug=False, port=5050)