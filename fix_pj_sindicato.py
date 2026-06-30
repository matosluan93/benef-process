"""
fix_pj_sindicato.py
Adiciona o campo 'Nome Sindicato' como critério de identificação de PJ
em TODOS os processos: TotalPass, Wellhub e New Value.

Regra: se 'Nome Sindicato' contiver 'Somente Pj' (case-insensitive),
o profissional é excluído como Pessoa Jurídica (PJ).

A verificação do tipo_vinculo existente é mantida como fallback.
Execute dentro da pasta do projeto:
    py fix_pj_sindicato.py
"""
import os, ast

app_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'app.py')

with open(app_path, encoding='utf-8') as f:
    code = f.read()

changed = False

# ── 1. Adiciona 'nome_sindicato' ao COL_ALIASES ───────────────────────────────
old_aliases = """    'tipo_vinculo': [
        # FIX: Stefanini tem 'Desc. Tipo de Vínculo' = 'CONTRATADO'/'ESTAGIÁRIO'.
        # A coluna 'Vínculo' contém apenas o código de letra ('C'), que nunca
        # bateria nas keywords de ESTAGIO_KW / PJ_KW.
        'desc. tipo de vínculo', 'desc. tipo de vinculo',
        'desc. vínculo', 'desc. vinculo',
        'vínculo', 'vinculo',
        'tipo de vínculo', 'tipo de vinculo',
        'regime', 'modalidade',
    ],"""

new_aliases = """    'tipo_vinculo': [
        # FIX: Stefanini tem 'Desc. Tipo de Vínculo' = 'CONTRATADO'/'ESTAGIÁRIO'.
        # A coluna 'Vínculo' contém apenas o código de letra ('C'), que nunca
        # bateria nas keywords de ESTAGIO_KW / PJ_KW.
        'desc. tipo de vínculo', 'desc. tipo de vinculo',
        'desc. vínculo', 'desc. vinculo',
        'vínculo', 'vinculo',
        'tipo de vínculo', 'tipo de vinculo',
        'regime', 'modalidade',
    ],
    'nome_sindicato': [
        'nome sindicato', 'sindicato', 'nome_sindicato',
        'desc. sindicato', 'desc sindicato',
    ],"""

if 'nome_sindicato' not in code:
    if old_aliases in code:
        code = code.replace(old_aliases, new_aliases)
        print("✓ 'nome_sindicato' adicionado ao COL_ALIASES")
        changed = True
    else:
        print("✗ Trecho COL_ALIASES não encontrado")
else:
    print("  'nome_sindicato' já existe no COL_ALIASES, pulando")

# ── 2. Adiciona verificação de sindicato em apply_rules ───────────────────────
old_apply = """        if not email or not is_valid_email(email):
            excluded_rows.append({**row.to_dict(), '_motivo': 'E-mail inválido ou ausente', '_processo': process_name})
            continue

        if adm is not None and is_future_date(adm):
            excluded_rows.append({**row.to_dict(), '_motivo': 'Data de admissão futura', '_processo': process_name})
            continue

        reason = check_fn(row, col_map, vinculo)"""

new_apply = """        if not email or not is_valid_email(email):
            excluded_rows.append({**row.to_dict(), '_motivo': 'E-mail inválido ou ausente', '_processo': process_name})
            continue

        if adm is not None and is_future_date(adm):
            excluded_rows.append({**row.to_dict(), '_motivo': 'Data de admissão futura', '_processo': process_name})
            continue

        # Verifica coluna Nome Sindicato: "Somente Pj" = PJ em qualquer processo
        sindicato_col = col_map.get('nome_sindicato')
        sindicato_val = clean_val(row[sindicato_col]) if sindicato_col and sindicato_col in row.index else ''
        if sindicato_val and 'somente pj' in norm(sindicato_val):
            excluded_rows.append({**row.to_dict(), '_motivo': 'Pessoa Jurídica (PJ)', '_processo': process_name})
            continue

        reason = check_fn(row, col_map, vinculo)"""

if 'sindicato_col' not in code:
    if old_apply in code:
        code = code.replace(old_apply, new_apply)
        print("✓ Verificação de 'Nome Sindicato' adicionada em apply_rules")
        changed = True
    else:
        print("✗ Trecho apply_rules não encontrado")
else:
    print("  Verificação de sindicato já existe em apply_rules, pulando")

# ── 3. Valida e salva ─────────────────────────────────────────────────────────
try:
    ast.parse(code)
    print("✓ Sintaxe Python válida")
except SyntaxError as e:
    print(f"✗ Erro de sintaxe: {e}"); exit(1)

with open(app_path, 'w', encoding='utf-8') as f:
    f.write(code)

if changed:
    print("\n✓ app.py atualizado com sucesso")
    print("  Agora TotalPass, Wellhub e New Value também")
    print("  excluem profissionais com Nome Sindicato = 'Somente Pj'")
else:
    print("\n  Nenhuma alteração necessária")
