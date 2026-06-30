"""
fix_dominios_stefanini.py
Atualiza a lista de domínios válidos para o grupo Stefanini.
A verificação muda de substring para domínio exato (case-insensitive).
Execute dentro da pasta do projeto:
    py fix_dominios_stefanini.py
"""
import os, ast

app_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'app.py')

with open(app_path, encoding='utf-8') as f:
    code = f.read()

# ── 1. Adiciona o set de domínios antes das regras de negócio ─────────────────
DOMINIOS_BLOCK = """
# ─── Domínios válidos por grupo ───────────────────────────────────────────────

STEFANINI_DOMINIOS = {
    'ar.stefanini.com', 'bankinginabox.site', 'caps.haus',
    'cl.stefanini.com', 'co.stefanini.com', 'datastorm.com.br',
    'ecglobal.com', 'ecglobal.haus', 'ecossistema.haus',
    'gauge.com.br', 'gauge.haus', 'huia.haus',
    'ihm.com.br', 'ihmsenior.com.br', 'inspiring.haus',
    'latam.stefanini.com', 'marcostefanini.com', 'marcostefanini.com.br',
    'mx.stefanini.com', 'necxt.com.br', 'necxtorbitall.com.br',
    'openstartups.com.br', 'orbitall.com.br', 'orbitallpay.com.br',
    'originacao.com', 'originacao.com.br', 'pe.stefanini.com',
    'perep.eu', 'perepanalytics.eu', 'pontocertificado.com',
    'pontocertificado.com.br', 'potenciaisstefanini.com.br',
    'reseller.stefanini.com', 'sbc091.teamsonevoice.com',
    'scalait.com', 'seniorengenharia.com.br', 'singulahr.com.br',
    'sophie.chat', 'stefanini.com', 'stefanini.com.br',
    'stefanini.org.br', 'stefaniniathome.com.br', 'stefaninicyber.com',
    'stefaninilatam.mail.onmicrosoft.com', 'stefaninilatam.onmicrosoft.com',
    'stefaninirafael.com', 'stefaninirafael.com.br',
    'stefaniniservico.com.br', 'stefaninitrends.com', 'stefaninitrends.com.br',
    'stfeacesso.com.br', 'sunrising.com.br', 'sv.stefanini.com',
    'techteam.biz', 'techteam.com', 'useniu.com.br', 'w3.haus',
}

"""

anchor_rules = "ESTAGIO_KW   = "

if 'STEFANINI_DOMINIOS' not in code:
    if anchor_rules in code:
        code = code.replace(anchor_rules, DOMINIOS_BLOCK + anchor_rules)
        print("✓ STEFANINI_DOMINIOS adicionado")
    else:
        print("✗ Âncora ESTAGIO_KW não encontrada")
        exit(1)
else:
    print("  STEFANINI_DOMINIOS já existe, pulando")

# ── 2. Substitui a verificação de domínio Stefanini em apply_rules ────────────
old_check = """        STEFANINI_KW = ['stefanini']
        if row.get('_grupo') == 'Stefanini' and '@' in email:
            if not any(kw in email.split('@')[-1] for kw in STEFANINI_KW):
                excluded_rows.append({**row.to_dict(), '_motivo': 'E-mail sem domínio Stefanini', '_processo': process_name})
                continue"""

new_check = """        if row.get('_grupo') == 'Stefanini' and '@' in email:
            if email.split('@')[-1].lower() not in STEFANINI_DOMINIOS:
                excluded_rows.append({**row.to_dict(), '_motivo': 'E-mail sem domínio Stefanini', '_processo': process_name})
                continue"""

if 'STEFANINI_DOMINIOS' in code and old_check in code:
    code = code.replace(old_check, new_check)
    print("✓ Verificação de domínio Stefanini atualizada")
elif 'STEFANINI_DOMINIOS' in code and old_check not in code:
    print("  Verificação já atualizada ou estrutura diferente")
    # Tenta variante sem STEFANINI_KW (já parcialmente modificada)
    old_v2 = """        if row.get('_grupo') == 'Stefanini' and '@' in email:
            if not any(kw in email.split('@')[-1] for kw in STEFANINI_KW):"""
    if old_v2 in code:
        code = code.replace(
            old_v2,
            """        if row.get('_grupo') == 'Stefanini' and '@' in email:
            if email.split('@')[-1].lower() not in STEFANINI_DOMINIOS:"""
        )
        print("✓ Verificação atualizada (variante)")

# ── 3. Valida e salva ─────────────────────────────────────────────────────────
try:
    ast.parse(code)
    print("✓ Sintaxe Python válida")
except SyntaxError as e:
    print(f"✗ Erro de sintaxe: {e}"); exit(1)

with open(app_path, 'w', encoding='utf-8') as f:
    f.write(code)

print(f"\n✓ app.py atualizado — {len('''ar.stefanini.com bankinginabox.site caps.haus cl.stefanini.com co.stefanini.com datastorm.com.br ecglobal.com ecglobal.haus ecossistema.haus gauge.com.br gauge.haus huia.haus ihm.com.br ihmsenior.com.br inspiring.haus latam.stefanini.com marcostefanini.com marcostefanini.com.br mx.stefanini.com necxt.com.br necxtorbitall.com.br openstartups.com.br orbitall.com.br orbitallpay.com.br originacao.com originacao.com.br pe.stefanini.com perep.eu perepanalytics.eu pontocertificado.com pontocertificado.com.br potenciaisstefanini.com.br reseller.stefanini.com sbc091.teamsonevoice.com scalait.com seniorengenharia.com.br singulahr.com.br sophie.chat stefanini.com stefanini.com.br stefanini.org.br stefaniniathome.com.br stefaninicyber.com stefaninilatam.mail.onmicrosoft.com stefaninilatam.onmicrosoft.com stefaninirafael.com stefaninirafael.com.br stefaniniservico.com.br stefaninitrends.com stefaninitrends.com.br stfeacesso.com.br sunrising.com.br sv.stefanini.com techteam.biz techteam.com useniu.com.br w3.haus'''.split())} domínios válidos para o grupo Stefanini")
