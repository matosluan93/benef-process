"""
add_ppt_route.py
Adiciona a geração de PPT ao BenefProcess sem tocar na análise.
Execute UMA VEZ dentro da pasta do projeto:
    py add_ppt_route.py
"""
import os, ast

app_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'app.py')

with open(app_path, encoding='utf-8') as f:
    code = f.read()

# 1. Adiciona import do gerador no topo
import_line = "from ppt_generator import generate as generate_ppt"
if import_line not in code:
    code = code.replace(
        "from flask import Flask",
        f"{import_line}\nfrom flask import Flask"
    )
    print("✓ Import adicionado")
else:
    print("  Import já existe, pulando")

# 2. Salva total_base.txt logo após gravar auditoria.xlsx
trigger = "write_styled_excel(audit_rows, os.path.join(session_dir, 'auditoria.xlsx'), 'Auditoria',  AUDIT_COLS)"
save_total = """write_styled_excel(audit_rows, os.path.join(session_dir, 'auditoria.xlsx'), 'Auditoria',  AUDIT_COLS)

        # Salva total da base para o PPT de auditoria
        with open(os.path.join(session_dir, 'total_base.txt'), 'w') as _f:
            _f.write(str(len(consolidated)))"""

if 'total_base.txt' not in code:
    if trigger in code:
        code = code.replace(trigger, save_total)
        print("✓ Salvamento de total_base.txt adicionado")
    else:
        print("  AVISO: trigger não encontrado — adicione manualmente")
else:
    print("  total_base.txt já existe, pulando")

# 3. Adiciona rota de download do PPT antes do bloco if __name__
new_route = '''
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
        return f'Erro ao gerar PPT: {str(e)}\\n{traceback.format_exc()}', 500

'''

if 'download_audit_ppt' not in code:
    code = code.replace("if __name__ == '__main__':", new_route + "if __name__ == '__main__':")
    print("✓ Rota /api/download/.../auditppt adicionada")
else:
    print("  Rota já existe, pulando")

# Valida sintaxe antes de salvar
try:
    ast.parse(code)
    print("✓ Sintaxe validada")
except SyntaxError as e:
    print(f"✗ Erro de sintaxe: {e}")
    exit(1)

with open(app_path, 'w', encoding='utf-8') as f:
    f.write(code)

print(f"\n✓ app.py atualizado em: {app_path}")
print("\nPróximos passos:")
print("  1. Copie o ppt_generator.py para a pasta do projeto")
print("  2. pip install python-pptx lxml")
print("  3. py app.py")
