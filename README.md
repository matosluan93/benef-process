# BenefProcess
**Tratamento automatizado de bases de benefícios · Stefanini**

---

## O que faz

Recebe **1 arquivo Excel com 3 abas** (Stefanini, Topaz, IHM) e gera
**3 planilhas tratadas** prontas para envio aos fornecedores:

| Saída | Regra principal |
|---|---|
| `Base_TotalPass_Tratada.xlsx` | Remove estagiários, aprendizes, terceiros e PJ |
| `Base_Wellhub_Tratada.xlsx` | Mantém CLT com desconto em folha habilitado |
| `Base_NewValue_Tratada.xlsx` | Remove apenas PJ |
| `Base_Auditoria_Exclusoes.xlsx` | Todos os removidos com motivo + data |

**Regras aplicadas em todos os processos:**
- Remove admissões futuras (data > hoje)
- Remove e-mails inválidos ou ausentes
- Remove registros duplicados por e-mail

---

## Pré-requisito

Python 3.10 ou superior instalado → [python.org](https://www.python.org/downloads/)

---

## Como usar

### Opção A — Duplo clique (mais fácil)
Execute o arquivo **`iniciar.bat`** (na primeira vez, instala as dependências automaticamente).

### Opção B — Git Bash
```bash
bash iniciar.sh
```

### Opção C — Manual
```bash
# Criar e ativar ambiente virtual
py -m venv venv
source venv/Scripts/activate   # Git Bash
# ou: venv\Scripts\activate.bat  (cmd.exe)

# Instalar dependências
pip install -r requirements.txt

# Iniciar
py app.py
```

Acesse no navegador: **http://localhost:5050**

---

## Estrutura do projeto

```
benef-process/
├── app.py              → Servidor Flask (regras de negócio)
├── requirements.txt    → Dependências Python
├── iniciar.bat         → Inicialização Windows (cmd.exe)
├── iniciar.sh          → Inicialização Git Bash
├── templates/
│   └── index.html      → Interface web (sem dependências externas)
└── temp/               → Arquivos temporários (criado automaticamente)
```

---

## Fluxo de uso

```
1. Upload      → Envie o .xlsx com as 3 abas
2. Abas        → Confirme qual aba = Stefanini / Topaz / IHM
3. Colunas     → Confirme o mapeamento de campos (auto-detectado)
4. Revisão     → Veja a prévia consolidada
5. Resultado   → Baixe os 3 arquivos tratados + auditoria
```

---

## Formato esperado do Excel de entrada

| Empresa | CNPJ | Nome | Matrícula | E-mail | Tipo de Vínculo | Data de Admissão | Desconto em Folha |
|---|---|---|---|---|---|---|---|

Os nomes exatos das colunas podem variar — o sistema detecta automaticamente.

**Desconto em folha (Wellhub):** valores aceitos como habilitado:
`Sim`, `S`, `Ativo`, `Habilitado`, `1`, `Yes`, `OK` (não diferencia maiúsculas/minúsculas)

---

## Observações

- Os arquivos temporários ficam na pasta `temp/` e podem ser apagados manualmente
- Nenhum dado é enviado para a internet — processamento 100% local
- Para limpar a pasta temp: apague o conteúdo de `temp/` manualmente
