# BenefProcess

Tratamento automatizado de bases de benefícios da Stefanini, com processamento 100% local.

## O Que Faz

O sistema roda em Flask e oferece três fluxos pela interface web:

- Base principal: Stefanini, Topaz e IHM para TotalPass, Wellhub e New Value.
- Mar Saúde: gera elegíveis e auditoria de excluídos PJ.
- Expert: separa ativos, inativos e removidos, com Excel, CSV e apresentação PPT.

## Saídas

| Fluxo | Arquivos gerados |
|---|---|
| TotalPass | Excel tratado e CSV |
| Wellhub | ZIP com Excel por empresa e ZIP com CSVs |
| New Value | Excel tratado e CSV |
| Auditoria | Excel de exclusões e PPT de auditoria |
| Mar Saúde | Excel de elegíveis, CSV de elegíveis e Excel de excluídos |
| Expert | Excel/CSV de ativos, Excel/CSV de inativos e PPT executivo |

## Regras Principais

Base principal:

- Remove admissões futuras.
- Remove e-mails inválidos ou ausentes.
- Remove duplicidades por CPF.
- TotalPass remove estagiários, aprendizes, terceiros e PJ.
- Wellhub mantém colaboradores elegíveis com desconto em folha habilitado.
- New Value remove PJ.

Mar Saúde:

- Remove PJ por `Nome Sindicato = Somente PJ`.
- Remove PJ por `Desc. Tipo de Vínculo = OUTROS` ou termos equivalentes a PJ.

Expert:

- Somente `Desc. Situação = Ativo` entra em ativos.
- Qualquer outra situação entra em inativos, mantendo a situação original no arquivo.
- Remove ativos que sejam PJ, estagiários ou aprendizes conforme o vínculo/categoria.

## Pré-Requisitos

- Python 3.10 ou superior.
- Navegador para acessar `http://localhost:5050`.

## Como Usar

### Windows

Dê duplo clique em `iniciar.bat`.

O script cria o ambiente virtual, confere as dependências e inicia o servidor.

### Git Bash

```bash
bash iniciar.sh
```

### Manual

```bash
py -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
py app.py
```

Depois acesse:

```text
http://localhost:5050
```

## Estrutura

```text
benef-process/
├── app.py              Servidor Flask e regras de negócio
├── ppt_generator.py    Geração do PPT de auditoria
├── requirements.txt    Dependências Python
├── iniciar.bat         Inicialização Windows
├── iniciar.sh          Inicialização Git Bash/Linux
├── templates/
│   └── index.html      Interface web
└── temp/               Arquivos temporários gerados pelo processamento
```

## Observações

- A pasta `temp/` pode crescer bastante e pode ser limpa manualmente quando não houver sessões em uso.
- Nenhum arquivo é enviado para a internet.
- Arquivos `app_backup*.py`, `fix_*.py` e `add_*.py` são históricos de manutenção e não fazem parte do fluxo principal de execução.
