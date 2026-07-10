# BenefProcess

![Python](https://img.shields.io/badge/Python-3.10%2B-0F2A56)
![Flask](https://img.shields.io/badge/Flask-Web%20local-00A98B)
![Status](https://img.shields.io/badge/status-produ%C3%A7%C3%A3o%20local-1D4E89)
![Privacidade](https://img.shields.io/badge/processamento-100%25%20local-64748B)

Aplicação web local para tratamento, auditoria e exportação de bases de benefícios da Stefanini.

O sistema consolida planilhas, aplica regras de elegibilidade por cliente e gera arquivos finais em Excel, CSV, ZIP e apresentações executivas em PowerPoint.

## Visão Geral

BenefProcess foi criado para transformar bases operacionais em entregáveis prontos para envio e apresentação executiva.

Principais objetivos:

- Padronizar bases de benefícios.
- Reduzir análise manual.
- Registrar motivos de exclusão.
- Gerar arquivos nos layouts dos clientes.
- Manter o processamento local, sem envio de arquivos para a internet.

## Fluxos Disponíveis

| Fluxo | Entrada | Saídas |
|---|---|---|
| TotalPass | Base consolidada Stefanini, Topaz e IHM | Excel tratado e CSV no layout do cliente |
| Wellhub | Base consolidada Stefanini, Topaz e IHM | ZIP com Excel por empresa e ZIP com CSVs |
| New Value | Base consolidada Stefanini, Topaz e IHM | Excel tratado e CSV no layout do cliente |
| Auditoria | Resultado das exclusões da base principal | Excel de auditoria e PPT executivo |
| Mar Saúde | Base Mar Saúde | Excel, CSV e auditoria de exclusões |
| Expert | Base Expert | Ativos, inativos, removidos e PPT executivo |

## Regras Principais

### Base Principal

- Remove admissões futuras.
- Remove e-mails ausentes ou inválidos.
- Remove registros duplicados por CPF.
- Prioriza ocorrência não-PJ quando o mesmo CPF aparece em mais de uma aba.
- TotalPass remove estagiários, aprendizes, terceiros e PJ.
- Wellhub mantém registros elegíveis por empresa/CNPJ habilitado.
- New Value remove apenas PJ, além das validações gerais.

### Mar Saúde

- Remove PJ por sindicato.
- Remove PJ por vínculo/categoria quando houver indicação equivalente.
- Gera arquivo final de elegíveis e auditoria de removidos.

### Expert

- Usa `Desc. Situação` como coluna de referência quando disponível.
- Envia para ativos somente registros com situação `Ativo`.
- Envia para inativos qualquer situação diferente de `Ativo`.
- Remove PJ, estagiários e aprendizes conforme vínculo/categoria.
- Gera apresentação executiva para diretoria.

## Como Rodar Localmente

### Windows

Execute:

```powershell
.\iniciar.bat
```

O script cria/ativa o ambiente virtual, instala dependências e inicia o servidor.

Depois acesse:

```text
http://localhost:5050
```

### Git Bash ou Linux

```bash
bash iniciar.sh
```

### Execução Manual

```powershell
py -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt
py app.py
```

## Estrutura do Projeto

```text
benef-process/
├── app.py                 Regras, rotas Flask e geração dos arquivos
├── ppt_generator.py       Geração do PPT executivo da auditoria
├── requirements.txt       Dependências Python
├── iniciar.bat            Inicialização no Windows
├── iniciar.sh             Inicialização via shell
├── templates/
│   └── index.html         Interface web da aplicação
└── temp/                  Arquivos temporários gerados localmente
```

## Arquivos Gerados

A pasta `temp/` é criada automaticamente durante o uso da aplicação.

Ela pode conter:

- Bases enviadas pelo usuário.
- Excel tratado.
- CSV de clientes.
- ZIPs por empresa.
- PPTs gerados.
- Arquivos de auditoria.

Esses arquivos não fazem parte do código-fonte e não devem ser enviados para o GitHub.

## Segurança e Privacidade

- O processamento é local.
- Nenhuma planilha é enviada para serviços externos.
- Arquivos de saída ficam na pasta `temp/` do projeto.
- O repositório mantém apenas código, dependências e documentação.

## Dependências Principais

- Flask
- pandas
- openpyxl
- xlsxwriter
- python-pptx
- lxml

## Manutenção do Repositório

O GitHub deve manter apenas os arquivos essenciais da aplicação.

Não devem ser versionados:

- `temp/`
- `__pycache__/`
- `venv/`
- arquivos `.pyc`
- scripts antigos de correção `fix_*.py`
- scripts antigos de adição `add_*.py`
- backups locais

Isso mantém o repositório leve, limpo e seguro para baixar em outra máquina.
