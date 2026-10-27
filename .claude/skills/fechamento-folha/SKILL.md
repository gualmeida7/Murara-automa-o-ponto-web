---
name: fechamento-folha
description: Contexto e regras do projeto de automação do fechamento de folha (Maçaneiro e Gonzaga / Cianorte Tubos) — Python, pandas, Streamlit, SQLite. Use ao alterar ou depurar pipeline.py, business_rules.py, afd_parser.py, banco_horas.py, feriados, app_web.py, db_sessao.py, rodar testes ou gerar a Relação de Valores / fila de exceções.
---

# Fechamento de folha — guia de trabalho

Automatiza o fechamento mensal (período padrão **26 a 25**): lê o ponto (AFD `.txt` ou cartão-ponto PDF), cruza com a jornada do cadastro, calcula HE/atraso/falta/COPR/VA/DSR, separa o que precisa de decisão do RH e gera um `.xlsx` por empresa ("Relação de Valores") para a contabilidade. O README.md tem a documentação completa; consulte-o antes de mudar comportamento.

## Mapa dos módulos

| Módulo | Responsabilidade |
|---|---|
| `afd_parser.py` | Parse do AFD. Layout em `LAYOUT_TIPO3` (único lugar a ajustar). |
| `cartao_ponto_pdf.py`, `pdf_espelho_ponto.py` | Leitura de PDFs de ponto (pdfplumber). |
| `leitores_ponto.py` | Escolhe o leitor por extensão; mapeia identificador → PIS. |
| `business_rules.py` | `Jornada`, `apurar_dia`, tolerância CLT, falta parcial, COPR, VA, DSR. |
| `banco_horas.py` | Banco de horas Secullum (Sidney) e planilha do Adriano — só para quem está no banco. |
| `feriados.py` / `gerar_feriados.py` / `feriados.csv` | Calendário offline (2025–2199). |
| `validacao.py`, `leitura_arquivos.py` | Validação defensiva, `ArquivoInvalidoError`, leitura resiliente csv/xlsx. |
| `pipeline.py` | Orquestra: `rodada_1_gerar_fila_de_excecoes` → decisões do RH → `rodada_2_gerar_relacao_de_valores` → `exportar_por_empresa`. |
| `db_sessao.py` | SQLite (`fechamento_folha.db`) que guarda decisões/sessão por competência. |
| `app_web.py` | UI Streamlit (recomendada). `app_gui.py` é o desktop tkinter legado. |

Regra de ouro: **a lógica de cálculo mora nos módulos, não nos apps.** `app_web.py` e `app_gui.py` só chamam o `pipeline`; não duplique regras nas telas.

## Regras de negócio que não podem quebrar

- Tolerância CLT art. 58 §1º: **5 min por batida** e **10 min agregados no dia** (configuráveis por colaborador).
- **Feriado não é falta**: ausência não tira COPR, não desconta VA/DSR e não vai para o RH; trabalho em feriado = **HE 100%** (igual domingo). Só `nacional` e `municipal` contam como feriado; `estadual` e `facultativo` não.
- **COPR**: perde o prêmio integral com qualquer falta/ausência no mês, mesmo justificada.
- **VA**: desconto proporcional só por falta **sem** justificativa.
- **DSR (8794)**: perda por semana ISO com ≥1 falta de **dia inteiro** injustificada; falta parcial e atraso não contam. Lançado em dias ou valor.
- O código **nunca decide sozinho** se uma falta é justificada: tudo ambíguo vai para a fila do RH, com padrão conservador "falta injustificada".
- Banco de horas não pode afetar colaboradores fora do banco.
- PIS/matrícula: sempre normalizar com `canonicalizar_pis`; erros de arquivo viram `ArquivoInvalidoError` com mensagem em português apontando arquivo e PIS (nunca traceback cru para o usuário).

## Convenções de código

- Português em nomes, docstrings, mensagens ao usuário e commits; `from __future__ import annotations`; type hints modernos (`str | None`).
- Docstrings de módulo explicam o "porquê" das regras de negócio — mantenha-as atualizadas ao mudar uma regra.
- Dependências: streamlit, pandas, openpyxl, pdfplumber. Evite adicionar novas sem necessidade.
- Atualize o README.md quando mudar comportamento visível ou fluxo.

## Rodar e testar (Windows)

Use o venv do projeto:

```powershell
.\.venv\Scripts\python.exe -m pytest -q                       # todos os testes
.\.venv\Scripts\python.exe -m pytest teste_feriados.py -q     # um arquivo
.\.venv\Scripts\streamlit.exe run app_web.py                  # app em localhost:8501
```

Testes: `teste_sintetico.py` (ponta a ponta com dados fictícios), `teste_banco_dsr.py`, `teste_falta_parcial.py`, `teste_feriados.py`, `teste_export_excecoes.py`. Mudou regra em `business_rules.py` ou `pipeline.py` → rode a suíte inteira antes de dizer que terminou, e adicione/ajuste teste para o caso novo.

## Cuidados

- **Origem dos arquivos em `Dados e Docs Reais/`**: só `cadastro_sintetico.xlsx` foi feito pelo desenvolvedor (fictício). Todo o resto — Cartão Ponto e Espelho de Ponto (PDF), extratos de consignado (`<CNPJ>-<aaaamm>v1.4.xlsx`) e `Consignado Modificado.xlsx` — são documentos **reais do cliente**. Ao alterar o código, trate o cadastro como sintético (pode não refletir o real, p.ex. sem colunas opcionais) e os demais como formato real que o parser precisa continuar aceitando. Nunca cite nomes/PIS/valores deles em guias, testes ou saídas.
- **Dados reais e sensíveis** (PIS, nomes, folha): `Dados e Docs Reais/`, `dados/`, `saida/`, `fechamento_folha.db` e `.log` estão no `.gitignore` — nunca commite nem cole esse conteúdo em saídas. Os `.txt`/`.xlsx` soltos na raiz (AFD/espelho/planilha) são amostras não rastreadas; confirme com o usuário antes de commitá-los.
- Não apague nem recrie `fechamento_folha.db` sem pedir: contém decisões do RH.
- `feriados.csv` é gerado por `gerar_feriados.py`; edite o gerador (ou use seus args) em vez de editar o CSV à mão.
- `app_web.py` usa `st.session_state` + SQLite: ao mexer na tabela de aprovação, preserve a gravação das decisões (`db_sessao.salvar_decisao*`) e a limpeza de estado do editor (`limpar_estado_editor`).
- Se for mexer na UI Streamlit, carregue também a skill `developing-with-streamlit`.
