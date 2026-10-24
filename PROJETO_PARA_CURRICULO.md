# Projeto: Automação do Fechamento de Folha de Pagamento

> **Instrução para a IA:** este documento descreve um projeto real e serve de fonte para adicioná-lo ao currículo do autor. Use apenas os fatos aqui descritos. Não invente métricas, prazos ou resultados. Os itens marcados com `[PREENCHER]` são dados que só o autor sabe; se não forem preenchidos, omita a informação em vez de estimar. Escreva em português.

## 1. Resumo em uma linha

Sistema em Python que automatiza o fechamento mensal da folha de pagamento: lê o ponto eletrônico, aplica as regras trabalhistas (CLT), separa o que precisa de decisão humana do RH e gera as planilhas por empresa para a contabilidade.

## 2. Contexto e problema

- **Cliente/empresas atendidas:** Maçaneiro e Gonzaga e Cianorte Tubos (Cianorte, PR). `[CONFIRMAR com o autor se os nomes podem aparecer no currículo; alternativa: "escritório de contabilidade / indústria de tubos"]`
- **Processo anterior:** apontamentos feitos à mão (a lápis), consolidação gerencial digitada, aprovação de exceções em planilha impressa marcada à caneta. Isso gerava retrabalho e risco de erro de digitação.
- **Período de fechamento:** do dia 26 ao dia 25 de cada mês.
- **Papel do autor:** `[PREENCHER: ex. desenvolvedor único / desenvolvedor e analista de requisitos]`
- **Período do desenvolvimento:** a partir de setembro de 2026 (primeiro commit em 14/09/2026), 22 commits no repositório.
- **Status:** `[PREENCHER: em produção / em homologação / protótipo]`

## 3. O que o sistema faz

1. **Leitura do ponto:** lê o arquivo AFD (`.txt`, Portaria 671/2021), o cartão-ponto em PDF e o espelho de ponto em PDF. O leitor é escolhido pela extensão do arquivo.
2. **Cálculo automático:** cruza as batidas com a jornada de cada colaborador e calcula hora extra 50%, hora extra 100% (domingo e feriado), atraso, falta de dia inteiro e falta parcial (meio período).
3. **Regras da CLT:** tolerância do Art. 58 §1º (5 minutos por batida e 10 minutos agregados no dia, configuráveis por colaborador).
4. **Consolidação mensal:**
   - perda do prêmio COPR com qualquer falta no mês;
   - desconto proporcional de vale-alimentação só para falta sem justificativa;
   - desconto de DSR (cód. 8794) por semana ISO com falta injustificada de dia inteiro, lançado em dias ou em valor (R$).
5. **Calendário de feriados offline (2025–2199):** feriado não conta como falta, e trabalho em feriado vira HE 100%. Há um gerador que mescla dados novos sem sobrescrever ajustes do RH e confere os feriados nacionais contra a BrasilAPI.
6. **Banco de horas:** integra o relatório do Secullum e uma planilha semanal paralela. Isola quem está no banco para não pagar a mesma hora extra duas vezes, e a checagem pré-voo bloqueia saldo para quem não está no banco.
7. **Fila de exceções para o RH:** o sistema **nunca decide sozinho** se uma falta é justificada. Casos ambíguos vão para uma fila com o motivo explicado. Sem decisão humana, o padrão é conservador (falta injustificada).
8. **Saída:** um `.xlsx` por empresa no layout exigido pela contabilidade ("Relação de Valores"), com a aba "Contabilidade" e a aba "Conferência RH" (destaques visuais, comparação com os totais do Secullum).

## 4. Fluxo de uso (aplicação web)

Upload por arrastar-e-soltar, com detecção automática do tipo de arquivo pelas colunas → período e checagem pré-voo (valida colunas e PIS/matrícula antes do processamento pesado) → processamento do ponto → tabela de aprovação com dropdown de decisão por linha → geração e download do Excel por empresa (ou `.zip`).

## 5. Stack e tecnologias

- **Linguagem:** Python 3 (type hints modernos).
- **Bibliotecas:** pandas, openpyxl, pdfplumber, Streamlit.
- **Banco de dados:** SQLite, para persistir sessão e decisões do RH por competência (evita perder o progresso ao atualizar a página).
- **Interfaces:** web com Streamlit (principal) e desktop com tkinter (legado, empacotável em `.exe` com PyInstaller).
- **Testes:** pytest, com dados sintéticos de ponta a ponta e testes de banco de horas, DSR, falta parcial, feriados e exportação de exceções.
- **Ferramentas:** Git/GitHub, VS Code, Claude Code (desenvolvimento assistido por IA).

## 6. Destaques técnicos (o que diferencia o projeto)

- **Arquitetura em camadas:** toda a lógica de cálculo fica em módulos separados (`business_rules`, `pipeline`, `banco_horas`, `feriados`), e as telas só chamam o pipeline. Assim, web e desktop compartilham as mesmas regras sem duplicação.
- **Parser configurável:** o layout do AFD fica numa única constante, e a mudança de versão de relógio exige ajuste em um só lugar.
- **Validação defensiva:** uma exceção própria (`ArquivoInvalidoError`) traduz qualquer erro de arquivo em mensagem em português que aponta o arquivo, a linha e o PIS. O usuário nunca vê traceback.
- **Identificação flexível:** aceita PIS ou matrícula como chave de colaborador e normaliza para o PIS canônico.
- **Segurança de decisão:** o padrão conservador e a ausência de fechamento silencioso de exceções reduzem o risco trabalhista.
- **Tratamento de dados sensíveis:** dados reais (PIS, folha) ficam fora do versionamento via `.gitignore`.
- **Premissas de negócio explícitas:** o que ainda não foi confirmado com o cliente fica em constantes configuráveis e documentadas (ex.: unidade do DSR, regra do banco de horas).

## 7. Tamanho do projeto (verificável no repositório)

- Cerca de 6.400 linhas de Python em 18 arquivos, das quais cerca de 1.300 são testes automatizados.
- 22 commits.

## 8. Resultados e impacto

`[PREENCHER com números reais, se existirem. Exemplos do que informar: horas gastas por fechamento antes e depois; nº de colaboradores processados por mês; nº de empresas; redução de erros.]`

Sem esses números, descreva o impacto apenas de forma qualitativa: eliminou digitação manual de apontamentos, substituiu a aprovação em papel por aprovação digital e padronizou a entrega à contabilidade.

## 9. Sugestões de bullets prontos para o currículo

**Versão curta (1 linha):**
- Desenvolvi em Python/Streamlit um sistema de automação de fechamento de folha que apura ponto eletrônico (AFD/PDF), aplica regras da CLT e gera planilhas por empresa para a contabilidade.

**Versão em bullets (3 a 5 linhas):**
- Automatizei o fechamento mensal de folha de pagamento (Python, pandas, Streamlit, SQLite), substituindo apontamentos manuais e aprovação em papel por um fluxo digital.
- Implementei regras trabalhistas da CLT: tolerância do Art. 58, horas extras 50%/100%, DSR, faltas parciais, banco de horas e calendário de feriados offline (2025–2199).
- Projetei uma fila de exceções com validação humana e padrão conservador, evitando decisões automáticas sobre faltas e reduzindo risco trabalhista.
- Construí parsers para AFD (Portaria 671/2021) e PDFs de ponto (pdfplumber), com validação defensiva e mensagens de erro claras para usuários não técnicos.
- Estruturei o código em camadas e criei uma suíte de testes com pytest (dados sintéticos de ponta a ponta), reutilizando a mesma lógica em interface web e desktop.

**Palavras-chave (ATS):** Python, pandas, Streamlit, SQLite, openpyxl, pdfplumber, pytest, automação de processos, folha de pagamento, CLT, ponto eletrônico, AFD, DSR, banco de horas, regras de negócio, ETL, Excel, Git.

## 10. Habilidades demonstradas

Levantamento de requisitos com áreas de RH e contabilidade, modelagem de regras de negócio trabalhistas, ETL de arquivos heterogêneos (TXT, PDF, XLSX, CSV), testes automatizados, design de UX para usuários não técnicos e documentação técnica em português.
