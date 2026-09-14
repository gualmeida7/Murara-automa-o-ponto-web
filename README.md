# Automação do fechamento de folha — Maçaneiro e Gonzaga / Cianorte Tubos

Esta pasta contém uma implementação de referência, testada de ponta a ponta
com dados sintéticos, para os 5 pontos pedidos. Ela **não é plug-and-play**
com os arquivos reais do Secullum sem um ajuste inicial de layout (ponto 0
abaixo) — mas o pipeline inteiro roda e produz a Relação de Valores no
formato exigido pela contabilidade.

## Arquivos

| Arquivo | Papel |
|---|---|
| `afd_parser.py` | Lê o `.txt` AFD e extrai as batidas de ponto. |
| `business_rules.py` | Cruza batidas x jornada contratual; calcula HE 50/100, atraso, falta; aplica tolerância CLT, COPR, VA e DSR; separa o que precisa de validação humana. |
| `banco_horas.py` | Integra o banco de horas do Sidney (relatório Secullum) e do Adriano (planilha semanal paralela) sem afetar os demais colaboradores. |
| `leitura_arquivos.py` | Leitura única e resiliente de qualquer arquivo de entrada (.xlsx/.xls/.csv), com detecção automática de delimitador e codificação para CSV. |
| `validacao.py` | Validação defensiva compartilhada — colunas obrigatórias e checagem pré-voo de PIS/matrícula — com a exceção amigável `ArquivoInvalidoError`. |
| `pipeline.py` | Orquestra tudo e gera um arquivo `.xlsx` por empresa no layout da "Relação de Valores". |
| `dashboard_aprovacao.html` | Painel de aprovação — abra no navegador, carregue o CSV de exceções, classifique, exporte a decisão. Não depende de instalar nada. |
| `teste_sintetico.py` | Gera dados fictícios e roda o pipeline inteiro — use como prova de conceito e como referência do formato de cada planilha de entrada. |
| `app_gui.py` | Interface gráfica (tkinter) para quem não programa: botões para selecionar os arquivos e um botão "Processar Folha de Pagamento". Ver seção "Aplicativo visual (desktop)" abaixo. |
| `app_web.py` | Versão 100% web (Streamlit), numa única tela: upload arrastar-e-soltar, checagem pré-voo instantânea, tabela de aprovação embutida e download direto do Excel — sem nenhum arquivo intermediário. Ver seção "Aplicativo web (Streamlit)" abaixo. |
| `requirements_web.txt` | Dependências para rodar `app_web.py` (`pip install -r requirements_web.txt`). |
| `COMO_GERAR_EXE.md` | Passo a passo para empacotar `app_gui.py` num único `.exe` do Windows com PyInstaller. |

## Aplicativo web (Streamlit) — recomendado para uso diário

`app_web.py` é a versão mais recente e a que recomendamos para o dia a dia:
tudo acontece numa página só, no navegador, sem nenhum arquivo solto para
gerenciar (nada de `fila_validacao_rh.csv`, `dashboard_aprovacao.html` ou
`decisoes_rh.csv` — esses continuam existindo no projeto só para quem
preferir o fluxo desktop antigo, descrito na seção seguinte).

### Como rodar na sua máquina

```bash
pip install -r requirements_web.txt
streamlit run app_web.py
```

O terminal mostra algo como `Local URL: http://localhost:8501` — abra esse
endereço no navegador (Chrome, Edge, o que preferir). Enquanto o terminal
ficar aberto, a página funciona; feche o terminal (ou `Ctrl+C`) para
encerrar. Não precisa de internet — tudo roda localmente no computador.

### O fluxo, na ordem que aparece na tela

1. **Enviar os arquivos** — uma única área de arrastar-e-soltar aceita
   Cadastro, AFD, Consignados e os bancos de horas de uma vez só. O
   sistema identifica sozinho o tipo de cada arquivo pelas colunas que
   encontra (a mesma lógica de detecção já usada em `banco_horas.py` e
   `validacao.py`) e mostra um dropdown para corrigir manualmente, caso
   algum arquivo saia classificado errado.
2. **Período e checagem pré-voo** — assim que os arquivos são confirmados,
   um clique já roda a validação (colunas obrigatórias + todo
   PIS/matrícula citado existe no cadastro). Se algo não bate, a página
   mostra o erro na hora, apontando o arquivo e o PIS exatos, e trava a
   continuação — sem gastar tempo processando o AFD inteiro primeiro.
3. **Processar o ponto** — só essa etapa (a mais pesada) roda sob demanda,
   num botão separado, para a página não reprocessar o AFD inteiro a cada
   pequena interação com a tela.
4. **Tabela de aprovação embutida** — as exceções aparecem numa tabela
   interativa, com o nome do colaborador já cruzado e um dropdown de
   decisão (Falta injustificada / Falta justificada / Erro de relógio —
   abonar) direto na linha. Cada edição fica salva em
   `st.session_state` — não existe mais um CSV para baixar, classificar
   em outro programa e reimportar.
5. **Gerar e baixar** — um clique consolida tudo em memória e libera o
   download do `.xlsx` final por empresa (ou de todos juntos, num `.zip`),
   já com as abas "Contabilidade" e "Conferência RH" formatadas.

### Sobre o estado da sessão (o que perde e o que não perde)

Como tudo fica em `st.session_state`, alternar entre as seções da página,
demorar para preencher a tabela de exceções ou clicar em outros elementos
NÃO reinicia o progresso — o Streamlit reexecuta o script a cada
interação, mas o conteúdo do `session_state` sobrevive entre essas
execuções. A única coisa que reinicia o fechamento do zero é **atualizar a
página (F5) ou fechar a aba do navegador** — isso é uma limitação inerente
de como o Streamlit guarda sessão (não existe um "salvar rascunho" nessa
versão). Se isso for um problema na prática, dá para adicionar depois uma
gravação periódica do `session_state` em disco (ex.: um `st.session_state`
espelhado em SQLite) — pergunte se quiser que eu implemente isso.

### Uma decisão de arquitetura que vale explicar

Os arquivos enviados por arrastar-e-soltar chegam ao Streamlit como bytes
em memória. Em vez de reescrever `afd_parser.py`, `pipeline.py` e
`banco_horas.py` para ler direto desses bytes (duplicando e arriscando
divergir da lógica já testada), `app_web.py` grava cada upload numa pasta
temporária exclusiva da sessão do navegador e reaproveita, sem nenhuma
alteração, as mesmas funções de carregamento usadas pelo `app_gui.py`. Do
ponto de vista de quem usa, isso é invisível — nenhum arquivo aparece para
baixar, procurar ou apagar manualmente; a pasta é só um detalhe interno de
implementação.

## Aplicativo visual (desktop, para quem não programa)

`app_gui.py` é uma janela tkinter que roda por cima dos mesmos módulos
acima — nenhuma lógica de cálculo foi duplicada. Ela pede:

- **Obrigatórios**: Cadastro de Colaboradores, Arquivo AFD (Ponto),
  Extrato de Consignados. (O cadastro foi adicionado em relação ao pedido
  original porque é dele que vem a jornada e a empresa de cada
  colaborador — sem ele não há como calcular nada.)
- **Opcionais**: Banco de Horas do Secullum (Sidney) e a planilha semanal
  do Adriano (pedindo o PIS/matrícula dele ao lado, já que a planilha dele
  não tem essa informação).
- **Período do fechamento**, pré-preenchido com o padrão 26 a 25, editável.
- **Pasta de saída**, com um padrão sugerido, editável.

Um único clique em "Processar Folha de Pagamento" roda 6 etapas, cada uma
atualizando uma barra de progresso e um texto de status ("Etapa 3 de 6:
Validando dados (checagem pré-voo)...") para a usuária nunca ficar
olhando pra uma tela parada sem saber se travou. Ao final, gera na pasta
escolhida a Relação de Valores por empresa **e** `fila_validacao_rh.csv`
com as exceções (que continuam com a classificação conservadora "falta
injustificada" até serem revisadas em `dashboard_aprovacao.html` — o
aviso de quantas pendências existem, e quantos arquivos foram gerados e
onde, aparece na caixa de mensagem final).

A checagem pré-voo (etapa 3) roda **antes** da etapa mais pesada (ler e
apurar o AFD inteiro): ela confere se todo PIS/matrícula citado nos
arquivos de banco de horas e consignados existe de fato no Cadastro de
Colaboradores. Um PIS digitado errado aparece em segundos, com uma
mensagem apontando exatamente qual PIS e qual arquivo — em vez de a
usuária esperar o processamento inteiro (que pode levar um tempo com um
AFD grande) só para descobrir o erro no fim, ou pior, a folha de alguém
sair sem o banco de horas dela silenciosamente. Qualquer arquivo com
coluna faltando ou formato inesperado também é pego já no carregamento
e mostrado numa `messagebox` em português claro, nunca um traceback
técnico — inclusive erros verdadeiramente inesperados mostram só a causa
mais provável, sugerindo levar ao suporte técnico se persistir.

Para virar um `.exe` que roda em qualquer Windows sem precisar instalar
Python, siga `COMO_GERAR_EXE.md`.

## 0. Antes de usar com dados reais (passo obrigatório)

O layout binário do AFD (posição exata de cada campo dentro da linha) varia
por versão de portaria e de REP. `afd_parser.py` já vem configurado para o
layout mais comum (Portaria 671/2021, registro tipo 3), mas **abra um AFD
real do seu Secullum e confira**: a constante `LAYOUT_TIPO3` no topo do
arquivo concentra essa configuração de propósito, para ser o único lugar
que precisa de ajuste. O mesmo vale para os nomes de coluna esperados no
cadastro de colaboradores, no extrato de consignados e nas planilhas de
banco de horas — todos documentados no docstring de cada função em
`pipeline.py` e `banco_horas.py`.

## 1. Parser do AFD → cálculo automático de horas

`afd_parser.py` lê o arquivo e devolve uma batida por linha
(`pis, data_hora`). `business_rules.py` agrupa essas batidas por dia,
compara com a `Jornada` contratual de cada colaborador (cadastrada uma vez
por pessoa) e calcula, sem digitação:

- **Horas extra 50%** — minutos trabalhados além da carga horária diária,
  em dia útil.
- **Horas extra 100%** — qualquer trabalho em dia não útil (domingo, e
  você pode injetar feriados na mesma lógica).
- **Atraso** — minutos faltando para a carga horária, respeitando a
  tolerância configurável (10 min por padrão, ajustável por colaborador).
- **Falta candidata** — dia útil sem nenhuma batida.
- **DSR** — calculado na consolidação mensal, proporcional às semanas com
  falta injustificada.

Isso substitui as etapas "apontamentos manuais a lápis" e "consolidação
gerencial" digitada.

### Tolerância da CLT (Art. 58 §1º)

Duas camadas, aplicadas juntas:
- **5 minutos por marcação** — cada uma das 4 batidas do dia é comparada
  com o horário programado; se o desvio for menor ou igual a 5 minutos,
  a batida é tratada como se fosse exatamente no horário (não gera
  minuto de HE nem de atraso). Só o que exceder os 5 minutos entra na
  conta.
- **10 minutos agregados no dia** — mesmo depois de aplicar a tolerância
  por marcação, o saldo do dia só vira hora extra (0150) ou atraso se
  ultrapassar 10 minutos no total.

Ambos os limites são configuráveis por colaborador (`Jornada.tolerancia_por_batida_min`
e `Jornada.tolerancia_min`), caso a convenção coletiva da empresa preveja
valores diferentes.

## 2. Tratamento de inconsistências (o RH só valida exceções)

`business_rules.apurar_dia()` nunca decide sozinho se um dia sem batida é
falta real ou erro de relógio — ele **sinaliza** (`precisa_validacao_rh =
True`) e explica o motivo (`motivo_validacao`). `gerar_fila_validacao_rh()`
filtra só essas linhas. Hoje, os casos sinalizados são:

- Dia útil sem nenhuma batida (`FALTA_OU_AUSENCIA`)
- Número de batidas diferente de 2 ou 4 (`BATIDA_INCOMPLETA`)
- Intervalo de almoço menor que 60 minutos (`INTERVALO_CURTO`)

Essa fila é exportada como `fila_validacao_rh.csv` e carregada direto no
`dashboard_aprovacao.html` — o RH só vê essas linhas, não o arquivo
inteiro. Por segurança, uma exceção sem decisão humana **nunca fecha
automaticamente a favor do colaborador nem da empresa**: ela entra como
"falta injustificada" até ser revisada (ponto de partida conservador,
nunca silencioso).

O CSV é gerado por `pipeline.exportar_fila_validacao_rh()`, que:
- grava em `utf-8-sig` — o Excel abre os acentos direto, sem pedir
  importação de texto;
- converte a data para `dd/mm/aaaa` (padrão brasileiro);
- cruza com o cadastro e adiciona `nome_colaborador` e `razao_social`,
  para o RH não precisar abrir outro arquivo para saber quem é quem
  (o painel de aprovação também já mostra o nome, não só o PIS);
- ordena primeiro por nome do colaborador, depois cronologicamente pela
  data (a ordenação por data acontece antes da conversão para texto, ou
  "05/12" apareceria antes de "20/01" numa ordenação alfabética).

### Validação defensiva dos arquivos de entrada

Antes de processar qualquer coisa, `carregar_cadastro_colaboradores()`,
`carregar_consignados()`, `ler_banco_horas_secullum()` e
`ler_banco_horas_adriano()` conferem se todas as colunas obrigatórias
estão presentes (via `validacao.verificar_colunas()`). Se faltar uma
coluna — ou o arquivo nem abrir —, o processo para com uma
`ArquivoInvalidoError` dizendo exatamente qual arquivo e qual coluna
está faltando, em vez de um `KeyError` genérico no meio do cálculo.
Como `ArquivoInvalidoError` é uma subclasse de `ValueError`, o
`app_gui.py` já exibe essa mensagem numa `messagebox` sem precisar de
nenhum ajuste adicional.

## 3. Regras de negócio (COPR e Vale Alimentação)

Em `consolidar_mes()`:

- **COPR**: `perde_copr = dias_falta > 0` — qualquer falta no mês, mesmo
  classificada como `FALTA_JUSTIFICADA`, zera a flag. Isso está isolado em
  uma única linha de código de propósito, para ficar fácil de auditar ou
  ajustar se a política mudar.
- **Vale Alimentação**: `dias_desconto_va` conta **só** os dias
  classificados como `FALTA_INJUSTIFICADA` — faltas justificadas não
  descontam o benefício, seguindo a regra que você descreveu.

## 4. Banco de horas manual do Adriano (sem corromper os demais)

`banco_horas.py` trata as duas fontes como **exceções isoladas**:

- `ler_banco_horas_secullum()` lê o relatório de banco de horas que já
  existe dentro do próprio Secullum (cobre o "Sidney").
- `ler_banco_horas_adriano()` lê a planilha semanal paralela dele, soma
  horas extras menos débito de cada semana e soma o saldo acumulado do mês
  anterior.
- `consolidar_banco_horas()` faz um `LEFT MERGE` dessas duas fontes contra
  a lista completa de colaboradores — quem não está em nenhuma das duas
  simplesmente fica com saldo `0`, e um `ValueError` explícito impede que a
  mesma pessoa apareça em duas fontes ao mesmo tempo (evita duplicar saldo
  por engano).

## 5. Aprovação digital e exportação por empresa

- **Aprovação**: `dashboard_aprovacao.html` é um arquivo único, sem
  backend — o gestor abre no navegador, carrega o CSV de exceções, escolhe
  para cada linha "Falta injustificada" / "Falta justificada" / "Erro de
  relógio — abonar", opcionalmente anota uma observação, e baixa
  `decisoes_rh.csv`. Isso substitui a planilha impressa marcada à caneta.
- **Exportação por empresa**: `pipeline.exportar_por_empresa()` agrupa a
  Relação de Valores consolidada por `codigo_empresa` + `razao_social` e
  grava um `.xlsx` separado por empresa (Maçaneiro e Gonzaga / Cianorte
  Tubos), com **duas abas**:
  - **"Contabilidade"** — estritamente as colunas do layout exigido
    (Código Empresa, Razão Social, Competência, Código Folha, Nome do
    Colaborador, 0150, 0200, 8069, 8792, 8794, 0981, 0999), sem nenhuma
    coluna auxiliar interna.
  - **"Conferência RH"** — as mesmas colunas + Prêmio COPR
    (Perde/Não Perde), Dias de Falta, Dias Desconto VA e Semanas com
    Perda de DSR, com linhas de quem perde o prêmio ou tem falta elevada
    (3+ dias, ajustável em `LIMITE_FALTAS_ALTO`) destacadas em vermelho
    suave.

  Formatação aplicada em ambas as abas: cabeçalho verde corporativo
  (#1B4D3E) com texto branco em negrito, congelamento da linha 1, bordas
  finas cinza-claro em todas as células, alinhamento à esquerda para
  texto, centralizado para códigos/matrícula/competência e à direita
  para valores numéricos, com 2 casas decimais explícitas nas colunas de
  horas e valores.

## Como rodar (ordem real de uso no fechamento)

```bash
pip install pandas openpyxl

# 1. Gera a fila de exceções a partir do AFD do período (26 a 25, por ex.)
python pipeline.py          # gera saida/fila_validacao_rh.csv

# 2. Abra dashboard_aprovacao.html no navegador, carregue esse CSV,
#    classifique as exceções e baixe decisoes_rh.csv em saida/

# 3. Rode pipeline.py novamente (ele detecta o decisoes_rh.csv e já
#    gera a Relação de Valores final por empresa em saida/)
python pipeline.py
```

`teste_sintetico.py` reproduz esse fluxo inteiro com dados fictícios, útil
para validar o ambiente antes de apontar para os arquivos reais.

## O que ainda depende de decisão sua antes de ir para produção

1. Confirmar o layout exato do AFD do seu Secullum (ponto 0 acima).
2. Confirmar os nomes de coluna reais do cadastro de colaboradores, do
   extrato de consignados e das planilhas de banco de horas — os
   `ValueError` explícitos em `pipeline.py`/`banco_horas.py` avisam
   imediatamente se um nome não bater, em vez de gerar um número errado
   silenciosamente.
3. Definir a fonte do valor de um dia de DSR por colaborador
   (`valor_dia_dsr_por_pis` em `pipeline.py`) — isso normalmente vem do
   sistema de folha em si, fora do escopo do ponto.
4. Regra de feriados: hoje só domingo é tratado como "dia não útil" com HE
   100%; um calendário de feriados pode ser adicionado como uma lista de
   datas extra em `apurar_dia()`.
