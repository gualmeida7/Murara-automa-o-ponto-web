# Guia do Fechamento Mensal da Folha

**Para que serve.** Uma vez por mês você usa uma página na internet (o "sistema") para fechar a folha das empresas **Maçaneiro e Gonzaga LTDA** e **Cianorte Tubos de Concretos LTDA**. O sistema lê o ponto e calcula horas extras, faltas, prêmio de assiduidade (COPR), vale-alimentação (VA) e descanso semanal (DSR). Ele só pergunta para você os dias que não consegue decidir sozinho. No final, entrega **uma planilha de Excel por empresa** para a contabilidade. O período normal é **do dia 26 de um mês ao dia 25 do mês seguinte**.

## Lista do mês

- [ ] No Secullum, exportar o **Cartão Ponto** (PDF) ou o **AFD** (.txt) do período.
- [ ] No Secullum, exportar o **Espelho de Ponto** (PDF) do mesmo período (opcional, mas ajuda).
- [ ] Separar os **extratos de consignado**, um por empresa (opcional).
- [ ] Separar os arquivos de **banco de horas**, só se alguém usar.
- [ ] Atualizar o **Cadastro de Funcionários** (admissões, desligamentos, mudanças).
- [ ] Abrir o sistema e seguir as etapas 1 a 5 abaixo.

## Os arquivos

| Arquivo | Obrigatório? | De onde vem | Onde enviar | Quando mudar |
|---|---|---|---|---|
| Cadastro de Funcionários (ex.: `cadastro_sintetico.xlsx`) | **Sim** | Você mesma mantém | Etapa 1 | Sempre que um funcionário mudar |
| Cartão Ponto (PDF, ex.: `Cartão Ponto 09.2026.pdf`) **ou** AFD (.txt) | **Sim** (um dos dois) | Secullum, do período | Etapa 1 | Todo mês |
| Espelho de Ponto (PDF, ex.: `Espelho de Ponto 09.2026.pdf`) | Não (recomendado) | Secullum, mesmo período | Etapa 1 | Todo mês |
| Extratos de consignado (ex.: `27488404000122-202609v1.4.xlsx` para a Cianorte Tubos e `17155411000159-202609v1.4.xlsx` para a Maçaneiro e Gonzaga; o número no início é o CNPJ da empresa) | Não | Banco/financeira, um por empresa | Etapa 1 | Todo mês |
| Banco de horas | Só se alguém usar | Exportação do Secullum ou planilha semanal | Etapa 1 | Todo mês |

### Como preencher o Cadastro de Funcionários

Uma linha por funcionário. Os **nomes das colunas precisam ser exatamente estes**:

| Coluna | O que escrever | Exemplo |
|---|---|---|
| pis | PIS com 11 números | 12345678901 |
| matricula | Número do funcionário (o mesmo do Secullum e do consignado) | 1001 |
| nome | Nome | Maria Exemplo |
| codigo_empresa | Código da empresa | 1 |
| razao_social | Nome da empresa | Cianorte Tubos de Concretos LTDA |
| entrada, saida_almoco, retorno_almoco, saida | Horários, no formato HH:MM | 07:30, 11:30, 12:30, 17:30 |
| dias_trabalho | Dias de trabalho em números, separados por vírgula: 0=segunda … 6=domingo | 0,1,2,3,4 (segunda a sexta) |
| banco_horas *(opcional)* | sim/não: usa banco de horas | não |
| registra_ponto *(opcional)* | sim/não: bate ponto. "não" = fica fora do cálculo de faltas | sim |
| valor_dia_dsr *(opcional)* | Valor em R$ de um dia de descanso | 85,00 |
| saida_sexta *(opcional)* | Horário de saída na sexta, quando for diferente dos outros dias | 16:30 |
| admissao *(opcional)* | Data de admissão (dd/mm/aaaa). Dias antes dela não contam como falta, e o vale do mês de admissão sai proporcional | 14/10/2026 |
| cnpj *(opcional)* | CNPJ da empresa, para o cabeçalho da planilha do escritório | 17.155.411/0001-59 |

Colunas opcionais usadas só na **Planilha do Thales** (em branco = célula vazia):

| Coluna | O que escrever | Exemplo |
|---|---|---|
| centro_custo | CC, como você escreve hoje | 07 e 25 |
| valor_va | Valor do vale do mês (Alelo/Ticket) em R$ | 960 |
| beneficio | Nome do vale no título da empresa | Ticket |
| cooper | Valor da Cooper em R$ | 200 |
| desconto_folha | Desconto em folha do mês em R$ (vai também no "Adiant.Salario" da planilha do escritório) | 224 |
| observacao | Texto livre para a coluna Observações | *10X R$ 224,00 VENDA 21779 (Parc 07) |
| nota_rodape | Nota escrita abaixo da tabela da empresa, como "Nome: texto" (o sistema coloca o primeiro nome) | atestado em 28/05, detenção desde 24/06 |

`desconto_folha`, `observacao` e `nota_rodape` mudam todo mês: atualize antes de processar.

## Passo a passo

**1. Abrir o sistema.** No navegador, entre no endereço que o responsável de TI informar (normalmente `http://localhost:8501`). [Print: tela inicial]

**2. Etapa 1 – Arquivos.** Arraste **todos** os arquivos do mês de uma vez para a área de envio. O sistema reconhece sozinho o tipo de cada arquivo. Se ele errar, corrija na lista ao lado do arquivo. Depois clique em **"Confirmar arquivos e continuar"**. [Print: área de envio]

**3. Etapa 2 – Período e validação.** Escolha a data inicial e a final (normalmente 26 a 25) e clique em **"Validar dados"**.
Se aparecer um PIS ou matrícula que não está no cadastro, o sistema mostra **o arquivo e o número exatos**. Corrija no Cadastro de Funcionários e envie os arquivos de novo. [Print: aviso de validação]

**4. Etapa 3 – Processamento.** Clique em **"Processar ponto e calcular fechamento"**. Marque a caixa dos feriados facultativos (Carnaval, Corpus Christi, 24 e 31 de dezembro) **somente se a empresa der esse dia de folga**. [Print: botão de processar]

> **Aviso possível: "colaboradores estão no cadastro mas não têm nenhuma batida no ponto".**
> Se você não fizer nada, o período inteiro deles vira falta. Selecione as pessoas e escolha uma ação:
>
> | Opção | Quando usar | O que acontece |
> |---|---|---|
> | **Desligado: tirar do fechamento** | A pessoa foi desligada | Sai deste fechamento |
> | **Não registra ponto** | A pessoa não bate ponto | Fica no relatório com zeros e COPR "Conferir" |
>
> Depois clique em **"Aplicar e reprocessar"** e clique em **processar** de novo. **Atenção:** essa escolha vale só enquanto você está com o sistema aberto. Depois, **atualize também o Cadastro de Funcionários**, senão o aviso volta no mês que vem. [Print: aviso]

**5. Etapa 4 – Revisão de exceções.** Aparece uma tabela só com os dias que o sistema não soube decidir. Em cada linha, escolha:

| Decisão | Significa |
|---|---|
| Falta injustificada | Faltou sem justificativa |
| Falta justificada | Faltou com justificativa |
| Erro de relógio — abonar | O relógio falhou; o dia é perdoado |
| Atestado médico | Tem atestado |

Linha em branco conta como falta injustificada (é a opção mais segura). Seu progresso é salvo sozinho como rascunho, então você pode voltar depois. [Print: tabela de exceções]

**6. Etapa 5 – Download.** Clique em **"Gerar fechamento final"**. Baixe os arquivos (botão **"Baixar"**) ou tudo junto (**"Baixar todos os arquivos (.zip)"**). Saem: **uma "Relação de Valores" por empresa**, igual à planilha que você envia ao escritório (horas no formato hh.mm, como você digita; aba **"Contabilidade"** para enviar e aba **"Conferência RH"** para você conferir), e a **Planilha do Thales**, com o resumo das duas empresas. O banco de horas (0999) sai em branco na planilha do escritório, como você faz hoje; o saldo aparece nas Observações da planilha do Thales.

**Dia em que o relógio falhou para todos:** na Etapa 2, preencha **"Dias abonados para todos"** (ex.: `14/10/2026, 15/10/2026`) antes de processar. Quem não fechou as 4 batidas nesses dias não leva falta nem perde DSR, VA ou COPR. [Print: botões de download]

## Quando os funcionários mudam

| Situação | O que fazer |
|---|---|
| Funcionário novo | Acrescente uma linha no cadastro, com o mesmo PIS e matrícula do Secullum e do consignado |
| Funcionário desligado | Apague a linha do cadastro **antes do próximo fechamento**. Se ele ainda aparecer no ponto mas não no cadastro, o sistema ignora e só mostra um aviso (não dá erro). Se continuar no cadastro sem batidas, use o aviso da etapa 3 |
| Mudou horário ou dias de trabalho | Edite os horários e `dias_trabalho` |
| Entrou ou saiu do banco de horas | Mude `banco_horas` |
| Funcionário que não bate ponto | Coloque `registra_ponto = não` |
| Novo empréstimo consignado | Nada a mudar no cadastro; ele vem no extrato do mês |
| Mudou de empresa | Atualize `codigo_empresa` e `razao_social` |

## O que o sistema já faz por você

> - Feriado **não** é falta. Trabalhar em feriado ou domingo paga **hora extra de 100%**.
> - **Qualquer** falta no mês, mesmo justificada, faz perder o prêmio **COPR**.
> - Só a falta **injustificada** desconta o vale-alimentação (VA) e o descanso semanal (DSR).
> - O sistema **nunca** decide sozinho se uma falta é justificada. Essa decisão é sempre sua.
> - Evite atualizar a página (F5) ou fechar a aba no meio do trabalho. O rascunho fica salvo, mas é mais seguro terminar a revisão de uma vez.

## Problemas e soluções

| Problema | O que fazer |
|---|---|
| A página não abre | Veja se o sistema está ligado (peça ao responsável de TI) e se o endereço está certo |
| "PIS/matrícula não encontrado no cadastro" | Anote o arquivo e o número mostrados, corrija o cadastro (ou o número) e envie de novo |
| Um arquivo foi reconhecido como tipo errado | Troque o tipo na lista ao lado do arquivo |
| Erro no extrato de consignado | Uma matrícula do extrato não está no cadastro. Inclua ou corrija e envie de novo |
| O cadastro é recusado | Confira se os nomes das colunas estão exatamente como na tabela, sem espaços sobrando |
| Alguém aparece com o mês inteiro de faltas | Provavelmente não tem batidas. Use o aviso da etapa 3 e depois corrija o cadastro |
| Um desligado aparece no relatório | Tire-o do cadastro (ou use "Desligado" no aviso da etapa 3) |
| A página foi atualizada ou fechada sem querer | Abra o sistema de novo; suas decisões da etapa 4 foram salvas como rascunho |
| Qualquer outro problema | Anote o que estava fazendo e o que a tela mostrou (ou tire um print) e envie ao responsável de TI |
