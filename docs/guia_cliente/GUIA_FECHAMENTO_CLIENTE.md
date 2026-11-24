# Guia do Fechamento Mensal da Folha

**Para que serve.** Uma vez por mês você usa uma página na internet (o "sistema") para fechar a folha das empresas **Maçaneiro e Gonzaga LTDA** e **Cianorte Tubos de Concretos LTDA**. O sistema lê o ponto e calcula horas extras, faltas, prêmio de assiduidade (COPR), vale-alimentação (VA) e descanso semanal (DSR). Ele só pergunta para você os dias que não consegue decidir sozinho. No final, entrega **uma planilha de Excel por empresa** para a contabilidade. O período normal é **do dia 26 de um mês ao dia 25 do mês seguinte**.

## Lista do mês

- [ ] No Secullum, exportar o **Espelho de Ponto** (PDF) do período. Ele sozinho basta. Se tiver o **Cartão Ponto** (PDF) ou o **AFD** (.txt), pode enviar também.
- [ ] Separar os **extratos de consignado**, um por empresa (opcional).
- [ ] Separar os arquivos de **banco de horas**, só se alguém usar.
- [ ] Atualizar o **Cadastro de Funcionários** (admissões, desligamentos, mudanças).
- [ ] Abrir o sistema e seguir as etapas 1 a 5 abaixo.

## Os arquivos

| Arquivo | Obrigatório? | De onde vem | Onde enviar | Quando mudar |
|---|---|---|---|---|
| Cadastro de Funcionários (modelo: `Cadastro_colaboradores_11_2026.xlsx`) | **Sim** | Você mesma mantém | Etapa 1 | Sempre que um funcionário mudar |
| Espelho de Ponto (PDF, ex.: `Espelho de Ponto 11.26.pdf`), **ou** Cartão Ponto (PDF), **ou** AFD (.txt) | **Sim** (pelo menos um dos três) | Secullum, do período | Etapa 1 | Todo mês |
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

Colunas que a tela entende além das acima (também podem ser marcadas diretamente na tabela):

| Coluna | O que escrever | Exemplo |
|---|---|---|
| va_fixo | sim/não: vale fixo, não sofre desconto por falta | sim |
| cooper_fixo | sim/não: Cooper fixa todo mês | sim |
| va_desconto_manual | sim/não: o desconto do vale é feito manualmente | não |

**Use a planilha-modelo (`Cadastro_colaboradores_11_2026.xlsx`).** Ela já vem com os dados de novembro e com as três
colunas acima preenchidas com **sim** ou **não**. Para marcar uma exceção, escreva **sim** na coluna; **não pinte
a célula**, porque a cor se perde quando a planilha é salva ou copiada por outro programa. O sistema ainda entende
as cores da planilha antiga (verde forte = valor fixo; verde claro = desconto manual), mas só no envio.

As colunas de valor são três: **Cooper Card Assiduidade** (valor da Cooper), **Alelo Alimentação** (vale da empresa
001, Maçaneiro e Gonzaga) e **Ticket Alimentação** (vale da empresa 002, Cianorte Tubos). Cada colaborador tem valor só
na coluna da sua empresa; a outra fica em branco ou zero.

**Quando marcar cada coluna**

| Situação | Coluna a marcar com "sim" |
|---|---|
| A Cooper é um valor fixo todo mês, que não é perdido por faltas | cooper_fixo |
| O vale é um valor fixo, sem desconto por falta | va_fixo |
| O desconto do vale por falta será feito à mão | va_desconto_manual |

Quem não está marcado recebe a Cooper e o vale pelo valor da planilha, com as regras de desconto descritas mais abaixo.

### Revisar o cadastro dentro do sistema

Ao confirmar os arquivos, o cadastro aparece numa tabela na própria tela. Nada é alterado na sua planilha original.

- Corrija qualquer célula, acrescente linhas no fim da tabela ou apague linhas. Use os botões acima da tabela
  (Todas, Identificação, Jornada, Benefícios, Opcionais) para ver menos colunas por vez.
- A coluna **Situação** mostra **OK**, **Aviso** ou **Erro** para cada colaborador. O quadro logo abaixo da tabela
  explica cada ocorrência, sempre indicando a **matrícula**.
- **Erro** impede o fechamento: o botão "Validar dados" só libera quando não houver nenhum erro.
- **Aviso** não impede; serve para você conferir.
- No alto do quadro, a **Conferência da leitura do cadastro** resume quantos colaboradores usam Alelo ou Ticket em
  cada empresa, quantos têm Cooper, quantos estão marcados como vale fixo, Cooper fixo ou vale com desconto manual, e
  como as datas de admissão foram lidas (com exemplos em dd/mm/aaaa). Confira esses números antes de seguir.
- **Valores fixos e desconto manual:** com a planilha-modelo, o sistema lê as colunas **Vale fixo**, **Cooper fixo** e
  **Vale c/ desconto manual** (sim/não). Se enviar uma planilha antiga que só tem cores, ele converte as cores nessas
  colunas. Se ela não tiver nem colunas nem cores, a tela avisa e você marca as três colunas diretamente na tabela.
  Confira sempre o quadro da **Conferência da leitura do cadastro**: ele mostra quantos estão marcados em cada coluna.
- **Datas:** digite as admissões no formato dd/mm/aaaa (por exemplo, 16/11/2026) para evitar dúvidas de leitura.
- **Salvar cadastro** guarda o cadastro no sistema. Na próxima vez, se você enviar outra planilha, o sistema
  mostra o que mudou (colaborador novo, ausente ou campo diferente) e só aplica o que você marcar em **Aceitar**.
  Sem marcar, vale o que já está salvo no sistema.
- **Baixar cadastro corrigido (.xlsx)** entrega uma planilha limpa, com os nomes de coluna corretos e sem legendas ou
  linhas em branco, que pode ser usada nos próximos meses.

**Erros (precisam ser corrigidos)**

| Mensagem | O que fazer |
|---|---|
| Razão social, código da empresa ou horário não preenchido | Preencha o campo indicado para aquela matrícula |
| Hora inválida | Escreva no formato hh:mm, por exemplo 08:00 |
| Dias de trabalho fora do formato | Use números de 0 a 6 separados por vírgula, por exemplo 0,1,2,3,4 (0 = segunda) |
| Matrícula ou PIS repetido | Cada colaborador precisa de matrícula e PIS próprios; corrija ou apague a linha repetida |
| Valor negativo ou inválido | Escreva um valor em R$ maior ou igual a zero, por exemplo 200,00 |

**Avisos (conferir)**

| Mensagem | O que significa |
|---|---|
| CNPJ com dígitos inválidos | O CNPJ digitado não passa na conferência. Se existir outro CNPJ válido da mesma empresa, ele aparece como sugestão |
| PIS igual à matrícula com zeros | Funciona somente com o Cartão Ponto (PDF). Com o arquivo AFD, o PIS precisa ser o real |
| Entrada/almoço diferentes da maioria da empresa | Confirme se o horário daquele colaborador realmente é diferente |
| Admissão depois do fim do período | O colaborador não terá dias apurados neste fechamento |
| Saída de sexta não preenchida | A sexta usará a mesma saída dos outros dias |
| Coluna banco_horas ausente | Ninguém será tratado como colaborador em banco de horas |
| N linhas ignoradas | Linhas sem matrícula/PIS, legendas e anotações da planilha não foram lidas; confira se não era um colaborador |

## Passo a passo

**1. Abrir o sistema.** No navegador, entre no endereço que o responsável de TI informar (normalmente `http://localhost:8501`). [Print: tela inicial]

**2. Etapa 1 – Arquivos.** Arraste **todos** os arquivos do mês de uma vez para a área de envio. O sistema reconhece sozinho o tipo de cada arquivo. Se ele errar, corrija na lista ao lado do arquivo. Depois clique em **"Confirmar arquivos e continuar"**. [Print: área de envio]

**3. Etapa 2 – Período e validação.** Escolha a data inicial e a final (normalmente 26 a 25) e clique em **"Validar dados"**.
Se aparecer um PIS ou matrícula que não está no cadastro, o sistema mostra **o arquivo e o número exatos**. Corrija no Cadastro de Funcionários e envie os arquivos de novo. [Print: aviso de validação]

**4. Etapa 3 – Processamento.** Marque a caixa dos feriados facultativos (Carnaval, Corpus Christi, 24 e 31 de dezembro) **somente se a empresa der esse dia de folga**. Se o relógio falhou para todos em algum dia, preencha **"Dias abonados para todos"** (veja abaixo). Depois clique em **"Processar ponto e calcular fechamento"**. [Print: botão de processar]

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

**Dia em que o relógio falhou para todos:** na Etapa 3, antes de processar, preencha **"Dias abonados para todos"** (ex.: `14/10/2026, 15/10/2026`). Quem não fechou as 4 batidas nesses dias não leva falta nem perde DSR, VA, COPR ou Cooper. Se mudar o campo depois de processar, clique em processar de novo. [Print: botões de download]

**Fechamento parcial (ponto só até uma data):** se o Espelho de Ponto ainda não cobre o período inteiro (por exemplo, o mês não terminou), abone os dias que faltam no mesmo campo. O resultado vale só até a última data com ponto e deve ser refeito quando o arquivo completo chegar.

## Quando os funcionários mudam

| Situação | O que fazer |
|---|---|
| Funcionário novo | Acrescente uma linha no cadastro, com o mesmo PIS e matrícula do Secullum e do consignado |
| Funcionário desligado | Apague a linha do cadastro **antes do próximo fechamento**. Se ele ainda aparecer no ponto mas não no cadastro, o sistema ignora e só mostra um aviso (não dá erro). Se continuar no cadastro sem batidas, use o aviso da etapa 3 |
| Mudou horário ou dias de trabalho | Edite os horários e `dias_trabalho` |
| Entrou ou saiu do banco de horas | Mude `banco_horas` |
| Funcionário que não bate ponto | Coloque `registra_ponto = não` |
| Novo empréstimo consignado | Nada a mudar no cadastro; ele vem no extrato do mês |
| Colaborador desligado (consignado) | Colaborador desligado: basta retirá-lo do cadastro. Se o extrato de consignados ainda listar essa matrícula, o sistema ignora e avisa na tela |
| Mudou de empresa | Atualize `codigo_empresa` e `razao_social` |

## O que o sistema já faz por você

> - Feriado **não** é falta. Trabalhar em feriado ou domingo paga **hora extra de 100%**.
> - **Qualquer** falta no mês, mesmo justificada, faz perder o prêmio **COPR**.
> - A **Cooper** (R$ 200 para a maioria) é perdida quando as **faltas do mês passam de 2 horas no total**, justificadas ou não. Exatamente 2 horas ainda mantém. Contam faltas de dia inteiro (a carga do dia) e faltas parciais; atraso não conta. Não é afetado quem está com **cooper_fixo = sim**, quem não tem Cooper e quem tem **registra_ponto = não**. Quando a Cooper é zerada, a Planilha do Thales sai com R$ 0,00 e uma observação com o total de faltas, e a aba "Conferência RH" mostra as colunas "Cooper zerado" e "Faltas p/ Cooper (min)".
> - Só a falta **injustificada** desconta o vale-alimentação (VA, 1/30 por falta de dia inteiro) e o descanso semanal (DSR). Vale fixo nunca é descontado. No vale com desconto manual, o valor sai cheio e a Observação da Planilha do Thales informa quantos dias de falta você deve descontar à mão.
> - O sistema **nunca** decide sozinho se uma falta é justificada. Essa decisão é sempre sua.
> - Evite atualizar a página (F5) ou fechar a aba no meio do trabalho. O rascunho fica salvo, mas é mais seguro terminar a revisão de uma vez.

## Problemas e soluções

| Problema | O que fazer |
|---|---|
| A página não abre | Veja se o sistema está ligado (peça ao responsável de TI) e se o endereço está certo |
| "PIS/matrícula não encontrado no cadastro" | Anote o arquivo e o número mostrados, corrija o cadastro (ou o número) e envie de novo |
| Um arquivo foi reconhecido como tipo errado | Troque o tipo na lista ao lado do arquivo |
| Aviso "Consignado ignorado para a(s) matrícula(s)..." | A matrícula está no extrato, mas não no cadastro. Se a pessoa foi desligada, não precisa fazer nada. Se não foi, corrija a matrícula no extrato ou inclua no cadastro e envie de novo |
| O cadastro tem erros na tela | Corrija os campos indicados na tabela (a mensagem cita a matrícula). Enquanto houver erro, "Validar dados" fica desligado |
| Alguém aparece com o mês inteiro de faltas | Provavelmente não tem batidas. Use o aviso da etapa 3 e depois corrija o cadastro |
| Um desligado aparece no relatório | Tire-o do cadastro (ou use "Desligado" no aviso da etapa 3) |
| A página foi atualizada ou fechada sem querer | Abra o sistema de novo; suas decisões da etapa 4 foram salvas como rascunho |
| Qualquer outro problema | Anote o que estava fazendo e o que a tela mostrou (ou tire um print) e envie ao responsável de TI |
