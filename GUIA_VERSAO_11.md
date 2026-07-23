# Guia didático — MPC Parecer versão 11

## 1. O que você deve copiar

Copie para `C:\Dev` todos os arquivos `mpc_*.py`, o arquivo principal
`MPC Parecer Application.py`, `banco_paragrafos.json` e `requirements.txt`.

Mantenha o seu `.env` atual em `C:\Dev`. Não substitua esse arquivo pelo
`.env.example`, porque o `.env` contém suas configurações locais e a chave da
API Gemini.

## 2. Como abrir

No PowerShell:

```powershell
cd C:\Dev
C:\Users\andre\AppData\Local\Programs\Python\Python313\python.exe ".\MPC Parecer Application.py"
```

## 3. Como funciona a recuperação automática

- Continue usando **SALVAR DADOS** normalmente.
- A cada 30 segundos, o programa verifica se existem alterações não salvas.
- Se o programa fechar antes do salvamento, a próxima abertura perguntará se
  você deseja recuperar o preenchimento.
- Escolher **SIM** restaura a tela. Depois, clique em **SALVAR DADOS** para
  gravar o processo definitivamente.
- Escolher **NÃO** apaga somente a cópia automática, sem excluir documentos,
  PDFs, Word, banco de dados ou JSONs salvos manualmente.

## 4. Guia do fluxo na tela

O topo mostra a próxima etapa sugerida e o progresso. As sete etapas são:

1. preencher processo e responsáveis;
2. executar Relatório de Auditoria;
3. executar Análise de Esclarecimentos;
4. gerar o e-Parecer;
5. revisar e certificar os apontamentos;
6. construir cabeçalho, introdução e conclusão;
7. registrar a produção.

O guia é informativo: ele não impede que você execute uma rotina fora dessa
ordem quando o caso concreto exigir.

### O que permanece visível durante o trabalho

- O cabeçalho superior continua na tela mesmo quando você muda para a aba
  **Apontamentos** ou **Pesquisa Jurisprudencial**.
- Processo, exercício e órgão aparecem juntos e se ajustam à largura da janela.
- A barra colorida representa graficamente o progresso das sete etapas.
- O status de validação muda de cor: amarelo indica pendências e verde indica
  que os dados essenciais estão prontos.
- A barra inferior mostra a última operação registrada, sem abrir uma nova
  caixa de diálogo.

## 5. Histórico

Na aba **3. Validação e comandos**, o quadro **Histórico das operações desta
tarefa** informa horário, operação e situação. Esse histórico é salvo no JSON
do processo e volta a aparecer quando os dados são carregados.

## 6. O que não mudou

- Seus modelos Word continuam sendo usados da mesma forma.
- O `banco_paragrafos.json` continua ao lado do programa.
- Os JSONs salvos pelas versões anteriores continuam aceitos.
- A chave Gemini permanece somente no `.env`.
- PDFs, documentos Word, banco de dados e planilhas não são enviados para o
  Git e não fazem parte do código-fonte.

## 7. Conferência técnica opcional

```powershell
cd C:\Dev
C:\Users\andre\AppData\Local\Programs\Python\Python313\python.exe -m unittest discover -s tests -v
```

O resultado esperado termina com `OK`.

## 8. Como alterar textos-modelo com segurança

O modo recomendado é utilizar o botão **GERENCIAR TEXTOS-MODELO**, localizado
na aba **3. Validação e comandos**, no quadro **Dados e utilidades**.

1. Digite uma palavra para localizar o texto desejado.
2. Clique no modelo encontrado.
3. Altere somente a redação no painel à direita.
4. Clique em **SALVAR ALTERAÇÃO**.
5. O programa cria uma cópia automática antes de gravar a alteração.

Os backups ficam em `C:\Dev\backups\banco_paragrafos`. Se uma alteração não
ficar como esperado, abra o gerenciador e use **RESTAURAR ÚLTIMA CÓPIA**. A
versão atual também será preservada antes da restauração.

O arquivo `banco_paragrafos.json` continua disponível para edição manual, mas
é preferível usar o gerenciador. Se precisar editar manualmente, faça antes uma
cópia com outro nome, por exemplo: `banco_paragrafos - cópia antes da alteração.json`.

Abra o arquivo no VS Code e altere somente o conteúdo depois de `"texto":`.
Não altere os nomes entre aspas à esquerda, as chaves `{` e `}`, vírgulas ou os
campos entre chaves, como `{gestores}`, `{valor_total}` e `{consequencias}`.
Esses campos são preenchidos automaticamente pelo programa conforme o processo.

Os novos modelos mais importantes ficam nestes caminhos:

- `fundamentacao_especial > contas_anuais > ressalvas > texto`;
- `fundamentacao_especial > contas_anuais > desfavoravel > texto`;
- `fundamentacao_especial > fernanda_contas_ordinarias > sem_responsabilidade > texto`;
- `fundamentacao_especial > fernanda_contas_ordinarias > sem_irregularidade > texto`;
- `fundamentacao_especial > fernanda_contas_ordinarias > irregularidade > texto`;
- `dispositivo > paragrafo_debito > texto`;
- `dispositivo > texto_ressalvas`.

Depois de salvar, feche e abra novamente o programa. Se uma vírgula, aspas ou
chave for removida por engano, o programa avisará que o banco de parágrafos não
foi carregado. Nesse caso, restaure a cópia que você fez antes da alteração.

## 9. Certificação antes de alterar o Word

Os botões que escrevem no Word agora fazem uma conferência automática antes de
qualquer modificação. Você não precisa acionar outro botão previamente.

O funcionamento é o seguinte:

1. clique normalmente na rotina desejada, por exemplo **INTRODUÇÃO**,
   **CONCLUSÃO** ou **EMENTA**;
2. o programa verifica somente os campos necessários para aquela operação;
3. se houver pendências, uma mensagem explica o que corrigir e o Word permanece
   intacto;
4. se a certificação for aprovada, a confirmação informa o nome do documento
   ativo e quantos responsáveis e apontamentos serão considerados;
5. somente depois da sua confirmação é criado o backup e iniciada a escrita.

A certificação é proporcional à etapa. **CABEÇALHO (E-PARECER)**, por exemplo,
não exige que as conclusões individuais já estejam preenchidas. Já
**CONCLUSÃO**, **EMENTA**, **RESULTADO DAS VERIFICAÇÕES** e **ANALISAR PARECER
COM IA** exigem a revisão completa das conclusões, das associações, da multa,
da repercussão e do débito.

Quando a rotina depende de um documento já aberto, também são verificados:

- a existência de um documento Word ativo;
- a possibilidade de edição do arquivo;
- os marcadores exigidos pela rotina, como `[INTRODUÇÃO]`, `[CONCLUSÃO]`,
  `[DISPOSITIVO]` e `[EMENTA]`;
- a linha-base `Gestor_1` quando for utilizado **CABEÇALHO (E-PARECER)**;
- uma possível diferença entre o processo da GUI e o nome do documento ativo.

A diferença no nome é exibida como aviso para sua conferência, pois alguns
modelos não incluem o número do processo no nome. Documento ausente ou aberto
somente para leitura bloqueia a operação.

## 10. Construção de Prompt

O botão **CONSTRUÇÃO DE PROMPT** fica no quadro **Apontamentos** da aba
**3. Validação e comandos**. Ele não chama a IA e não transmite nenhum arquivo:
apenas prepara localmente o texto que você poderá copiar e usar onde desejar.

Ao abrir a janela:

1. marque os apontamentos que deseja analisar;
2. mantenha marcados somente os arquivos que serão efetivamente anexados;
3. escolha os Administradores que devem aparecer no contexto;
4. ative ou desative contexto, mapa de responsabilidade, referências de
   página/peça e separação por item;
5. se desejar, escreva orientações adicionais;
6. clique em **ATUALIZAR PRÉVIA**;
7. revise livremente o texto apresentado à direita;
8. use **COPIAR PROMPT** ou **SALVAR COMO TXT**.

A lista de documentos é formada automaticamente com os campos preenchidos na
GUI: Relatório de Auditoria, Análise de Esclarecimentos, PDFs individuais dos
Administradores, Relatório e Voto e, como referência opcional, o e-Parecer já
existente. Peça e páginas também são incluídas quando estiverem informadas.

O texto é dividido em quatro partes: Relatório de Auditoria, Esclarecimentos
dos Administradores, Conclusões do Órgão Instrutivo e Análise/Minuta de Parecer
do MPC. Para reduzir erros da IA, o prompt também determina que não sejam
inventados fatos ou referências e que as posições da auditoria, da defesa, do
Órgão Instrutivo e do MPC permaneçam claramente separadas.

### Rolagem da aba de comandos

A aba **3. Validação e comandos** agora possui barra vertical própria. Quando a
tela não comportar todos os quadros, utilize essa barra ou a roda do mouse para
chegar ao histórico e aos últimos botões de **Dados e utilidades**.

## 11. Segunda etapa visual

Na aba **Apontamentos**, as rotinas usadas durante a classificação das falhas
ficam reunidas na barra superior da grade. Nela você encontra, sem precisar
rolar a coluna lateral:

- **PREENCHIMENTO EM LOTE**;
- **DESFAZER ÚLTIMO LOTE**;
- **FORMATAÇÃO DA CAMINO**;
- **PESQUISA DE FALHA POR IA**;
- pesquisa de uma peça na pasta **Notebook**, informando seu número e
  pressionando **BUSCAR E ABRIR** ou a tecla `Enter`.

Na aba **3. Validação e comandos**, os grupos continuam separados em **Fluxo
principal**, **Construção do documento**, **Apontamentos** e **Dados e
utilidades**. O que mudou foi a disposição interna: o programa calcula o espaço
disponível e distribui os botões em até três colunas. Ao reduzir a janela, a
grade retorna automaticamente para duas ou uma coluna para que os nomes não
sejam cortados.

A rolagem vertical permanece disponível em resoluções menores. Em telas
maiores, a grade compacta deixa uma área mais ampla para o **Histórico das
operações desta tarefa**.

## 12. Estado central da interface

Responsáveis, apontamentos e suas quatro associações de responsabilidade agora
são reunidos em uma fotografia central antes da validação ou do salvamento.
Essa mudança é interna: você continua preenchendo os mesmos campos e os arquivos
JSON anteriores continuam compatíveis.

Na prática, o programa deixa de montar listas independentes para cada comando.
O salvamento manual, o salvamento automático e o painel de validação consultam o
mesmo conjunto normalizado. Isso reduz o risco de uma associação aparecer na
tela, mas ser interpretada de modo diferente por outra rotina.

Esta é a primeira parte da modernização estrutural. As listas de widgets antigas
foram mantidas temporariamente como uma camada de compatibilidade, permitindo
que as demais funções sejam migradas gradualmente e com testes a cada etapa.

## 13. Separação entre interface e regras

A interface continua com os mesmos campos e botões, mas deixou de decidir
diretamente como os apontamentos devem ser classificados. O módulo de regras
agora calcula:

- quais itens ficam em **Com Resp.**;
- quais itens ficam em **Sem Resp.**;
- quais itens constituem recomendações genuínas;
- quais itens ainda estão com conclusão pendente;
- se uma combinação escolhida no preenchimento em lote é permitida;
- como ficam as associações de falha, multa e repercussão depois do lote.

Na prática, o uso permanece igual. A diferença é que as regras podem ser
testadas sem abrir a janela e sem alterar seus documentos. A GUI limita-se a
recolher as escolhas, solicitar a confirmação e mostrar o resultado calculado.

O modo **Substituir a associação do campo alterado** continua protegendo os
responsáveis que precisam permanecer vinculados por multa, repercussão ou
débito não alterados pelo lote.

## 14. Camada segura do banco de dados

O funcionamento visual continua igual, mas todas as operações do banco SQLite
passaram a ficar no módulo `mpc_banco.py`. A tela não monta mais comandos SQL e
não precisa decidir como abrir, confirmar ou fechar uma conexão.

Na abertura do programa, a camada de dados executa quatro cuidados:

1. migra bancos antigos de jurisprudência sem descartar os registros;
2. cria tabelas ou colunas que ainda não existam;
3. prepara e sincroniza os índices de pesquisa textual e seus gatilhos;
4. certifica a integridade e a estrutura final antes do uso.

O banco também registra internamente a versão de sua estrutura. Esse número não
precisa ser preenchido pelo usuário; ele serve para que atualizações futuras
apliquem somente as migrações necessárias e na ordem correta.

O botão de diagnóstico agora informa **estrutura e integridade verificadas**
quando o arquivo está completo. Se houver corrupção, tabela, coluna ou gatilho
ausente, a linha **Banco SQLite** aparece como pendência e descreve o problema.

Cadastros de órgãos, inclusão e pesquisa de jurisprudência e registro de
pareceres usam a mesma camada transacional. Isso reduz o risco de uma conexão
ficar aberta ou de uma gravação parcial permanecer no arquivo após um erro.

## 15. Alertas de associações provisórias

É permitido associar previamente um Administrador a uma falha que ainda esteja
marcada como **Análise Pendente**. O programa não apaga essa escolha e não a
trata como erro, pois ela pode representar uma organização provisória do exame.

Para que essa situação não passe despercebida, o programa apresenta:

- no controle das falhas, a quantidade de itens pendentes já associados;
- no topo, o resumo **Dados prontos com alerta(s)**, quando não houver outro
  problema bloqueante;
- no painel **Validação do preenchimento**, o número da falha, os nomes dos
  Administradores e o tipo de associação encontrada.

O aviso é apenas informativo. Os demais comandos continuam disponíveis. Ao
executar **CONCLUSÃO**, entretanto, a proteção já existente exige que todos os
itens tenham uma conclusão definitiva; portanto, nenhum texto será inserido no
Word enquanto permanecer **Análise Pendente**.

Essa apresentação é calculada no módulo `mpc_controladores.py`. A interface
apenas mostra o texto e a cor devolvidos por ele, facilitando novos alertas e
painéis sem espalhar regras pelos widgets.

## 16. Conferência automática das extrações de PDF

O uso dos botões **RELATÓRIO DE AUDITORIA** e **ANÁLISE DE ESCLARECIMENTOS**
continua igual. Internamente, porém, a resposta da IA passa por uma camada de
conferência antes de preencher a tela.

No Relatório de Auditoria, o programa:

1. padroniza o número do processo, o tipo, o órgão e o Serviço de Auditoria;
2. conserva somente numerações válidas de apontamentos e recomendações;
3. elimina numerações repetidas mantendo a ordem original;
4. recalcula as quantidades a partir das listas efetivamente aceitas;
5. simplifica cargos como **Prefeito Municipal** para **Prefeito**;
6. reconhece **Relatório Sem Falhas** mesmo quando a IA omitir o acento;
7. informa na mensagem final qualquer correção automática que mereça revisão.

Se a leitura nativa do PDF falhar, continua disponível a tentativa pelo texto
local. Quando o documento for apenas uma imagem, a mensagem explica que é
necessário aplicar OCR.

Na Análise de Esclarecimentos, a lista de processos em tramitação também é
normalizada. O programa aceita a resposta em JSON puro ou cercada por marcações
Markdown, remove duplicidades e utiliza as duas primeiras ocorrências válidas.
Se nenhuma for encontrada, marca **Não** e limpa as duas linhas de tramitação,
evitando a reutilização acidental dos dados do processo anterior.

Essas regras estão no módulo `mpc_extracao.py` e são testadas sem abrir a GUI ou
transmitir documentos reais à IA.

## 17. Operações seguras em segundo plano

As rotinas mais demoradas não devem congelar a janela enquanto aguardam a
leitura de um PDF ou uma resposta da IA. A versão 11.16 centraliza esse trabalho
no módulo `mpc_tarefas.py`.

O fluxo passou a funcionar assim:

1. a GUI abre a janela de progresso;
2. o trabalho demorado é iniciado em uma thread separada;
3. sucesso, erro e detalhes técnicos são colocados em uma fila segura;
4. a thread principal recebe o resultado e somente então atualiza a tela;
5. a barra de progresso é interrompida e destruída uma única vez.

Se a janela principal for fechada enquanto uma resposta ainda estiver sendo
produzida, o resultado tardio é descartado. O programa não tenta interromper
violentamente uma comunicação ou automação em andamento, mas impede que ela
acesse controles que já deixaram de existir. Essa proteção evita erros como
`invalid command name` em barras de progresso destruídas.

As tarefas que utilizam automação COM também passam por preparação e limpeza
garantidas. Na função **ANALISAR PARECER COM IA**, a leitura e a geração ficam
no worker, mas a inserção final no Word volta à thread principal. Para o usuário,
o botão e o resultado permanecem iguais; a mudança é de segurança interna.

## 18. Controlador das operações do Word e histórico

A versão 11.17 reúne em um único controlador o ciclo das rotinas que alteram o
Word. A sequência agora é sempre a mesma:

1. coletar o estado atual da interface;
2. certificar os campos, responsáveis, apontamentos e o documento ativo;
3. apresentar a confirmação ao usuário;
4. criar o backup de segurança;
5. executar a rotina solicitada;
6. registrar o resultado no histórico.

Se qualquer certificação falhar, o Word não é modificado e o histórico registra
a operação como **bloqueada**. Se o usuário desistir, o estado será
**cancelada**. Exceções são convertidas no estado **erro** e continuam sendo
gravadas no arquivo técnico de log.

Na aba **3. Validação e comandos**, o quadro de histórico mostra data, operação,
situação e detalhe. Um duplo clique abre o detalhe completo. Acima da tabela há
um resumo com o total de eventos, conclusões, problemas e cancelamentos da
tarefa atual. O histórico continua sendo salvo juntamente com os dados do
processo e arquivos antigos permanecem compatíveis.

## 19. Biblioteca Jurídica Local

A aba **Biblioteca Jurídica Local** pesquisa documentos existentes em uma ou
mais pastas do computador ou da rede institucional. Ela não interfere no botão
**Registro de Produção**, que continua inserindo os dados do parecer na tabela
própria do banco.

### Cadastrar e indexar uma pasta

1. Clique em **+ ADICIONAR PASTA**.
2. Escolha a pasta no Explorador do Windows.
3. Classifique-a como Pareceres, Legislação, Decisões e Acórdãos, Relatórios e
   Votos ou Outros.
4. Selecione a linha cadastrada e clique em **INDEXAR SELECIONADA**; para várias
   pastas, use **ATUALIZAR TODOS OS ÍNDICES**.

O botão **GERENCIAR CATEGORIAS** permite criar classificações próprias,
renomeá-las ou excluir as que não estejam sendo utilizadas. Antes de excluir
uma categoria em uso, altere a classificação das pastas correspondentes. A
categoria **Outros** é obrigatória e permanece protegida pelo programa.

A primeira indexação pode demorar conforme a quantidade de documentos. As
seguintes são incrementais: somente arquivos novos ou modificados são lidos.
Arquivos removidos deixam de aparecer na pesquisa após a atualização.

São aceitos `.doc`, `.docx`, `.docm` e PDFs que já possuam texto selecionável.
PDFs compostos apenas por imagens aparecem no contador **Sem texto** e não são
submetidos a OCR. Todo o processamento é local e nenhum documento é transmitido
à API Gemini.

### Pesquisar e utilizar um trecho

Digite as palavras, escolha uma categoria ou **Todos** e, no campo **Modo**,
selecione uma das opções:

- **Todas as palavras**: exige todos os termos, mas permite que apareçam
  separados e em qualquer ordem;
- **Expressão exata**: exige que as palavras apareçam juntas e na mesma ordem.

Ao abrir um resultado, as ocorrências são realçadas em amarelo e o painel se
posiciona automaticamente na primeira delas.

### Textos em controles de conteúdo do Word

A partir da versão 11.21, a indexação de arquivos `.docx` e `.docm` também
lê os textos inseridos em controles de conteúdo do Word (as caixas usadas em
muitos modelos institucionais). A leitura continua inteiramente local e não
utiliza OCR nem serviços de inteligência artificial.

Na primeira utilização desta versão, clique em **ATUALIZAR TODOS OS ÍNDICES**.
O programa reconhecerá a mudança do extrator e reprocessará automaticamente
os documentos já cadastrados; não é necessário remover e adicionar as pastas.

Selecione com o mouse o trecho exato e use:

- **COPIAR TRECHO**, para levá-lo à área de transferência;
- **INSERIR TRECHO NO WORD**, para inseri-lo no cursor do documento ativo após
  confirmação e backup de segurança;
- **ABRIR DOCUMENTO**, para consultar o arquivo original.

Se uma pasta for realocada, selecione seu cadastro, clique em **ALTERAR LOCAL**
e informe o novo endereço. Em seguida, atualize o índice. A classificação pode
ser corrigida com **ALTERAR CATEGORIA**. O botão **REMOVER DA BIBLIOTECA**
elimina somente o catálogo daquela pasta, nunca os documentos.
