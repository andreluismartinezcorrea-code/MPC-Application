# Alterações desta revisão

## Versão 11.21.0 — 22/07/2026

- A Biblioteca Jurídica Local passou a ler diretamente a estrutura OOXML dos
  arquivos `.docx` e `.docm`, incluindo textos armazenados em controles de
  conteúdo do Word.
- A nova extração preserva a pesquisa em parágrafos, tabelas, cabeçalhos,
  rodapés, notas de rodapé, notas de fim e comentários.
- O banco agora registra a versão do extrator utilizada em cada documento.
  Assim, a primeira atualização do acervo nesta versão reprocessa
  automaticamente os arquivos antigos, mesmo quando tamanho e data não mudaram.

## Versão 11.20.0 — 22/07/2026

- A Biblioteca Jurídica Local ganhou os modos **Todas as palavras** e
  **Expressão exata**.
- No modo abrangente, todos os termos são obrigatórios, mas podem aparecer em
  qualquer ordem e posição do documento; a pesquisa por prefixo foi preservada.
- No modo exato, os termos devem aparecer juntos e na mesma ordem informada.
- As ocorrências encontradas são destacadas em amarelo no texto completo, com
  posicionamento automático na primeira ocorrência.
- Quando a expressão exata difere apenas em pontuação ou quebra de linha, as
  palavras correspondentes são realçadas individualmente para facilitar a
  localização visual.
- Adicionados dois testes dos modos de busca, elevando a suíte de 157 para 159.

## Versão 11.19.0 — 22/07/2026

- Incluído o botão **Gerenciar Categorias** na Biblioteca Jurídica Local.
- As categorias deixaram de ser uma lista fixa e passaram a ser armazenadas no
  banco, permitindo criar classificações próprias para cada ambiente de trabalho.
- O gerenciador permite adicionar, renomear e excluir categorias sem sair da GUI.
- Ao renomear uma categoria, todas as pastas associadas são atualizadas na mesma
  transação, preservando a organização existente.
- Categorias em uso não podem ser excluídas até que as respectivas pastas sejam
  reclassificadas. A categoria **Outros** permanece protegida como padrão seguro.
- Os filtros da pesquisa e as caixas de seleção são atualizados imediatamente
  depois de qualquer alteração no cadastro.
- Adicionados quatro testes de gerenciamento, elevando a suíte de 153 para 157.

## Versão 11.18.0 — 22/07/2026

- Criada a aba **Biblioteca Jurídica Local**, independente da Pesquisa
  Jurisprudencial já existente e sem alterar o Registro de Produção.
- É possível cadastrar múltiplas pastas e classificá-las como Pareceres,
  Legislação, Decisões e Acórdãos, Relatórios e Votos ou Outros.
- A localização de cada pasta pode ser alterada pela GUI. A remoção de um
  acervo apaga somente seu índice; os arquivos originais nunca são excluídos.
- Criado o módulo `mpc_biblioteca.py`, com extração local de `.doc`, `.docx`,
  `.docm` e PDFs que já contenham texto selecionável.
- A indexação é incremental: arquivos inalterados não são relidos; arquivos
  novos, modificados ou removidos são sincronizados com o índice SQLite FTS5.
- PDFs sem texto são sinalizados como **Sem texto**, sem executar OCR.
- A busca pode abranger todos os acervos ou uma categoria específica e mostra
  o documento completo para seleção precisa de trechos.
- Foram incluídos os comandos **Abrir documento**, **Copiar trecho** e
  **Inserir trecho no Word**. A inserção passa pela certificação e pelo backup
  preventivo das demais operações Word.
- As conexões SQLite da biblioteca têm fechamento garantido, inclusive após
  erros, preservando backups e movimentações do banco no Windows.
- Adicionados dez testes automatizados, elevando a suíte de 143 para 153.

## Versão 11.17.0 — 22/07/2026

- Criado o módulo `mpc_operacoes_word.py`, que centraliza a certificação dos
  dados e do documento ativo, a confirmação, o backup, a execução e o resultado
  final das rotinas que escrevem no Word.
- As operações passaram a devolver estados explícitos — concluída, cancelada,
  bloqueada ou erro — evitando que a GUI precise deduzir o resultado a partir
  de exceções ou variáveis dispersas.
- A mensagem de confirmação e a inspeção dos marcadores do Word foram retiradas
  do arquivo principal e passaram a ser reutilizadas pelo controlador.
- Criado o módulo `mpc_historico.py`, responsável por normalizar registros
  antigos, limitar o histórico, calcular indicadores e descrever a última ação.
- O painel de histórico ganhou a coluna **Detalhes**, contadores de operações
  concluídas, bloqueadas/com erro e canceladas, além da consulta por duplo clique.
- O carregamento de arquivos JSON agora normaliza o histórico persistido antes
  de apresentá-lo na interface.
- Adicionados oito testes para os novos controladores e para o histórico,
  elevando a suíte de 135 para 143 testes automatizados.

## Versão 11.16.0 — 22/07/2026

- Criado o módulo `mpc_tarefas.py`, responsável pela execução padronizada de
  trabalhos demorados fora da thread da interface.
- Resultado, exceção e traceback passam por uma fila interna e somente são
  entregues à GUI pelo agendador da thread principal.
- A infraestrutura verifica se a janela ainda existe antes de executar
  callbacks; resultados tardios são descartados sem acessar widgets destruídos.
- O fechamento da barra de progresso tornou-se idempotente e protegido contra
  `TclError`, inclusive quando o callback de atualização também falhar.
- Incluídos controle de cancelamento da entrega, callbacks distintos para
  sucesso, erro e finalização, e ciclo opcional de preparação/limpeza do worker.
- A inicialização e liberação COM das tarefas do Word foram centralizadas na
  mesma infraestrutura, assim como as tarefas isoladas de minuta e versionamento
  automático de modelos.
- **Analisar parecer com IA** deixou de chamar `after` a partir do worker: a
  geração ocorre em segundo plano e a escrita final no Word retorna pela thread
  principal, com a mesma janela segura de progresso.
- Adicionados sete testes de concorrência e ciclo de vida, elevando a suíte de
  128 para 135 testes automatizados.

## Versão 11.15.0 — 22/07/2026

- Criado o módulo `mpc_extracao.py`, separando da GUI o esquema e o prompt do
  Relatório de Auditoria, a leitura textual alternativa e a interpretação das
  respostas estruturadas da IA.
- Processo, tipo de processo, órgão, Serviço de Auditoria, apontamentos,
  recomendações e gestores passaram a ser normalizados antes do preenchimento
  dos widgets.
- O total de falhas e recomendações agora é recalculado a partir dos itens
  válidos. Divergências nas contagens informadas pela IA e itens sem numeração
  geram avisos claros para conferência.
- As marcações **Relatório Sem Falhas** com ou sem acento são reconhecidas sem
  serem apresentadas como conteúdo inválido.
- A leitura dos processos em tramitação da Análise de Esclarecimentos aceita
  JSON puro ou em bloco Markdown, remove duplicidades, ignora registros
  incompletos e limita o resultado às duas linhas disponíveis na GUI.
- Quando nenhum processo em tramitação é encontrado, os campos anteriores são
  limpos, impedindo a permanência de dados pertencentes a outro processo.
- A simplificação de cargos dos gestores e a tentativa de leitura textual com
  PyMuPDF/PyPDF2 também foram retiradas do arquivo principal.
- Adicionados oito testes, inclusive com a criação e leitura de um PDF textual
  real, elevando a suíte de 120 para 128 testes automatizados.

## Versão 11.14.0 — 22/07/2026

- Criado o módulo `mpc_controladores.py`, iniciando a separação entre as ações
  da tela e a tradução das regras em estados visuais.
- O controlador do painel passou a distinguir pendências bloqueantes de alertas
  informativos e devolve à GUI texto, resumo, cor e estado de liberação.
- Falhas marcadas como **Análise Pendente** que já possuam associação de falha,
  multa, repercussão ou débito agora são identificadas com item, Administradores
  e natureza do vínculo.
- O aviso aparece no status da classificação, no resumo superior e no painel
  detalhado de validação, sem remover ou modificar a associação provisória.
- A associação pendente não cria impedimento novo para outros comandos;
  **Conclusão** continua protegida pela exigência já existente de que todos os
  apontamentos tenham análise definitiva.
- O fluxo de atualização do painel e o fluxo do status de classificação foram
  retirados da lógica direta do Tkinter e passaram a consumir o controlador.
- Adicionados seis testes de regressão e controle, elevando a suíte de 114 para
  120 testes automatizados.

## Versão 11.13.0 — 22/07/2026

- Criado o módulo `mpc_banco.py`, que centraliza conexão, transações, schema,
  migração, tabelas auxiliares, jurisprudência, pareceres e pesquisa FTS.
- A interface deixou de abrir conexões SQLite e executar comandos SQL
  diretamente; agora apenas solicita operações à camada de dados e apresenta
  seus resultados ao usuário.
- A migração da tabela antiga de jurisprudência permanece compatível e passa a
  ser testada com comprovação de preservação do conteúdo e preenchimento do
  tipo de registro legado.
- O diagnóstico interno do programa passou a executar `quick_check` e a
  certificar tabelas, colunas adicionais e os seis gatilhos dos índices FTS.
- O schema recebeu versão explícita (`user_version`), preparando migrações
  futuras para serem aplicadas em sequência e de forma verificável.
- A criação de diretório, o tempo de espera em banco ocupado, as transações e o
  fechamento das conexões foram padronizados em um único ponto.
- Bancos criados por uma versão futura do programa são recusados com mensagem
  clara, evitando que uma versão antiga modifique uma estrutura desconhecida.
- Adicionados sete testes de banco, elevando a suíte de 107 para 114 testes
  automatizados.

## Versão 11.12.0 — 22/07/2026

- Iniciada a segunda etapa estrutural: separar a interface das regras de
  negócio, mantendo a migração gradual e compatível com a GUI atual.
- A consolidação de **Com Resp.**, **Sem Resp.** e **Recomendações** foi
  transferida para `mpc_regras.py`, inclusive ordenação, remoção de duplicatas,
  preservação das recomendações nativas e identificação de itens pendentes.
- As incompatibilidades do preenchimento em lote passaram a ser verificadas
  fora do Tkinter antes de qualquer alteração dos campos.
- A aplicação do lote também foi extraída: conclusão, multa, repercussão e suas
  associações são calculadas em cópias seguras dos apontamentos; a GUI apenas
  apresenta o resultado.
- Preservada a proteção que, no modo **Substituir**, mantém na responsabilidade
  pela falha os administradores vinculados a multa, repercussão ou débito que
  não estejam sendo alterados pelo lote.
- Adicionados sete testes de regressão, elevando a suíte de 100 para 107 testes
  automatizados.

## Versão 11.11.0 — 22/07/2026

- A busca de documentos da pasta **Notebook** foi transferida para a barra de
  **Preenchimento repetitivo**, ao lado das demais ações usadas durante a
  classificação das falhas.
- Removido o quadro lateral exclusivo da busca, liberando mais espaço vertical
  para análise, voto, arquivos auxiliares e controle das falhas.
- Criado o módulo `mpc_estado.py`, primeira etapa da centralização do estado da
  interface, reunindo responsáveis, apontamentos e associações em uma fotografia
  normalizada e independente dos widgets Tkinter.
- Validação, salvamento manual e salvamento automático passaram a consumir a
  mesma fotografia central, preservando a compatibilidade com os JSON antigos.
- As fotografias devolvidas são cópias seguras: uma rotina consumidora não pode
  modificar acidentalmente o estado mantido pela interface.
- Eliminada uma segunda montagem, redundante e descartada, de todos os dados no
  comando **Salvar dados**.
- Adicionados seis testes do estado central, elevando a suíte de 94 para 100
  testes automatizados.

## Versão 11.10.0 — 22/07/2026

- Concluída a segunda etapa visual da interface com a aproximação dos comandos
  ao ponto em que são usados e a redução da rolagem desnecessária.
- **Formatação DA CAMINO** e **Pesquisa de Falha por IA** foram transferidos
  para a mesma barra de **Preenchimento repetitivo**, no topo da lista de falhas.
- O espaço vertical antes ocupado por esses comandos na coluna lateral foi
  eliminado, aproximando o quadro de busca de documentos dos controles.
- Os quatro quadros da aba **3. Validação e comandos** agora reorganizam seus
  botões automaticamente em uma, duas ou três colunas de acordo com a largura
  disponível, preservando a separação por etapa e liberando espaço para o
  histórico das operações.
- A barra de rolagem vertical foi mantida como proteção para telas menores.
- Adicionados quatro testes para a grade responsiva, elevando a suíte de 90
  para 94 testes automatizados.

## Versão 11.9.0 — 21/07/2026

- Criado o módulo `mpc_prompt.py` para reunir e formatar, sem dependência da
  interface, os apontamentos, arquivos, responsáveis e associações escolhidos.
- Adicionado o botão **Construção de Prompt** ao quadro **Apontamentos** da aba
  **3. Validação e comandos**.
- A nova janela permite selecionar por botões deslizantes os itens, Relatório
  de Auditoria, Análise de Esclarecimentos, defesas individualizadas, Relatório
  e Voto, e-Parecer existente e administradores.
- O prompt adapta automaticamente singular/plural, mantém as quatro partes
  solicitadas e pode acrescentar contexto do processo, responsabilização,
  referências de peça/página e orientações livres.
- A prévia permanece editável e pode ser copiada para a área de transferência
  ou salva como TXT. A construção é integralmente local e não transmite PDFs
  ou textos a serviços externos.
- A aba **3. Validação e comandos** passou a possuir barra vertical e suporte à
  roda do mouse, garantindo acesso aos últimos botões em telas menores.
- Adicionados oito testes do construtor, elevando a suíte de 82 para 90 testes.

## Versão 11.8.0 — 21/07/2026

- Criado o módulo `mpc_certificacao.py`, independente de Tkinter e Word, para
  certificar os dados conforme a operação solicitada.
- Todas as onze rotinas que podem escrever no Word agora passam pela
  certificação antes da confirmação, do backup e de qualquer alteração.
- A certificação tornou-se proporcional à etapa: cabeçalho exige os dados do
  cabeçalho, introdução verifica esclarecimentos e revisão de Contas Anuais, e
  conclusão/ementa/resultado/IA verificam integralmente conclusões, falhas,
  multa, repercussão, débito, valores e associações.
- Documento Word ausente ou aberto somente para leitura bloqueia a rotina; uma
  possível divergência entre o processo da GUI e o nome do arquivo é mostrada
  como aviso confirmável.
- Antes da escrita, o programa também confere os marcadores exigidos pela
  operação e a linha-base `Gestor_1` do cabeçalho dinâmico, evitando alterações
  parciais em um modelo incompatível.
- A caixa de confirmação passou a identificar o documento alvo e os totais de
  responsáveis e apontamentos considerados.
- Adicionados doze testes da nova certificação, elevando a suíte de 70 para 82
  testes automatizados.

## Versão 11.7.0 — 21/07/2026

- Criado o módulo `mpc_templates.py`, responsável por localizar, listar,
  validar, salvar e restaurar o banco de textos-modelo.
- Adicionado o botão **Gerenciar textos-modelo** na interface, com pesquisa,
  lista de modelos, edição assistida, indicação dos campos automáticos e
  mensagens claras de validação.
- Cada salvamento cria uma cópia em `backups\\banco_paragrafos` antes de gravar
  o JSON por substituição atômica; também foi incluída a restauração da última
  cópia, preservando a versão atual.
- Centralizados no banco os cinco modelos do comando **Resultado das
  Verificações Procedidas**, incluindo não responsabilização,
  responsabilização, fluxo Fernanda e recomendações singular/plural.
- Adicionados quatro testes do serviço de modelos, elevando a suíte de 66 para
  70 testes, além do teste controlado da própria janela do gerenciador.

## Versão 11.6.0 — 21/07/2026

- Centralizados no `banco_paragrafos.json` os modelos editáveis de fixação de
  débito, cláusula de ressalvas e fundamentações especiais de Contas Anuais e
  de Fernanda Ismael em Contas Ordinárias.
- A lógica de associações, responsáveis, gênero, plural, itens de falha,
  multa, débito e negrito permanece no código, evitando que uma alteração de
  redação interfira no cálculo jurídico.
- Criado mecanismo de compatibilidade: bancos JSON anteriores continuam
  funcionando com os textos de segurança do programa até serem atualizados.
- Acrescentado guia didático com os caminhos editáveis e os cuidados para
  preservar os campos automáticos entre chaves.
- Acrescentados dois testes que comprovam a leitura dos novos textos diretamente
  do banco, elevando a suíte de 64 para 66 testes.

## Versão 11.5.0 — 21/07/2026

- Criado o módulo `mpc_visual.py` para centralizar tipografia, espaçamentos,
  estilos da interface e compatibilidade visual dos componentes Tk antigos.
- O cabeçalho com processo, exercício, órgão, validação e versão foi retirado
  da primeira aba e passou a permanecer visível em toda a aplicação.
- Acrescentada uma barra gráfica que acompanha automaticamente o progresso das
  sete etapas já calculado pelo guia textual do fluxo.
- Criada uma barra inferior permanente para exibir a última operação e sua
  situação sem interromper o trabalho com caixas adicionais.
- Ampliados os espaçamentos e a legibilidade das abas, botões, tabelas e
  cabeçalhos da grade de apontamentos, preservando o tema escuro atual.
- O contexto do processo agora ajusta automaticamente a quebra de linha quando
  a janela é redimensionada.
- Acrescentados três testes do mecanismo visual de progresso, elevando a suíte
  de 61 para 64 testes, além da compilação e do teste controlado da interface.

## Versão 11.4.0 — 21/07/2026

- Transferido para `mpc_word.py` todo o tratamento dos marcadores gramaticais
  de falhas, inclusive singular, plural, responsabilização e não
  responsabilização.
- Substituído no arquivo principal um bloco repetitivo de automação Word por
  uma única chamada ao serviço especializado.
- Criado um plano testável de concordância para as quantidades total, com
  responsabilidade e sem responsabilidade.
- A função **Conclusão** agora verifica o documento ao final e bloqueia a
  confirmação de sucesso se algum marcador conhecido permanecer sem processar.
- Acrescentados cinco testes Word, elevando a suíte de 56 para 61 testes, além
  da compilação e do teste controlado da interface.

## Versão 11.3.0 — 21/07/2026

- Transferidas da função principal para `mpc_word.py` as operações de busca e
  substituição dos marcadores `[CONCLUSÃO]` e `[DISPOSITIVO]`.
- Centralizadas no serviço Word a exclusão de marcador sem texto, a numeração
  dos parágrafos, a aplicação dos negritos e a preservação da conjunção entre
  itens destacados.
- A substituição de `Isto posto, opina este` por `Diante do exposto, opina
  este` também passou para o serviço especializado.
- A GUI ficou limitada à coordenação entre o motor jurídico e o serviço Word,
  sem executar diretamente os laços de formatação desses dois marcadores.
- Criados sete testes com objetos Word simulados, elevando a suíte de 49 para
  56 testes automatizados, além da compilação e do teste controlado da GUI.

## Versão 11.2.1 — 21/07/2026

- Corrigido o parágrafo pré-dispositivo especial de Fernanda Ismael em Contas
  Ordinárias para identificar nominalmente, com nome e cargo, somente os
  responsáveis cuja conclusão individual seja `Contas Irregulares`.
- A expressão genérica `dos Administradores` foi substituída por construção com
  concordância, como `do Sr. Ulisses Cecchin (Prefeito)`.
- Multa, fixação de débito e julgamento passaram a formar uma enumeração com
  vírgulas e apenas uma conjunção `e` antes do último elemento.
- Os números dos itens com repercussão passaram a receber negrito também nesse
  fluxo especial.
- Acrescentados testes do caso concreto e das combinações com somente multa,
  somente débito ou julgamento sem sanção, elevando a suíte para 49 testes.

## Versão 11.2.0 — 21/07/2026

- Extraídos da função principal os três fluxos que constroem o parágrafo
  explicativo anterior ao dispositivo, vinculado ao marcador `[CONCLUSÃO]`.
- Centralizadas no módulo `mpc_conclusao.py` as regras especiais de Fernanda
  Ismael em Contas Ordinárias, de Contas Anuais e do fluxo geral aplicável aos
  demais procuradores e processos.
- Preservados literalmente os textos jurídicos, a concordância, a separação
  entre multa e débito, os itens com repercussão e os destaques em negrito.
- A função da GUI agora apenas coleta os dados, recebe o texto certificado,
  substitui o marcador no Word e aplica as instruções de formatação.
- Acrescentados cinco testes de fundamentação, elevando a suíte de 42 para 47
  testes automatizados, além da compilação e do teste controlado da interface.

## Versão 11.1.0 — 21/07/2026

- Criado o módulo `mpc_conclusao.py`, responsável pela construção testável
  do núcleo dispositivo, sem dependência de Tkinter ou Microsoft Word.
- A função **Conclusão** passou a delegar ao novo mecanismo a redação das
  conclusões individuais, da multa única e dos débitos agrupados por devedor.
- Preservadas a ordem dos parágrafos, a redação dos modelos, os fundamentos
  legais, os destaques em negrito e as regras especiais posteriores de cada
  procurador.
- Acrescentados testes para concordância masculina, feminina e empresarial,
  Contas Anuais, Contas Ordinárias, Processo de Contas Especiais, Tomada de
  Contas Especial, multa conjunta e débitos com diferentes devedores.
- A suíte passou de 35 para 42 testes automatizados, todos aprovados, além do
  teste controlado de abertura e fechamento da interface.

## Versão 11.0.1 — 21/07/2026

- Corrigida a função **Conclusão**, que aceitava o campo individual vazio e
  silenciosamente omitia o respectivo administrador do texto inserido no Word.
- Todo responsável com nome preenchido passa a exigir uma conclusão individual
  antes da abertura dos textos-modelo e antes de qualquer inserção no Word.
- Contas Anuais, Contas Ordinárias, Processo de Contas Especiais e Tomada de
  Contas Especial também certificam se a conclusão escolhida é compatível com
  o tipo do processo.
- Linhas realmente vazias da tabela continuam dispensadas da validação.
- O painel **Validação do preenchimento** agora apresenta antecipadamente os
  nomes e as linhas que estão sem conclusão individual.
- A suíte foi ampliada de 31 para 35 testes automatizados.

## Versão 11.0.0 — 21/07/2026

- Criado um modelo central versionado para os dados do processo, com migração
  automática dos JSONs antigos e preservação de campos legados desconhecidos.
- O salvamento JSON passou a ser atômico: um arquivo temporário completo é
  validado antes de substituir o destino oficial.
- Implantado salvamento automático silencioso a cada 30 segundos, somente
  quando houver alteração ainda não salva, com recuperação oferecida na
  próxima abertura após interrupção.
- Criado o módulo independente de regras jurídicas para intimação,
  esclarecimentos, classificação, multa, repercussão, débito e associações.
- Separados serviços próprios para Gemini e Microsoft Word, incluindo limite
  de documentos, tentativas por cota, fechamento de cliente e conexão legível
  ao documento ativo.
- Adicionado ao topo da GUI um guia que informa a próxima etapa do fluxo e o
  progresso entre sete etapas operacionais.
- Adicionado à aba de comandos o histórico das operações da tarefa; o histórico
  é preservado junto com o JSON e usado para calcular o andamento.
- Uma instalação com banco SQLite novo agora cria também as tabelas auxiliares
  das caixas de seleção, eliminando erros de “tabela inexistente”.
- Inicializado um repositório Git exclusivamente local, com proteção explícita
  contra versionamento de `.env`, PDFs, Word, bancos, planilhas, backups e logs.
- Criada uma suíte de 31 testes automatizados, acrescida de compilação dos
  módulos e teste de abertura/fechamento da interface em banco temporário.

## Versão 10.12.0 — 21/07/2026

- Descontinuada e removida da interface a função **Preparar Esqueleto**, por duplicar comandos já disponíveis no fluxo normal.
- Descontinuada e removida da interface a **Fundamentação Automática** baseada no banco Word legado com marcadores `[START]` e `[END]`.
- O botão **2. Analisar Parecer com IA** passou a se chamar **Analisar Parecer com IA** e foi mantido como gerador opcional de minuta preliminar.
- A análise com IA agora valida previamente as associações, cria backup de segurança, escreve exatamente no documento Word capturado no início, não sobrescreve minutas anteriores e apresenta falhas da IA na interface.
- A função **Listar Apontes** passou a solicitar escolha quando houver vários PDFs e-Parecer, detectar PDFs sem texto/OCR, validar a numeração devolvida pela IA, ignorar duplicatas e linhas explicativas e informar corretamente o limite de 50 itens.
- A recarga da lista agora protege conclusões e associações já revisadas e atualiza imediatamente a caixa de seleção de apontamentos.

## Versão 10.11.1 — 21/07/2026

- A função **Gerar Log de Apontes** deixou de abrir a janela de seleção de destino.
- O documento é salvo automaticamente na pasta de trabalho indicada no campo **Pasta** da aba Parecer do MPC.
- Quando já existir um log com o mesmo nome, um sufixo numérico é acrescentado para preservar o arquivo anterior.
- Se o campo Pasta estiver vazio ou apontar para um local inexistente, o programa apresenta uma orientação e não cria o log em outro diretório.

## Versão 10.11.0 — 21/07/2026

- A função **Gerar Log de Apontes** passou a registrar todos os campos de dados da interface, preservando a estrutura em seções, a fonte e o espaçamento já utilizados.
- Incluídos no log os dados completos do processo, todos os responsáveis preenchidos, Parecer do MPC, documentos, esclarecimentos, tramitação, registro de produção, todos os apontamentos preenchidos e suas associações independentes.
- Incluídos também análise de apontamentos, voto, arquivos auxiliares, controle de falhas e recomendações, pesquisa jurisprudencial e estado da validação.
- Campos simples vazios passam a aparecer como **(não preenchido)**; linhas dinâmicas vazias de responsáveis e apontamentos são omitidas para evitar poluição visual.
- Corrigido o salvamento final explícito do documento Word e incluída a confirmação do caminho gerado.

## Versão 10.10.1 — 21/07/2026

- Corrigida a transferência das marcações da janela **Associar PDFs de esclarecimentos** para os `Combobox` em modo somente leitura do quadro **Responsáveis**.
- A confirmação agora certifica os valores efetivamente gravados em Intimação, Esclarecimentos e PDF antes de fechar a janela.
- Caso alguma transferência falhe, a janela permanece aberta e informa o administrador e o campo divergente.

## Versão 10.10.0 — 21/07/2026

- A Triagem Inteligente de Pasta passou a abrir automaticamente os PDFs documentais extraídos para a pasta raiz de cada processo.
- A janela da Análise de Esclarecimentos agora permite editar, por administrador, a Intimação, a Situação dos Esclarecimentos e o respectivo PDF.
- Foram separados os estados **Não Apresentou Esclarecimentos**, **Responsável Não Intimado** e **Esclarecimentos Espontâneos Desconsiderados**.
- A escolha da situação harmoniza automaticamente Intimação, Falhas, Multa, Débito e, para Contas Anuais ou Contas Ordinárias, a Conclusão favorável correspondente.
- Incluídas validações contra combinações contraditórias e contra a ausência de PDF quando houver esclarecimentos apresentados.
- Dados antigos com a marcação **Não Apresentou Defesa** são convertidos ao carregar, sem perda dos arquivos associados.

## Versão 10.9.1 — 21/07/2026

- A limpeza pós-registro passou a remover também todos os arquivos ZIP diretamente existentes na pasta de trabalho do processo.
- ZIPs abertos ou bloqueados são preservados e informados na mensagem final, sem interromper o Registro de Produção.

## Versão 10.9.0 — 21/07/2026

- O Registro de Produção agora remove o documento Word original após confirmar o salvamento do documento final e o registro em Excel.
- Eliminada a cópia adicional criada na pasta Mesa de Trabalho; permanece apenas a cópia preservada na pasta Produção.
- Incluída limpeza pós-registro dos PDFs da pasta de trabalho e da subpasta Notebook.
- Arquivos abertos ou não removíveis são preservados, listados na mensagem final e não interrompem o registro.

## Versão 10.8.7 — 21/07/2026

- Corrigida a redação do parágrafo de ressalvas em Contas Anuais para utilizar a referência a dispositivos legais e constitucionais e às normas de administração financeira e orçamentária.

## Versão 10.8.6 — 21/07/2026

- Criada certificação bidirecional entre a Repercussão das falhas e a conclusão individual dos administradores.
- Em Contas Anuais, **Parecer Desfavorável** exige associação a ao menos uma falha com Repercussão = Sim, e toda repercussão exige esse parecer para o administrador associado.
- Em Contas Ordinárias, a mesma regra passou a vincular **Contas Irregulares** à Repercussão = Sim.
- Removida a antiga exigência de Multa ou Débito como condição automática para Contas Irregulares, pois tais efeitos são independentes da repercussão que fundamenta a conclusão.

## Versão 10.8.5 — 21/07/2026

- Corrigida a seleção dos administradores do novo parágrafo de ressalvas em Contas Anuais.
- O parágrafo passa a considerar diretamente todos os administradores cuja conclusão seja **Parecer Favorável, com Ressalvas**, independentemente da marcação da coluna Repercussão.

## Versão 10.8.4 — 21/07/2026

- Acrescentado, como primeiro parágrafo da Conclusão de Contas Anuais, o texto específico para responsáveis com **Parecer Favorável, com Ressalvas**, quando o procurador for Fernanda Ismael, Ângelo Gräbin Borghetti ou Daniela Wendt Toniazzo.
- A relação de responsáveis adapta automaticamente tratamento, gênero, singular, plural e grupos mistos.
- Quando coexistirem responsáveis com ressalvas e com parecer desfavorável, os dois parágrafos são preservados na ordem adequada.

## Versão 10.8.3 — 21/07/2026

- Removida da função **Formatação DA CAMINO** a rotina obsoleta de busca e inserção de apontamentos.
- O botão agora executa exclusivamente a limpeza e a formatação do texto selecionado, seguida de sua cópia para a área de transferência.

## Versão 10.8.2 — 21/07/2026

- A caixa **Conclusão** do Preenchimento em lote passou a oferecer todos os tipos previstos na aba Apontamentos.
- Centralizada a lista de conclusões para manter o quadro e a operação em lote sincronizados.
- Incluída validação para impedir conclusões sem responsabilidade quando Multa, Repercussão ou Débito ainda estiverem marcados como Sim.

## Versão 10.8.1 — 21/07/2026

- Eliminada a segunda origem do erro COM `-2147467259` na função **Formatação DA CAMINO**.
- O recuo de primeira linha de 2 cm agora é convertido localmente para pontos, sem utilizar `CentimetersToPoints` do Word.

## Versão 10.8.0 — 21/07/2026

- Padronizados em maiúsculas os textos dos botões da janela principal e dos diálogos auxiliares.
- Ampliado o resumo no topo da interface para mostrar processo, exercício e órgão em fonte maior.
- Reorganizada a janela **Preenchimento em lote**, eliminando a sobreposição visual entre títulos, divisória e barras de rolagem.
- Separadas visualmente as opções de tratamento das associações existentes.

## Versão 10.7.3 — 21/07/2026

- Corrigido o erro COM `-2147467259` exibido após a execução de **Formatação DA CAMINO**.
- O espaçamento de 1,5 linha passa a ser aplicado pela regra nativa do Word, sem utilizar a conversão instável `LinesToPoints`.

## Versão 10.7.0 — 20/07/2026

- Adicionada a janela `Preenchimento em lote…` na aba Apontamentos.
- O usuário pode selecionar todas as falhas preenchidas ou apenas parte delas
  e aplicar Conclusão, Multa e Repercussão simultaneamente.
- A mesma janela permite escolher um ou vários administradores e optar por
  adicionar às associações existentes ou substituir a associação do campo
  alterado.
- Multa ou Repercussão em lote somente são aceitas quando a conclusão
  resultante for `Mantido` ou `Mantido Parcialmente`.
- O comando sincroniza automaticamente os campos gerais `Falhas` e `Multa` da
  tabela de administradores.
- Incluídos resumo e confirmação antes da aplicação e o botão
  `Desfazer último lote` para recuperação imediata.
- Débito e Valor permanecem individuais como proteção contra atribuição
  patrimonial indevida.
- A suíte foi ampliada para 43 testes automatizados.

## Versão 10.6.1 — 20/07/2026

- Consolidados todos os administradores associados a `Multa = Sim` em um
  único parágrafo do dispositivo, preservando a ordem da tabela de
  responsáveis e eliminando nomes repetidos.
- Mantida a separação dos parágrafos de débito por conjunto de devedores e
  valor total, pois essa distinção continua necessária para individualizar a
  responsabilidade patrimonial.

## Versão 10.6.0 — 20/07/2026

- Adicionada a coluna `Valor` ao lado de `Débito` nas 50 linhas da aba
  Apontamentos, com formatação monetária brasileira, como `R$ 1.000,00`.
- `Débito = Sim` passou a exigir valor maior que zero, além da associação dos
  administradores responsáveis.
- A Conclusão agora totaliza os débitos por conjunto exato de devedores e gera
  parágrafos separados quando as responsabilidades não forem idênticas.
- Os valores totais de débito também recebem negrito no texto inserido no
  Word.
- O dispositivo de multa voltou a indicar apenas os administradores multados,
  sem repetir os itens de falha.
- A caixa `Recomendações` tornou-se multilinha, com rolagem vertical, para
  manter visível uma lista extensa de itens.
- Salvamento JSON, banco de dados, snapshots, carregamento e validação passaram
  a preservar o valor individual de cada débito.
- A suíte foi ampliada para 42 testes automatizados.

## Versão 10.5.0 — 20/07/2026

- Separadas, em cada apontamento, quatro associações independentes:
  responsável pela falha, pela multa, pela repercussão e pelo débito.
- A janela **Selecionar…** passou a exibir todos os administradores em uma
  matriz de chaves deslizantes, habilitadas conforme a conclusão e os campos
  `Sim/Não`.
- A certificação agora respeita a cadeia de responsabilização: multa,
  repercussão e débito somente podem alcançar administradores previamente
  associados à própria falha.
- Uma falha compartilhada pode atribuir multa ou débito a somente parte dos
  administradores, sem estender automaticamente a consequência aos demais.
- A função **Conclusão** agrupa os itens pelo conjunto exato de responsáveis e
  cria parágrafos separados de multa e fixação de débito com a numeração das
  falhas.
- O parágrafo pré-dispositivo de repercussão passou a considerar os
  administradores especificamente associados a essa consequência.
- Salvamento JSON, banco de dados, snapshots, carregamento e análise assistida
  preservam as quatro associações.
- Arquivos da versão 10.4 continuam compatíveis e têm a associação única
  convertida automaticamente quando são carregados.
- A suíte foi ampliada para 40 testes automatizados.

## Versão 10.4.0 — 20/07/2026

- `Convertido em Alerta` volta a ser tratado como falha `Sem Resp.`, pois
  permanece sujeito ao julgamento colegiado e não se confunde com uma
  recomendação genuína.
- Adicionados `PDF dos esclarecimentos` e `Débito` a cada administrador.
- Ao marcar `Intimação = Não`, o programa sugere automaticamente:
  `Responsável Não Intimado`, `Não Apresentou Defesa`, ausência de falhas,
  multa e débito, além da conclusão favorável compatível com o tipo do
  processo. Todos os campos continuam editáveis.
- A função `Análise de Esclarecimentos` agora mantém um PDF técnico geral e
  abre uma janela para associar um PDF de defesa diferente a cada
  administrador.
- Os documentos individuais também são utilizados, identificados pelo nome do
  administrador, na análise assistida por IA.
- Adicionadas as colunas `Débito` e `Administradores associados` às 50 linhas
  da aba Apontamentos.
- Cada falha permite selecionar um ou vários administradores, mesmo quando não
  há multa nem débito.
- Implementada certificação bidirecional: multa e débito das falhas precisam
  coincidir com os administradores associados, e administradores sancionados
  precisam estar vinculados a uma falha correspondente.
- Falhas com multa ou débito são classificadas `Com Resp.`; alertas
  convertidos permanecem `Sem Resp.`.
- Salvamento JSON, banco de dados, snapshots, carregamento, limpeza, ementa e
  log passaram a preservar os novos campos.
- Adicionada rolagem horizontal à aba Apontamentos e ampliada a suíte para 35
  testes automatizados.

## Versão 10.3.0 — 20/07/2026

- Corrigida a classificação preliminar do Relatório de Auditoria para impedir
  que o mesmo item apareça simultaneamente em `Sem Resp.` e `Recomendações`.
- O e-Parecer agora transfere as recomendações preliminares para as linhas da
  aba Apontamentos e recalcula o quadro de controle.
- Adicionado indicador de revisão pendente e botão para recalcular a
  classificação.
- Em Contas Anuais, a Introdução aguarda a conclusão da revisão dos
  apontamentos, evitando inserir no Word um parágrafo ainda provisório.
- Itens `Convertido em Alerta` passam a integrar o grupo de recomendações e
  alertas, sem dupla contagem.
- Adicionados os comandos `Subir` e `Descer` para reordenar responsáveis,
  preservando todos os campos e renumerando automaticamente as linhas.
- Adicionados testes de sobreposição de recomendações e movimentação de
  responsáveis.

## Versão 10.2.0 — 20/07/2026

- Adicionada limpeza automática de backups vencidos, com retenção padrão de 90 dias.
- A limpeza remove somente arquivos dentro da pasta de backups e ignora links.
- O prazo pode ser configurado por `MPC_BACKUP_RETENCAO_DIAS`; valor `0` desativa a rotina.
- Adicionado botão de limpeza manual, com confirmação, na tela Diagnóstico.
- Adicionados testes para retenção, arquivos recentes e proteção de arquivos externos.

## Versão 10.0.0 — 20/07/2026

- Criada a primeira separação modular por meio de `mpc_infra.py`.
- Adicionados logs rotativos e captura de erros da interface e de threads.
- Adicionados snapshots JSON antes de operações críticas.
- Adicionado backup do último estado salvo do documento Word ativo.
- Adicionadas confirmações antes de rotinas que modificam o Word e da limpeza.
- Adicionada tela de diagnóstico do ambiente.
- Adicionada versão visível no título e no cabeçalho.
- Adicionado versionamento automático, por conteúdo, dos modelos Word.
- Ampliados testes de regras jurídicas, backups e versionamento.

## Interface moderna e responsáveis dinâmicos — 20/07/2026

- A tela principal foi reorganizada em etapas: processo, parecer/documentos e validação/comandos.
- Incluídos cabeçalho de contexto, painel de pendências, cartões de comandos, tooltips e comportamento responsivo.
- A tabela de responsáveis agora possui numeração, cabeçalho uniforme, linhas alternadas e botões para adicionar ou remover linhas.
- Eliminado o limite fixo de cinco responsáveis na extração, geração textual, conclusão, salvamento, carregamento, limpeza e logs.
- O cabeçalho do Word replica a última linha formatada do modelo quando houver mais responsáveis que linhas disponíveis.
- Linhas não utilizadas do modelo continuam sendo removidas de baixo para cima.

## Correção da introdução — 20/07/2026

- Responsáveis intimados que não apresentaram defesa agora são agrupados por tratamento.
- A concordância passa automaticamente para o plural (`intimados` e `apresentaram`) quando houver mais de um responsável.
- Removidas as vírgulas duplicadas entre os nomes e o trecho `regularmente intimados`.
- Adicionado teste com os quatro responsáveis do caso informado.

## Hotfix de dependência Gemini — 20/07/2026

- A ausência de `google-genai` não impede mais a inicialização da interface.
- Ao usar um recurso de IA sem o SDK, a aplicação informa o comando correto de instalação.
- A documentação agora esclarece que o programa deve ser extraído do ZIP e executado com o mesmo interpretador usado para instalar as dependências.

## Hotfix de recurso incorporado — 20/07/2026

- Centralizado o carregamento de `banco_paragrafos.json`.
- Adicionada uma cópia comprimida do JSON dentro do próprio arquivo Python.
- Quando o programa é aberto diretamente do ZIP e o Windows usa uma pasta
  temporária, conclusão e ementa continuam funcionando pelo fallback.
- O arquivo JSON externo continua tendo prioridade e pode ser configurado por
  `MPC_BANCO_PARAGRAFOS`.
- Adicionados testes que comparam integralmente o fallback com o JSON original.

## Hotfix de inicialização — 20/07/2026

- Removido `from tkinter import ttk` de dentro de `main()`. Essa importação
  tornava `ttk` uma variável local e provocava `UnboundLocalError` antes da
  criação da janela.
- Removido também um `import re` interno redundante.
- Adicionado teste que proíbe imports dentro de `main()`, evitando nova sombra
  de módulos já importados globalmente.

## Segurança

- Remoção da chave Gemini gravada no código.
- Configuração por `.env`, excluído do versionamento.
- Migração de `google.generativeai` para `google-genai`.
- Confirmação explícita antes do envio de documentos à IA.
- Instrução uniforme contra prompt injection documental.
- Resposta JSON solicitada pelo MIME correto na extração estruturada.
- Limites configuráveis para PDFs e conteúdo descompactado de ZIPs.
- Remoção do uso do clipboard para extrair o texto integral do parecer.

## Integridade de dados

- O original deixa de ser apagado no registro de produção.
- Cópias nunca sobrescrevem arquivos existentes.
- O documento Word é salvo uma vez antes da distribuição das cópias.
- Workbooks recebem fechamento garantido em caso de erro.
- Correção da criação do cabeçalho em uma planilha anual vazia.

## Banco SQLite

- Inicialização ocorre antes das consultas feitas pela interface.
- Remoção do `commit` intermediário durante migração.
- Tabela temporária residual é eliminada antes de uma nova tentativa.
- FTS é verificada/reconstruída sem apagar as tabelas em todo início.
- Colunas novas são detectadas por `PRAGMA table_info`.
- Nomes de tabelas de lookup são validados por allowlist.

## Estrutura

- Interface encapsulada em `main()` com guarda de execução.
- Importar o módulo não abre GUI, não migra o banco e não limpa cache.
- Limpeza de `gen_py` tornou-se opt-in.
- Sete redefinições silenciosas foram eliminadas.
- Código inalcançável após `return` foi removido.
- Threads de Word inicializam e liberam COM.
- Caminhos principais passaram a aceitar configuração externa.
- Importações e dependências sem uso foram removidas.

## Validações executadas

- Compilação do código Python.
- Doze testes automatizados.
- Importação sem criação da GUI.
- Importação sem alteração do banco real.
- Inicialização de banco novo em diretório temporário.
- Migração sobre cópia do banco real com contagens preservadas.
- `PRAGMA integrity_check`: `ok`.
- Verificação de dependências do ambiente isolado: sem conflitos.

## Limitações conhecidas

- A interface ainda utiliza muitas variáveis globais. A migração completa para
  classes/MVC deve ser feita em uma segunda etapa com testes funcionais da GUI.
- As funções jurídicas mais extensas ainda precisam ser divididas e cobertas por
  testes de regressão baseados em documentos anonimizados.
- Automação Word/Excel e aparência da interface exigem teste manual no desktop,
  pois não foram acionadas contra documentos reais nesta revisão.
