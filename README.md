# MPC Parecer Application — versão 11.22.0

Esta é a versão de produção reorganizada da aplicação em `C:\Dev`. A versão
11 mantém os arquivos JSON anteriores compatíveis, mas separa as regras, a IA,
o Word, a persistência e a recuperação de sessão em módulos menores.

Extraia todo o conteúdo do ZIP antes de executar. Abrir somente o arquivo `.py`
diretamente do ZIP faz o Windows usar uma pasta temporária, onde o `.env` não
estará disponível. O banco de parágrafos possui fallback incorporado, mas a
configuração da API continua dependendo do `.env` externo.

Leia também `GUIA_VERSAO_11.md`.

## Arquivos necessários na pasta de trabalho

Para executar a versão 11, mantenha juntos:

- `MPC Parecer Application.py`;
- `mpc_banco.py`;
- `mpc_biblioteca.py`;
- `mpc_biblioteca_gui.py`;
- `mpc_conclusao.py`;
- `mpc_certificacao.py`;
- `mpc_controladores.py`;
- `mpc_estado.py`;
- `mpc_extracao.py`;
- `mpc_fluxo.py`;
- `mpc_ia.py`;
- `mpc_infra.py`;
- `mpc_historico.py`;
- `mpc_modelos.py`;
- `mpc_operacoes_word.py`;
- `mpc_persistencia.py`;
- `mpc_prompt.py`;
- `mpc_regras.py`;
- `mpc_sessao.py`;
- `mpc_templates.py`;
- `mpc_tarefas.py`;
- `mpc_visual.py`;
- `mpc_word.py`;
- `banco_paragrafos.json`;
- `requirements.txt`;
- `.env` — arquivo local que contém a chave e não deve ser compartilhado.

Os arquivos da pasta `tests` são recomendados para conferência, mas não são
necessários para abrir a GUI. `README.md`, `CHANGELOG.md`, `.env.example` e os
guias são documentação.

Os scripts antigos `database.py`, `gemini_automation.py` e `preparar_bd.py`
não são utilizados pela versão 11.

## Antes do primeiro uso

1. Revogue a antiga chave Gemini que estava gravada no código.
2. Copie `.env.example` para `.env`.
3. Preencha `GEMINI_API_KEY` no `.env`.
4. Confirme `MPC_TCE_ROOT` e, se necessário, os caminhos opcionais.
5. Faça um backup do banco SQLite, da planilha de produção e dos documentos.

## Instalação

No PowerShell, dentro desta pasta:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python ".\MPC Parecer Application.py"
```

Use o mesmo interpretador em todos os comandos. Se o terminal mostrar
`cannot import name 'genai' from 'google'`, o pacote `google-genai` não está
instalado no Python que executou o programa. Com a pasta já extraída, execute:

```powershell
C:\Users\andre\AppData\Local\Programs\Python\Python313\python.exe -m pip install -r requirements.txt
C:\Users\andre\AppData\Local\Programs\Python\Python313\python.exe ".\MPC Parecer Application.py"
```

Não instale o pacote chamado apenas `google`: ele não fornece `google.genai`.

## Principais proteções aplicadas

- A chave da API foi removida do código.
- A integração foi migrada para o SDK oficial `google-genai`.
- A aplicação pode ser importada sem abrir a interface ou executar migrações.
- A limpeza do cache `gen_py` tornou-se opt-in.
- O banco é inicializado antes de as comboboxes fazerem consultas.
- Índices FTS deixam de ser apagados e recriados em toda inicialização.
- As redefinições silenciosas de classes e funções foram eliminadas.
- Threads que usam Word/COM inicializam e liberam COM corretamente.
- Fluxos de IA exibem confirmação explícita de transmissão externa.
- O texto do parecer deixa de passar pela área de transferência.
- O registro de produção preserva o original e evita sobrescrever destinos.
- Caminhos principais podem ser configurados pelo `.env`.
- O salvamento dos dados usa substituição atômica, evitando JSON incompleto.
- Uma recuperação automática é atualizada somente quando há alteração não
  salva; o padrão é verificar a cada 30 segundos.
- Ao reabrir o programa depois de uma interrupção, a GUI oferece a restauração
  do último preenchimento não salvo.
- Arquivos antigos são migrados por um modelo central, sem descartar campos
  desconhecidos de versões anteriores.
- Regras de intimação, repercussão, multa, débito e associações têm testes
  automatizados independentes da interface.
- A redação do dispositivo e dos parágrafos explicativos anteriores da
  função **Conclusão** é produzida em um módulo independente do Tkinter e do
  Word, com testes de concordância, tipos de processo, procuradores, repercussão,
  multa e débito.
- A localização dos marcadores `[CONCLUSÃO]` e `[DISPOSITIVO]`, a numeração
  dos itens e os respectivos negritos passaram a ser responsabilidade exclusiva
  do serviço `mpc_word.py`, testado com documentos Word simulados.
- Os marcadores gramaticais de falhas, responsabilização e não
  responsabilização também são processados pelo serviço Word; a conclusão é
  bloqueada caso algum marcador conhecido permaneça no documento.
- A apresentação visual possui um módulo próprio: o cabeçalho com processo,
  exercício e órgão permanece visível em todas as abas, o fluxo conta com uma
  barra gráfica de progresso e a última operação aparece na barra inferior.
- A segunda etapa visual reúne **PREENCHIMENTO EM LOTE**, **DESFAZER ÚLTIMO
  LOTE**, **FORMATAÇÃO DA CAMINO**, **PESQUISA DE FALHA POR IA** e a busca de
  peças na pasta **Notebook** na barra superior dos apontamentos. Na aba de
  comandos, cada grupo distribui suas ações automaticamente em uma, duas ou
  três colunas conforme a largura.
- O estado dinâmico de responsáveis e apontamentos é centralizado no módulo
  `mpc_estado.py`. Validação, salvamento manual e salvamento automático usam a
  mesma fotografia normalizada da tela, evitando coletas divergentes.
- O painel de preenchimento e o status da classificação usam o controlador
  independente `mpc_controladores.py`, separando bloqueios reais de alertas
  informativos e deixando a GUI responsável apenas pela apresentação.
- A interpretação das respostas de Relatório de Auditoria e tramitação da
  Análise de Esclarecimentos fica em `mpc_extracao.py`. Processo, tipo, serviço,
  itens, recomendações, gestores e contagens são saneados antes de chegar aos
  campos da interface.
- Contagens divergentes devolvidas pela IA são recalculadas a partir das listas;
  itens sem numeração válida são desconsiderados com aviso para conferência.
- Tarefas demoradas usam `mpc_tarefas.py`, que centraliza worker, fila de
  resultados, retorno à thread da interface, tratamento de exceções, ciclo COM
  e descarte seguro de respostas recebidas depois do fechamento da janela.
- A barra de progresso é encerrada de forma idempotente em sucesso ou falha,
  evitando chamadas a widgets que já tenham sido destruídos.
- Uma falha em **Análise Pendente** que já possua Administrador associado é
  exibida como alerta no topo, no controle de falhas e no painel detalhado; a
  associação provisória não é apagada nem bloqueia os demais comandos.
- A classificação entre **Com Resp.**, **Sem Resp.** e **Recomendações**, a
  identificação de pendências e as regras do preenchimento em lote foram
  retiradas da GUI e centralizadas em `mpc_regras.py`, com testes independentes.
- Os textos especiais de fundamentação e de débito passaram a ter modelos
  editáveis no `banco_paragrafos.json`; a lógica jurídica e a formatação Word
  continuam protegidas no código.
- O botão **Gerenciar textos-modelo** permite editar o banco diretamente pela
  aplicação, com pesquisa, validação dos campos automáticos, backup e
  restauração da última cópia.
- Toda rotina que escreve no Word passa primeiro por uma certificação adequada
  à etapa. Pendências bloqueiam a operação antes do backup e antes de qualquer
  alteração; a confirmação mostra o documento ativo e os totais efetivamente
  considerados.
- O botão **Construção de Prompt** reúne seletivamente apontamentos, arquivos,
  administradores e associações da GUI, adapta a concordância e produz uma
  instrução jurídica editável, copiável ou salvável em TXT sem enviar dados.
- A IA e o Word possuem serviços próprios, com mensagens de erro mais claras.
- O banco SQLite possui uma camada própria para conexão, migração, estrutura,
  pesquisas FTS e gravações; a interface deixou de executar comandos SQL.
- O diagnóstico da aplicação verifica também a integridade, as tabelas, as
  colunas de compatibilidade e os gatilhos de pesquisa do banco SQLite.
- O topo indica a próxima etapa do fluxo e a aba de comandos conserva o
  histórico das operações da tarefa no próprio JSON.

## Recuperação automática

A recuperação não substitui **Salvar dados**. Ela serve apenas para proteger o
preenchimento contra fechamento inesperado, travamento ou queda de energia.

1. A cada 30 segundos, o programa compara a tela com o último salvamento.
2. Se nada mudou, nenhum arquivo é gravado.
3. Se houve alteração, é criada uma cópia local em `.mpc_session`.
4. Depois de **Salvar dados**, essa cópia é removida.
5. Se o programa for encerrado com alterações não salvas, na próxima abertura
   será exibida a opção **Recuperar trabalho não salvo**.

Para alterar o intervalo, defina `MPC_AUTOSAVE_SEGUNDOS` no `.env`. O valor
padrão é `30`.

## Responsáveis e tabela do Word

A quantidade de responsáveis não está mais limitada a cinco:

1. Use **Adicionar responsável** para criar quantas linhas forem necessárias.
2. Escolha o número em **Linha selecionada** e use **Remover selecionado**; as
   linhas seguintes serão renumeradas sem deixar lacunas.
3. Use **Subir** ou **Descer** para mudar a ordem de qualquer responsável.
   Nome, cargo e todos os demais campos acompanham a linha.
4. Ao salvar um processo, todas as linhas preenchidas são armazenadas.
5. Ao carregar, a interface cria automaticamente as linhas necessárias.
6. No cabeçalho do Word, as linhas não usadas do modelo são removidas.
7. Se houver mais responsáveis que linhas no modelo, a aplicação replica a
   última linha de responsável, preserva sua formatação e insere os nomes
   adicionais antes do restante da tabela.

O modelo deve manter pelo menos uma linha de responsável com controle de
conteúdo intitulado `Gestor_1`. Recomenda-se conservar as cinco linhas atuais,
pois elas servem como base de formatação e dão compatibilidade aos documentos
já existentes.

## Ordem segura para revisar os apontamentos

1. Execute **Relatório de Auditoria**, **Análise de Esclarecimentos** e
   **e-Parecer**.
2. Na aba **Apontamentos**, revise `Conclusão`, `Multa`, `Repercussão` e
   `Débito` de todos os itens.
3. Em **Associações detalhadas**, informe separadamente quem responde pela
   falha, pela multa, pela repercussão e pelo débito.
4. Confira o indicador no quadro **Controle de Falhas e Sugestões**. Ele deve
   informar que a classificação está consolidada.
5. Depois execute **Cabeçalho (e-Parecer)** e **Introdução**.

Em Contas Anuais, a Introdução não é gerada enquanto existir item em branco ou
com `Análise Pendente`. Essa proteção impede que uma lista provisória de
recomendações seja gravada no documento Word.

Na distribuição dos três grupos:

- somente `Conclusão = Recomendação` integra o grupo de recomendações;
- `Convertido em Alerta` permanece como falha `Sem Resp.`;
- `Multa = Sim` ou `Débito = Sim` classifica a falha como `Com Resp.`;
- as demais falhas ficam em `Sem Resp.`.

`Repercussão` é usada nos textos próprios da conclusão e também dispara a
atualização do quadro, mas não altera sozinha essa separação.

## Preenchimento em lote

Na parte superior da tabela de apontamentos, use **Preenchimento em lote…**
para aplicar a várias falhas:

- `Conclusão = Mantido` ou `Mantido Parcialmente`;
- `Multa = Sim/Não`;
- `Repercussão = Sim/Não`;
- os administradores associados a essas alterações.

É possível selecionar todas as falhas preenchidas ou somente algumas e optar
por adicionar os administradores às associações existentes ou substituir a
associação do campo alterado. Antes da aplicação, o programa apresenta um
resumo e pede confirmação. **Desfazer último lote** restaura o estado anterior
enquanto a aplicação permanecer aberta.

Débito e Valor continuam fora do comando em lote, pois exigem conferência
individual dos devedores e dos valores de cada item.

## Esclarecimentos individualizados

Cada administrador possui a coluna **PDF dos esclarecimentos**:

1. Use o botão `…` na própria linha para escolher o PDF daquele administrador.
2. Ou clique em **Análise de Esclarecimentos**. Depois da seleção do PDF
   técnico geral, uma janela apresentará todos os administradores para a
   associação dos documentos individuais.
3. Para quem não apresentou defesa, use **Sem defesa**.
4. Se `Intimação` for alterada para `Não`, o programa preencherá a situação
   lógica automaticamente, mas os campos continuarão editáveis.

O quadro **Documentos, esclarecimentos e tramitação** conserva um resumo dos
arquivos associados para compatibilidade com os modelos e registros antigos.

## Débito e associação das falhas

Cada apontamento possui:

- `Multa`;
- `Repercussão`;
- `Débito`;
- `Valor`;
- `Associações detalhadas`.

Use **Selecionar…** na linha da falha. A mesma janela apresenta quatro
seleções independentes:

- responsável pela existência da falha;
- responsável pela multa;
- responsável pela repercussão;
- responsável pelo débito.

A associação da falha fica disponível quando a conclusão é `Mantido` ou
`Mantido Parcialmente`. As demais associações ficam disponíveis quando a
respectiva coluna está marcada como `Sim`. Quem for associado à multa, à
repercussão ou ao débito também precisa estar associado à falha.

Quando `Débito = Sim`, preencha a coluna **Valor** com o total apurado para
aquela falha, por exemplo, `1.000,00`. Ao sair do campo, o programa o exibe
como `R$ 1.000,00`. Um débito sem valor é apontado como pendência e impede a
geração da Conclusão.

O painel de validação confere os dois sentidos da informação. Por exemplo, uma
falha com Débito = Sim associada a determinado administrador exige que esse
administrador também esteja com Débito = Sim. A geração do Resultado das
Verificações e da Conclusão fica bloqueada enquanto houver divergências.

Na função **Conclusão**, as falhas com os mesmos administradores multados ou
devedores são agrupadas. O dispositivo identifica os itens e os respectivos
responsáveis em parágrafos separados de Multa e Fixação de débito. A
Repercussão continua sendo usada no parágrafo explicativo anterior ao
dispositivo.

Nos parágrafos de débito, o programa soma os valores dos itens que tenham o
mesmo conjunto de devedores. Caso os conjuntos sejam diferentes, gera
parágrafos independentes, sem presumir solidariedade. O parágrafo de multa
identifica, em um único item do dispositivo, todos os administradores
multados, sem repetir os itens de falha nem duplicar nomes.

## Validação recomendada

Execute primeiro com cópias de documentos e banco. Valide especialmente:

- leitura do Relatório de Auditoria;
- geração de introdução, conclusão e ementa;
- registro na planilha de produção;
- gravação e pesquisa no banco de jurisprudência;
- salvamento nas três pastas esperadas.

Textos produzidos por IA continuam sendo minutas e exigem revisão humana.

## Testes automatizados

Para executar a conferência técnica sem abrir a GUI:

```powershell
C:\Users\andre\AppData\Local\Programs\Python\Python313\python.exe -m unittest discover -s tests -v
```

A versão 11 possui 107 testes, além de compilação de todos os módulos
e um teste de abertura/fechamento da interface usando banco temporário novo.
