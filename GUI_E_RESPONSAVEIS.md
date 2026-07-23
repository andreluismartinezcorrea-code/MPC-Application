# Guia da nova interface e dos responsáveis dinâmicos

## Organização da tela

A primeira área do programa foi dividida em três etapas:

1. **Processo e responsáveis**: dados do RAG e tabela dinâmica.
2. **Parecer e documentos**: informações do parecer, produção, arquivos e tramitação.
3. **Validação e comandos**: pendências de preenchimento e rotinas agrupadas por finalidade.

O cabeçalho superior mostra o processo atual e o estado dos campos essenciais.

## Adicionar e remover responsáveis

- Clique em **Adicionar responsável** para criar uma nova linha.
- Escolha o número em **Linha selecionada** e clique em
  **Remover selecionado** para retirar qualquer linha.
- Se a linha contiver um nome, o programa pedirá confirmação antes de removê-la.
- Use **Subir** ou **Descer** para reordenar a linha selecionada. Todos os
  campos acompanham o responsável e a numeração é refeita automaticamente.
- Use o botão `…` na coluna **PDF dos esclarecimentos** para associar o
  documento individual daquele administrador.
- A coluna **Débito** é certificada com as falhas vinculadas na aba
  Apontamentos.
- O programa começa com cinco linhas por compatibilidade, mas não há limite fixo.
- Salvamento, carregamento, introdução, conclusão, ementa e logs consideram todas as linhas preenchidas.

## Associações detalhadas dos apontamentos

Na aba **Apontamentos**, clique em **Selecionar…** na linha desejada. A janela
mostra cada administrador e quatro chaves deslizantes:

1. **Falha**: quem é responsável pela ocorrência mantida ou parcialmente
   mantida.
2. **Multa**: quem recebe a multa daquela falha.
3. **Repercussão**: quem está relacionado à falha que fundamenta a conclusão
   desfavorável ou irregular.
4. **Débito**: quem responde pelo débito daquele item.

Uma falha pode ser compartilhada por dois administradores e ter débito
atribuído a apenas um deles. O programa preserva essa distinção ao salvar os
dados e ao construir a Conclusão.

Quando marcar **Débito = Sim**, informe também o **Valor** do débito daquela
linha no formato `1.000,00`. A aplicação soma os valores apenas dos itens que
tenham exatamente os mesmos devedores, formando parágrafos de débito
separados quando as responsabilidades forem diferentes.

## Funcionamento no Word

Não é necessário aumentar manualmente o modelo para dez linhas.

O modelo atual pode conservar as cinco linhas com controles de conteúdo
`Gestor_1` a `Gestor_5`.

- Com menos de cinco responsáveis, o botão **Cabeçalho (e-Parecer)** remove as linhas não usadas.
- Com cinco responsáveis, preenche as cinco linhas existentes.
- Com mais de cinco, replica a última linha de responsável, preserva sua formatação e insere as novas linhas antes do restante da tabela.

Faça os primeiros testes usando uma cópia do documento:

1. Teste com três responsáveis e confirme a remoção de duas linhas.
2. Abra novamente uma cópia limpa do modelo.
3. Teste com sete responsáveis e confirme a criação de duas linhas.
4. Verifique bordas, fonte, alinhamento e posição das linhas acrescentadas.

Se o modelo for alterado futuramente, mantenha pelo menos o controle
`Gestor_1` dentro da linha que servirá como referência visual.
