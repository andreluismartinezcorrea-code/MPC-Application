# Guia didático — MPC Parecer versão 10.6

Esta versão acrescenta proteções de segurança e manutenção sem alterar
intencionalmente as regras jurídicas já validadas.

## 1. Instalação ou atualização

1. Feche o MPC Parecer e o Word.
2. Não apague a pasta que funciona atualmente.
3. Renomeie-a, por exemplo, para `MPC Parecer - versão 9 - backup`.
4. Extraia integralmente o ZIP da versão 10 em uma nova pasta.
5. Copie o seu arquivo secreto `.env` para a nova pasta.
6. Não envie o `.env` a outras pessoas e não o coloque em ZIPs.
7. Execute `MPC Parecer Application.py`.
8. Confirme se aparece `Versão 10.6.0` no alto da janela.
9. Clique em **Diagnóstico** e confira os itens marcados com `⚠`.

O arquivo `mpc_infra.py` deve permanecer na mesma pasta de
`MPC Parecer Application.py`. A partir desta versão, não se deve copiar somente
o programa principal.

## 2. Backups automáticos

Antes das rotinas que alteram ou geram conteúdo no Word, o programa:

1. mostra uma confirmação;
2. cria um arquivo JSON com os dados essenciais da tela;
3. copia o último estado salvo do documento Word ativo;
4. somente depois executa a rotina solicitada.

Os backups ficam, por padrão, na pasta `backups`, ao lado do programa.

Backups com mais de 90 dias são removidos automaticamente na abertura do
programa. Para alterar o prazo, inclua no `.env`, por exemplo:

`MPC_BACKUP_RETENCAO_DIAS=180`

Use `MPC_BACKUP_RETENCAO_DIAS=0` apenas se quiser desativar a limpeza
automática. A limpeza atua exclusivamente dentro da pasta de backups, registra
o resultado no log e também pode ser executada pelo botão **Diagnóstico**.

Importante: o backup do Word corresponde ao último estado salvo em disco.
Se você tiver digitado algo no Word e ainda não tiver salvado, pressione
`Ctrl+S` antes de executar uma rotina.

Ao usar **Limpar dados**, também será criado um snapshot antes da limpeza.

## 3. Tela de diagnóstico

O botão **Diagnóstico**, no alto da janela, verifica:

- versão do programa e do Python;
- existência do `.env`;
- configuração da chave Gemini;
- instalação do SDK de IA;
- banco de parágrafos e banco SQLite;
- pastas do TCE, modelos e produção;
- disponibilidade da pasta de backups;
- arquivo de log;
- existência de documento Word ativo.

`✓` significa disponível. `⚠` significa que o item deve ser conferido — não
significa necessariamente que todo o programa está impedido de funcionar.

## 4. Logs de erro

Os erros técnicos são registrados em `logs\mpc_parecer.log`.

O arquivo não registra sua chave Gemini nem o conteúdo integral dos documentos.
Quando ocorrer um erro, envie o texto da mensagem e, se necessário, uma cópia
do log — sempre confira antes se não há informação processual sensível.

## 5. Modelos Word versionados

Na abertura do programa, os modelos Word são examinados em segundo plano.
Uma nova cópia é criada somente quando o conteúdo de um modelo mudou.

As versões ficam em `modelos_versionados`.

Também é possível abrir **Diagnóstico** e clicar em
**Versionar modelos agora**.

## 6. Arquivos que devem permanecer juntos

Obrigatórios para a distribuição:

- `MPC Parecer Application.py`;
- `mpc_infra.py`;
- `banco_paragrafos.json`;
- `requirements.txt`.

Local e secreto:

- `.env`.

Documentação e suporte:

- `README.md`;
- `GUIA_VERSAO_10.md`;
- `.env.example`;
- pasta `tests`.

Criados automaticamente:

- pasta `backups`;
- pasta `logs`;
- pasta `modelos_versionados`.

## 7. Primeira validação recomendada

Faça os testes em cópias:

1. abra um documento Word de teste e salve-o;
2. preencha um processo simples;
3. execute **Introdução** e confirme a mensagem de backup;
4. confira se foi criado um JSON em `backups\dados`;
5. confira se foi criada uma cópia do Word em uma subpasta de `backups`;
6. abra **Diagnóstico**;
7. teste um relatório PDF;
8. teste adicionar e remover um responsável intermediário;
9. compare os textos produzidos com os modelos institucionais.

Minutas produzidas com IA continuam exigindo revisão humana.
