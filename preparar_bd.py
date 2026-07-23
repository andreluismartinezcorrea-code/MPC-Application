import sqlite3
import os

# Caminho para o seu banco de dados.
DB_PATH = r"E:\TCE\Database\BD-MPC.db"

def preparar_banco_de_dados_completo():
    """
    Cria e popula todas as tabelas de lookup necessárias para a aplicação,
    garantindo a integridade dos dados.
    """
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        print("--- INICIANDO PREPARAÇÃO COMPLETA DO BANCO DE DADOS ---")

        # Dicionário com todas as tabelas a serem criadas e seus dados.
        # A chave 'check_upper' define se o nome deve ser sempre maiúsculo.
        tabelas_a_criar = {
            "orgaos": {"check_upper": True, "dados": []},
            "tipos_processo": {"check_upper": True, "dados": [
                'CONTAS ANUAIS', 'CONTAS ORDINÁRIAS', 'PROCESSO DE CONTAS ESPECIAIS',
                'TOMADA DE CONTAS ESPECIAL', 'RETIFICAÇÃO DE CERTIDÃO',
                'CONTAS DE GESTÃO', 'CONTAS DE GOVERNO'
            ]},
            "relator": {"check_upper": True, "dados": [
                "CONSELHEIRO ALEXANDRE POSTAL", "CONSELHEIRO CEZAR MIOLA", "CONSELHEIRO EDSON BRUM",
                "CONSELHEIRO ESTILAC MARTINS RODRIGUES XAVIER", "CONSELHEIRO IRADIR PIETROSKI",
                "CONSELHEIRO MARCO PEIXOTO", "CONSELHEIRO RENATO LUÍS BORDIN DE AZEREDO",
                "CONSELHEIRO-SUBSTITUTO ALEXANDRE MARIOTTI", "CONSELHEIRA-SUBSTITUTA ANA CRISTINA MORAES",
                "CONSELHEIRA-SUBSTITUTA DANIELA ZAGO GONÇALVES DA CUNDA",
                "CONSELHEIRA-SUBSTITUTA HELOISA T. GOULART PICCININI", "CONSELHEIRA-SUBSTITUTA LETÍCIA AYRES RAMOS",
                "CONSELHEIRO-SUBSTITUTO ROBERTO DEBACCO LOUREIRO"
            ]},
            "servico_de_auditoria": {"check_upper": False, "dados": [
                "Serviço de Auditoria da Região de Porto Alegre I - SRPA I",
                "Serviço de Auditoria da Região de Porto Alegre II - SRPA II",
                "Serviço Regional de Auditoria de Caxias do Sul", "Serviço Regional de Auditoria de Erechim",
                "Serviço Regional de Auditoria de Frederico Westphalen", "Serviço Regional de Auditoria de Passo Fundo",
                "Serviço Regional de Auditoria de Santa Cruz do Sul", "Serviço Regional de Auditoria de Santa Maria",
                "Serviço Regional de Auditoria de Santo Ângelo", "Serviço de Instrução Estadual e Municipal - SIEM"
            ]},
            "procurador": {"check_upper": True, "dados": [
                "ÂNGELO GRÄBIN BORGHETTI", "DANIELA WENDT TONIAZZO", "FERNANDA ISMAEL", "GERALDO COSTA DA CAMINO"
            ]},
            "tipo_parecer": {"check_upper": True, "dados": [
                "PARECER", "PROMOÇÃO", "RECURSO", "REPRESENTAÇÃO"
            ]},
            "cargo": {"check_upper": False, "dados": [
                "Prefeito", "Prefeita", "Vice-Prefeito", "Vice-Prefeita", "Diretor Presidente",
                "Diretora Presidente", "Presidente", "Presidente da Câmara Municipal", "Vereador",
                "Vereadora", "Fiscal de Contrato", "CNPJ:"
            ]},
            "conclusao": {"check_upper": False, "dados": [
                "Parecer Favorável", "Parecer Favorável, com Ressalvas", "Parecer Desfavorável",
                "Contas Regulares", "Contas Regulares, com Ressalvas", "Contas Irregulares", "Arquivamento"
            ]},
            "esclarecimentos": {"check_upper": False, "dados": [
                "Advogada", "Advogado", "Advogadas", "Advogados", "Pessoalmente", "Não Apresentou Defesa"
            ]}
        }

        for nome_tabela, info in tabelas_a_criar.items():
            print(f"\nVerificando tabela '{nome_tabela}'...")
            
            # Monta o SQL para criação da tabela
            sql_create = f"""
                CREATE TABLE IF NOT EXISTS {nome_tabela} (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nome TEXT NOT NULL UNIQUE { "CHECK(nome = UPPER(nome))" if info['check_upper'] else "" }
                )
            """
            cursor.execute(sql_create)
            print(f" -> Estrutura da tabela '{nome_tabela}' garantida.")

            # Popula a tabela se houver dados
            if info['dados']:
                # Transforma a lista de strings em uma lista de tuplas para o executemany
                dados_para_inserir = [(d,) for d in info['dados']]
                cursor.executemany(f"INSERT OR IGNORE INTO {nome_tabela} (nome) VALUES (?)", dados_para_inserir)
                print(f" -> {len(dados_para_inserir)} registros garantidos na tabela '{nome_tabela}'.")


# Dentro de preparar_bd.py -> preparar_banco_de_dados_completo()

        print("\nVerificando tabela de 'decisoes' para pesquisa jurisprudencial...")
        # Usamos uma VIRTUAL TABLE com o motor FTS5 para pesquisa de texto otimizada.
        cursor.execute("""
            CREATE VIRTUAL TABLE IF NOT EXISTS decisoes USING fts5(
                tema,                    -- Título ou resumo curto da decisão
                texto_decisao,           -- O texto completo da ementa ou decisão
                processo_relacionado,    -- Nº do processo que originou a decisão
                orgao_julgador,          -- Ex: TCE-RS, TJRS, STF
                data_decisao,            -- Data da decisão (formato AAAA-MM-DD para ordenação)
                palavras_chave,          -- Tags ou palavras-chave para facilitar a busca
                
                -- Opções do FTS5
                tokenize = 'porter unicode61' -- Melhora a busca em português
            );
        """)
        print(" -> Estrutura da tabela 'jurisprudencia' (FTS5) garantida.")


        # Dentro de preparar_bd.py -> preparar_banco_de_dados_completo()

        print("\nVerificando tabela FTS para a tabela 'pareceres'...")
        # Cria uma tabela FTS "sombra" para indexar o texto da tabela 'pareceres'
        # content='pareceres' -> Diz ao FTS que o conteúdo real está na tabela 'pareceres'
        # content_rowid='id' -> Liga o índice ao 'id' da tabela 'pareceres'
        cursor.execute("""
            CREATE VIRTUAL TABLE IF NOT EXISTS pareceres_fts USING fts5(
                texto_parecer,
                content='pareceres',
                content_rowid='id',
                tokenize = 'porter unicode61'
            );
        """)

        # Este gatilho (TRIGGER) automatiza a sincronização.
        # Ele será executado sempre que um registro for alterado na tabela 'pareceres'.
        print(" -> Garantindo gatilhos (triggers) de sincronização para 'pareceres_fts'...")
        cursor.execute("""
            CREATE TRIGGER IF NOT EXISTS pareceres_ai AFTER INSERT ON pareceres BEGIN
                INSERT INTO pareceres_fts(rowid, texto_parecer) VALUES (new.id, new.texto_parecer);
            END;
        """)
        cursor.execute("""
            CREATE TRIGGER IF NOT EXISTS pareceres_ad AFTER DELETE ON pareceres BEGIN
                INSERT INTO pareceres_fts(pareceres_fts, rowid, texto_parecer) VALUES ('delete', old.id, old.texto_parecer);
            END;
        """)
        cursor.execute("""
            CREATE TRIGGER IF NOT EXISTS pareceres_au AFTER UPDATE ON pareceres BEGIN
                INSERT INTO pareceres_fts(pareceres_fts, rowid, texto_parecer) VALUES ('delete', old.id, old.texto_parecer);
                INSERT INTO pareceres_fts(rowid, texto_parecer) VALUES (new.id, new.texto_parecer);
            END;
        """)
        
        print(" -> Estrutura e sincronização de 'pareceres_fts' garantidas.")

        print("\nSincronizando dados existentes com o índice de busca 'pareceres_fts'...")
        # Popula o índice FTS com todos os dados já existentes na tabela 'pareceres'
        cursor.execute("""
            INSERT INTO pareceres_fts(rowid, texto_parecer)
            SELECT id, texto_parecer FROM pareceres;
        """)
        print(" -> Sincronização concluída.")
        conn.commit()
        print("\n--- BANCO DE DADOS ATUALIZADO COM SUCESSO! ---")

    except sqlite3.Error as e:
        print(f"\n[ERRO] Ocorreu um erro no banco de dados: {e}")
    finally:
        if conn:
            conn.close()

if __name__ == '__main__':
    preparar_banco_de_dados_completo()