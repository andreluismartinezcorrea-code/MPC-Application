import sqlite3
import os
from tkinter import messagebox

DB_PATH = r"E:\TCE\Database\BD-MPC.db"

def _run_migration():
    """Executes the database schema migration for the 'jurisprudencia' table.

    This function updates the schema of the 'jurisprudencia' table by using a
    "copy and replace" strategy to avoid errors with virtual tables. It adds
    the 'tipo_registro' column if it does not exist.
    """
    conn = None
    try:
        print("[DB_MIGRATION] Verificando necessidade de migração de schema...")
        os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        # 1. Verifica se a tabela 'jurisprudencia' existe
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='jurisprudencia'")
        if not cursor.fetchone():
            print("[DB_MIGRATION] Tabela 'jurisprudencia' não existe. Nenhuma migração necessária.")
            return

        # 2. Se a tabela existe, verifica se a coluna 'tipo_registro' existe
        cursor.execute("PRAGMA table_info(jurisprudencia)")
        columns = [info[1] for info in cursor.fetchall()]
        
        if 'tipo_registro' in columns:
            print("[DB_MIGRATION] Schema 'jurisprudencia' já está atualizado.")
            return

        # 3. Se a coluna NÃO existe, executa a migração (Estratégia "Copy and Replace")
        print("[DB_MIGRATION] Schema antigo detectado. Iniciando migração (estratégia 'Copy/Replace')...")
        
        # 3a. Remove a FTS antiga para destravar tudo.
        print("[DB_MIGRATION] Removendo tabela 'jurisprudencia_fts' antiga...")
        cursor.execute("DROP TABLE IF EXISTS jurisprudencia_fts")
        conn.commit()

        # 3b. Cria uma tabela temporária com o NOVO schema
        print("[DB_MIGRATION] Criando tabela 'jurisprudencia_temp' com novo schema...")
        cursor.execute('''
            CREATE TABLE jurisprudencia_temp (
                tema TEXT NOT NULL,
                texto_decisao TEXT NOT NULL,
                processo_relacionado TEXT,
                orgao_julgador TEXT,
                data_decisao TEXT,
                palavras_chave TEXT,
                tipo_registro TEXT NOT NULL DEFAULT 'Integral'
            )
        ''')

        # 3c. Copia os dados da tabela antiga para a nova
        # (Note que 'tipo_registro' receberá o valor DEFAULT 'Integral' para todos os registros antigos)
        print("[DB_MIGRATION] Copiando dados da tabela antiga para a nova...")
        cursor.execute('''
            INSERT INTO jurisprudencia_temp (
                tema, texto_decisao, processo_relacionado, 
                orgao_julgador, data_decisao, palavras_chave
            ) 
            SELECT 
                tema, texto_decisao, processo_relacionado, 
                orgao_julgador, data_decisao, palavras_chave 
            FROM jurisprudencia
        ''')
        
        # 3d. Remove a tabela antiga
        print("[DB_MIGRATION] Removendo tabela 'jurisprudencia' antiga...")
        cursor.execute("DROP TABLE jurisprudencia")

        # 3e. Renomeia a tabela nova para o nome original
        print("[DB_MIGRATION] Renomeando 'jurisprudencia_temp' para 'jurisprudencia'...")
        cursor.execute("ALTER TABLE jurisprudencia_temp RENAME TO jurisprudencia")
        
        conn.commit()
        print("[DB_MIGRATION] Migração do schema concluída com sucesso.")

    except sqlite3.Error as e:
        if conn:
            conn.rollback() # Desfaz qualquer alteração parcial se um erro ocorrer
        print(f"[DB_MIGRATION_ERROR] Erro durante a migração do schema: {e}")
        messagebox.showerror("Erro de Migração do BD", f"Falha ao atualizar o banco de dados:\n{e}")
    finally:
        if conn:
            conn.close()

def inicializar_bd():
    """Initializes the database and ensures the correct schema.

    This function guarantees that the database and its tables exist with the
    correct schema. It uses a "drop and recreate" strategy for Full-Text Search
    (FTS) indexes and triggers to ensure they are correctly recreated and
    synchronized with existing data upon each initialization.
    """
    print("[DB_INFO] Verificando e inicializando o banco de dados...")
    
    # Garante que o diretório do banco de dados exista.
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    
    conn = None
    try:
        # Conecta ao arquivo do banco de dados (cria o arquivo se não existir)
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        # --- Tabela PARECERES ---
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS pareceres (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                texto_parecer TEXT,
                modo_insercao TEXT NOT NULL,
                data_hora_insercao TEXT NOT NULL,
                
                -- Campos da Aba Principal
                exercicio TEXT, processo TEXT, tipo_rag TEXT, orgao TEXT, servico_auditoria TEXT,
                rag_arquivo TEXT, peca_rag TEXT, apontes_resumo TEXT,
                
                -- Campos do Parecer MPC
                tipo_parecer TEXT, num_parecer TEXT, ano_parecer TEXT, relator TEXT,
                num_proc_parecer TEXT, tipo_proc_parecer TEXT, ano_exercicio_parecer TEXT,
                orgao_parecer TEXT, procurador TEXT, arquivo_parecer TEXT,
                
                -- Campos da Análise de Esclarecimentos
                advogados TEXT, problema_procuracao TEXT, arq_analise_escl TEXT, pasta TEXT,
                peca_ae TEXT, arq_esclarecimentos TEXT, peca_esclarecimentos TEXT, municipio TEXT,
                tramitacao_status TEXT, responsavel_tramitacao TEXT, tramitacao_proc1_tipo TEXT, tramitacao_proc1_num TEXT,
                tramitacao_proc2_tipo TEXT, tramitacao_proc2_num TEXT,
                
                -- Campos da Aba Apontamentos (Análise e Voto)
                apontamento_selecionado TEXT, paginas_rag TEXT, paginas_escl TEXT, paginas_ae TEXT,
                voto TEXT, paginas_voto TEXT, peca_voto TEXT,
                
                -- Campos da Aba Parâmetros
                sexo_relator TEXT, genero_orgao TEXT, qtd_total_apontamentos TEXT,
                falhas_com_resp TEXT, qtd_com_resp TEXT, falhas_sem_resp TEXT, qtd_sem_resp TEXT,
                gestor2_intimado TEXT,
                
                -- Dados Estruturados (JSON)
                responsaveis_json TEXT,
                apontamentos_detalhados_json TEXT,
                
                -- Dados de Registro
                registro_id TEXT,
                registro_data TEXT
            )
        ''')
        
        # --- Tabela JURISPRUDENCIA ---
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS jurisprudencia (
                tema TEXT NOT NULL,
                texto_decisao TEXT NOT NULL,
                processo_relacionado TEXT,
                orgao_julgador TEXT,
                data_decisao TEXT,
                palavras_chave TEXT,
                tipo_registro TEXT NOT NULL DEFAULT 'Integral'
            )
        ''')

        # ==============================================================================
        # ### INÍCIO DO BLOCO DE CORREÇÃO (ESTRATÉGIA DROP/RECREATE) ###
        # ==============================================================================
        
        # --- 1. Recriação do índice FTS para 'jurisprudencia' ---
        
        print("[DB_SYNC] Removendo índice FTS (jurisprudencia_fts) e gatilhos antigos para recriação...")
        cursor.execute("DROP TRIGGER IF EXISTS jurisprudencia_ai")
        cursor.execute("DROP TRIGGER IF EXISTS jurisprudencia_ad")
        cursor.execute("DROP TRIGGER IF EXISTS jurisprudencia_au")
        cursor.execute("DROP TABLE IF EXISTS jurisprudencia_fts")
        
        print("[DB_SYNC] Recriando tabela 'jurisprudencia_fts'...")
        cursor.execute('''
            CREATE VIRTUAL TABLE jurisprudencia_fts USING fts5(
                tema,
                texto_decisao,
                palavras_chave,
                content='jurisprudencia',
                content_rowid='rowid'
            )
        ''')
        
        print("[DB_SYNC] Recriando gatilhos para 'jurisprudencia'...")
        cursor.execute('''
            CREATE TRIGGER jurisprudencia_ai AFTER INSERT ON jurisprudencia BEGIN
                INSERT INTO jurisprudencia_fts(rowid, tema, texto_decisao, palavras_chave) 
                VALUES (new.rowid, new.tema, new.texto_decisao, new.palavras_chave);
            END;
        ''')
        cursor.execute('''
            CREATE TRIGGER jurisprudencia_ad AFTER DELETE ON jurisprudencia BEGIN
                INSERT INTO jurisprudencia_fts(jurisprudencia_fts, rowid, tema, texto_decisao, palavras_chave) 
                VALUES ('delete', old.rowid, old.tema, old.texto_decisao, old.palavras_chave);
            END;
        ''')
        cursor.execute('''
            CREATE TRIGGER jurisprudencia_au AFTER UPDATE ON jurisprudencia BEGIN
                INSERT INTO jurisprudencia_fts(jurisprudencia_fts, rowid, tema, texto_decisao, palavras_chave) 
                VALUES ('delete', old.rowid, old.tema, old.texto_decisao, old.palavras_chave);
                INSERT INTO jurisprudencia_fts(rowid, tema, texto_decisao, palavras_chave) 
                VALUES (new.rowid, new.tema, new.texto_decisao, new.palavras_chave);
            END;
        ''')
        
        print("[DB_SYNC] Reconstruindo o índice 'jurisprudencia_fts' com todos os dados...")
        # Reconstrói o índice FTS com TODOS os dados da tabela principal
        cursor.execute('''
            INSERT INTO jurisprudencia_fts (rowid, tema, texto_decisao, palavras_chave)
            SELECT rowid, tema, texto_decisao, palavras_chave FROM jurisprudencia;
        ''')
        print("[DB_SYNC] Sincronização de 'jurisprudencia_fts' concluída.")
        
        # --- 2. Recriação do índice FTS para 'pareceres' (para consistência) ---

        print("[DB_SYNC] Removendo índice FTS (pareceres_fts) e gatilhos antigos para recriação...")
        cursor.execute("DROP TRIGGER IF EXISTS pareceres_ai")
        cursor.execute("DROP TRIGGER IF EXISTS pareceres_ad")
        cursor.execute("DROP TRIGGER IF EXISTS pareceres_au")
        cursor.execute("DROP TABLE IF EXISTS pareceres_fts")
        
        print("[DB_SYNC] Recriando tabela 'pareceres_fts'...")
        cursor.execute('''
            CREATE VIRTUAL TABLE pareceres_fts USING fts5(
                texto_parecer,
                content='pareceres',
                content_rowid='id'
            )
        ''')
        
        print("[DB_SYNC] Recriando gatilhos para 'pareceres'...")
        cursor.execute('''
            CREATE TRIGGER pareceres_ai AFTER INSERT ON pareceres BEGIN
                INSERT INTO pareceres_fts(rowid, texto_parecer) VALUES (new.id, new.texto_parecer);
            END;
        ''')
        cursor.execute('''
            CREATE TRIGGER pareceres_ad AFTER DELETE ON pareceres BEGIN
                INSERT INTO pareceres_fts(pareceres_fts, rowid, texto_parecer) VALUES ('delete', old.id, old.texto_parecer);
            END;
        ''')
        cursor.execute('''
            CREATE TRIGGER pareceres_au AFTER UPDATE ON pareceres BEGIN
                INSERT INTO pareceres_fts(pareceres_fts, rowid, texto_parecer) VALUES ('delete', old.id, old.texto_parecer);
                INSERT INTO pareceres_fts(rowid, texto_parecer) VALUES (new.id, new.texto_parecer);
            END;
        ''')
        
        print("[DB_SYNC] Reconstruindo o índice 'pareceres_fts' com todos os dados...")
        # Reconstrói o índice FTS com TODOS os dados da tabela principal
        cursor.execute('''
            INSERT INTO pareceres_fts (rowid, texto_parecer)
            SELECT id, texto_parecer FROM pareceres;
        ''')
        print("[DB_SYNC] Sincronização de 'pareceres_fts' concluída.")

        # ==============================================================================
        # ### FIM DO BLOCO DE CORREÇÃO ###
        # ==============================================================================

        conn.commit()
        print("[DB_SUCCESS] Banco de dados e tabelas 'pareceres' e 'jurisprudencia' prontos para uso.")

    except sqlite3.Error as e:
        print(f"[DB_ERROR] Erro ao inicializar o banco de dados: {e}")
        messagebox.showerror("Erro de Banco de Dados", f"Não foi possível criar ou verificar o banco de dados:\n{e}")
    finally:
        if conn:
            conn.close()

def carregar_dados_lookup(nome_tabela):
    """Loads a list of names from a lookup table in the database.

    Args:
        nome_tabela (str): The name of the table to query.

    Returns:
        list: A list of names from the specified table, or an empty list if an
              error occurs.
    """
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute(f"SELECT nome FROM {nome_tabela} ORDER BY nome")
        return [item[0] for item in cursor.fetchall()]
    except sqlite3.Error as e:
        messagebox.showerror("Erro de Leitura do BD", f"Não foi possível carregar dados da tabela '{nome_tabela}':\n{e}")
        return []
    finally:
        if conn:
            conn.close()

def adicionar_novo_orgao_bd(nome_orgao):
    """Adds a new government body to the database programmatically.

    Args:
        nome_orgao (str): The name of the government body to add.

    Returns:
        bool: True if the operation was successful (or if the body already
              exists), False otherwise.
    """
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        # O nome já vem processado, mas garantimos o UPPER aqui
        cursor.execute("INSERT INTO orgaos (nome) VALUES (?)", (nome_orgao.upper(),))
        conn.commit()
        print(f"[DB_SUCCESS] Novo órgão '{nome_orgao}' adicionado via automação.")
        return True
    except sqlite3.IntegrityError:
        print(f"[DB_WARN] Tentativa de adicionar órgão duplicado: '{nome_orgao}'.")
        # Mesmo que já exista, consideramos um sucesso para o fluxo.
        return True
    except sqlite3.Error as e:
        print(f"[DB_ERROR] Falha ao adicionar novo órgão: {e}")
        return False
    finally:
        if conn:
            conn.close()
