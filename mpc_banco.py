"""Camada de acesso, estrutura e diagnóstico do banco SQLite do MPC Parecer."""

from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any, Iterator, Mapping, Sequence


VERSAO_SCHEMA = 1

TABELAS_LOOKUP = frozenset(
    {
        "orgaos",
        "tipos_processo",
        "servico_de_auditoria",
        "cargo",
        "esclarecimentos",
        "conclusao",
        "tipo_parecer",
        "relator",
        "procurador",
    }
)

TABELAS_PRINCIPAIS = frozenset(
    {
        *TABELAS_LOOKUP,
        "pareceres",
        "jurisprudencia",
        "pareceres_fts",
        "jurisprudencia_fts",
    }
)

GATILHOS_FTS = frozenset(
    {
        "jurisprudencia_ai",
        "jurisprudencia_ad",
        "jurisprudencia_au",
        "pareceres_ai",
        "pareceres_ad",
        "pareceres_au",
    }
)

COLUNAS_ADICIONAIS_PARECERES = {
    "doc_probatoria": "TEXT",
    "falhas_sugestao_rec": "TEXT",
    "qtd_sugestao_rec": "TEXT",
    "aux_1": "TEXT",
    "aux_2": "TEXT",
    "aux_3": "TEXT",
    "aux_4": "TEXT",
    "aux_5": "TEXT",
}

# A ordem é a do comando INSERT e, por isso, constitui parte do contrato de
# persistência. Campos ausentes no dicionário são gravados como NULL.
CAMPOS_REGISTRO_PARECER = (
    "exercicio",
    "processo",
    "tipo_rag",
    "orgao",
    "servico_auditoria",
    "rag_arquivo",
    "peca_rag",
    "apontes_resumo",
    "tipo_parecer",
    "num_parecer",
    "ano_parecer",
    "relator",
    "num_proc_parecer",
    "tipo_proc_parecer",
    "ano_exercicio_parecer",
    "orgao_parecer",
    "procurador",
    "arquivo_parecer",
    "arq_analise_escl",
    "pasta",
    "peca_ae",
    "arq_esclarecimentos",
    "peca_esclarecimentos",
    "municipio",
    "tramitacao_status",
    "responsavel_tramitacao",
    "tramitacao_proc1_tipo",
    "tramitacao_proc1_num",
    "tramitacao_proc2_tipo",
    "tramitacao_proc2_num",
    "apontamento_selecionado",
    "paginas_rag",
    "paginas_escl",
    "paginas_ae",
    "voto",
    "paginas_voto",
    "peca_voto",
    "sexo_relator",
    "genero_orgao",
    "qtd_total_apontamentos",
    "falhas_com_resp",
    "qtd_com_resp",
    "falhas_sem_resp",
    "qtd_sem_resp",
    "gestor2_intimado",
    "responsaveis_json",
    "apontamentos_detalhados_json",
    "registro_id",
    "registro_data",
    "doc_probatoria",
    "falhas_sugestao_rec",
    "qtd_sugestao_rec",
    "aux_1",
    "aux_2",
    "aux_3",
    "aux_4",
    "aux_5",
)


class VersaoBancoIncompativel(RuntimeError):
    """Indica que o banco foi criado por uma versão mais nova do programa."""


class BancoMPC:
    """Serviço SQLite sem dependência de Tkinter ou da interface principal."""

    def __init__(self, caminho: str | Path):
        self.caminho = Path(caminho).expanduser()

    def _garantir_diretorio(self) -> None:
        self.caminho.parent.mkdir(parents=True, exist_ok=True)

    def _conectar(self) -> sqlite3.Connection:
        self._garantir_diretorio()
        conexao = sqlite3.connect(str(self.caminho), timeout=30)
        conexao.execute("PRAGMA busy_timeout = 5000")
        conexao.execute("PRAGMA foreign_keys = ON")
        return conexao

    @contextmanager
    def _transacao(self) -> Iterator[sqlite3.Connection]:
        conexao = self._conectar()
        try:
            with conexao:
                yield conexao
        finally:
            conexao.close()

    def migrar_jurisprudencia(self) -> str:
        """Atualiza o schema antigo, preservando integralmente os registros."""
        with self._transacao() as conexao:
            cursor = conexao.cursor()
            existe = cursor.execute(
                "SELECT 1 FROM sqlite_master "
                "WHERE type='table' AND name='jurisprudencia'"
            ).fetchone()
            if not existe:
                return "ausente"

            colunas = {
                linha[1]
                for linha in cursor.execute(
                    "PRAGMA table_info(jurisprudencia)"
                ).fetchall()
            }
            if "tipo_registro" in colunas:
                return "atualizado"

            cursor.execute("DROP TABLE IF EXISTS jurisprudencia_fts")
            cursor.execute("DROP TABLE IF EXISTS jurisprudencia_temp")
            cursor.execute(
                """
                CREATE TABLE jurisprudencia_temp (
                    tema TEXT NOT NULL,
                    texto_decisao TEXT NOT NULL,
                    processo_relacionado TEXT,
                    orgao_julgador TEXT,
                    data_decisao TEXT,
                    palavras_chave TEXT,
                    tipo_registro TEXT NOT NULL DEFAULT 'Integral'
                )
                """
            )
            cursor.execute(
                """
                INSERT INTO jurisprudencia_temp (
                    tema, texto_decisao, processo_relacionado,
                    orgao_julgador, data_decisao, palavras_chave
                )
                SELECT tema, texto_decisao, processo_relacionado,
                       orgao_julgador, data_decisao, palavras_chave
                FROM jurisprudencia
                """
            )
            cursor.execute("DROP TABLE jurisprudencia")
            cursor.execute(
                "ALTER TABLE jurisprudencia_temp RENAME TO jurisprudencia"
            )
        return "migrado"

    def inicializar(self) -> dict[str, Any]:
        """Cria tabelas, completa colunas e sincroniza os índices FTS."""
        if self.caminho.is_file():
            with self._transacao() as conexao:
                versao_encontrada = conexao.execute(
                    "PRAGMA user_version"
                ).fetchone()[0]
            if versao_encontrada > VERSAO_SCHEMA:
                raise VersaoBancoIncompativel(
                    "O banco usa o schema v"
                    f"{versao_encontrada}, mas esta versão do programa "
                    f"reconhece somente até o schema v{VERSAO_SCHEMA}. "
                    "Atualize o MPC Parecer antes de abrir este banco."
                )

        # Torna a inicialização segura mesmo quando chamada isoladamente.
        self.migrar_jurisprudencia()
        with self._transacao() as conexao:
            cursor = conexao.cursor()
            for tabela in sorted(TABELAS_LOOKUP):
                cursor.execute(
                    f"CREATE TABLE IF NOT EXISTS {tabela} ("
                    "id INTEGER PRIMARY KEY AUTOINCREMENT, "
                    "nome TEXT NOT NULL UNIQUE)"
                )

            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS pareceres (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    texto_parecer TEXT,
                    modo_insercao TEXT NOT NULL,
                    data_hora_insercao TEXT NOT NULL,
                    exercicio TEXT, processo TEXT, tipo_rag TEXT, orgao TEXT,
                    servico_auditoria TEXT, rag_arquivo TEXT, peca_rag TEXT,
                    apontes_resumo TEXT, tipo_parecer TEXT, num_parecer TEXT,
                    ano_parecer TEXT, relator TEXT, num_proc_parecer TEXT,
                    tipo_proc_parecer TEXT, ano_exercicio_parecer TEXT,
                    orgao_parecer TEXT, procurador TEXT, arquivo_parecer TEXT,
                    advogados TEXT, problema_procuracao TEXT,
                    arq_analise_escl TEXT, pasta TEXT, peca_ae TEXT,
                    arq_esclarecimentos TEXT, peca_esclarecimentos TEXT,
                    municipio TEXT, tramitacao_status TEXT,
                    responsavel_tramitacao TEXT, tramitacao_proc1_tipo TEXT,
                    tramitacao_proc1_num TEXT, tramitacao_proc2_tipo TEXT,
                    tramitacao_proc2_num TEXT, apontamento_selecionado TEXT,
                    paginas_rag TEXT, paginas_escl TEXT, paginas_ae TEXT,
                    voto TEXT, paginas_voto TEXT, peca_voto TEXT,
                    sexo_relator TEXT, genero_orgao TEXT,
                    qtd_total_apontamentos TEXT, falhas_com_resp TEXT,
                    qtd_com_resp TEXT, falhas_sem_resp TEXT, qtd_sem_resp TEXT,
                    gestor2_intimado TEXT, responsaveis_json TEXT,
                    apontamentos_detalhados_json TEXT, registro_id TEXT,
                    registro_data TEXT, doc_probatoria TEXT,
                    falhas_sugestao_rec TEXT, qtd_sugestao_rec TEXT,
                    aux_1 TEXT, aux_2 TEXT, aux_3 TEXT, aux_4 TEXT, aux_5 TEXT
                )
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS jurisprudencia (
                    tema TEXT NOT NULL,
                    texto_decisao TEXT NOT NULL,
                    processo_relacionado TEXT,
                    orgao_julgador TEXT,
                    data_decisao TEXT,
                    palavras_chave TEXT,
                    tipo_registro TEXT NOT NULL DEFAULT 'Integral'
                )
                """
            )

            colunas = {
                linha[1]
                for linha in cursor.execute("PRAGMA table_info(pareceres)")
            }
            for nome, tipo in COLUNAS_ADICIONAIS_PARECERES.items():
                if nome not in colunas:
                    cursor.execute(
                        f"ALTER TABLE pareceres ADD COLUMN {nome} {tipo}"
                    )

            self._criar_fts(cursor)
            versao_atual = cursor.execute("PRAGMA user_version").fetchone()[0]
            if versao_atual < VERSAO_SCHEMA:
                cursor.execute(f"PRAGMA user_version = {VERSAO_SCHEMA}")
        return self.diagnosticar()

    @staticmethod
    def _criar_fts(cursor: sqlite3.Cursor) -> None:
        cursor.execute(
            """
            CREATE VIRTUAL TABLE IF NOT EXISTS jurisprudencia_fts USING fts5(
                tema, texto_decisao, palavras_chave,
                content='jurisprudencia', content_rowid='rowid'
            )
            """
        )
        cursor.execute(
            """
            CREATE TRIGGER IF NOT EXISTS jurisprudencia_ai
            AFTER INSERT ON jurisprudencia BEGIN
                INSERT INTO jurisprudencia_fts(
                    rowid, tema, texto_decisao, palavras_chave
                ) VALUES (new.rowid, new.tema, new.texto_decisao,
                          new.palavras_chave);
            END
            """
        )
        cursor.execute(
            """
            CREATE TRIGGER IF NOT EXISTS jurisprudencia_ad
            AFTER DELETE ON jurisprudencia BEGIN
                INSERT INTO jurisprudencia_fts(
                    jurisprudencia_fts, rowid, tema, texto_decisao,
                    palavras_chave
                ) VALUES ('delete', old.rowid, old.tema,
                          old.texto_decisao, old.palavras_chave);
            END
            """
        )
        cursor.execute(
            """
            CREATE TRIGGER IF NOT EXISTS jurisprudencia_au
            AFTER UPDATE ON jurisprudencia BEGIN
                INSERT INTO jurisprudencia_fts(
                    jurisprudencia_fts, rowid, tema, texto_decisao,
                    palavras_chave
                ) VALUES ('delete', old.rowid, old.tema,
                          old.texto_decisao, old.palavras_chave);
                INSERT INTO jurisprudencia_fts(
                    rowid, tema, texto_decisao, palavras_chave
                ) VALUES (new.rowid, new.tema, new.texto_decisao,
                          new.palavras_chave);
            END
            """
        )
        cursor.execute(
            """
            CREATE VIRTUAL TABLE IF NOT EXISTS pareceres_fts USING fts5(
                texto_parecer, content='pareceres', content_rowid='id'
            )
            """
        )
        cursor.execute(
            """
            CREATE TRIGGER IF NOT EXISTS pareceres_ai
            AFTER INSERT ON pareceres BEGIN
                INSERT INTO pareceres_fts(rowid, texto_parecer)
                VALUES (new.id, new.texto_parecer);
            END
            """
        )
        cursor.execute(
            """
            CREATE TRIGGER IF NOT EXISTS pareceres_ad
            AFTER DELETE ON pareceres BEGIN
                INSERT INTO pareceres_fts(
                    pareceres_fts, rowid, texto_parecer
                ) VALUES ('delete', old.id, old.texto_parecer);
            END
            """
        )
        cursor.execute(
            """
            CREATE TRIGGER IF NOT EXISTS pareceres_au
            AFTER UPDATE ON pareceres BEGIN
                INSERT INTO pareceres_fts(
                    pareceres_fts, rowid, texto_parecer
                ) VALUES ('delete', old.id, old.texto_parecer);
                INSERT INTO pareceres_fts(rowid, texto_parecer)
                VALUES (new.id, new.texto_parecer);
            END
            """
        )
        cursor.execute(
            "INSERT INTO jurisprudencia_fts(jurisprudencia_fts) "
            "VALUES('rebuild')"
        )
        cursor.execute(
            "INSERT INTO pareceres_fts(pareceres_fts) VALUES('rebuild')"
        )

    def listar_lookup(self, tabela: str) -> list[str]:
        if tabela not in TABELAS_LOOKUP:
            raise ValueError(f"Tabela de consulta não permitida: {tabela}")
        with self._transacao() as conexao:
            linhas = conexao.execute(
                f"SELECT nome FROM {tabela} ORDER BY nome"
            ).fetchall()
        return [linha[0] for linha in linhas]

    def adicionar_lookup(self, tabela: str, nome: str) -> bool:
        """Insere um item; retorna False quando ele já existia."""
        if tabela not in TABELAS_LOOKUP:
            raise ValueError(f"Tabela de consulta não permitida: {tabela}")
        try:
            with self._transacao() as conexao:
                conexao.execute(
                    f"INSERT INTO {tabela} (nome) VALUES (?)",
                    (str(nome).strip(),),
                )
        except sqlite3.IntegrityError:
            return False
        return True

    def inserir_jurisprudencia(self, dados: Mapping[str, Any]) -> int:
        valores = (
            dados.get("tema"),
            dados.get("texto_decisao"),
            dados.get("processo", ""),
            dados.get("orgao", ""),
            dados.get("data", ""),
            dados.get("palavras", ""),
            dados.get("tipo_registro", "Integral"),
        )
        with self._transacao() as conexao:
            cursor = conexao.execute(
                """
                INSERT INTO jurisprudencia (
                    tema, texto_decisao, processo_relacionado,
                    orgao_julgador, data_decisao, palavras_chave,
                    tipo_registro
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                valores,
            )
            return int(cursor.lastrowid)

    def pesquisar_decisoes(
        self,
        termo: str,
        *,
        incluir_pareceres: bool = True,
    ) -> list[tuple[Any, ...]]:
        termo_fts = f"{str(termo).strip()}*"
        if incluir_pareceres:
            consulta = """
                SELECT id, fonte, tema, processo, data FROM (
                    SELECT j.rowid AS id,
                           CASE j.tipo_registro
                               WHEN 'Nota' THEN 'Nota de Jurisprudência'
                               ELSE 'Jurisprudência'
                           END AS fonte,
                           j.tema,
                           j.processo_relacionado AS processo,
                           j.data_decisao AS data,
                           jurisprudencia_fts.rank AS relevancia
                    FROM jurisprudencia j
                    JOIN jurisprudencia_fts
                      ON j.rowid = jurisprudencia_fts.rowid
                    WHERE jurisprudencia_fts MATCH ?
                    UNION ALL
                    SELECT p.id,
                           'Parecer Antigo' AS fonte,
                           substr(p.texto_parecer, 1, 150) AS tema,
                           p.processo,
                           p.data_hora_insercao AS data,
                           pareceres_fts.rank AS relevancia
                    FROM pareceres p
                    JOIN pareceres_fts ON p.id = pareceres_fts.rowid
                    WHERE pareceres_fts MATCH ?
                ) ORDER BY relevancia
            """
            parametros: Sequence[Any] = (termo_fts, termo_fts)
        else:
            consulta = """
                SELECT j.rowid AS id,
                       CASE j.tipo_registro
                           WHEN 'Nota' THEN 'Nota de Jurisprudência'
                           ELSE 'Jurisprudência'
                       END AS fonte,
                       j.tema,
                       j.processo_relacionado AS processo,
                       j.data_decisao AS data
                FROM jurisprudencia j
                JOIN jurisprudencia_fts
                  ON j.rowid = jurisprudencia_fts.rowid
                WHERE jurisprudencia_fts MATCH ?
                ORDER BY jurisprudencia_fts.rank
            """
            parametros = (termo_fts,)
        with self._transacao() as conexao:
            return conexao.execute(consulta, parametros).fetchall()

    def obter_texto_decisao(self, item_id: int | str, fonte: str) -> str | None:
        if fonte in {"Jurisprudência", "Nota de Jurisprudência"}:
            consulta = (
                "SELECT texto_decisao FROM jurisprudencia WHERE rowid = ?"
            )
        elif fonte == "Parecer Antigo":
            consulta = "SELECT texto_parecer FROM pareceres WHERE id = ?"
        else:
            raise ValueError(f"Fonte de decisão desconhecida: {fonte}")
        with self._transacao() as conexao:
            resultado = conexao.execute(consulta, (item_id,)).fetchone()
        return None if resultado is None else resultado[0]

    def inserir_parecer(
        self,
        dados: Mapping[str, Any],
        texto_completo: str,
        modo: str,
        *,
        data_hora: str | None = None,
    ) -> int:
        colunas = (
            "texto_parecer",
            "modo_insercao",
            "data_hora_insercao",
            *CAMPOS_REGISTRO_PARECER,
        )
        valores = (
            texto_completo,
            modo,
            data_hora or datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            *(dados.get(campo) for campo in CAMPOS_REGISTRO_PARECER),
        )
        marcadores = ", ".join("?" for _ in colunas)
        with self._transacao() as conexao:
            cursor = conexao.execute(
                f"INSERT INTO pareceres ({', '.join(colunas)}) "
                f"VALUES ({marcadores})",
                valores,
            )
            return int(cursor.lastrowid)

    def diagnosticar(self) -> dict[str, Any]:
        """Retorna uma fotografia segura da estrutura e integridade do banco."""
        resultado: dict[str, Any] = {
            "caminho": str(self.caminho),
            "existe": self.caminho.is_file(),
            "integridade": "ausente",
            "versao_schema": 0,
            "tabelas_ausentes": sorted(TABELAS_PRINCIPAIS),
            "gatilhos_ausentes": sorted(GATILHOS_FTS),
            "colunas_pareceres_ausentes": sorted(
                COLUNAS_ADICIONAIS_PARECERES
            ),
            "pronto": False,
        }
        if not resultado["existe"]:
            return resultado
        conexao = self._conectar()
        try:
            tabelas = {
                linha[0]
                for linha in conexao.execute(
                    "SELECT name FROM sqlite_master "
                    "WHERE type IN ('table', 'view')"
                ).fetchall()
            }
            gatilhos = {
                linha[0]
                for linha in conexao.execute(
                    "SELECT name FROM sqlite_master WHERE type='trigger'"
                ).fetchall()
            }
            colunas = {
                linha[1]
                for linha in conexao.execute(
                    "PRAGMA table_info(pareceres)"
                ).fetchall()
            }
            integridade = conexao.execute("PRAGMA quick_check").fetchone()[0]
            versao_schema = conexao.execute("PRAGMA user_version").fetchone()[0]
            resultado.update(
                {
                    "integridade": integridade,
                    "versao_schema": versao_schema,
                    "tabelas_ausentes": sorted(TABELAS_PRINCIPAIS - tabelas),
                    "gatilhos_ausentes": sorted(GATILHOS_FTS - gatilhos),
                    "colunas_pareceres_ausentes": sorted(
                        set(COLUNAS_ADICIONAIS_PARECERES) - colunas
                    ),
                }
            )
            resultado["pronto"] = (
                integridade == "ok"
                and versao_schema >= VERSAO_SCHEMA
                and not resultado["tabelas_ausentes"]
                and not resultado["gatilhos_ausentes"]
                and not resultado["colunas_pareceres_ausentes"]
            )
            return resultado
        finally:
            conexao.close()
