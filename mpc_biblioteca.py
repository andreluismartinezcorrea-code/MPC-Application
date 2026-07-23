"""Biblioteca jurídica local baseada em pastas e índice SQLite FTS5.

Os arquivos originais permanecem nas pastas escolhidas pelo usuário. O banco
armazena somente metadados e texto extraído para permitir pesquisas rápidas.
Nenhum conteúdo é enviado a serviços externos.
"""

from __future__ import annotations

import os
import re
import sqlite3
import zipfile
import xml.etree.ElementTree as ET
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterable


EXTENSOES_SUPORTADAS = frozenset({".doc", ".docx", ".docm", ".pdf"})
CATEGORIAS_ACERVO = (
    "Pareceres",
    "Legislação",
    "Decisões e Acórdãos",
    "Relatórios e Votos",
    "Outros",
)
MODOS_PESQUISA = ("Todas as palavras", "Expressão exata")
VERSAO_EXTRATOR = 2

_NAMESPACE_WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_TAG_PARAGRAFO = f"{{{_NAMESPACE_WORD}}}p"
_TAG_TEXTO = f"{{{_NAMESPACE_WORD}}}t"
_TAG_TABULACAO = f"{{{_NAMESPACE_WORD}}}tab"
_TAGS_QUEBRA = {
    f"{{{_NAMESPACE_WORD}}}br",
    f"{{{_NAMESPACE_WORD}}}cr",
}


@dataclass(slots=True)
class ResultadoIndexacao:
    acervos: int = 0
    encontrados: int = 0
    novos: int = 0
    atualizados: int = 0
    inalterados: int = 0
    removidos: int = 0
    sem_texto: int = 0
    erros: int = 0
    indisponiveis: int = 0
    avisos: list[str] = field(default_factory=list)

    def incorporar(self, outro: "ResultadoIndexacao") -> None:
        for campo in (
            "acervos", "encontrados", "novos", "atualizados", "inalterados",
            "removidos", "sem_texto", "erros", "indisponiveis",
        ):
            setattr(self, campo, getattr(self, campo) + getattr(outro, campo))
        self.avisos.extend(outro.avisos)

    def resumo(self) -> str:
        return (
            f"{self.encontrados} arquivo(s) localizado(s): {self.novos} novo(s), "
            f"{self.atualizados} atualizado(s), {self.inalterados} inalterado(s), "
            f"{self.removidos} removido(s) do índice, {self.sem_texto} sem texto "
            f"pesquisável e {self.erros} com erro."
        )


class ExtratorTextoLocal:
    """Extrai texto de PDF/Word; Word legado usa automação local opcional."""

    def __init__(self) -> None:
        self._word = None
        self._com_inicializado = False

    @staticmethod
    def _docx(caminho: Path) -> str:
        """Lê o OOXML diretamente, incluindo controles de conteúdo ``w:sdt``.

        ``python-docx`` não expõe em ``Document.paragraphs`` todos os
        parágrafos encapsulados em controles de conteúdo. Os elementos
        ``w:t`` presentes no XML, porém, representam tanto o texto comum como
        o texto desses controles. A leitura direta também evita depender do
        Word instalado para arquivos DOCX e DOCM.
        """

        partes: list[str] = []
        with zipfile.ZipFile(caminho) as pacote:
            nomes = pacote.namelist()
            partes_word = ["word/document.xml"]
            partes_word.extend(
                nome
                for nome in nomes
                if re.fullmatch(
                    r"word/(?:header\d+|footer\d+|footnotes|endnotes|comments)\.xml",
                    nome,
                    flags=re.IGNORECASE,
                )
            )

            for nome_parte in dict.fromkeys(partes_word):
                if nome_parte not in nomes:
                    continue
                raiz = ET.fromstring(pacote.read(nome_parte))
                encontrou_paragrafo = False
                for paragrafo in raiz.iter(_TAG_PARAGRAFO):
                    encontrou_paragrafo = True
                    fragmentos: list[str] = []
                    for elemento in paragrafo.iter():
                        if elemento.tag == _TAG_TEXTO and elemento.text:
                            fragmentos.append(elemento.text)
                        elif elemento.tag == _TAG_TABULACAO:
                            fragmentos.append("\t")
                        elif elemento.tag in _TAGS_QUEBRA:
                            fragmentos.append("\n")
                    texto = "".join(fragmentos).strip()
                    if texto:
                        partes.append(texto)

                # Proteção para uma parte OOXML atípica sem ``w:p``.
                if not encontrou_paragrafo:
                    texto = "".join(
                        elemento.text or ""
                        for elemento in raiz.iter(_TAG_TEXTO)
                    ).strip()
                    if texto:
                        partes.append(texto)
        return "\n".join(partes)

    @staticmethod
    def _pdf(caminho: Path) -> str:
        import fitz

        partes = []
        with fitz.open(str(caminho)) as documento:
            for pagina in documento:
                texto = pagina.get_text("text").strip()
                if texto:
                    partes.append(texto)
        return "\n\n".join(partes)

    def _obter_word(self):
        if self._word is not None:
            return self._word
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        self._com_inicializado = True
        self._word = win32com.client.DispatchEx("Word.Application")
        self._word.Visible = False
        self._word.DisplayAlerts = 0
        try:
            self._word.AutomationSecurity = 3
        except Exception:
            pass
        return self._word

    def _doc_legado(self, caminho: Path) -> str:
        word = self._obter_word()
        documento = word.Documents.Open(
            str(caminho),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )
        try:
            return str(documento.Content.Text or "")
        finally:
            documento.Close(SaveChanges=False)

    def extrair(self, caminho: str | Path) -> str:
        arquivo = Path(caminho)
        extensao = arquivo.suffix.casefold()
        if extensao in {".docx", ".docm"}:
            return self._docx(arquivo)
        if extensao == ".pdf":
            return self._pdf(arquivo)
        if extensao == ".doc":
            return self._doc_legado(arquivo)
        raise ValueError(f"Formato não suportado: {extensao}")

    def fechar(self) -> None:
        if self._word is not None:
            try:
                self._word.Quit()
            except Exception:
                # O processo de indexação já encerrou seus documentos. Uma
                # falha tardia do Word ao sair não deve invalidar o índice.
                pass
            finally:
                self._word = None
        if self._com_inicializado:
            import pythoncom

            try:
                pythoncom.CoUninitialize()
            finally:
                self._com_inicializado = False

    def __enter__(self) -> "ExtratorTextoLocal":
        return self

    def __exit__(self, _tipo, _valor, _traceback) -> None:
        self.fechar()


class BibliotecaLocal:
    """Gerencia fontes locais, indexação incremental e pesquisa textual."""

    def __init__(self, caminho_banco: str | Path):
        self.caminho_banco = Path(caminho_banco).expanduser()

    def _conectar(self) -> sqlite3.Connection:
        self.caminho_banco.parent.mkdir(parents=True, exist_ok=True)
        conexao = sqlite3.connect(str(self.caminho_banco), timeout=30)
        conexao.row_factory = sqlite3.Row
        conexao.execute("PRAGMA busy_timeout = 5000")
        conexao.execute("PRAGMA foreign_keys = ON")
        return conexao

    @contextmanager
    def _transacao(self):
        conexao = self._conectar()
        try:
            with conexao:
                yield conexao
        finally:
            conexao.close()

    def inicializar(self) -> None:
        with self._transacao() as conexao:
            fts_existia = conexao.execute(
                "SELECT 1 FROM sqlite_master WHERE type='table' "
                "AND name='documentos_locais_fts'"
            ).fetchone() is not None
            conexao.executescript(
                """
                CREATE TABLE IF NOT EXISTS categorias_acervo_local (
                    nome TEXT PRIMARY KEY COLLATE NOCASE,
                    ordem INTEGER NOT NULL DEFAULT 0
                );

                CREATE TABLE IF NOT EXISTS acervos_locais (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    caminho TEXT NOT NULL UNIQUE,
                    categoria TEXT NOT NULL,
                    ativo INTEGER NOT NULL DEFAULT 1,
                    data_cadastro TEXT NOT NULL,
                    ultima_indexacao TEXT,
                    status TEXT NOT NULL DEFAULT 'Aguardando indexação',
                    mensagem TEXT NOT NULL DEFAULT ''
                );

                CREATE TABLE IF NOT EXISTS documentos_locais (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    acervo_id INTEGER NOT NULL,
                    caminho TEXT NOT NULL,
                    nome TEXT NOT NULL,
                    extensao TEXT NOT NULL,
                    tamanho INTEGER NOT NULL,
                    modificado_ns INTEGER NOT NULL,
                    texto TEXT NOT NULL DEFAULT '',
                    indexado_em TEXT NOT NULL,
                    status TEXT NOT NULL,
                    erro TEXT NOT NULL DEFAULT '',
                    versao_extrator INTEGER NOT NULL DEFAULT 1,
                    FOREIGN KEY(acervo_id) REFERENCES acervos_locais(id)
                        ON DELETE CASCADE,
                    UNIQUE(acervo_id, caminho)
                );

                CREATE INDEX IF NOT EXISTS idx_documentos_locais_acervo
                    ON documentos_locais(acervo_id);
                CREATE INDEX IF NOT EXISTS idx_documentos_locais_caminho
                    ON documentos_locais(caminho);

                CREATE VIRTUAL TABLE IF NOT EXISTS documentos_locais_fts
                    USING fts5(
                        nome, texto,
                        content='documentos_locais', content_rowid='id',
                        tokenize='unicode61 remove_diacritics 2'
                    );

                CREATE TRIGGER IF NOT EXISTS documentos_locais_ai
                AFTER INSERT ON documentos_locais BEGIN
                    INSERT INTO documentos_locais_fts(rowid, nome, texto)
                    VALUES (new.id, new.nome, new.texto);
                END;

                CREATE TRIGGER IF NOT EXISTS documentos_locais_ad
                AFTER DELETE ON documentos_locais BEGIN
                    INSERT INTO documentos_locais_fts(
                        documentos_locais_fts, rowid, nome, texto
                    ) VALUES ('delete', old.id, old.nome, old.texto);
                END;

                CREATE TRIGGER IF NOT EXISTS documentos_locais_au
                AFTER UPDATE ON documentos_locais BEGIN
                    INSERT INTO documentos_locais_fts(
                        documentos_locais_fts, rowid, nome, texto
                    ) VALUES ('delete', old.id, old.nome, old.texto);
                    INSERT INTO documentos_locais_fts(rowid, nome, texto)
                    VALUES (new.id, new.nome, new.texto);
                END;
                """
            )
            colunas_documentos = {
                str(linha[1])
                for linha in conexao.execute("PRAGMA table_info(documentos_locais)")
            }
            if "versao_extrator" not in colunas_documentos:
                conexao.execute(
                    "ALTER TABLE documentos_locais ADD COLUMN "
                    "versao_extrator INTEGER NOT NULL DEFAULT 1"
                )
            conexao.executemany(
                "INSERT OR IGNORE INTO categorias_acervo_local(nome, ordem) "
                "VALUES (?, ?)",
                ((nome, ordem) for ordem, nome in enumerate(CATEGORIAS_ACERVO)),
            )
            if not fts_existia:
                conexao.execute(
                    "INSERT INTO documentos_locais_fts(documentos_locais_fts) "
                    "VALUES('rebuild')"
                )

    @staticmethod
    def _caminho_normalizado(caminho: str | Path) -> str:
        return str(Path(caminho).expanduser().resolve())

    def adicionar_acervo(self, caminho: str | Path, categoria: str) -> int:
        pasta = Path(caminho).expanduser()
        if not pasta.is_dir():
            raise FileNotFoundError(f"A pasta não existe ou não está acessível: {pasta}")
        categoria = self.resolver_categoria(categoria)
        normalizado = self._caminho_normalizado(pasta)
        try:
            with self._transacao() as conexao:
                cursor = conexao.execute(
                    """
                    INSERT INTO acervos_locais(
                        caminho, categoria, data_cadastro, status
                    ) VALUES (?, ?, ?, 'Aguardando indexação')
                    """,
                    (normalizado, categoria, datetime.now().isoformat(timespec="seconds")),
                )
                return int(cursor.lastrowid)
        except sqlite3.IntegrityError as erro:
            raise ValueError("Esta pasta já está cadastrada na Biblioteca Local.") from erro

    def listar_acervos(self) -> list[dict[str, Any]]:
        with self._transacao() as conexao:
            linhas = conexao.execute(
                """
                SELECT a.*, COUNT(d.id) AS documentos
                FROM acervos_locais a
                LEFT JOIN documentos_locais d ON d.acervo_id = a.id
                GROUP BY a.id
                ORDER BY a.categoria, a.caminho
                """
            ).fetchall()
        return [dict(linha) for linha in linhas]

    def listar_categorias(self) -> list[str]:
        with self._transacao() as conexao:
            linhas = conexao.execute(
                "SELECT nome FROM categorias_acervo_local "
                "ORDER BY ordem, nome COLLATE NOCASE"
            ).fetchall()
        return [str(linha["nome"]) for linha in linhas]

    def resolver_categoria(self, categoria: str | None) -> str:
        desejada = str(categoria or "").strip()
        localizada = self._localizar_categoria(desejada)
        return localizada or "Outros"

    def _localizar_categoria(self, categoria: str | None) -> str | None:
        desejada = str(categoria or "").strip()
        with self._transacao() as conexao:
            linha = conexao.execute(
                "SELECT nome FROM categorias_acervo_local "
                "WHERE nome = ? COLLATE NOCASE",
                (desejada,),
            ).fetchone()
        return str(linha["nome"]) if linha else None

    def adicionar_categoria(self, nome: str) -> str:
        nome = str(nome or "").strip()
        if not nome:
            raise ValueError("Informe o nome da nova categoria.")
        if len(nome) > 80:
            raise ValueError("O nome da categoria deve ter no máximo 80 caracteres.")
        try:
            with self._transacao() as conexao:
                proxima_ordem = conexao.execute(
                    "SELECT COALESCE(MAX(ordem), -1) + 1 "
                    "FROM categorias_acervo_local"
                ).fetchone()[0]
                conexao.execute(
                    "INSERT INTO categorias_acervo_local(nome, ordem) VALUES (?, ?)",
                    (nome, proxima_ordem),
                )
        except sqlite3.IntegrityError as erro:
            raise ValueError("Já existe uma categoria com esse nome.") from erro
        return nome

    def renomear_categoria(self, nome_atual: str, novo_nome: str) -> str:
        nome_atual = self._localizar_categoria(nome_atual)
        if nome_atual is None:
            raise ValueError("A categoria selecionada não foi encontrada.")
        novo_nome = str(novo_nome or "").strip()
        if nome_atual.casefold() == "outros":
            raise ValueError("A categoria 'Outros' é necessária e não pode ser renomeada.")
        if not novo_nome:
            raise ValueError("Informe o novo nome da categoria.")
        if len(novo_nome) > 80:
            raise ValueError("O nome da categoria deve ter no máximo 80 caracteres.")
        try:
            with self._transacao() as conexao:
                conexao.execute(
                    "UPDATE categorias_acervo_local SET nome=? "
                    "WHERE nome=? COLLATE NOCASE",
                    (novo_nome, nome_atual),
                )
                conexao.execute(
                    "UPDATE acervos_locais SET categoria=? "
                    "WHERE categoria=? COLLATE NOCASE",
                    (novo_nome, nome_atual),
                )
        except sqlite3.IntegrityError as erro:
            raise ValueError("Já existe uma categoria com esse nome.") from erro
        return novo_nome

    def remover_categoria(self, nome: str) -> None:
        nome = self._localizar_categoria(nome)
        if nome is None:
            raise ValueError("A categoria selecionada não foi encontrada.")
        if nome.casefold() == "outros":
            raise ValueError("A categoria 'Outros' é necessária e não pode ser excluída.")
        with self._transacao() as conexao:
            quantidade = conexao.execute(
                "SELECT COUNT(*) FROM acervos_locais "
                "WHERE categoria=? COLLATE NOCASE",
                (nome,),
            ).fetchone()[0]
            if quantidade:
                raise ValueError(
                    f"A categoria '{nome}' está atribuída a {quantidade} pasta(s). "
                    "Altere a categoria dessas pastas antes de excluí-la."
                )
            conexao.execute(
                "DELETE FROM categorias_acervo_local WHERE nome=? COLLATE NOCASE",
                (nome,),
            )

    def alterar_acervo(
        self, acervo_id: int, *, caminho: str | Path | None = None,
        categoria: str | None = None, ativo: bool | None = None,
    ) -> None:
        atual = self.obter_acervo(acervo_id)
        if atual is None:
            raise ValueError("O acervo selecionado não foi encontrado.")
        novo_caminho = atual["caminho"]
        mudou_caminho = False
        if caminho is not None:
            pasta = Path(caminho).expanduser()
            if not pasta.is_dir():
                raise FileNotFoundError(
                    f"A pasta não existe ou não está acessível: {pasta}"
                )
            novo_caminho = self._caminho_normalizado(pasta)
            mudou_caminho = os.path.normcase(novo_caminho) != os.path.normcase(
                atual["caminho"]
            )
        nova_categoria = (
            self.resolver_categoria(categoria)
            if categoria is not None
            else atual["categoria"]
        )
        novo_ativo = atual["ativo"] if ativo is None else int(bool(ativo))
        try:
            with self._transacao() as conexao:
                conexao.execute(
                    """
                    UPDATE acervos_locais
                    SET caminho=?, categoria=?, ativo=?,
                        status=?, mensagem='', ultima_indexacao=?
                    WHERE id=?
                    """,
                    (
                        novo_caminho,
                        nova_categoria,
                        novo_ativo,
                        "Aguardando indexação" if mudou_caminho else atual["status"],
                        None if mudou_caminho else atual["ultima_indexacao"],
                        acervo_id,
                    ),
                )
                if mudou_caminho:
                    conexao.execute(
                        "DELETE FROM documentos_locais WHERE acervo_id=?",
                        (acervo_id,),
                    )
        except sqlite3.IntegrityError as erro:
            raise ValueError("A nova pasta já está cadastrada.") from erro

    def obter_acervo(self, acervo_id: int) -> dict[str, Any] | None:
        with self._transacao() as conexao:
            linha = conexao.execute(
                "SELECT * FROM acervos_locais WHERE id=?", (acervo_id,)
            ).fetchone()
        return None if linha is None else dict(linha)

    def remover_acervo(self, acervo_id: int) -> None:
        with self._transacao() as conexao:
            conexao.execute("DELETE FROM acervos_locais WHERE id=?", (acervo_id,))

    @staticmethod
    def _listar_arquivos(pasta: Path) -> Iterable[Path]:
        for raiz, _diretorios, arquivos in os.walk(pasta, followlinks=False):
            for nome in arquivos:
                if nome.startswith("~$"):
                    continue
                caminho = Path(raiz) / nome
                if caminho.suffix.casefold() in EXTENSOES_SUPORTADAS:
                    yield caminho

    def indexar_acervo(
        self,
        acervo_id: int,
        *,
        cancelado: Callable[[], bool] | None = None,
    ) -> ResultadoIndexacao:
        resultado = ResultadoIndexacao(acervos=1)
        acervo = self.obter_acervo(acervo_id)
        if acervo is None:
            raise ValueError("O acervo solicitado não existe.")
        pasta = Path(acervo["caminho"])
        if not pasta.is_dir():
            resultado.indisponiveis = 1
            resultado.avisos.append(f"Pasta indisponível: {pasta}")
            with self._transacao() as conexao:
                conexao.execute(
                    "UPDATE acervos_locais SET status='Indisponível', mensagem=? "
                    "WHERE id=?",
                    ("Pasta não localizada. Selecione o novo local.", acervo_id),
                )
            return resultado

        with self._transacao() as conexao:
            existentes = {
                os.path.normcase(linha["caminho"]): dict(linha)
                for linha in conexao.execute(
                    "SELECT * FROM documentos_locais WHERE acervo_id=?",
                    (acervo_id,),
                )
            }

        vistos: set[str] = set()
        agora = datetime.now().isoformat(timespec="seconds")
        gravacoes_pendentes = 0
        with ExtratorTextoLocal() as extrator, self._transacao() as conexao:
            for arquivo in self._listar_arquivos(pasta):
                if cancelado is not None and cancelado():
                    resultado.avisos.append("Indexação interrompida antes da conclusão.")
                    break
                resultado.encontrados += 1
                caminho = self._caminho_normalizado(arquivo)
                chave = os.path.normcase(caminho)
                vistos.add(chave)
                try:
                    stat = arquivo.stat()
                except OSError as erro:
                    resultado.erros += 1
                    resultado.avisos.append(f"Não foi possível ler: {arquivo.name}")
                    continue
                anterior = existentes.get(chave)
                if (
                    anterior
                    and anterior["tamanho"] == stat.st_size
                    and anterior["modificado_ns"] == stat.st_mtime_ns
                    and anterior["status"] in {"Indexado", "Sem texto"}
                    and int(anterior["versao_extrator"]) == VERSAO_EXTRATOR
                ):
                    resultado.inalterados += 1
                    continue

                try:
                    texto = extrator.extrair(arquivo).strip()
                    status = "Indexado" if texto else "Sem texto"
                    erro_texto = ""
                    if texto:
                        if anterior:
                            resultado.atualizados += 1
                        else:
                            resultado.novos += 1
                    else:
                        resultado.sem_texto += 1
                except Exception as erro:
                    texto = ""
                    status = "Erro"
                    erro_texto = f"{type(erro).__name__}: {erro}"
                    resultado.erros += 1

                valores = (
                    acervo_id, caminho, arquivo.name, arquivo.suffix.casefold(),
                    stat.st_size, stat.st_mtime_ns, texto, agora, status, erro_texto,
                    VERSAO_EXTRATOR,
                )
                conexao.execute(
                    """
                    INSERT INTO documentos_locais(
                        acervo_id, caminho, nome, extensao, tamanho,
                        modificado_ns, texto, indexado_em, status, erro,
                        versao_extrator
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(acervo_id, caminho) DO UPDATE SET
                        nome=excluded.nome, extensao=excluded.extensao,
                        tamanho=excluded.tamanho,
                        modificado_ns=excluded.modificado_ns,
                        texto=excluded.texto, indexado_em=excluded.indexado_em,
                        status=excluded.status, erro=excluded.erro,
                        versao_extrator=excluded.versao_extrator
                    """,
                    valores,
                )
                gravacoes_pendentes += 1
                if gravacoes_pendentes >= 50:
                    # Em acervos com milhares de arquivos, confirma lotes para
                    # reduzir o tempo de bloqueio e preservar o progresso caso
                    # o computador seja desligado durante a primeira leitura.
                    conexao.commit()
                    gravacoes_pendentes = 0

            removidos = [
                item["id"] for chave, item in existentes.items() if chave not in vistos
            ]
            if removidos:
                conexao.executemany(
                    "DELETE FROM documentos_locais WHERE id=?",
                    ((item_id,) for item_id in removidos),
                )
                resultado.removidos = len(removidos)
            status_acervo = "Concluído com avisos" if resultado.erros else "Concluído"
            conexao.execute(
                """
                UPDATE acervos_locais
                SET ultima_indexacao=?, status=?, mensagem=? WHERE id=?
                """,
                (agora, status_acervo, resultado.resumo(), acervo_id),
            )
        return resultado

    def indexar_todos(self) -> ResultadoIndexacao:
        total = ResultadoIndexacao()
        for acervo in self.listar_acervos():
            if acervo["ativo"]:
                total.incorporar(self.indexar_acervo(int(acervo["id"])))
        return total

    @staticmethod
    def termos_pesquisa(termo: str) -> list[str]:
        tokens = re.findall(r"[^\W_]+(?:[-./][^\W_]+)*", termo, flags=re.UNICODE)
        if not tokens:
            raise ValueError("Digite ao menos uma palavra ou número para pesquisar.")
        return tokens

    @classmethod
    def _consulta_fts(cls, termo: str, modo: str) -> str:
        tokens = cls.termos_pesquisa(termo)
        escapados = [token.replace('"', '""') for token in tokens]
        if modo == "Expressão exata":
            return f'"{" ".join(escapados)}"'
        return " AND ".join(f'"{token}"*' for token in escapados)

    def pesquisar(
        self,
        termo: str,
        *,
        categoria: str = "Todos",
        modo: str = "Todas as palavras",
        limite: int = 200,
    ) -> list[dict[str, Any]]:
        if modo not in MODOS_PESQUISA:
            raise ValueError("Modo de pesquisa desconhecido.")
        consulta_fts = self._consulta_fts(str(termo).strip(), modo)
        filtros = [
            "documentos_locais_fts MATCH ?",
            "a.ativo = 1",
            "a.status <> 'Indisponível'",
            "d.status = 'Indexado'",
        ]
        parametros: list[Any] = [consulta_fts]
        if categoria and categoria != "Todos":
            filtros.append("a.categoria = ?")
            parametros.append(categoria)
        parametros.append(max(1, min(int(limite), 1000)))
        sql = f"""
            SELECT d.id, d.nome, d.caminho, d.extensao, a.categoria,
                   snippet(documentos_locais_fts, 1, '«', '»', ' … ', 36) AS trecho,
                   bm25(documentos_locais_fts, 1.2, 1.0) AS relevancia
            FROM documentos_locais_fts
            JOIN documentos_locais d ON d.id = documentos_locais_fts.rowid
            JOIN acervos_locais a ON a.id = d.acervo_id
            WHERE {' AND '.join(filtros)}
            ORDER BY relevancia
            LIMIT ?
        """
        with self._transacao() as conexao:
            linhas = conexao.execute(sql, parametros).fetchall()
        return [dict(linha) for linha in linhas]

    def obter_documento(self, documento_id: int) -> dict[str, Any] | None:
        with self._transacao() as conexao:
            linha = conexao.execute(
                """
                SELECT d.*, a.categoria
                FROM documentos_locais d
                JOIN acervos_locais a ON a.id=d.acervo_id
                WHERE d.id=?
                """,
                (documento_id,),
            ).fetchone()
        return None if linha is None else dict(linha)

    def estatisticas(self) -> dict[str, int]:
        with self._transacao() as conexao:
            linha = conexao.execute(
                """
                SELECT COUNT(DISTINCT a.id) AS acervos,
                       COUNT(d.id) AS documentos,
                       SUM(CASE WHEN d.status='Indexado' THEN 1 ELSE 0 END) AS indexados,
                       SUM(CASE WHEN d.status='Sem texto' THEN 1 ELSE 0 END) AS sem_texto,
                       SUM(CASE WHEN d.status='Erro' THEN 1 ELSE 0 END) AS erros
                FROM acervos_locais a
                LEFT JOIN documentos_locais d ON d.acervo_id=a.id
                """
            ).fetchone()
        return {chave: int(linha[chave] or 0) for chave in linha.keys()}
