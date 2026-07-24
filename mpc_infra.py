"""Infraestrutura local do MPC Parecer: versão, logs, backups e modelos."""

from __future__ import annotations

import hashlib
import json
import logging
import os
import re
import shutil
import sys
import threading
from datetime import datetime, timedelta
from logging.handlers import RotatingFileHandler
from pathlib import Path


APP_VERSION = "11.22.0"
APP_RELEASE_DATE = "2026-07-22"
LOGGER_NAME = "mpc_parecer"


def configurar_logging(script_dir: str) -> tuple[logging.Logger, str]:
    """Cria um log rotativo, sem registrar chaves ou conteúdo de documentos."""
    logs_dir = Path(script_dir) / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    log_path = logs_dir / "mpc_parecer.log"
    logger = logging.getLogger(LOGGER_NAME)
    logger.setLevel(logging.INFO)
    if not any(
        isinstance(handler, RotatingFileHandler)
        and Path(handler.baseFilename) == log_path
        for handler in logger.handlers
    ):
        handler = RotatingFileHandler(
            log_path,
            maxBytes=2_000_000,
            backupCount=5,
            encoding="utf-8",
        )
        handler.setFormatter(
            logging.Formatter(
                "%(asctime)s | %(levelname)s | %(threadName)s | %(message)s"
            )
        )
        logger.addHandler(handler)
    logger.info("Aplicação iniciada | versão %s", APP_VERSION)
    return logger, str(log_path)


def instalar_ganchos_de_erro(logger: logging.Logger) -> None:
    """Registra exceções não tratadas do Python e de threads."""
    original_sys_hook = sys.excepthook

    def sys_hook(exc_type, exc_value, exc_traceback):
        logger.critical(
            "Exceção não tratada",
            exc_info=(exc_type, exc_value, exc_traceback),
        )
        original_sys_hook(exc_type, exc_value, exc_traceback)

    sys.excepthook = sys_hook

    if hasattr(threading, "excepthook"):
        original_thread_hook = threading.excepthook

        def thread_hook(args):
            logger.critical(
                "Exceção não tratada em thread",
                exc_info=(args.exc_type, args.exc_value, args.exc_traceback),
            )
            original_thread_hook(args)

        threading.excepthook = thread_hook


def nome_seguro(texto: str) -> str:
    texto = re.sub(r"[^0-9A-Za-zÀ-ÿ._-]+", "_", str(texto).strip())
    return texto.strip("._")[:80] or "operacao"


def _destino_unico(pasta: Path, nome: str) -> Path:
    pasta.mkdir(parents=True, exist_ok=True)
    carimbo = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    return pasta / f"{carimbo}_{nome}"


def criar_backup_arquivo(
    caminho_origem: str,
    backup_root: str,
    motivo: str,
) -> str | None:
    """Copia um arquivo existente para uma pasta de backup datada."""
    origem = Path(caminho_origem)
    if not origem.is_file():
        return None
    destino = _destino_unico(
        Path(backup_root) / nome_seguro(motivo),
        origem.name,
    )
    shutil.copy2(origem, destino)
    return str(destino)


def gerar_caminho_backup_arquivo(
    nome_original: str,
    backup_root: str,
    motivo: str,
) -> str:
    """Gera um destino único para APIs que salvam diretamente uma cópia."""
    nome = Path(nome_original).name or "documento.docx"
    destino = _destino_unico(
        Path(backup_root) / nome_seguro(motivo),
        nome,
    )
    return str(destino)


def salvar_snapshot_json(
    dados: dict,
    backup_root: str,
    motivo: str,
) -> str:
    """Salva os dados essenciais da tela antes de uma operação importante."""
    destino = _destino_unico(
        Path(backup_root) / "dados",
        f"{nome_seguro(motivo)}.json",
    )
    destino.write_text(
        json.dumps(dados, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    return str(destino)


def limpar_backups_antigos(
    backup_root: str,
    retencao_dias: int,
    agora: datetime | None = None,
) -> dict:
    """Remove apenas arquivos de backup vencidos, nunca pastas ou arquivos externos."""
    resultado = {
        "retencao_dias": retencao_dias,
        "removidos": 0,
        "bytes_liberados": 0,
        "erros": [],
        "desativada": retencao_dias <= 0,
    }
    if retencao_dias <= 0:
        return resultado

    raiz = Path(backup_root)
    if not raiz.is_dir():
        return resultado

    raiz_resolvida = raiz.resolve()
    limite = (agora or datetime.now()) - timedelta(days=retencao_dias)
    for arquivo in raiz.rglob("*"):
        try:
            # Links são ignorados para que a limpeza nunca alcance outro local.
            if not arquivo.is_file() or arquivo.is_symlink():
                continue
            if raiz_resolvida not in arquivo.resolve().parents:
                continue
            modificado = datetime.fromtimestamp(arquivo.stat().st_mtime)
            if modificado >= limite:
                continue
            tamanho = arquivo.stat().st_size
            arquivo.unlink()
            resultado["removidos"] += 1
            resultado["bytes_liberados"] += tamanho
        except Exception as erro:
            resultado["erros"].append(f"{arquivo}: {erro}")
    return resultado


def verificar_gravacao(pasta: str) -> tuple[bool, str]:
    """Verifica se a pasta pode ser criada e usada, sem deixar arquivo de teste."""
    try:
        destino = Path(pasta)
        destino.mkdir(parents=True, exist_ok=True)
        teste = destino / f".teste_{os.getpid()}.tmp"
        teste.write_text("ok", encoding="utf-8")
        teste.unlink()
        return True, "Disponível para gravação"
    except Exception as erro:
        return False, str(erro)


def _sha256(caminho: Path) -> str:
    digest = hashlib.sha256()
    with caminho.open("rb") as arquivo:
        for bloco in iter(lambda: arquivo.read(1024 * 1024), b""):
            digest.update(bloco)
    return digest.hexdigest()


def versionar_modelos_word(modelos_dir: str, destino_root: str) -> dict:
    """Arquiva somente versões ainda não registradas dos modelos Word."""
    origem = Path(modelos_dir)
    destino = Path(destino_root)
    resultado = {"novos": 0, "inalterados": 0, "erros": [], "destino": str(destino)}
    if not origem.is_dir():
        resultado["erros"].append(f"Pasta de modelos não encontrada: {origem}")
        return resultado

    destino.mkdir(parents=True, exist_ok=True)
    manifesto_path = destino / "manifesto.json"
    try:
        manifesto = json.loads(manifesto_path.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        manifesto = {}

    extensoes = {".docx", ".docm", ".dotx", ".dotm"}
    for arquivo in origem.rglob("*"):
        if not arquivo.is_file() or arquivo.suffix.lower() not in extensoes:
            continue
        try:
            relativo = str(arquivo.relative_to(origem))
            hash_atual = _sha256(arquivo)
            if manifesto.get(relativo) == hash_atual:
                resultado["inalterados"] += 1
                continue
            pasta_relativa = destino / arquivo.relative_to(origem).parent
            nome_versao = (
                f"{arquivo.stem}__{datetime.now():%Y%m%d_%H%M%S}"
                f"__{hash_atual[:10]}{arquivo.suffix}"
            )
            pasta_relativa.mkdir(parents=True, exist_ok=True)
            shutil.copy2(arquivo, pasta_relativa / nome_versao)
            manifesto[relativo] = hash_atual
            resultado["novos"] += 1
        except Exception as erro:
            resultado["erros"].append(f"{arquivo}: {erro}")

    manifesto_path.write_text(
        json.dumps(manifesto, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return resultado
