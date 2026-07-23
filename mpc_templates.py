"""Serviço seguro para consulta e edição do banco de textos-modelo."""

from __future__ import annotations

import json
import os
import shutil
import tempfile
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from string import Formatter
from typing import Any


class ErroModelo(ValueError):
    """Indica uma alteração inválida em um texto-modelo."""


def extrair_campos(texto: str) -> set[str]:
    """Retorna os campos entre chaves usados pela formatação do Python."""
    try:
        campos = {
            str(nome_campo).split(".", 1)[0].split("[", 1)[0]
            for _, nome_campo, _, _ in Formatter().parse(str(texto or ""))
            if nome_campo
        }
    except ValueError as erro:
        raise ErroModelo(
            "O texto contém chaves inválidas. Preserve campos como "
            "{gestores} e {valor_total} exatamente como foram definidos."
        ) from erro
    return campos


def _percorrer_textos(
    valor: Any,
    caminho: tuple[str, ...] = (),
) -> list[dict[str, Any]]:
    encontrados: list[dict[str, Any]] = []
    if isinstance(valor, dict):
        for chave, item in valor.items():
            if chave == "descricao":
                continue
            encontrados.extend(_percorrer_textos(item, caminho + (str(chave),)))
    elif isinstance(valor, str):
        encontrados.append(
            {
                "caminho": caminho,
                "identificador": " > ".join(caminho),
                "texto": valor,
                "campos": sorted(extrair_campos(valor)),
            }
        )
    return encontrados


def listar_textos_editaveis(banco: dict) -> list[dict[str, Any]]:
    """Lista todos os textos do banco que podem ser alterados pela interface."""
    if not isinstance(banco, dict):
        raise ErroModelo("O banco de parágrafos deve ter a estrutura de um objeto JSON.")
    return sorted(_percorrer_textos(banco), key=lambda item: item["identificador"].casefold())


def obter_texto(banco: dict, caminho: tuple[str, ...]) -> str:
    """Obtém um texto a partir de seu caminho hierárquico."""
    atual: Any = banco
    for chave in caminho:
        if not isinstance(atual, dict) or chave not in atual:
            raise ErroModelo("O modelo selecionado não existe mais no banco atual.")
        atual = atual[chave]
    if not isinstance(atual, str):
        raise ErroModelo("O modelo selecionado não contém um texto editável.")
    return atual


def atualizar_texto(banco: dict, caminho: tuple[str, ...], novo_texto: str) -> dict:
    """Devolve uma cópia do banco com um único texto validado e atualizado."""
    texto_atual = obter_texto(banco, caminho)
    campos_atuais = extrair_campos(texto_atual)
    campos_novos = extrair_campos(novo_texto)
    if campos_atuais != campos_novos:
        removidos = sorted(campos_atuais - campos_novos)
        adicionados = sorted(campos_novos - campos_atuais)
        detalhes = []
        if removidos:
            detalhes.append("campos removidos: " + ", ".join(removidos))
        if adicionados:
            detalhes.append("campos não reconhecidos: " + ", ".join(adicionados))
        raise ErroModelo(
            "Os campos automáticos entre chaves devem ser preservados ("
            + "; ".join(detalhes)
            + ")."
        )

    atualizado = deepcopy(banco)
    destino: Any = atualizado
    for chave in caminho[:-1]:
        destino = destino[chave]
    destino[caminho[-1]] = str(novo_texto)
    return atualizado


def localizar_banco_externo(script_dir: str | os.PathLike[str]) -> Path | None:
    """Localiza o banco externo editável, respeitando a configuração opcional."""
    candidatos = [
        os.getenv("MPC_BANCO_PARAGRAFOS", "").strip(),
        str(Path(script_dir) / "banco_paragrafos.json"),
        str(Path.cwd() / "banco_paragrafos.json"),
    ]
    vistos: set[Path] = set()
    for candidato in candidatos:
        if not candidato:
            continue
        caminho = Path(candidato).expanduser().resolve()
        if caminho in vistos:
            continue
        vistos.add(caminho)
        if caminho.is_file():
            return caminho
    return None


def carregar_banco_arquivo(caminho: str | os.PathLike[str]) -> dict:
    """Lê e valida a estrutura básica de um banco externo."""
    try:
        with Path(caminho).open("r", encoding="utf-8") as arquivo:
            dados = json.load(arquivo)
    except (OSError, json.JSONDecodeError) as erro:
        raise ErroModelo(f"Não foi possível ler o banco de parágrafos: {erro}") from erro
    if not isinstance(dados, dict):
        raise ErroModelo("O banco de parágrafos não contém um objeto JSON válido.")
    return dados


def salvar_banco_com_backup(
    caminho_banco: str | os.PathLike[str],
    banco: dict,
    diretorio_backups: str | os.PathLike[str],
) -> Path:
    """Cria backup e grava o JSON por substituição atômica."""
    destino = Path(caminho_banco).resolve()
    if not destino.is_file():
        raise ErroModelo("O arquivo externo banco_paragrafos.json não foi localizado.")

    # Confere toda a estrutura antes de modificar o arquivo em disco.
    listar_textos_editaveis(banco)
    diretorio = Path(diretorio_backups).resolve()
    diretorio.mkdir(parents=True, exist_ok=True)
    data_hora = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    backup = diretorio / f"{data_hora}_banco_paragrafos.json"
    shutil.copy2(destino, backup)

    temporario = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            newline="\n",
            dir=destino.parent,
            prefix=".banco_paragrafos_",
            suffix=".tmp",
            delete=False,
        ) as arquivo:
            temporario = Path(arquivo.name)
            json.dump(banco, arquivo, ensure_ascii=False, indent=2)
            arquivo.write("\n")
        os.replace(temporario, destino)
    except OSError as erro:
        raise ErroModelo(f"Não foi possível salvar o banco de parágrafos: {erro}") from erro
    finally:
        if temporario and temporario.exists():
            temporario.unlink(missing_ok=True)
    return backup


def listar_backups(diretorio_backups: str | os.PathLike[str]) -> list[Path]:
    """Lista cópias disponíveis, da mais recente para a mais antiga."""
    diretorio = Path(diretorio_backups)
    if not diretorio.is_dir():
        return []
    return sorted(
        diretorio.glob("*_banco_paragrafos.json"),
        key=lambda item: item.name,
        reverse=True,
    )


def restaurar_backup(
    caminho_banco: str | os.PathLike[str],
    caminho_backup: str | os.PathLike[str],
    diretorio_backups: str | os.PathLike[str],
) -> Path:
    """Restaura uma cópia validada, preservando backup da versão atual."""
    banco_backup = carregar_banco_arquivo(caminho_backup)
    return salvar_banco_com_backup(caminho_banco, banco_backup, diretorio_backups)
