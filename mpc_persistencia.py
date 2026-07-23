"""Leitura e gravação seguras dos dados do MPC Parecer."""

from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path
from typing import Any

from mpc_modelos import normalizar_dados_persistidos


def carregar_json_normalizado(caminho: str | os.PathLike[str]) -> dict[str, Any]:
    """Carrega um JSON e aplica a migração compatível do modelo central."""
    arquivo = Path(caminho)
    with arquivo.open("r", encoding="utf-8-sig") as origem:
        dados = json.load(origem)
    return normalizar_dados_persistidos(dados)


def salvar_json_atomico(
    caminho: str | os.PathLike[str],
    dados: dict[str, Any],
    *,
    normalizar: bool = True,
) -> Path:
    """
    Grava o JSON por substituição atômica.

    Primeiro é criado um arquivo temporário completo na mesma pasta. Somente
    depois de ``flush`` e ``fsync`` ele substitui o destino. Assim, uma queda
    durante a gravação não deixa o arquivo oficial truncado.
    """
    destino = Path(caminho)
    destino.parent.mkdir(parents=True, exist_ok=True)
    conteudo = normalizar_dados_persistidos(dados) if normalizar else dados

    temporario: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            newline="\n",
            dir=destino.parent,
            prefix=f".{destino.name}.",
            suffix=".tmp",
            delete=False,
        ) as arquivo:
            temporario = Path(arquivo.name)
            json.dump(conteudo, arquivo, indent=4, ensure_ascii=False)
            arquivo.write("\n")
            arquivo.flush()
            os.fsync(arquivo.fileno())
        os.replace(temporario, destino)
        temporario = None
        return destino
    finally:
        if temporario is not None:
            try:
                temporario.unlink(missing_ok=True)
            except OSError:
                pass

