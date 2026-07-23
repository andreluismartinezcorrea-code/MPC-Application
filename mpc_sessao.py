"""Salvamento automático e recuperação de sessão do MPC Parecer."""

from __future__ import annotations

import hashlib
import json
from datetime import datetime
from pathlib import Path
from typing import Any

from mpc_modelos import normalizar_dados_persistidos
from mpc_persistencia import carregar_json_normalizado, salvar_json_atomico


CHAVE_METADADOS_SESSAO = "_sessao_automatica"


def _dados_sem_metadados(dados: dict[str, Any]) -> dict[str, Any]:
    normalizados = normalizar_dados_persistidos(dados)
    normalizados.pop(CHAVE_METADADOS_SESSAO, None)
    return normalizados


def assinatura_dados(dados: dict[str, Any]) -> str:
    """Cria uma assinatura estável para detectar alterações na tela."""
    conteudo = json.dumps(
        _dados_sem_metadados(dados),
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(conteudo).hexdigest()


def possui_conteudo_relevante(dados: dict[str, Any]) -> bool:
    """Evita criar recuperação para uma tela completamente vazia."""
    normalizados = _dados_sem_metadados(dados)
    campos = (
        "exercicio",
        "processo",
        "tipo",
        "orgao",
        "rag",
        "arquivo_parecer",
        "pasta",
    )
    if any(str(normalizados.get(campo, "")).strip() for campo in campos):
        return True
    if any(
        str(item.get("nome", "")).strip()
        for item in normalizados.get("responsaveis", [])
    ):
        return True
    return any(
        str(item.get("irregularidade", "")).strip()
        for item in normalizados.get("apontamentos_detalhado", [])
    )


class ControleSessao:
    """Controla a cópia de recuperação sem interferir no salvamento manual."""

    def __init__(self, caminho: str | Path):
        self.caminho = Path(caminho)
        self._assinatura_salva: str | None = None
        self._assinatura_autosalva: str | None = None

    def definir_estado_inicial(self, dados: dict[str, Any]) -> None:
        """Registra o estado inicial da GUI sem apagar eventual recuperação."""
        self._assinatura_salva = assinatura_dados(dados)

    def registrar_estado_salvo(self, dados: dict[str, Any]) -> None:
        """Marca a tela como salva no JSON oficial e remove a recuperação."""
        self._assinatura_salva = assinatura_dados(dados)
        self._assinatura_autosalva = None
        self.descartar_recuperacao()

    def registrar_estado_recuperado(self, dados: dict[str, Any]) -> None:
        """Evita regravar continuamente uma sessão que acabou de ser aberta."""
        self._assinatura_autosalva = assinatura_dados(dados)

    def autosalvar_se_necessario(
        self,
        dados: dict[str, Any],
        *,
        versao_aplicacao: str = "",
    ) -> bool:
        """Grava somente quando há conteúdo e alteração ainda não salva."""
        if not possui_conteudo_relevante(dados):
            self._assinatura_autosalva = None
            self.descartar_recuperacao()
            return False

        assinatura = assinatura_dados(dados)
        if assinatura == self._assinatura_salva:
            self._assinatura_autosalva = None
            self.descartar_recuperacao()
            return False
        if assinatura == self._assinatura_autosalva and self.caminho.is_file():
            return False

        conteudo = _dados_sem_metadados(dados)
        conteudo[CHAVE_METADADOS_SESSAO] = {
            "salvo_em": datetime.now().isoformat(timespec="seconds"),
            "versao_aplicacao": str(versao_aplicacao or ""),
        }
        salvar_json_atomico(self.caminho, conteudo)
        self._assinatura_autosalva = assinatura
        return True

    def carregar_recuperacao(self) -> dict[str, Any] | None:
        if not self.caminho.is_file():
            return None
        return carregar_json_normalizado(self.caminho)

    def descartar_recuperacao(self) -> None:
        try:
            self.caminho.unlink(missing_ok=True)
        except OSError:
            pass

