"""Estado central da interface do MPC Parecer.

Esta primeira etapa concentra responsáveis e apontamentos fora dos widgets.
A interface continua compatível com o código legado, mas validações, salvamento
e demais consumidores passam a receber uma fotografia única e normalizada.
"""

from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass, field
from typing import Any, Iterable

from mpc_modelos import Apontamento, Responsavel


def _normalizar_responsaveis(
    dados: Iterable[dict[str, Any]],
) -> list[dict[str, Any]]:
    resultado = []
    for item in dados:
        if not isinstance(item, dict):
            continue
        normalizado = Responsavel.from_dict(item).to_dict()
        resultado.append({**deepcopy(item), **normalizado})
    return resultado


def _normalizar_apontamentos(
    dados: Iterable[dict[str, Any]],
) -> list[dict[str, Any]]:
    resultado = []
    for item in dados:
        if not isinstance(item, dict):
            continue
        normalizado = Apontamento.from_dict(item).to_dict()
        # Campos exclusivamente visuais, como o resumo das associações,
        # permanecem disponíveis embora não façam parte da regra jurídica.
        resultado.append({**deepcopy(item), **normalizado})
    return resultado


@dataclass(slots=True)
class EstadoInterface:
    """Mantém uma fotografia coerente dos dados dinâmicos da tela."""

    _responsaveis: list[dict[str, Any]] = field(default_factory=list)
    _apontamentos: list[dict[str, Any]] = field(default_factory=list)
    revisao: int = 0

    def sincronizar(
        self,
        responsaveis: Iterable[dict[str, Any]],
        apontamentos: Iterable[dict[str, Any]],
    ) -> dict[str, Any]:
        novos_responsaveis = _normalizar_responsaveis(responsaveis)
        novos_apontamentos = _normalizar_apontamentos(apontamentos)
        if (
            novos_responsaveis != self._responsaveis
            or novos_apontamentos != self._apontamentos
        ):
            self._responsaveis = novos_responsaveis
            self._apontamentos = novos_apontamentos
            self.revisao += 1
        return self.snapshot()

    def atualizar_responsaveis(
        self,
        responsaveis: Iterable[dict[str, Any]],
    ) -> list[dict[str, Any]]:
        novos = _normalizar_responsaveis(responsaveis)
        if novos != self._responsaveis:
            self._responsaveis = novos
            self.revisao += 1
        return self.responsaveis()

    def atualizar_apontamentos(
        self,
        apontamentos: Iterable[dict[str, Any]],
    ) -> list[dict[str, Any]]:
        novos = _normalizar_apontamentos(apontamentos)
        if novos != self._apontamentos:
            self._apontamentos = novos
            self.revisao += 1
        return self.apontamentos()

    def responsaveis(self) -> list[dict[str, Any]]:
        return deepcopy(self._responsaveis)

    def apontamentos(self) -> list[dict[str, Any]]:
        return deepcopy(self._apontamentos)

    def snapshot(self) -> dict[str, Any]:
        responsaveis = self.responsaveis()
        apontamentos = self.apontamentos()
        return {
            "revisao": self.revisao,
            "responsaveis": responsaveis,
            "apontamentos_detalhado": apontamentos,
            "apontamentos_lista": [
                item.get("irregularidade", "") for item in apontamentos
            ],
        }

    def limpar(self) -> None:
        if self._responsaveis or self._apontamentos:
            self._responsaveis = []
            self._apontamentos = []
            self.revisao += 1
