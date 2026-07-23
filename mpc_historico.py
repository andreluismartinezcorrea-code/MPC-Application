"""Histórico operacional e resumos de status da aplicação MPC Parecer.

Este módulo não depende de Tkinter. A interface apenas exibe os eventos que
ele normaliza, registra e resume.
"""

from __future__ import annotations

from collections import Counter
from datetime import datetime
from typing import Any, Callable, Iterable


STATUS_CONCLUIDOS = frozenset({"concluída", "concluida"})
STATUS_PROBLEMA = frozenset({"erro", "bloqueada"})


def normalizar_evento(evento: Any) -> dict[str, str] | None:
    """Converte um registro persistido para o formato operacional atual."""
    if not isinstance(evento, dict):
        return None
    operacao = str(evento.get("operacao", "") or "").strip()
    status = str(evento.get("status", "") or "").strip()
    if not operacao or not status:
        return None
    return {
        "data_hora": str(evento.get("data_hora", "") or "").strip(),
        "operacao": operacao,
        "status": status,
        "detalhe": str(evento.get("detalhe", "") or "").strip(),
    }


def registrar_evento(
    historico: list[dict[str, str]],
    operacao: str,
    status: str,
    detalhe: Any = "",
    *,
    limite: int = 100,
    agora: Callable[[], datetime] = datetime.now,
) -> dict[str, str]:
    """Acrescenta um evento e limita o crescimento do histórico em memória."""
    evento = {
        "data_hora": agora().isoformat(timespec="seconds"),
        "operacao": str(operacao or "").strip() or "Operação",
        "status": str(status or "").strip() or "informação",
        "detalhe": str(detalhe or "").strip(),
    }
    historico.append(evento)
    limite = max(1, int(limite))
    del historico[:-limite]
    return evento


def carregar_historico(eventos: Iterable[Any], limite: int = 100) -> list[dict[str, str]]:
    """Normaliza eventos de arquivos antigos e preserva somente os recentes."""
    normalizados = []
    for evento in eventos or ():
        item = normalizar_evento(evento)
        if item is not None:
            normalizados.append(item)
    return normalizados[-max(1, int(limite)):]


def resumir_historico(eventos: Iterable[Any]) -> dict[str, Any]:
    """Produz os indicadores utilizados pelo painel operacional."""
    normalizados = [
        item for evento in eventos if (item := normalizar_evento(evento)) is not None
    ]
    contagens = Counter(item["status"].casefold() for item in normalizados)
    concluidas = sum(contagens.get(status, 0) for status in STATUS_CONCLUIDOS)
    problemas = sum(contagens.get(status, 0) for status in STATUS_PROBLEMA)
    ultimo = normalizados[-1] if normalizados else None
    if ultimo:
        texto_ultimo = (
            f"ÚLTIMA OPERAÇÃO • {ultimo['operacao'].upper()} — "
            f"{ultimo['status'].upper()}"
        )
    else:
        texto_ultimo = "PRONTO • Aguardando a próxima operação"
    return {
        "total": len(normalizados),
        "concluidas": concluidas,
        "problemas": problemas,
        "canceladas": contagens.get("cancelada", 0),
        "ultimo": ultimo,
        "texto_ultimo": texto_ultimo,
    }
