"""Cálculo do andamento do fluxo operacional do parecer."""

from __future__ import annotations

from typing import Any


ETAPAS_FLUXO = (
    ("dados", "Preencher processo e responsáveis"),
    ("rag", "Executar Relatório de Auditoria"),
    ("esclarecimentos", "Executar Análise de Esclarecimentos"),
    ("eparecer", "Gerar o e-Parecer"),
    ("revisao", "Revisar e certificar os apontamentos"),
    ("documento", "Construir cabeçalho, introdução e conclusão"),
    ("registro", "Registrar a produção"),
)


def _acoes_concluidas(dados: dict[str, Any]) -> set[str]:
    resultado = set()
    for evento in dados.get("historico_operacoes", []):
        if not isinstance(evento, dict):
            continue
        if str(evento.get("status", "")).lower() in {
            "concluída",
            "concluida",
            "concluído",
            "concluido",
        }:
            resultado.add(str(evento.get("operacao", "")).strip())
    return resultado


def avaliar_fluxo(dados: dict[str, Any]) -> dict[str, Any]:
    """Retorna etapas concluídas, percentual e a próxima ação sugerida."""
    responsaveis = [
        item
        for item in dados.get("responsaveis", [])
        if isinstance(item, dict) and str(item.get("nome", "")).strip()
    ]
    apontamentos = [
        item
        for item in dados.get("apontamentos_detalhado", [])
        if isinstance(item, dict)
        and str(item.get("irregularidade", "")).strip()
    ]
    acoes = _acoes_concluidas(dados)

    dados_prontos = all(
        str(dados.get(campo, "")).strip()
        for campo in ("exercicio", "processo", "tipo", "orgao")
    ) and bool(responsaveis)
    rag_pronto = dados_prontos and bool(
        str(dados.get("rag", "")).strip() and apontamentos
    )
    # As operações de PDF podem trabalhar em segundo plano. Por isso, o guia
    # só avança quando o campo resultante foi realmente preenchido, e não
    # apenas quando o botão foi acionado.
    esclarecimentos_prontos = rag_pronto and bool(
        str(dados.get("arq_anal_escl", "")).strip()
    )
    eparecer_pronto = esclarecimentos_prontos and bool(
        str(dados.get("arquivo_parecer", "")).strip()
    )
    revisao_pronta = eparecer_pronto and bool(apontamentos) and all(
        str(item.get("conclusao", "")).strip()
        not in {"", "Análise Pendente"}
        for item in apontamentos
    )
    documento_pronto = revisao_pronta and bool(
        {"Introdução", "Conclusão"}.issubset(acoes)
        and ({"Cabeçalho", "Cabeçalho (e-Parecer)"} & acoes)
    )
    registro_pronto = documento_pronto and bool(
        str(dados.get("registro_id", "")).strip()
    )

    estados = (
        dados_prontos,
        rag_pronto,
        esclarecimentos_prontos,
        eparecer_pronto,
        revisao_pronta,
        documento_pronto,
        registro_pronto,
    )
    etapas = [
        {"id": identificador, "titulo": titulo, "concluida": concluida}
        for (identificador, titulo), concluida in zip(ETAPAS_FLUXO, estados)
    ]
    concluidas = sum(1 for estado in estados if estado)
    proxima = next(
        (etapa for etapa in etapas if not etapa["concluida"]),
        None,
    )
    return {
        "etapas": etapas,
        "concluidas": concluidas,
        "total": len(etapas),
        "percentual": round(concluidas * 100 / len(etapas)),
        "proxima": proxima,
    }


def resumir_fluxo(dados: dict[str, Any]) -> str:
    avaliacao = avaliar_fluxo(dados)
    if avaliacao["proxima"] is None:
        return "FLUXO CONCLUÍDO — produção registrada"
    return (
        f"PRÓXIMA ETAPA: {avaliacao['proxima']['titulo']}  •  "
        f"PROGRESSO: {avaliacao['concluidas']}/{avaliacao['total']} "
        f"({avaliacao['percentual']}%)"
    )
