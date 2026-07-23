"""Controladores puros que traduzem regras em estados exibidos pela GUI."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Mapping, Sequence

from mpc_regras import (
    obter_responsaveis_apontamento,
    validar_conclusoes_responsaveis,
    validar_vinculos_responsabilidade,
)


@dataclass(frozen=True, slots=True)
class AlertaAssociacaoPendente:
    """Associação provisória existente antes da conclusão de uma falha."""

    item: str
    responsaveis: tuple[str, ...]
    naturezas: tuple[str, ...]

    @property
    def mensagem(self) -> str:
        nomes = ", ".join(self.responsaveis)
        naturezas = ", ".join(self.naturezas)
        return (
            f"Item {self.item}: permanece em 'Análise Pendente', mas já "
            f"possui associação com {nomes} ({naturezas})."
        )


@dataclass(frozen=True, slots=True)
class StatusClassificacao:
    texto: str
    estilo: str
    alertas: tuple[AlertaAssociacaoPendente, ...] = ()


@dataclass(frozen=True, slots=True)
class ResultadoPainelPreenchimento:
    pronto: bool
    resumo: str
    detalhes: str
    estilo: str
    pendencias: tuple[str, ...]
    erros: tuple[str, ...]
    alertas: tuple[AlertaAssociacaoPendente, ...]
    total_responsaveis: int


def _texto(dados: Mapping[str, Any], chave: str) -> str:
    return str(dados.get(chave, "") or "").strip()


def _identificar_item(apontamento: Mapping[str, Any], linha: int) -> str:
    descricao = str(
        apontamento.get("irregularidade")
        or apontamento.get("item")
        or ""
    ).strip()
    numero = re.match(r"\d+(?:\.\d+)*", descricao)
    return numero.group(0) if numero else f"linha {linha}"


def detectar_associacoes_em_analise_pendente(
    apontamentos: Sequence[Mapping[str, Any]],
) -> tuple[AlertaAssociacaoPendente, ...]:
    """Localiza vínculos provisórios sem tratá-los como erro bloqueante."""
    alertas = []
    rotulos = (
        ("falha", "falha"),
        ("multa", "multa"),
        ("repercussao", "repercussão"),
        ("debito", "débito"),
    )
    for linha, apontamento in enumerate(apontamentos, start=1):
        if _texto(apontamento, "conclusao") != "Análise Pendente":
            continue
        nomes = []
        naturezas = []
        for natureza, rotulo in rotulos:
            associados = obter_responsaveis_apontamento(
                apontamento,
                natureza,
            )
            if associados:
                naturezas.append(rotulo)
                nomes.extend(associados)
        nomes = list(dict.fromkeys(nomes))
        if nomes:
            alertas.append(
                AlertaAssociacaoPendente(
                    item=_identificar_item(apontamento, linha),
                    responsaveis=tuple(nomes),
                    naturezas=tuple(naturezas),
                )
            )
    return tuple(alertas)


def construir_status_classificacao(
    classificacao: Mapping[str, Any],
    apontamentos: Sequence[Mapping[str, Any]],
) -> StatusClassificacao:
    """Monta o texto e a cor do status sem conhecer widgets Tkinter."""
    pendentes = list(classificacao.get("pendentes", []))
    alertas = detectar_associacoes_em_analise_pendente(apontamentos)
    if pendentes:
        texto = (
            f"Revisão pendente: {len(pendentes)} item(ns) ainda "
            "precisa(m) de conclusão."
        )
        if alertas:
            texto += (
                f" Atenção: {len(alertas)} item(ns) pendente(s) já "
                "possui(em) Administrador associado."
            )
        return StatusClassificacao(texto, "warning", alertas)
    if classificacao.get("tem_apontamentos"):
        return StatusClassificacao(
            "Classificação consolidada: nenhum item pendente.",
            "success",
        )
    return StatusClassificacao(
        "Nenhum apontamento carregado.",
        "secondary",
    )

def avaliar_painel_preenchimento(
    dados: Mapping[str, Any],
) -> ResultadoPainelPreenchimento:
    """Separa bloqueios reais de alertas úteis para o painel da aplicação."""
    pendencias = []
    for chave, rotulo in (
        ("exercicio", "Exercício"),
        ("processo", "Processo"),
        ("tipo", "Tipo de processo"),
        ("orgao", "Órgão"),
    ):
        if not _texto(dados, chave):
            pendencias.append(rotulo)

    responsaveis = [
        dict(item)
        for item in dados.get("responsaveis", [])
        if isinstance(item, Mapping) and _texto(item, "nome")
    ]
    apontamentos = [
        dict(item)
        for item in dados.get("apontamentos", [])
        if isinstance(item, Mapping)
        and (_texto(item, "irregularidade") or _texto(item, "item"))
    ]
    if not responsaveis:
        pendencias.append("ao menos um responsável")
    elif any(not _texto(item, "cargo") for item in responsaveis):
        pendencias.append("cargo de todos os responsáveis")

    tipo = _texto(dados, "tipo")
    erros_conclusoes = validar_conclusoes_responsaveis(
        responsaveis,
        tipo,
    )
    if erros_conclusoes:
        pendencias.append(
            "conclusão individual dos administradores "
            f"({len(erros_conclusoes)} pendência(s))"
        )

    erros_vinculos = validar_vinculos_responsabilidade(
        responsaveis,
        apontamentos,
        tipo,
    )
    if erros_vinculos:
        pendencias.append(
            "certificação entre falhas e administradores "
            f"({len(erros_vinculos)} inconsistência(s))"
        )

    erros = tuple(erros_conclusoes + erros_vinculos)
    alertas = detectar_associacoes_em_analise_pendente(apontamentos)
    if pendencias:
        detalhes = "Ainda falta preencher ou corrigir: " + ", ".join(
            pendencias
        ) + "."
        amostra_erros = erros[:4]
        if amostra_erros:
            detalhes += "\n\n" + "\n".join(
                f"• {erro}" for erro in amostra_erros
            )
            if len(erros) > len(amostra_erros):
                detalhes += f"\n• ... e mais {len(erros) - len(amostra_erros)}."
        if alertas:
            detalhes += "\n\nALERTAS INFORMATIVOS:\n" + "\n".join(
                f"• {alerta.mensagem}" for alerta in alertas[:3]
            )
        return ResultadoPainelPreenchimento(
            pronto=False,
            resumo=f"{len(pendencias)} pendência(s) de preenchimento",
            detalhes=detalhes,
            estilo="warning",
            pendencias=tuple(pendencias),
            erros=erros,
            alertas=alertas,
            total_responsaveis=len(responsaveis),
        )

    if alertas:
        detalhes = (
            "Dados essenciais preenchidos. As associações abaixo são "
            "provisórias e não impedem o trabalho:\n\n"
            + "\n".join(f"• {alerta.mensagem}" for alerta in alertas[:4])
        )
        if len(alertas) > 4:
            detalhes += f"\n• ... e mais {len(alertas) - 4}."
        detalhes += (
            "\n\nRevise esses itens antes de executar CONCLUSÃO."
        )
        return ResultadoPainelPreenchimento(
            pronto=True,
            resumo=f"Dados prontos com {len(alertas)} alerta(s)",
            detalhes=detalhes,
            estilo="warning",
            pendencias=(),
            erros=(),
            alertas=alertas,
            total_responsaveis=len(responsaveis),
        )

    total = len(responsaveis)
    return ResultadoPainelPreenchimento(
        pronto=True,
        resumo="Dados essenciais prontos",
        detalhes=(
            f"Dados essenciais preenchidos. {total} responsável(is) "
            "será(ão) considerado(s) nas rotinas e no Word."
        ),
        estilo="success",
        pendencias=(),
        erros=(),
        alertas=(),
        total_responsaveis=total,
    )
