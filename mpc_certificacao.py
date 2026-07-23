"""Certificação preventiva dos dados antes de qualquer escrita no Word.

O módulo não conhece Tkinter nem a automação COM. Ele recebe um retrato dos
dados da interface e informa, de forma testável, se a operação pode prosseguir.
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from mpc_regras import (
    arquivo_real_esclarecimentos,
    validar_coerencia_esclarecimentos,
    validar_conclusoes_responsaveis,
    validar_vinculos_responsabilidade,
)


OPERACOES_WORD = frozenset(
    {
        "Analisar parecer com IA",
        "Modelo de Parecer",
        "e-Parecer",
        "Cabeçalho (e-Parecer)",
        "Cabeçalho",
        "Introdução",
        "Conclusão",
        "Ementa",
        "Resultado das Verificações",
        "Fundamentação Individual",
        "Inserir apontes no Word",
        "Inserir trecho da Biblioteca Local",
    }
)

OPERACOES_COM_WORD_ATIVO = OPERACOES_WORD - {"e-Parecer"}

_CAMPOS_POR_OPERACAO = {
    "Analisar parecer com IA": (
        ("exercicio", "Exercício"),
        ("processo", "Processo"),
        ("tipo", "Tipo de processo"),
        ("orgao", "Órgão"),
        ("pasta", "Pasta de trabalho"),
    ),
    "Modelo de Parecer": (
        ("tipo", "Tipo de processo"),
        ("procurador", "Procurador(a)"),
    ),
    "e-Parecer": (
        ("exercicio", "Exercício"),
        ("processo", "Processo"),
        ("tipo", "Tipo de processo"),
        ("orgao", "Órgão"),
        ("relator", "Relator(a)"),
        ("procurador", "Procurador(a)"),
        ("pasta", "Pasta de trabalho"),
    ),
    "Cabeçalho (e-Parecer)": (
        ("exercicio", "Exercício"),
        ("processo", "Processo"),
        ("orgao", "Órgão"),
        ("relator", "Relator(a)"),
    ),
    "Introdução": (
        ("tipo", "Tipo de processo"),
        ("orgao", "Órgão"),
        ("municipio", "Município"),
        ("relator", "Relator(a)"),
    ),
    "Conclusão": (
        ("tipo", "Tipo de processo"),
        ("municipio", "Município"),
        ("procurador", "Procurador(a)"),
    ),
    "Ementa": (
        ("tipo", "Tipo de processo"),
        ("procurador", "Procurador(a)"),
    ),
    "Resultado das Verificações": (
        ("tipo", "Tipo de processo"),
        ("procurador", "Procurador(a)"),
    ),
    "Fundamentação Individual": (
        ("processo", "Processo"),
        ("orgao", "Órgão"),
    ),
}

_OPERACOES_COM_RESPONSAVEIS = {
    "Analisar parecer com IA",
    "e-Parecer",
    "Cabeçalho (e-Parecer)",
    "Introdução",
    "Conclusão",
    "Ementa",
    "Resultado das Verificações",
}

_OPERACOES_COM_REVISAO_COMPLETA = {
    "Analisar parecer com IA",
    "Conclusão",
    "Ementa",
    "Resultado das Verificações",
}


@dataclass(frozen=True, slots=True)
class ResultadoCertificacao:
    """Resultado imutável da verificação preventiva."""

    operacao: str
    erros: tuple[str, ...]
    avisos: tuple[str, ...]
    verificacoes: tuple[str, ...]
    total_responsaveis: int
    total_apontamentos: int

    @property
    def aprovada(self) -> bool:
        return not self.erros


def _texto(dados: dict[str, Any], chave: str) -> str:
    return str(dados.get(chave, "") or "").strip()


def _responsaveis(dados: dict[str, Any]) -> list[dict[str, Any]]:
    return [
        item
        for item in dados.get("responsaveis", [])
        if isinstance(item, dict) and _texto(item, "nome")
    ]


def _apontamentos(dados: dict[str, Any]) -> list[dict[str, Any]]:
    origem = dados.get("apontamentos_detalhado", [])
    return [
        item
        for item in origem
        if isinstance(item, dict)
        and str(item.get("irregularidade") or item.get("item") or "").strip()
    ]


def _adicionar_sem_repetir(destino: list[str], mensagens: list[str]) -> None:
    for mensagem in mensagens:
        if mensagem not in destino:
            destino.append(mensagem)


def _validar_responsaveis_basicos(
    responsaveis: list[dict[str, Any]],
) -> list[str]:
    if not responsaveis:
        return ["Informe ao menos um administrador no quadro Responsáveis."]

    erros = []
    nomes_vistos: set[str] = set()
    for linha, responsavel in enumerate(responsaveis, start=1):
        nome = _texto(responsavel, "nome")
        cargo = _texto(responsavel, "cargo")
        sexo = _texto(responsavel, "sexo")
        chave_nome = nome.casefold()
        if not cargo:
            erros.append(f"Linha {linha} — '{nome}': informe o Cargo.")
        if sexo not in {"M", "F"}:
            erros.append(
                f"Linha {linha} — '{nome}': selecione o Sexo (M ou F) para "
                "a concordância do texto."
            )
        if chave_nome in nomes_vistos:
            erros.append(
                f"Linha {linha} — o administrador '{nome}' aparece mais de "
                "uma vez; as associações ficariam ambíguas."
            )
        nomes_vistos.add(chave_nome)
    return erros


def _validar_revisao_apontamentos(
    apontamentos: list[dict[str, Any]],
) -> list[str]:
    erros = []
    pendentes = []
    for linha, apontamento in enumerate(apontamentos, start=1):
        conclusao = _texto(apontamento, "conclusao")
        if conclusao in {"", "Análise Pendente"}:
            descricao = str(
                apontamento.get("irregularidade")
                or apontamento.get("item")
                or f"linha {linha}"
            ).strip()
            numero = re.match(r"\d+(?:\.\d+)*", descricao)
            pendentes.append(numero.group(0) if numero else f"linha {linha}")
    if pendentes:
        amostra = ", ".join(pendentes[:8])
        complemento = f" e mais {len(pendentes) - 8}" if len(pendentes) > 8 else ""
        erros.append(
            "Conclua a análise de todos os apontamentos. Permanecem "
            f"pendentes: {amostra}{complemento}."
        )
    return erros


def _validar_arquivos_ia(
    dados: dict[str, Any],
    responsaveis: list[dict[str, Any]],
) -> list[str]:
    """Confere somente arquivos que a operação de IA realmente tentará ler."""
    erros = []
    pasta = _texto(dados, "pasta")
    if pasta and not os.path.isdir(pasta):
        erros.append(
            f"A Pasta de trabalho não existe ou não está acessível: {pasta}."
        )

    for responsavel in responsaveis:
        arquivo = _texto(responsavel, "arquivo_esclarecimentos")
        if not arquivo_real_esclarecimentos(arquivo):
            continue
        candidato = Path(arquivo)
        if not candidato.is_absolute() and pasta:
            candidato = Path(pasta) / candidato
        if not candidato.is_file():
            erros.append(
                f"O PDF de esclarecimentos de '{_texto(responsavel, 'nome')}' "
                f"não foi localizado: {arquivo}."
            )
    return erros


def certificar_dados_pre_word(
    operacao: str,
    dados: dict[str, Any],
) -> ResultadoCertificacao:
    """Certifica o retrato atual da GUI conforme a operação solicitada."""
    erros: list[str] = []
    avisos: list[str] = []
    verificacoes: list[str] = []
    responsaveis = _responsaveis(dados)
    apontamentos = _apontamentos(dados)
    tipo_processo = _texto(dados, "tipo")

    for chave, rotulo in _CAMPOS_POR_OPERACAO.get(operacao, ()):
        if not _texto(dados, chave):
            erros.append(f"Preencha o campo '{rotulo}'.")
    if _CAMPOS_POR_OPERACAO.get(operacao):
        verificacoes.append("campos essenciais da operação")

    if operacao in _OPERACOES_COM_RESPONSAVEIS:
        _adicionar_sem_repetir(
            erros,
            _validar_responsaveis_basicos(responsaveis),
        )
        verificacoes.append("identificação e concordância dos responsáveis")

    if operacao == "Introdução":
        _adicionar_sem_repetir(
            erros,
            validar_coerencia_esclarecimentos(responsaveis, tipo_processo),
        )
        if tipo_processo.upper() == "CONTAS ANUAIS":
            _adicionar_sem_repetir(
                erros,
                _validar_revisao_apontamentos(apontamentos),
            )
        verificacoes.append("intimação e esclarecimentos individualizados")

    if operacao in _OPERACOES_COM_REVISAO_COMPLETA:
        _adicionar_sem_repetir(
            erros,
            validar_conclusoes_responsaveis(responsaveis, tipo_processo),
        )
        _adicionar_sem_repetir(
            erros,
            _validar_revisao_apontamentos(apontamentos),
        )
        _adicionar_sem_repetir(
            erros,
            validar_vinculos_responsabilidade(
                responsaveis,
                apontamentos,
                tipo_processo,
            ),
        )
        verificacoes.append(
            "conclusões, falhas, multa, repercussão, débito e associações"
        )

    if operacao == "Inserir apontes no Word":
        if not apontamentos:
            erros.append(
                "Não há apontamentos preenchidos para inserir no documento."
            )
        verificacoes.append("existência de apontamentos para inserção")

    if operacao == "Analisar parecer com IA":
        _adicionar_sem_repetir(
            erros,
            _validar_arquivos_ia(dados, responsaveis),
        )
        verificacoes.append("pasta e PDFs individualizados usados pela IA")

    if operacao in {"Conclusão", "Ementa", "Resultado das Verificações"}:
        if not apontamentos:
            avisos.append(
                "Nenhum apontamento foi informado. A certificação considera "
                "que o processo pode não possuir falhas."
            )

    return ResultadoCertificacao(
        operacao=operacao,
        erros=tuple(erros),
        avisos=tuple(avisos),
        verificacoes=tuple(dict.fromkeys(verificacoes)),
        total_responsaveis=len(responsaveis),
        total_apontamentos=len(apontamentos),
    )


def possivel_divergencia_documento_processo(
    nome_documento: str,
    numero_processo: str,
) -> bool:
    """Sinaliza, sem bloquear, quando o processo não aparece no nome do Word."""
    documento = re.sub(r"\D", "", str(nome_documento or ""))
    processo = re.sub(r"\D", "", str(numero_processo or ""))
    if not documento or len(processo) < 6:
        return False
    # Os nomes podem omitir separadores, dígito final ou parte do exercício.
    prefixo = processo[: min(10, len(processo))]
    return prefixo not in documento


def certificar_estrutura_documento(
    operacao: str,
    texto_documento: str,
    titulos_controles: list[str] | tuple[str, ...] = (),
) -> tuple[list[str], list[str], list[str]]:
    """Confere marcadores e controles sem modificar o documento Word."""
    erros: list[str] = []
    avisos: list[str] = []
    verificacoes: list[str] = []
    texto = str(texto_documento or "")
    marcadores_obrigatorios = {
        "Introdução": ("[INTRODUÇÃO]",),
        "Conclusão": ("[CONCLUSÃO]", "[DISPOSITIVO]"),
        "Ementa": ("[EMENTA]",),
    }.get(operacao, ())
    ausentes = [
        marcador for marcador in marcadores_obrigatorios if marcador not in texto
    ]
    if ausentes:
        erros.append(
            "O documento ativo não contém o(s) marcador(es) obrigatório(s) "
            f"para {operacao}: {', '.join(ausentes)}. Abra o modelo correto "
            "ou restaure os marcadores antes de continuar."
        )
    if marcadores_obrigatorios and not ausentes:
        verificacoes.append("marcadores obrigatórios do modelo Word")

    marcador_resultado = "[RESULTADO DAS VERIFICAÇÕES PROCEDIDAS]"
    if operacao == "Resultado das Verificações":
        if marcador_resultado not in texto:
            avisos.append(
                f"O marcador {marcador_resultado} não foi encontrado. Se "
                "prosseguir, a própria rotina oferecerá a inserção na posição "
                "atual do cursor."
            )
        else:
            verificacoes.append("marcador do Resultado das Verificações")

    titulos = {str(titulo or "").strip() for titulo in titulos_controles}
    if operacao == "Cabeçalho (e-Parecer)":
        if not any(re.fullmatch(r"Gestor_\d+", titulo) for titulo in titulos):
            erros.append(
                "O modelo ativo não contém um controle 'Gestor_1', "
                "necessário para preencher e ajustar a tabela de responsáveis."
            )
        else:
            verificacoes.append("linha-base da tabela de responsáveis no Word")
        gerais_ausentes = [
            titulo
            for titulo in ("Processo", "Relator", "Ano", "Órgão")
            if titulo not in titulos
        ]
        if gerais_ausentes:
            avisos.append(
                "O modelo não possui estes controles gerais do cabeçalho: "
                f"{', '.join(gerais_ausentes)}. Confira se o arquivo aberto é "
                "o modelo esperado."
            )

    if operacao == "Cabeçalho":
        controles_conhecidos = {
            "Nome do relator",
            "Tipo de processo",
            "Nome do interessado",
            "Órgão",
        }
        if not (titulos & controles_conhecidos):
            avisos.append(
                "Nenhum dos controles normalmente tratados pelo Cabeçalho foi "
                "localizado. A rotina pode não produzir alteração visível."
            )
        else:
            verificacoes.append("controles de conteúdo do cabeçalho")

    return erros, avisos, verificacoes
