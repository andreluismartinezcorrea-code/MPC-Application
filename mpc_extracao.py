"""Extração local e normalização de respostas estruturadas de documentos."""

from __future__ import annotations

import os
import re
import unicodedata
from pathlib import Path
from typing import Any, Mapping, Sequence

import fitz
import PyPDF2

from mpc_ia import carregar_json_resposta_ia


ESQUEMA_RELATORIO_AUDITORIA = {
    "type": "object",
    "properties": {
        "Processo": {"type": "string"},
        "Exercicio": {"type": "string"},
        "Orgao": {"type": "string"},
        "Tipo": {"type": "string"},
        "ServicoAuditoria": {"type": "string"},
        "Apontamentos": {"type": "array", "items": {"type": "string"}},
        "QuantidadeApontamentos": {"type": "integer"},
        "SugestoesRecomendacoes": {
            "type": "array",
            "items": {"type": "string"},
        },
        "Alertas": {
            "type": "array",
            "items": {"type": "string"},
        },
        "Recomendacoes": {
            "type": "array",
            "items": {"type": "string"},
        },
        "QuantidadeSugestoes": {"type": "integer"},
        "Gestores": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "nome": {"type": "string"},
                    "cargo": {"type": "string"},
                },
                "required": ["nome", "cargo"],
            },
        },
    },
    "required": [
        "Processo",
        "Exercicio",
        "Orgao",
        "Tipo",
        "ServicoAuditoria",
        "Apontamentos",
        "QuantidadeApontamentos",
        "SugestoesRecomendacoes",
        "Alertas",
        "Recomendacoes",
        "QuantidadeSugestoes",
        "Gestores",
    ],
}


ESQUEMA_LISTA_APONTAMENTOS_RAG = {
    "type": "object",
    "properties": {
        "Itens": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "Numero": {"type": "string"},
                    "Descricao": {"type": "string"},
                },
                "required": ["Numero", "Descricao"],
            },
        },
    },
    "required": ["Itens"],
}


ESQUEMA_ALERTAS_RECOMENDACOES_RAG = {
    "type": "object",
    "properties": {
        "Alertas": {
            "type": "array",
            "items": {"type": "string"},
        },
        "Recomendacoes": {
            "type": "array",
            "items": {"type": "string"},
        },
    },
    "required": ["Alertas", "Recomendacoes"],
}


PROMPT_RELATORIO_AUDITORIA = """
Examine minuciosamente o Relatório de Auditoria fornecido e extraia as
informações solicitadas pelo esquema de resposta. Todas as strings devem ser
limpas.

REGRAS ESTREITAS E OBRIGATÓRIAS PARA CADA CAMPO:

* **Processo:** número formatado estritamente como "000000-0200/00-0".
* **Exercicio:** ano do exercício examinado, com quatro algarismos.
* **Orgao:** nome do órgão sempre em MAIÚSCULAS.
* **Tipo:** em MAIÚSCULAS. Para contas, retorne apenas "CONTAS ANUAIS" ou
  "CONTAS ORDINÁRIAS". Nunca inclua a palavra "RELATÓRIO".
* **ServicoAuditoria:** nome localizado no cabeçalho do documento.
* **Apontamentos:** somente as numerações dos itens da tabela de
  falhas/irregularidades. Sem descrição. Se não houver falhas, retorne
  ["Relatório Sem Falhas"] e quantidade zero.
* **Alertas:** somente as numerações dos achados que o próprio Relatório de
  Auditoria classifica ou propõe tratar como alerta. Examine o documento
  inteiro, especialmente conclusão, síntese e proposta de encaminhamento.
* **Recomendacoes:** somente as numerações dos achados que o próprio Relatório
  classifica ou propõe tratar como recomendação. Examine também tabelas,
  quadros, notas e textos que continuem na página seguinte.
* **SugestoesRecomendacoes:** união de Alertas e Recomendacoes, sem repetição.
* **QuantidadeSugestoes:** quantidade de numerações distintas da união.
  Não confunda números de páginas, exercícios, processos, leis, artigos,
  valores monetários ou prazos com a numeração dos achados de auditoria.
* **Gestores:** objetos com nome e cargo; simplifique "Prefeito Municipal"
  para "Prefeito".
"""


PROMPT_LISTA_APONTAMENTOS_RAG = """
Examine exclusivamente o Relatório de Auditoria fornecido e identifique todos
os achados de auditoria numerados: falhas, irregularidades, apontamentos,
alertas e recomendações. Para cada achado, devolva a numeração original em
Numero e o respectivo título ou descrição curta em Descricao.

REGRAS OBRIGATÓRIAS:

* Examine o relatório inteiro, inclusive tabelas, quadros, anexos, conclusão,
  síntese dos apontamentos e proposta de encaminhamento.
* Reúna títulos quebrados entre linhas ou páginas.
* Não invente numeração ou descrição.
* Não inclua títulos meramente estruturais do documento.
* Não confunda páginas, exercícios, processos, leis, artigos, valores,
  percentuais, prazos ou referências citadas com achados de auditoria.
* Se o mesmo achado reaparecer na conclusão, devolva-o apenas uma vez.
* Se não houver achados numerados, devolva Itens como lista vazia.
"""


PROMPT_ALERTAS_RECOMENDACOES_RAG = """
Faça uma varredura especializada e integral do Relatório de Auditoria
fornecido. Localize os achados que o próprio relatório classifica, converte,
propõe ou encaminha como ALERTA ou RECOMENDAÇÃO.

REGRAS OBRIGATÓRIAS:

* Examine especialmente tabelas, quadros, conclusão, síntese dos apontamentos
  e proposta de encaminhamento, inclusive continuações na página seguinte.
* Em Alertas, devolva somente a numeração original de cada achado tratado como
  alerta.
* Em Recomendacoes, devolva somente a numeração original de cada achado
  tratado como recomendação.
* Não inclua uma falha apenas porque o texto recomenda sua manutenção,
  responsabilização, multa ou correção; é preciso que o encaminhamento seja
  efetivamente alerta ou recomendação.
* Não confunda páginas, exercícios, processos, leis, artigos, valores,
  percentuais ou prazos com números de achados.
* Não invente numerações. Elimine repetições.
* Se nenhuma ocorrência existir, devolva as duas listas vazias.
"""


def _texto(dados: Mapping[str, Any], chave: str) -> str:
    return str(dados.get(chave, "") or "").strip()


def _inteiro_seguro(valor: Any) -> int | None:
    try:
        return int(valor)
    except (TypeError, ValueError):
        return None


def _chave_sem_acentos(valor: Any) -> str:
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    return "".join(
        caractere for caractere in texto if not unicodedata.combining(caractere)
    ).strip().casefold()


def formatar_lista_portugues(itens: Sequence[str]) -> str:
    itens = [str(item).strip() for item in itens if str(item).strip()]
    if not itens:
        return ""
    if len(itens) == 1:
        return itens[0]
    return ", ".join(itens[:-1]) + " e " + itens[-1]


def _normalizar_numeracoes(valor: Any) -> tuple[list[str], list[str]]:
    if isinstance(valor, (list, tuple, set)):
        origem = list(valor)
    elif valor is None:
        origem = []
    else:
        origem = [valor]
    numeros = []
    ignorados = []
    for item in origem:
        texto = str(item or "").strip()
        encontrados = re.findall(r"(?<!\d)(\d+(?:\.\d+)+)(?!\d)", texto)
        if not encontrados:
            if texto:
                ignorados.append(texto)
            continue
        for numero in encontrados:
            if numero not in numeros:
                numeros.append(numero)
    return numeros, ignorados


def _combinar_numeracoes_campos(
    dados: Mapping[str, Any],
    *campos: str,
) -> tuple[list[str], list[str]]:
    numeros: list[str] = []
    ignorados: list[str] = []
    vistos: set[str] = set()
    for campo in campos:
        encontrados, descartados = _normalizar_numeracoes(dados.get(campo, []))
        ignorados.extend(descartados)
        for numero in encontrados:
            if numero not in vistos:
                vistos.add(numero)
                numeros.append(numero)
    return numeros, ignorados


def normalizar_tipo_processo(valor: Any) -> str:
    tipo = re.sub(r"\s+", " ", str(valor or "").strip().upper())
    tipo = re.sub(r"\bRELAT[ÓO]RIO\b", "", tipo)
    tipo = re.sub(r"\s+", " ", tipo).strip(" -–—")
    if "CONTAS ANUAIS" in tipo:
        return "CONTAS ANUAIS"
    if "CONTAS ORDINÁRIAS" in tipo or "CONTAS ORDINARIAS" in tipo:
        return "CONTAS ORDINÁRIAS"
    return tipo


def formatar_servico_auditoria(valor: Any) -> str:
    texto = re.sub(r"\s+", " ", str(valor or "").strip())
    preposicoes = {"de", "do", "da", "dos", "das", "e"}
    palavras = [
        palavra.lower()
        if indice > 0 and palavra.casefold() in preposicoes
        else palavra.capitalize()
        for indice, palavra in enumerate(texto.split())
    ]
    return " ".join(palavras)


def normalizar_numero_processo(valor: Any) -> str:
    processo = re.sub(r"\s+", "", str(valor or "").strip())
    correspondencia = re.fullmatch(r"(\d+)-(\d{3,4})/(\d{2})-(\d)", processo)
    if correspondencia:
        primeiro, unidade, ano, digito = correspondencia.groups()
        return f"{primeiro.zfill(6)}-{unidade}/{ano}-{digito}"
    return processo


def normalizar_gestores(valor: Any) -> list[dict[str, str]]:
    if not isinstance(valor, list):
        return []
    gestores = []
    nomes_vistos = set()
    for item in valor:
        if not isinstance(item, Mapping):
            continue
        nome = _texto(item, "nome")
        cargo = _texto(item, "cargo")
        if not nome or nome.casefold() in nomes_vistos:
            continue
        if cargo.upper() == "PREFEITO MUNICIPAL":
            cargo = "Prefeito"
        elif cargo.upper() == "VICE-PREFEITO MUNICIPAL":
            cargo = "Vice-Prefeito"
        nomes_vistos.add(nome.casefold())
        gestores.append({"nome": nome, "cargo": cargo})
    return gestores


def normalizar_relatorio_auditoria(
    dados: Mapping[str, Any],
    caminho_arquivo: str | os.PathLike[str],
    *,
    metodo_extracao: str,
) -> dict[str, Any]:
    """Converte a resposta da IA no contrato consumido pela interface."""
    if not isinstance(dados, Mapping):
        raise ValueError("A resposta estruturada do relatório não é um objeto JSON.")

    nome_arquivo = Path(caminho_arquivo).name
    apontamentos_brutos = dados.get("Apontamentos", [])
    origem_apontamentos = (
        apontamentos_brutos
        if isinstance(apontamentos_brutos, list)
        else [apontamentos_brutos]
    )
    sem_falhas = any(
        _chave_sem_acentos(item) == "relatorio sem falhas"
        for item in (
            origem_apontamentos
        )
    )
    apontamentos, ignorados_apontamentos = _normalizar_numeracoes(
        [
            item
            for item in origem_apontamentos
            if _chave_sem_acentos(item) != "relatorio sem falhas"
        ]
    )
    sugestoes, ignoradas_sugestoes = _combinar_numeracoes_campos(
        dados,
        "SugestoesRecomendacoes",
        "Alertas",
        "Recomendacoes",
    )
    if sem_falhas:
        apontamentos = []

    avisos = []
    quantidade_informada = _inteiro_seguro(
        dados.get("QuantidadeApontamentos")
    )
    if quantidade_informada is not None and quantidade_informada != len(apontamentos):
        avisos.append(
            "A quantidade de apontamentos informada pela IA divergia da lista; "
            "o programa recalculou o total."
        )
    quantidade_sugestoes = _inteiro_seguro(dados.get("QuantidadeSugestoes"))
    if quantidade_sugestoes is not None and quantidade_sugestoes != len(sugestoes):
        avisos.append(
            "A quantidade de recomendações informada pela IA divergia da lista; "
            "o programa recalculou o total."
        )
    if ignorados_apontamentos:
        avisos.append(
            f"{len(ignorados_apontamentos)} apontamento(s) sem numeração válida "
            "foi(ram) desconsiderado(s)."
        )
    if ignoradas_sugestoes:
        avisos.append(
            f"{len(ignoradas_sugestoes)} recomendação(ões) sem numeração válida "
            "foi(ram) desconsiderada(s)."
        )

    return {
        "processo": normalizar_numero_processo(dados.get("Processo")),
        "exercicio": _texto(dados, "Exercicio"),
        "orgao": _texto(dados, "Orgao").upper(),
        "tipo": normalizar_tipo_processo(dados.get("Tipo")),
        "servico": formatar_servico_auditoria(
            dados.get("ServicoAuditoria")
        ),
        "nome_arquivo": nome_arquivo,
        "peca": nome_arquivo[5:13],
        "apontes": (
            "Relatório Sem Falhas"
            if sem_falhas
            else formatar_lista_portugues(apontamentos) or "Não encontrado"
        ),
        "quantidade_de_apontamentos": str(len(apontamentos)),
        "sugestoes_rec": formatar_lista_portugues(sugestoes),
        "qtd_sugestoes": str(len(sugestoes)),
        "gestores_cargos": normalizar_gestores(dados.get("Gestores", [])),
        "metodo_extracao": str(metodo_extracao),
        "avisos_extracao": avisos,
    }


def normalizar_lista_apontamentos_rag(
    dados: Mapping[str, Any],
    *,
    limite: int | None = None,
) -> dict[str, Any]:
    """Normaliza a lista estruturada de achados extraída do RAG."""
    if not isinstance(dados, Mapping):
        raise ValueError("A resposta da lista de apontamentos não é um objeto JSON.")
    origem = dados.get("Itens", [])
    if not isinstance(origem, list):
        raise ValueError("O campo 'Itens' da resposta deve ser uma lista.")

    itens: list[str] = []
    descartadas: list[str] = []
    duplicadas: list[str] = []
    vistos: set[str] = set()
    for item in origem:
        if not isinstance(item, Mapping):
            descartadas.append(str(item or "").strip())
            continue
        numero_bruto = _texto(item, "Numero")
        descricao = re.sub(r"\s+", " ", _texto(item, "Descricao")).strip(" .;-–—")
        numeros, _ignorados = _normalizar_numeracoes(numero_bruto)
        if len(numeros) != 1 or not descricao:
            descartadas.append(
                " ".join(parte for parte in (numero_bruto, descricao) if parte)
            )
            continue
        numero = numeros[0]
        if numero in vistos:
            duplicadas.append(numero)
            continue
        vistos.add(numero)
        itens.append(f"{numero} {descricao}")

    excedentes: list[str] = []
    if limite is not None and len(itens) > max(0, limite):
        excedentes = itens[max(0, limite):]
        itens = itens[:max(0, limite)]
    return {
        "itens": itens,
        "descartadas": descartadas,
        "duplicadas": duplicadas,
        "excedentes": excedentes,
    }


def normalizar_alertas_recomendacoes_rag(
    dados: Mapping[str, Any],
) -> dict[str, Any]:
    """Une alertas e recomendações do RAG, sem duplicar suas numerações."""
    if not isinstance(dados, Mapping):
        raise ValueError(
            "A resposta de alertas e recomendações não é um objeto JSON."
        )
    numeros, ignorados = _combinar_numeracoes_campos(
        dados,
        "Alertas",
        "Recomendacoes",
        "SugestoesRecomendacoes",
    )
    return {
        "numeracoes": numeros,
        "texto": formatar_lista_portugues(numeros),
        "quantidade": len(numeros),
        "descartadas": ignorados,
    }


def normalizar_processos_tramitacao(
    resposta: str | Mapping[str, Any],
    *,
    limite: int = 2,
) -> dict[str, Any]:
    """Aceita JSON puro/Markdown e limita o resultado aos campos da GUI."""
    dados = (
        dict(resposta)
        if isinstance(resposta, Mapping)
        else carregar_json_resposta_ia(str(resposta or ""))
    )
    origem = dados.get("processos", [])
    if not isinstance(origem, list):
        raise ValueError("O campo 'processos' da resposta deve ser uma lista.")
    processos = []
    ignorados = 0
    vistos = set()
    for item in origem:
        if not isinstance(item, Mapping):
            ignorados += 1
            continue
        tipo = _texto(item, "tipo")
        numero = normalizar_numero_processo(item.get("numero"))
        if not tipo or not numero:
            ignorados += 1
            continue
        chave = (tipo.casefold(), numero)
        if chave in vistos:
            continue
        vistos.add(chave)
        processos.append({"tipo": tipo, "numero": numero})
    excedentes = max(0, len(processos) - max(0, limite))
    processos = processos[: max(0, limite)]
    return {
        "tem_processos": bool(processos),
        "processos": processos,
        "ignorados": ignorados,
        "excedentes": excedentes,
    }


def extrair_texto_pdf_seguro(caminho_pdf: str | os.PathLike[str]) -> str:
    """Extrai texto com PyMuPDF, retornando vazio somente se a leitura falhar."""
    try:
        with fitz.open(caminho_pdf) as documento:
            if documento.page_count == 0:
                return ""
            return "\n".join(
                pagina.get_text("text") for pagina in documento
            )
    except Exception:
        return ""


def extrair_texto_pdf_para_ia(caminho_pdf: str | os.PathLike[str]) -> str:
    """Tenta duas bibliotecas e explica PDFs de imagem ou protegidos."""
    erros = []
    texto = extrair_texto_pdf_seguro(caminho_pdf).strip()
    if texto:
        return texto
    erros.append("PyMuPDF não encontrou texto selecionável")
    try:
        with open(caminho_pdf, "rb") as arquivo_pdf:
            leitor = PyPDF2.PdfReader(arquivo_pdf)
            if leitor.is_encrypted:
                try:
                    leitor.decrypt("")
                except Exception as erro:
                    raise ValueError("O PDF está protegido por senha.") from erro
            texto = "\n".join(
                (pagina.extract_text() or "") for pagina in leitor.pages
            ).strip()
            if texto:
                return texto
        erros.append("PyPDF2 não encontrou texto selecionável")
    except Exception as erro:
        erros.append(f"PyPDF2: {erro}")
    raise ValueError(
        "Não foi possível obter texto do PDF. Ele pode ser digitalizado "
        "(formado apenas por imagens) ou estar protegido por senha. Abra-o "
        "no Adobe Reader; se não for possível selecionar uma frase, aplique "
        "OCR ou use a versão textual. Detalhes técnicos: "
        + "; ".join(erros)
    )
