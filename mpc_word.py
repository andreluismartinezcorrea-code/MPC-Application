"""Serviços comuns para automação segura do Microsoft Word."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


MARCADORES_CONCLUSAO_WORD = (
    "[FALHAS/CR]",
    "[FALHAS/SR]",
    "[S-FALHAS]",
    "[M-FALHAS]",
    "[NS-FALHAS]",
    "[EM-FALHAS]",
    "[S-FALHAS/CR]",
    "[M-FALHAS/CR]",
    "[NS-FALHAS/CR]",
    "[EM-FALHAS/CR]",
    "[S-FALHAS/SR]",
    "[M-FALHAS/SR]",
    "[NS-FALHAS/SR]",
    "[EM-FALHAS/SR]",
    "[RESPONSABILIZAÇÃO]",
    "[NÃO RESPONSABILIZAÇÃO]",
    "[CONCLUSÃO]",
    "[DISPOSITIVO]",
)


@dataclass
class PlanoMarcadoresFalhas:
    """Substituições e exclusões necessárias para os marcadores de falhas."""

    substituicoes: dict[str, str] = field(default_factory=dict)
    excluir_paragrafos_com: tuple[str, ...] = ()


def obter_aplicacao_word(
    cliente_com: Any,
    *,
    criar_se_necessario: bool = False,
    visivel: bool = True,
):
    """Obtém a instância ativa e, opcionalmente, abre o Word."""
    try:
        word = cliente_com.GetActiveObject("Word.Application")
    except Exception as erro_ativo:
        if not criar_se_necessario:
            raise RuntimeError(
                "O Microsoft Word não está aberto ou não pôde ser acessado."
            ) from erro_ativo
        word = cliente_com.Dispatch("Word.Application")
    if visivel:
        word.Visible = True
    return word


def obter_documento_word_ativo(
    cliente_com: Any,
    *,
    criar_word_se_necessario: bool = False,
    criar_documento_se_necessario: bool = False,
    visivel: bool = True,
):
    """Retorna ``(aplicação, documento)`` com mensagens de erro legíveis."""
    word = obter_aplicacao_word(
        cliente_com,
        criar_se_necessario=criar_word_se_necessario,
        visivel=visivel,
    )
    quantidade = int(getattr(word.Documents, "Count", 0))
    if quantidade == 0 and criar_documento_se_necessario:
        word.Documents.Add()
        quantidade = int(getattr(word.Documents, "Count", 0))
    if quantidade == 0:
        raise RuntimeError(
            "O Microsoft Word está aberto, mas não há documento ativo."
        )
    try:
        documento = word.ActiveDocument
    except Exception as erro:
        raise RuntimeError(
            "Não foi possível acessar o documento ativo do Microsoft Word."
        ) from erro
    return word, documento


def planejar_ajuste_linhas_word(
    quantidade_responsaveis: int,
    linhas_modelo: int,
) -> dict[str, int]:
    """Calcula linhas a preencher, acrescentar ou excluir no modelo Word."""
    quantidade_responsaveis = max(0, int(quantidade_responsaveis))
    linhas_modelo = max(0, int(linhas_modelo))
    return {
        "preencher": min(quantidade_responsaveis, linhas_modelo),
        "adicionar": max(0, quantidade_responsaveis - linhas_modelo),
        "remover": max(0, linhas_modelo - quantidade_responsaveis),
    }


def _quantidade_inteira(valor) -> int:
    texto = str(valor or "").strip()
    return int(texto) if texto.isdigit() else 0


def _sufixos_falhas(quantidade: int, escopo: str = "") -> dict[str, str]:
    sufixo = f"/{escopo}" if escopo else ""
    if quantidade > 1:
        valores = {"S": "s", "M": "m", "NS": "ns", "EM": "em"}
    else:
        valores = {"S": "", "M": "", "NS": "m", "EM": ""}
    return {
        f"[{chave}-FALHAS{sufixo}]": valor
        for chave, valor in valores.items()
    }


def planejar_marcadores_falhas(
    quantidade_total,
    quantidade_com_responsabilidade,
    quantidade_sem_responsabilidade,
) -> PlanoMarcadoresFalhas:
    """Calcula a concordância e os parágrafos ausentes sem acessar o Word."""
    total = _quantidade_inteira(quantidade_total)
    com_responsabilidade = _quantidade_inteira(
        quantidade_com_responsabilidade
    )
    sem_responsabilidade = _quantidade_inteira(
        quantidade_sem_responsabilidade
    )
    substituicoes = _sufixos_falhas(total)
    excluir = []

    if com_responsabilidade > 0:
        substituicoes.update(
            _sufixos_falhas(com_responsabilidade, "CR")
        )
        substituicoes["[RESPONSABILIZAÇÃO] "] = ""
    else:
        excluir.append("[RESPONSABILIZAÇÃO]")

    if sem_responsabilidade > 0:
        substituicoes.update(
            _sufixos_falhas(sem_responsabilidade, "SR")
        )
        substituicoes["[NÃO RESPONSABILIZAÇÃO] "] = ""
    else:
        excluir.append("[NÃO RESPONSABILIZAÇÃO] A[S-FALHAS/SR]")

    return PlanoMarcadoresFalhas(
        substituicoes=substituicoes,
        excluir_paragrafos_com=tuple(excluir),
    )


def aplicar_negrito_exceto_conjuncao(intervalo: Any, texto: str) -> None:
    """Insere uma lista em negrito, preservando a conjunção ``e`` normal."""
    intervalo.Text = ""
    for caractere in str(texto or ""):
        intervalo.InsertAfter(caractere)
        fim = intervalo.End
        trecho = intervalo.Document.Range(fim - 1, fim)
        trecho.Font.Bold = caractere.lower() != "e"


def substituir_lista_falhas(
    documento: Any,
    marcador: str,
    texto: str,
) -> None:
    """Substitui todas as ocorrências de uma lista com o negrito histórico."""
    busca = documento.Content.Find
    busca.ClearFormatting()
    busca.Text = marcador
    while busca.Execute(FindText=marcador, Forward=True):
        aplicar_negrito_exceto_conjuncao(busca.Parent, texto)


def aplicar_marcadores_falhas(
    documento: Any,
    *,
    quantidade_total,
    quantidade_com_responsabilidade,
    quantidade_sem_responsabilidade,
    falhas_com_responsabilidade: str,
    falhas_sem_responsabilidade: str,
    wd_replace_all: int = 2,
) -> PlanoMarcadoresFalhas:
    """Aplica listas, concordância e exclusões dos parágrafos de falhas."""
    quantidade_com = _quantidade_inteira(quantidade_com_responsabilidade)
    quantidade_sem = _quantidade_inteira(quantidade_sem_responsabilidade)
    if quantidade_com > 0:
        substituir_lista_falhas(
            documento,
            "[FALHAS/CR]",
            falhas_com_responsabilidade,
        )
    if quantidade_sem > 0:
        substituir_lista_falhas(
            documento,
            "[FALHAS/SR]",
            falhas_sem_responsabilidade,
        )

    plano = planejar_marcadores_falhas(
        quantidade_total,
        quantidade_com,
        quantidade_sem,
    )
    for paragrafo in documento.Paragraphs:
        texto_paragrafo = str(paragrafo.Range.Text or "")
        if any(
            marcador in texto_paragrafo
            for marcador in plano.excluir_paragrafos_com
        ):
            paragrafo.Range.Text = ""
            continue
        for marcador, substituicao in plano.substituicoes.items():
            if marcador not in str(paragrafo.Range.Text or ""):
                continue
            paragrafo.Range.Find.Execute(
                FindText=marcador,
                ReplaceWith=substituicao,
                Replace=wd_replace_all,
            )
    return plano


def localizar_marcadores_residuais(documento: Any) -> list[str]:
    """Lista marcadores conhecidos que ainda permaneceram no documento."""
    texto = str(documento.Content.Text or "")
    return [
        marcador
        for marcador in MARCADORES_CONCLUSAO_WORD
        if marcador in texto
    ]


def validar_ausencia_marcadores(documento: Any) -> None:
    """Impede concluir silenciosamente um documento com marcador conhecido."""
    residuais = localizar_marcadores_residuais(documento)
    if residuais:
        lista = ", ".join(residuais)
        raise RuntimeError(
            "O documento ainda contém marcadores não processados: "
            f"{lista}. Verifique se o modelo Word é compatível com esta versão."
        )


def substituir_marcador_por_texto(
    documento: Any,
    marcador: str,
    texto: str,
):
    """Substitui um marcador e devolve o intervalo exato do texto inserido."""
    busca = documento.Content.Find
    busca.ClearFormatting()
    busca.Text = marcador
    if not busca.Execute():
        return None

    intervalo_marcador = busca.Parent
    if not texto:
        intervalo_marcador.Paragraphs(1).Range.Delete()
        return None

    inicio = intervalo_marcador.Start
    intervalo_marcador.Text = texto
    return documento.Range(Start=inicio, End=inicio + len(texto))


def aplicar_destaques_fundamentacao(
    intervalo: Any,
    destaques,
    *,
    wd_find_stop: int = 0,
) -> None:
    """Aplica as instruções de negrito produzidas pelo motor jurídico."""
    if intervalo is None:
        return
    for destaque in destaques or []:
        texto = str(getattr(destaque, "texto", destaque) or "")
        if not texto:
            continue
        intervalo_busca = intervalo.Duplicate
        busca = intervalo_busca.Find
        busca.ClearFormatting()
        busca.Text = texto
        busca.Forward = True
        busca.Wrap = wd_find_stop
        if not busca.Execute():
            continue
        trecho = busca.Parent
        trecho.Font.Bold = True
        if getattr(destaque, "desnegritar_conjuncao", False):
            busca_conjuncao = trecho.Find
            busca_conjuncao.Text = " e "
            if busca_conjuncao.Execute():
                busca_conjuncao.Parent.Font.Bold = False


def inserir_fundamentacao(
    documento: Any,
    resultado: Any,
    *,
    marcador: str = "[CONCLUSÃO]",
    wd_find_stop: int = 0,
) -> bool:
    """Insere a fundamentação e aplica seus destaques no documento Word."""
    texto = str(getattr(resultado, "texto", "") or "")
    intervalo = substituir_marcador_por_texto(documento, marcador, texto)
    if intervalo is None:
        return False
    aplicar_destaques_fundamentacao(
        intervalo,
        getattr(resultado, "destaques", []),
        wd_find_stop=wd_find_stop,
    )
    return True


def aplicar_negritos_dispositivo(
    intervalo: Any,
    frases,
    *,
    wd_find_stop: int = 0,
    wd_collapse_end: int = 0,
) -> None:
    """Coloca em negrito todas as ocorrências indicadas no dispositivo."""
    if intervalo is None:
        return
    for frase in dict.fromkeys(frases or []):
        frase = str(frase or "")
        if not frase:
            continue
        intervalo_busca = intervalo.Duplicate
        busca = intervalo_busca.Find
        busca.ClearFormatting()
        busca.Text = frase
        busca.Forward = True
        busca.Wrap = wd_find_stop
        while busca.Execute():
            busca.Parent.Font.Bold = True
            intervalo_busca.Collapse(wd_collapse_end)


def inserir_dispositivo(
    documento: Any,
    paragrafos,
    frases_negrito,
    *,
    marcador: str = "[DISPOSITIVO]",
    wd_find_stop: int = 0,
    wd_collapse_end: int = 0,
) -> str:
    """Numera, insere e formata os parágrafos do dispositivo."""
    paragrafos_numerados = [
        f"{indice}º) {paragrafo}"
        for indice, paragrafo in enumerate(paragrafos or [], start=1)
    ]
    texto = "\r".join(paragrafos_numerados)
    intervalo = substituir_marcador_por_texto(documento, marcador, texto)
    if intervalo is not None and paragrafos_numerados:
        aplicar_negritos_dispositivo(
            intervalo,
            frases_negrito,
            wd_find_stop=wd_find_stop,
            wd_collapse_end=wd_collapse_end,
        )
    return texto


def substituir_expressao_expositiva(
    documento: Any,
    *,
    wd_replace_all: int = 2,
) -> None:
    """Troca a abertura padrão quando o fluxo não é o de Da Camino."""
    busca = documento.Content.Find
    busca.ClearFormatting()
    busca.Execute(
        FindText="Isto posto, opina este",
        ReplaceWith="Diante do exposto, opina este",
        Replace=wd_replace_all,
    )
