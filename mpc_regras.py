"""Regras de negócio independentes da interface gráfica do MPC Parecer."""

from __future__ import annotations

from copy import deepcopy
import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP


RESPONSAVEL_NAO_INTIMADO = "Responsável Não Intimado"
NAO_APRESENTOU_ESCLARECIMENTOS = "Não Apresentou Esclarecimentos"
ESCLARECIMENTOS_ESPONTANEOS_DESCONSIDERADOS = (
    "Esclarecimentos Espontâneos Desconsiderados"
)
NAO_APRESENTOU_DEFESA_LEGADO = "Não Apresentou Defesa"
SEPARADOR_RESPONSAVEIS_APONTAMENTO = " | "
NAO_ALTERAR = "Não alterar"
CONCLUSOES_COM_RESPONSABILIDADE = {"Mantido", "Mantido Parcialmente"}
CONCLUSOES_RESPONSAVEL_POR_TIPO = {
    "CONTAS ANUAIS": {
        "Parecer Favorável",
        "Parecer Favorável, com Ressalvas",
        "Parecer Desfavorável",
    },
    "CONTAS ORDINÁRIAS": {
        "Contas Regulares",
        "Contas Regulares, com Ressalvas",
        "Contas Irregulares",
    },
    "PROCESSO DE CONTAS ESPECIAIS": {
        "Contas Regulares",
        "Contas Regulares, com Ressalvas",
        "Contas Irregulares",
    },
    "TOMADA DE CONTAS ESPECIAL": {
        "Contas Regulares",
        "Contas Regulares, com Ressalvas",
        "Contas Irregulares",
    },
}


def normalizar_resposta_lista_apontamentos(resposta, limite=None):
    """Valida e normaliza a lista numerada devolvida pela IA."""
    itens = []
    numeracoes_vistas = set()
    descartadas = []
    duplicadas = []
    for linha_original in str(resposta or "").splitlines():
        linha = re.sub(r"\s+", " ", linha_original).strip()
        if not linha or linha.startswith("```"):
            continue
        if "relatório sem falhas" in linha.lower():
            continue
        linha = re.sub(r"^[-*•]+\s*", "", linha)
        linha = re.sub(r"^\d+[.)]\s+(?=\d+(?:\.\d+)+\s)", "", linha)
        correspondencia = re.match(
            r"^(\d+(?:\.\d+)+)\.?\s*[-–—:]?\s+(.+)$", linha
        )
        if not correspondencia:
            descartadas.append(linha_original.strip())
            continue
        numero = correspondencia.group(1)
        titulo = re.sub(r"\s+", " ", correspondencia.group(2)).strip()
        titulo = titulo.rstrip(". ")
        if not titulo:
            descartadas.append(linha_original.strip())
            continue
        if numero in numeracoes_vistas:
            duplicadas.append(numero)
            continue
        numeracoes_vistas.add(numero)
        itens.append(f"{numero} {titulo}")
    excedentes = []
    if limite is not None and len(itens) > limite:
        excedentes = itens[limite:]
        itens = itens[:limite]
    return {
        "itens": itens,
        "descartadas": descartadas,
        "duplicadas": duplicadas,
        "excedentes": excedentes,
    }


def extrair_numeracoes_apontamentos(valor):
    """Extrai numerações como 1.2.3, sem duplicar e mantendo a ordem."""
    if isinstance(valor, (list, tuple, set)):
        texto = " ".join(str(item) for item in valor)
    else:
        texto = str(valor or "")
    resultado = []
    vistos = set()
    for numero in re.findall(r"\d+(?:\.\d+)*", texto):
        if numero not in vistos:
            vistos.add(numero)
            resultado.append(numero)
    return resultado


def formatar_numeracoes_apontamentos(numeracoes):
    """Formata uma lista jurídica: ``1.1, 1.2 e 1.3``."""
    itens = [str(item).strip() for item in numeracoes if str(item).strip()]
    if not itens:
        return ""
    if len(itens) == 1:
        return itens[0]
    return ", ".join(itens[:-1]) + " e " + itens[-1]


def separar_falhas_de_recomendacoes(apontamentos, recomendacoes):
    """Separa falhas sem responsabilidade das recomendações genuínas."""
    numeros_apontamentos = extrair_numeracoes_apontamentos(apontamentos)
    numeros_recomendacoes = extrair_numeracoes_apontamentos(recomendacoes)
    conjunto_recomendacoes = set(numeros_recomendacoes)
    falhas = [
        numero
        for numero in numeros_apontamentos
        if numero not in conjunto_recomendacoes
    ]
    return falhas, numeros_recomendacoes


def conclusao_padrao_sem_intimacao(tipo_processo):
    """Escolhe a conclusão favorável compatível com o tipo do processo."""
    tipo = str(tipo_processo or "").strip().upper()
    if tipo == "CONTAS ANUAIS":
        return "Parecer Favorável"
    if tipo == "CONTAS ORDINÁRIAS":
        return "Contas Regulares"
    return None


def validar_conclusoes_responsaveis(responsaveis, tipo_processo=""):
    """Impede que um responsável preenchido seja omitido do dispositivo."""
    tipo = str(tipo_processo or "").strip().upper()
    conclusoes_validas = CONCLUSOES_RESPONSAVEL_POR_TIPO.get(tipo)
    erros = []
    for linha, responsavel in enumerate(responsaveis, start=1):
        nome = str(responsavel.get("nome", "")).strip()
        if not nome:
            continue
        conclusao = str(responsavel.get("conclusao", "")).strip()
        if not conclusao:
            erros.append(
                f"Linha {linha} — '{nome}': selecione a Conclusão individual."
            )
            continue
        if conclusoes_validas is not None and conclusao not in conclusoes_validas:
            opcoes = ", ".join(sorted(conclusoes_validas))
            erros.append(
                f"Linha {linha} — '{nome}': a conclusão '{conclusao}' é "
                f"incompatível com {tipo.title()}. Opções: {opcoes}."
            )
    return erros


def arquivo_real_esclarecimentos(valor):
    """Distingue um caminho de PDF dos marcadores lógicos da interface."""
    texto = str(valor or "").strip()
    return bool(
        texto
        and texto
        not in {
            RESPONSAVEL_NAO_INTIMADO,
            NAO_APRESENTOU_ESCLARECIMENTOS,
            NAO_APRESENTOU_DEFESA_LEGADO,
        }
    )


def validar_coerencia_esclarecimentos(responsaveis, tipo_processo=""):
    """Valida intimação, resposta, arquivo e efeitos na responsabilização."""
    erros = []
    tipo = str(tipo_processo or "").strip().upper()
    conclusao_esperada = conclusao_padrao_sem_intimacao(tipo)
    estados_sem_resposta = {
        NAO_APRESENTOU_ESCLARECIMENTOS,
        NAO_APRESENTOU_DEFESA_LEGADO,
    }
    for responsavel in responsaveis:
        nome = str(responsavel.get("nome", "")).strip()
        if not nome:
            continue
        intimacao = str(responsavel.get("intimacao", "")).strip()
        esclarecimentos = str(
            responsavel.get("esclarecimentos", "")
        ).strip()
        arquivo = str(
            responsavel.get("arquivo_esclarecimentos", "")
        ).strip()
        tem_pdf = arquivo_real_esclarecimentos(arquivo)

        if intimacao == "Não":
            if esclarecimentos not in {
                RESPONSAVEL_NAO_INTIMADO,
                ESCLARECIMENTOS_ESPONTANEOS_DESCONSIDERADOS,
            }:
                erros.append(
                    f"'{nome}' não foi intimado; selecione "
                    f"'{RESPONSAVEL_NAO_INTIMADO}' ou "
                    f"'{ESCLARECIMENTOS_ESPONTANEOS_DESCONSIDERADOS}'."
                )
            if (
                esclarecimentos == ESCLARECIMENTOS_ESPONTANEOS_DESCONSIDERADOS
                and not tem_pdf
            ):
                erros.append(
                    f"'{nome}' apresentou esclarecimentos espontâneos, mas o "
                    "respectivo PDF não foi associado."
                )
            if esclarecimentos == RESPONSAVEL_NAO_INTIMADO and tem_pdf:
                erros.append(
                    f"'{nome}' está como não intimado, mas possui PDF; use a "
                    "situação de esclarecimentos espontâneos desconsiderados."
                )
            for rotulo, campo in (
                ("Falhas", "falhas"),
                ("Multa", "multa"),
                ("Débito", "debito"),
            ):
                if str(responsavel.get(campo, "")).strip() == "Sim":
                    erros.append(
                        f"'{nome}' não foi intimado e deve permanecer com "
                        f"{rotulo} = Não."
                    )
            if (
                conclusao_esperada
                and str(responsavel.get("conclusao", "")).strip()
                != conclusao_esperada
            ):
                erros.append(
                    f"'{nome}' não foi intimado; para {tipo.title()}, a "
                    f"conclusão deve ser '{conclusao_esperada}'."
                )
        elif intimacao == "Sim":
            if esclarecimentos == RESPONSAVEL_NAO_INTIMADO:
                erros.append(
                    f"'{nome}' está intimado, mas os esclarecimentos indicam "
                    "'Responsável Não Intimado'."
                )
            if esclarecimentos == ESCLARECIMENTOS_ESPONTANEOS_DESCONSIDERADOS:
                erros.append(
                    f"'{nome}' está intimado; esclarecimentos espontâneos "
                    "desconsiderados somente se aplicam a quem não foi intimado."
                )
            if esclarecimentos in estados_sem_resposta and tem_pdf:
                erros.append(
                    f"'{nome}' está marcado como sem esclarecimentos, mas há "
                    "um PDF associado."
                )
            if (
                esclarecimentos
                and esclarecimentos not in estados_sem_resposta
                and esclarecimentos != RESPONSAVEL_NAO_INTIMADO
                and not tem_pdf
            ):
                erros.append(
                    f"'{nome}' está marcado como tendo apresentado "
                    "esclarecimentos, mas nenhum PDF foi associado."
                )
        else:
            erros.append(f"Informe se '{nome}' foi intimado (Sim ou Não).")
    return erros


def classificar_apontamento(conclusao, multa, debito):
    """Classifica uma falha sem confundir alerta convertido com recomendação."""
    conclusao = str(conclusao or "").strip()
    if conclusao == "Recomendação":
        return "recomendacao"
    if conclusao == "Convertido em Alerta":
        return "sem_responsabilidade"
    if "Sim" in {str(multa or "").strip(), str(debito or "").strip()}:
        return "com_responsabilidade"
    return "sem_responsabilidade"


def nomes_responsaveis_do_vinculo(valor):
    """Converte uma associação persistida em uma lista de nomes."""
    if isinstance(valor, (list, tuple, set)):
        return [str(nome).strip() for nome in valor if str(nome).strip()]
    return [
        nome.strip()
        for nome in str(valor or "").split(SEPARADOR_RESPONSAVEIS_APONTAMENTO)
        if nome.strip()
    ]


def formatar_vinculo_responsaveis(nomes):
    """Formata responsáveis para exibição em uma única célula."""
    return SEPARADOR_RESPONSAVEIS_APONTAMENTO.join(
        str(nome).strip() for nome in nomes if str(nome).strip()
    )


def combinar_vinculos_lote(atuais, novos, substituir=False):
    """Combina associações em lote, preservando ordem e sem duplicações."""
    if substituir:
        return list(dict.fromkeys(novos))
    return list(dict.fromkeys(list(atuais) + list(novos)))


def consolidar_classificacao_apontamentos(
    apontamentos,
    recomendacoes_existentes="",
):
    """Consolida as listas de controle sem depender dos widgets da GUI."""
    com_responsabilidade = []
    sem_responsabilidade = []
    recomendacoes_classificadas = []
    numeros_na_grade = set()
    pendentes = []

    for linha, apontamento in enumerate(apontamentos, start=1):
        texto = str(
            apontamento.get("irregularidade")
            or apontamento.get("item")
            or ""
        ).strip()
        if not texto:
            continue
        conclusao = str(apontamento.get("conclusao", "")).strip()
        correspondencia = re.match(r"^(\d+(?:\.\d+)*)", texto)
        if conclusao in {"", "Análise Pendente"}:
            pendentes.append(
                correspondencia.group(1)
                if correspondencia
                else f"linha {linha}"
            )
        if not correspondencia:
            continue
        numero = correspondencia.group(1)
        numeros_na_grade.add(numero)
        classificacao = classificar_apontamento(
            conclusao,
            apontamento.get("multa"),
            apontamento.get("debito"),
        )
        if classificacao == "recomendacao":
            recomendacoes_classificadas.append(numero)
        elif classificacao == "com_responsabilidade":
            com_responsabilidade.append(numero)
        else:
            sem_responsabilidade.append(numero)

    recomendacoes_nativas = [
        numero
        for numero in extrair_numeracoes_apontamentos(
            recomendacoes_existentes
        )
        if numero not in numeros_na_grade
    ]

    def ordenar(itens):
        unicos = list(dict.fromkeys(itens))
        return sorted(
            unicos,
            key=lambda item: [int(parte) for parte in item.split(".")],
        )

    com_responsabilidade = ordenar(com_responsabilidade)
    sem_responsabilidade = ordenar(sem_responsabilidade)
    recomendacoes = ordenar(
        recomendacoes_nativas + recomendacoes_classificadas
    )
    return {
        "com_responsabilidade": com_responsabilidade,
        "sem_responsabilidade": sem_responsabilidade,
        "recomendacoes": recomendacoes,
        "pendentes": pendentes,
        "total_falhas": (
            len(com_responsabilidade) + len(sem_responsabilidade)
        ),
        "tem_apontamentos": bool(numeros_na_grade),
    }


def validar_compatibilidade_preenchimento_lote(
    apontamentos,
    indices,
    conclusao=NAO_ALTERAR,
    multa=NAO_ALTERAR,
    repercussao=NAO_ALTERAR,
):
    """Valida combinações do lote antes de qualquer alteração visual."""
    erros = []
    indices_validos = [
        indice for indice in indices if 0 <= indice < len(apontamentos)
    ]
    if multa == "Sim" or repercussao == "Sim":
        invalidos = []
        for indice in indices_validos:
            atual = apontamentos[indice]
            conclusao_resultante = (
                conclusao
                if conclusao != NAO_ALTERAR
                else str(atual.get("conclusao", "")).strip()
            )
            if conclusao_resultante not in CONCLUSOES_COM_RESPONSABILIDADE:
                invalidos.append(str(indice + 1))
        if invalidos:
            erros.append(
                "Para aplicar Multa ou Repercussão = Sim, a conclusão deve "
                "ser Mantido ou Mantido Parcialmente. Corrija as falhas n.º "
                f"{', '.join(invalidos)} ou escolha também uma conclusão no lote."
            )

    if (
        conclusao != NAO_ALTERAR
        and conclusao not in CONCLUSOES_COM_RESPONSABILIDADE
    ):
        invalidos = []
        for indice in indices_validos:
            atual = apontamentos[indice]
            multa_resultante = (
                multa if multa != NAO_ALTERAR else atual.get("multa")
            )
            repercussao_resultante = (
                repercussao
                if repercussao != NAO_ALTERAR
                else atual.get("repercussao")
            )
            if "Sim" in {
                str(multa_resultante or "").strip(),
                str(repercussao_resultante or "").strip(),
                str(atual.get("debito", "")).strip(),
            }:
                invalidos.append(str(indice + 1))
        if invalidos:
            erros.append(
                f"A conclusão '{conclusao}' não admite Multa, Repercussão "
                "ou Débito = Sim. Defina Multa e Repercussão como Não "
                "neste lote e, se houver Débito, ajuste-o individualmente. "
                f"Verifique as falhas n.º {', '.join(invalidos)}."
            )
    return erros


def aplicar_preenchimento_lote(
    apontamentos,
    indices,
    nomes,
    *,
    conclusao=NAO_ALTERAR,
    multa=NAO_ALTERAR,
    repercussao=NAO_ALTERAR,
    substituir=False,
):
    """Aplica as regras do lote em cópias dos apontamentos informados."""
    resultado = deepcopy(list(apontamentos))
    novos_nomes = list(dict.fromkeys(
        str(nome).strip() for nome in nomes if str(nome).strip()
    ))

    for indice in indices:
        if not 0 <= indice < len(resultado):
            continue
        apontamento = resultado[indice]
        responsaveis_falha = nomes_responsaveis_do_vinculo(
            apontamento.get("responsaveis", [])
        )
        if conclusao != NAO_ALTERAR:
            apontamento["conclusao"] = conclusao
            if conclusao in CONCLUSOES_COM_RESPONSABILIDADE:
                responsaveis_falha = combinar_vinculos_lote(
                    responsaveis_falha,
                    novos_nomes,
                    substituir,
                )
                if substituir:
                    obrigatorios = []
                    if multa == NAO_ALTERAR and apontamento.get("multa") == "Sim":
                        obrigatorios.extend(
                            nomes_responsaveis_do_vinculo(
                                apontamento.get("responsaveis_multa", [])
                            )
                        )
                    if (
                        repercussao == NAO_ALTERAR
                        and apontamento.get("repercussao") == "Sim"
                    ):
                        obrigatorios.extend(
                            nomes_responsaveis_do_vinculo(
                                apontamento.get(
                                    "responsaveis_repercussao", []
                                )
                            )
                        )
                    if apontamento.get("debito") == "Sim":
                        obrigatorios.extend(
                            nomes_responsaveis_do_vinculo(
                                apontamento.get("responsaveis_debito", [])
                            )
                        )
                    responsaveis_falha = combinar_vinculos_lote(
                        responsaveis_falha,
                        obrigatorios,
                    )

        if multa == "Sim":
            apontamento["multa"] = "Sim"
            atuais = nomes_responsaveis_do_vinculo(
                apontamento.get("responsaveis_multa", [])
            )
            apontamento["responsaveis_multa"] = combinar_vinculos_lote(
                atuais,
                novos_nomes,
                substituir,
            )
            responsaveis_falha = combinar_vinculos_lote(
                responsaveis_falha,
                novos_nomes,
            )
        elif multa == "Não":
            apontamento["multa"] = "Não"
            apontamento["responsaveis_multa"] = []

        if repercussao == "Sim":
            apontamento["repercussao"] = "Sim"
            atuais = nomes_responsaveis_do_vinculo(
                apontamento.get("responsaveis_repercussao", [])
            )
            apontamento["responsaveis_repercussao"] = (
                combinar_vinculos_lote(
                    atuais,
                    novos_nomes,
                    substituir,
                )
            )
            responsaveis_falha = combinar_vinculos_lote(
                responsaveis_falha,
                novos_nomes,
            )
        elif repercussao == "Não":
            apontamento["repercussao"] = "Não"
            apontamento["responsaveis_repercussao"] = []

        apontamento["responsaveis"] = responsaveis_falha
    return resultado


def converter_valor_monetario_brl(valor):
    """Converte ``1.000,00`` ou ``R$ 1.000,00`` em Decimal seguro."""
    texto = str(valor or "").strip()
    if not texto:
        return None
    texto = texto.replace("R$", "").replace(" ", "")
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif texto.count(".") > 1:
        texto = texto.replace(".", "")
    elif texto.count(".") == 1 and len(texto.rsplit(".", 1)[1]) == 3:
        texto = texto.replace(".", "")
    try:
        return Decimal(texto).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
    except (InvalidOperation, ValueError) as erro:
        raise ValueError(
            "Informe o valor no formato brasileiro, por exemplo: 1.000,00."
        ) from erro


def formatar_valor_monetario_brl(valor):
    """Exibe Decimal como ``R$ 1.000,00`` sem usar ponto flutuante."""
    valor_decimal = converter_valor_monetario_brl(valor)
    if valor_decimal is None:
        return ""
    texto = f"{valor_decimal:,.2f}"
    return "R$ " + texto.replace(",", "X").replace(".", ",").replace("X", ".")


def obter_responsaveis_apontamento(apontamento, natureza):
    """Lê uma associação específica e aceita o formato antigo unificado."""
    chave = {
        "falha": "responsaveis",
        "multa": "responsaveis_multa",
        "repercussao": "responsaveis_repercussao",
        "debito": "responsaveis_debito",
    }[natureza]
    if chave in apontamento:
        return nomes_responsaveis_do_vinculo(apontamento.get(chave, []))
    antigos = nomes_responsaveis_do_vinculo(apontamento.get("responsaveis", []))
    if natureza == "falha":
        return antigos
    if str(apontamento.get(natureza, "")).strip() == "Sim":
        return antigos
    return []


def resumir_associacoes_apontamento(apontamento):
    """Produz o resumo das quatro associações para a tabela da GUI."""
    partes = []
    for rotulo, natureza in (
        ("Falha", "falha"),
        ("Multa", "multa"),
        ("Repercussão", "repercussao"),
        ("Débito", "debito"),
    ):
        nomes = obter_responsaveis_apontamento(apontamento, natureza)
        if nomes:
            partes.append(f"{rotulo}: {', '.join(nomes)}")
    return " | ".join(partes)


def agrupar_itens_por_responsaveis(apontamentos, natureza):
    """Agrupa itens que possuem a mesma lista de responsáveis."""
    grupos = []
    indice_por_chave = {}
    for posicao, apontamento in enumerate(apontamentos, start=1):
        if str(apontamento.get(natureza, "")).strip() != "Sim":
            continue
        nomes = obter_responsaveis_apontamento(apontamento, natureza)
        if not nomes:
            continue
        descricao = str(
            apontamento.get("irregularidade")
            or apontamento.get("item")
            or f"linha {posicao}"
        ).strip()
        numeros = extrair_numeracoes_apontamentos(descricao)
        numero = numeros[0] if numeros else f"linha {posicao}"
        chave = tuple(nomes)
        if chave not in indice_por_chave:
            indice_por_chave[chave] = len(grupos)
            grupos.append({"responsaveis": list(nomes), "itens": []})
        itens = grupos[indice_por_chave[chave]]["itens"]
        if numero not in itens:
            itens.append(numero)
    return grupos


def agrupar_debitos_por_responsaveis(apontamentos):
    """Agrupa débitos pelo conjunto exato de devedores e soma os valores."""
    grupos = []
    indice_por_chave = {}
    for posicao, apontamento in enumerate(apontamentos, start=1):
        if str(apontamento.get("debito", "")).strip() != "Sim":
            continue
        nomes = obter_responsaveis_apontamento(apontamento, "debito")
        if not nomes:
            continue
        valor = converter_valor_monetario_brl(apontamento.get("valor_debito", ""))
        if valor is None or valor <= 0:
            raise ValueError(
                f"Item {posicao}: informe um valor de débito maior que zero."
            )
        descricao = str(
            apontamento.get("irregularidade")
            or apontamento.get("item")
            or f"linha {posicao}"
        ).strip()
        numeros = extrair_numeracoes_apontamentos(descricao)
        numero = numeros[0] if numeros else f"linha {posicao}"
        chave = tuple(nomes)
        if chave not in indice_por_chave:
            indice_por_chave[chave] = len(grupos)
            grupos.append(
                {
                    "responsaveis": list(nomes),
                    "itens": [],
                    "valor_total": Decimal("0.00"),
                }
            )
        grupo = grupos[indice_por_chave[chave]]
        if numero not in grupo["itens"]:
            grupo["itens"].append(numero)
            grupo["valor_total"] += valor
    return grupos


def validar_coerencia_conclusao_repercussao(
    responsaveis, apontamentos, tipo_processo
):
    """Certifica a relação obrigatória entre repercussão e conclusão final."""
    tipo = str(tipo_processo or "").strip().upper()
    conclusao_exigida = {
        "CONTAS ANUAIS": "Parecer Desfavorável",
        "CONTAS ORDINÁRIAS": "Contas Irregulares",
    }.get(tipo)
    if not conclusao_exigida:
        return []
    itens_repercussao_por_nome = {}
    for posicao, apontamento in enumerate(apontamentos, start=1):
        if str(apontamento.get("repercussao", "")).strip() != "Sim":
            continue
        descricao = str(
            apontamento.get("irregularidade")
            or apontamento.get("item")
            or f"linha {posicao}"
        ).strip()
        numero = (
            extrair_numeracoes_apontamentos(descricao)[:1]
            or [f"linha {posicao}"]
        )[0]
        for nome in obter_responsaveis_apontamento(apontamento, "repercussao"):
            itens_repercussao_por_nome.setdefault(nome, []).append(numero)
    erros = []
    for responsavel in responsaveis:
        nome = str(responsavel.get("nome", "")).strip()
        if not nome:
            continue
        conclusao = str(responsavel.get("conclusao", "")).strip()
        itens = itens_repercussao_por_nome.get(nome, [])
        if conclusao == conclusao_exigida and not itens:
            erros.append(
                f"'{nome}' está com '{conclusao_exigida}', mas não foi "
                "associado a nenhuma falha com Repercussão = Sim."
            )
        elif itens and conclusao != conclusao_exigida:
            itens_formatados = formatar_numeracoes_apontamentos(itens)
            erros.append(
                f"'{nome}' foi associado à Repercussão dos item(ns) "
                f"{itens_formatados}, mas sua conclusão deve ser "
                f"'{conclusao_exigida}'."
            )
    return erros


def validar_vinculos_responsabilidade(
    responsaveis, apontamentos, tipo_processo=""
):
    """Certifica vínculos de falha, multa, repercussão e débito."""
    erros = validar_coerencia_esclarecimentos(responsaveis, tipo_processo)
    mapa = {}
    duplicados = set()
    for responsavel in responsaveis:
        nome = str(responsavel.get("nome", "")).strip()
        if not nome:
            continue
        if nome in mapa:
            duplicados.add(nome)
        mapa[nome] = responsavel
    for nome in sorted(duplicados):
        erros.append(
            f"Há mais de um administrador chamado '{nome}'; "
            "a associação das falhas ficaria ambígua."
        )

    itens_multa_por_nome = {nome: [] for nome in mapa}
    itens_debito_por_nome = {nome: [] for nome in mapa}
    for posicao, apontamento in enumerate(apontamentos, start=1):
        descricao = str(
            apontamento.get("irregularidade")
            or apontamento.get("item")
            or f"linha {posicao}"
        ).strip()
        numero = (
            extrair_numeracoes_apontamentos(descricao)[:1]
            or [f"linha {posicao}"]
        )[0]
        nomes_falha = obter_responsaveis_apontamento(apontamento, "falha")
        nomes_multa = obter_responsaveis_apontamento(apontamento, "multa")
        nomes_repercussao = obter_responsaveis_apontamento(
            apontamento, "repercussao"
        )
        nomes_debito = obter_responsaveis_apontamento(apontamento, "debito")
        multa_sim = str(apontamento.get("multa", "")).strip() == "Sim"
        repercussao_sim = str(apontamento.get("repercussao", "")).strip() == "Sim"
        debito_sim = str(apontamento.get("debito", "")).strip() == "Sim"
        conclusao = str(apontamento.get("conclusao", "")).strip()
        falha_mantida = conclusao in {"Mantido", "Mantido Parcialmente"}

        for natureza, nomes in {
            "falha": nomes_falha,
            "multa": nomes_multa,
            "repercussão": nomes_repercussao,
            "débito": nomes_debito,
        }.items():
            inexistentes = [nome for nome in nomes if nome not in mapa]
            if inexistentes:
                erros.append(
                    f"Item {numero}: administrador(es) de {natureza} não "
                    f"localizado(s): {', '.join(inexistentes)}."
                )

        if falha_mantida and not nomes_falha:
            erros.append(
                f"Item {numero}: a falha foi '{conclusao}', mas nenhum "
                "administrador foi associado à falha."
            )
        if (multa_sim or repercussao_sim or debito_sim) and not falha_mantida:
            erros.append(
                f"Item {numero}: Multa, Repercussão ou Débito somente podem "
                "ser associados quando a conclusão for 'Mantido' ou "
                "'Mantido Parcialmente'."
            )
        for natureza, ativo, nomes in (
            ("multa", multa_sim, nomes_multa),
            ("repercussão", repercussao_sim, nomes_repercussao),
            ("débito", debito_sim, nomes_debito),
        ):
            if ativo and not nomes:
                erros.append(
                    f"Item {numero}: {natureza} marcada como Sim, mas nenhum "
                    f"administrador foi associado à {natureza}."
                )

        if debito_sim:
            try:
                valor_debito = converter_valor_monetario_brl(
                    apontamento.get("valor_debito", "")
                )
                if valor_debito is None or valor_debito <= 0:
                    erros.append(
                        f"Item {numero}: Débito = Sim exige um Valor maior que zero."
                    )
            except ValueError as erro:
                erros.append(f"Item {numero}: {erro}")

        if conclusao in {"Recomendação", "Convertido em Alerta"} and (
            multa_sim or debito_sim
        ):
            erros.append(
                f"Item {numero}: a conclusão '{conclusao}' é incompatível "
                "com multa ou débito."
            )

        conjunto_falha = set(nomes_falha)
        for natureza, nomes in (
            ("multa", nomes_multa),
            ("repercussão", nomes_repercussao),
            ("débito", nomes_debito),
        ):
            fora_da_falha = [nome for nome in nomes if nome not in conjunto_falha]
            if fora_da_falha:
                erros.append(
                    f"Item {numero}: administrador(es) de {natureza} deve(m) "
                    "também estar associado(s) à falha: "
                    f"{', '.join(fora_da_falha)}."
                )

        todos_os_nomes = list(
            dict.fromkeys(
                nomes_falha + nomes_multa + nomes_repercussao + nomes_debito
            )
        )
        for nome in todos_os_nomes:
            responsavel = mapa.get(nome)
            if responsavel is None:
                continue
            if str(responsavel.get("falhas", "")).strip() == "Não":
                erros.append(
                    f"Item {numero}: '{nome}' foi associado à falha, mas "
                    "está marcado como sem falhas na tabela de administradores."
                )
            if multa_sim and nome in nomes_multa:
                itens_multa_por_nome[nome].append(numero)
                if str(responsavel.get("multa", "")).strip() != "Sim":
                    erros.append(
                        f"Item {numero}: há multa, mas '{nome}' não está "
                        "marcado com Multa = Sim na tabela de administradores."
                    )
            if debito_sim and nome in nomes_debito:
                itens_debito_por_nome[nome].append(numero)
                if str(responsavel.get("debito", "")).strip() != "Sim":
                    erros.append(
                        f"Item {numero}: há débito, mas '{nome}' não está "
                        "marcado com Débito = Sim na tabela de administradores."
                    )

    for nome, responsavel in mapa.items():
        if (
            str(responsavel.get("multa", "")).strip() == "Sim"
            and not itens_multa_por_nome[nome]
        ):
            erros.append(
                f"'{nome}' está com Multa = Sim, mas não foi associado a "
                "nenhum apontamento com multa."
            )
        if (
            str(responsavel.get("debito", "")).strip() == "Sim"
            and not itens_debito_por_nome[nome]
        ):
            erros.append(
                f"'{nome}' está com Débito = Sim, mas não foi associado a "
                "nenhum apontamento com débito."
            )
    erros.extend(
        validar_coerencia_conclusao_repercussao(
            responsaveis, apontamentos, tipo_processo
        )
    )
    return erros
