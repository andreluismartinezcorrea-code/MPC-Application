"""Construção testável da fundamentação e do dispositivo do MPC Parecer.

Este módulo não conhece Tkinter nem controla o Microsoft Word. Ele recebe
dados simples, aplica as regras de redação e devolve os textos e os fragmentos
que devem receber negrito. Essa separação permite testar a linguagem jurídica
sem abrir a interface ou um documento.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from mpc_regras import (
    agrupar_debitos_por_responsaveis,
    agrupar_itens_por_responsaveis,
    formatar_numeracoes_apontamentos,
    formatar_valor_monetario_brl,
    obter_responsaveis_apontamento,
)


@dataclass
class ResultadoConclusao:
    """Textos construídos e fragmentos que devem ser destacados no Word."""

    paragrafos: list[str] = field(default_factory=list)
    negritos: list[str] = field(default_factory=list)

    def incorporar(self, outro: "ResultadoConclusao") -> None:
        """Acrescenta outro resultado, preservando a ordem dos parágrafos."""
        self.paragrafos.extend(outro.paragrafos)
        self.negritos.extend(outro.negritos)


@dataclass(frozen=True)
class DestaqueFundamentacao:
    """Trecho a colocar em negrito no parágrafo pré-dispositivo."""

    texto: str
    desnegritar_conjuncao: bool = False


@dataclass
class ResultadoFundamentacao:
    """Texto pré-dispositivo e instruções de destaque para o Word."""

    texto: str = ""
    destaques: list[DestaqueFundamentacao] = field(default_factory=list)


def _texto_template(
    templates: dict | None,
    caminho: tuple[str, ...],
    padrao: str,
) -> str:
    """Obtém um texto editável do banco, preservando compatibilidade.

    O texto padrão mantém funcionando os JSONs salvos antes da inclusão de um
    modelo novo. Assim, atualizar o programa não obriga o usuário a substituir
    imediatamente um banco de parágrafos já personalizado.
    """
    atual = templates if isinstance(templates, dict) else {}
    for chave in caminho:
        if not isinstance(atual, dict):
            return padrao
        atual = atual.get(chave)
    if isinstance(atual, str) and atual.strip():
        return atual
    return padrao


def _nome_com_cargo(responsavel: dict) -> str:
    return f"{responsavel.get('nome', '').strip()} ({responsavel.get('cargo', '').strip()})"


def _separar_por_genero(responsaveis: list[dict]) -> tuple[list[str], list[str], list[str]]:
    masculinos = []
    femininos = []
    empresas = []
    for responsavel in responsaveis:
        nome_cargo = _nome_com_cargo(responsavel)
        cargo = str(responsavel.get("cargo", "")).strip().upper()
        sexo = str(responsavel.get("sexo", "")).strip().upper()
        if cargo.startswith("CNPJ"):
            empresas.append(nome_cargo)
        elif sexo == "F":
            femininos.append(nome_cargo)
        else:
            masculinos.append(nome_cargo)
    return masculinos, femininos, empresas


def _formatar_grupo(nomes: list[str], singular: str, plural: str) -> str:
    if not nomes:
        return ""
    if len(nomes) == 1:
        return f"{singular} {nomes[0]}"
    return f"{plural} {', '.join(nomes[:-1])} e {nomes[-1]}"


def _unir_grupos(grupos: list[str]) -> str:
    grupos = [grupo for grupo in grupos if grupo]
    if not grupos:
        return ""
    if len(grupos) == 1:
        return grupos[0]
    if len(grupos) == 2:
        return f"{grupos[0]} e {grupos[1]}"
    return "; ".join(grupos[:-1]) + f" e {grupos[-1]}"


def _formatar_enumeracao_juridica(itens: list[str]) -> str:
    """Une consequências com vírgulas e apenas uma conjunção final."""
    itens = [str(item).strip() for item in itens if str(item).strip()]
    if not itens:
        return ""
    if len(itens) == 1:
        return itens[0]
    return ", ".join(itens[:-1]) + f" e {itens[-1]}"


def formatar_lista_responsaveis_contas(
    responsaveis: list[dict],
) -> tuple[str, str, list[str]]:
    """Formata responsáveis como ``do Sr.``, ``da Sra.`` ou ``da empresa``."""
    masculinos, femininos, empresas = _separar_por_genero(responsaveis)
    texto = _unir_grupos(
        [
            _formatar_grupo(masculinos, "do Sr.", "dos Srs."),
            _formatar_grupo(femininos, "da Sra.", "das Sras."),
            _formatar_grupo(empresas, "da empresa", "das empresas"),
        ]
    )
    if len(responsaveis) == 1:
        titulo = "Administradora" if responsaveis[0].get("sexo") == "F" else "Administrador"
    else:
        tem_homem = any(responsavel.get("sexo") == "M" for responsavel in responsaveis)
        titulo = "Administradores" if tem_homem else "Administradoras"
    return texto, titulo, [_nome_com_cargo(item) for item in responsaveis]


def formatar_destinatarios_multa(responsaveis: list[dict]) -> tuple[str, list[str]]:
    """Formata a lista que sucede a palavra ``Multa``."""
    masculinos, femininos, empresas = _separar_por_genero(responsaveis)
    texto = _unir_grupos(
        [
            _formatar_grupo(masculinos, "ao Sr.", "aos Srs."),
            _formatar_grupo(femininos, "à Sra.", "às Sras."),
            _formatar_grupo(empresas, "à empresa", "às empresas"),
        ]
    )
    return texto, [_nome_com_cargo(item) for item in responsaveis]


def formatar_responsabilidade_de(responsaveis: list[dict]) -> tuple[str, list[str]]:
    """Formata a expressão posterior a ``de responsabilidade``."""
    texto, _, negritos = formatar_lista_responsaveis_contas(responsaveis)
    return texto, negritos


def construir_paragrafos_multa_e_debito(
    tipo_processo: str,
    responsaveis: list[dict],
    apontamentos: list[dict],
    templates: dict,
) -> ResultadoConclusao:
    """Constrói multa única e débitos agrupados pelo conjunto de devedores."""
    resultado = ResultadoConclusao()
    if str(tipo_processo or "").strip().upper() == "CONTAS ANUAIS":
        return resultado

    mapa_por_nome = {
        str(responsavel.get("nome", "")).strip(): responsavel
        for responsavel in responsaveis
        if str(responsavel.get("nome", "")).strip()
    }

    grupos_multa = agrupar_itens_por_responsaveis(apontamentos, "multa")
    nomes_multados = {
        nome
        for grupo in grupos_multa
        for nome in grupo["responsaveis"]
    }
    gestores_multa = [
        responsavel
        for responsavel in responsaveis
        if responsavel.get("nome") in nomes_multados
    ]
    texto_gestores, nomes_negrito = formatar_destinatarios_multa(gestores_multa)
    if texto_gestores:
        template_multa = templates["dispositivo"].get("paragrafo_multa", {}).get(
            "texto",
            "Multa {texto_gestores_multados} por infringência de normas de "
            "administração financeira e orçamentária, com base nos arts. 67 da "
            "Lei Estadual n.º 11.424/2000 e 135 do RITCE.",
        )
        resultado.paragrafos.append(
            template_multa.format(texto_gestores_multados=texto_gestores)
        )
        resultado.negritos.extend(["Multa", *nomes_negrito])

    for grupo in agrupar_debitos_por_responsaveis(apontamentos):
        gestores = [
            mapa_por_nome[nome]
            for nome in grupo["responsaveis"]
            if nome in mapa_por_nome
        ]
        texto_responsabilidade, nomes_negrito = formatar_responsabilidade_de(gestores)
        itens = formatar_numeracoes_apontamentos(grupo["itens"])
        if not texto_responsabilidade or not itens:
            continue
        valor_total = formatar_valor_monetario_brl(grupo["valor_total"])
        referencia = (
            f"ao item {itens}"
            if len(grupo["itens"]) == 1
            else f"aos itens {itens}"
        )
        template_debito = _texto_template(
            templates,
            ("dispositivo", "paragrafo_debito", "texto"),
            "Fixação de débito no valor de {valor_total}, correspondente "
            "{referencia} do Relatório de Auditoria, de responsabilidade "
            "{texto_responsabilidade}.",
        )
        resultado.paragrafos.append(
            template_debito.format(
                valor_total=valor_total,
                referencia=referencia,
                texto_responsabilidade=texto_responsabilidade,
            )
        )
        resultado.negritos.extend(
            ["Fixação de débito", valor_total, itens, *nomes_negrito]
        )
    return resultado


def construir_dispositivo_por_conclusoes(
    tipo_processo: str,
    responsaveis: list[dict],
    templates: dict,
    *,
    artigo_orgao: str,
    orgao_completo: str,
    ano_exercicio: str,
) -> ResultadoConclusao:
    """Agrupa os responsáveis por conclusão e redige o dispositivo principal."""
    resultado = ResultadoConclusao()
    tipo = str(tipo_processo or "").strip().upper()

    if tipo == "CONTAS ANUAIS":
        ordem = [
            "Parecer Favorável",
            "Parecer Favorável, com Ressalvas",
            "Parecer Desfavorável",
        ]
        fundamentos = {
            "Parecer Favorável": "I",
            "Parecer Favorável, com Ressalvas": "II",
            "Parecer Desfavorável": "III",
        }
        template = templates["dispositivo"]["contas_anuais"]["texto_base"]
        for conclusao in ordem:
            grupo = [item for item in responsaveis if item.get("conclusao") == conclusao]
            if not grupo:
                continue
            gestores, titulo, nomes_negrito = formatar_lista_responsaveis_contas(grupo)
            conclusao_capitalizada = conclusao.capitalize()
            ressalvas = (
                _texto_template(
                    templates,
                    ("dispositivo", "texto_ressalvas"),
                    ", tendo em vista os critérios estabelecidos pelos arts. 2º e 3º "
                    "da Resolução TCE n.º 1.142/2021",
                )
                if "Ressalvas" in conclusao or "Desfavorável" in conclusao
                else ""
            )
            resultado.paragrafos.append(
                template.format(
                    conclusao_capitalizada=conclusao_capitalizada,
                    gestores_formatados=gestores,
                    titulo_administrador=titulo,
                    artigo_orgao=artigo_orgao,
                    orgao_completo=orgao_completo,
                    ano_exercicio=ano_exercicio,
                    inciso_ritce=fundamentos[conclusao],
                    texto_ressalvas=ressalvas,
                )
            )
            resultado.negritos.extend([conclusao_capitalizada, *nomes_negrito])
        if any(item.get("conclusao") == "Parecer Desfavorável" for item in responsaveis):
            resultado.paragrafos.append(
                templates["dispositivo"]["contas_anuais"]["ciencia_mp"]
            )
        return resultado

    if tipo == "CONTAS ORDINÁRIAS":
        ordem = [
            "Contas Regulares",
            "Contas Regulares, com Ressalvas",
            "Contas Irregulares",
        ]
        fundamentos = {
            "Contas Regulares": "I",
            "Contas Regulares, com Ressalvas": "II",
            "Contas Irregulares": "III",
        }
        template = templates["dispositivo"]["contas_ordinarias"]["texto_base"]
        for conclusao in ordem:
            grupo = [item for item in responsaveis if item.get("conclusao") == conclusao]
            if not grupo:
                continue
            gestores, titulo, nomes_negrito = formatar_lista_responsaveis_contas(grupo)
            conclusao_capitalizada = conclusao.capitalize()
            ressalvas = (
                _texto_template(
                    templates,
                    ("dispositivo", "texto_ressalvas"),
                    ", tendo em vista os critérios estabelecidos pelos arts. 2º e 3º "
                    "da Resolução TCE n.º 1.142/2021",
                )
                if "Ressalvas" in conclusao or "Irregulares" in conclusao
                else ""
            )
            resultado.paragrafos.append(
                template.format(
                    conclusao_capitalizada=conclusao_capitalizada,
                    gestores_formatados=gestores,
                    titulo_administrador=titulo,
                    artigo_orgao=artigo_orgao,
                    orgao_completo=orgao_completo,
                    ano_exercicio=ano_exercicio,
                    inciso_ritce=fundamentos[conclusao],
                    texto_ressalvas=ressalvas,
                )
            )
            resultado.negritos.extend([conclusao_capitalizada, *nomes_negrito])
        if any(item.get("conclusao") == "Contas Irregulares" for item in responsaveis):
            resultado.paragrafos.append(
                templates["dispositivo"]["contas_ordinarias"]["ciencia_mp"]
            )
        return resultado

    if tipo in {"PROCESSO DE CONTAS ESPECIAIS", "TOMADA DE CONTAS ESPECIAL"}:
        ordem = [
            "Contas Regulares",
            "Contas Regulares, com Ressalvas",
            "Contas Irregulares",
        ]
        fundamentos = {
            "Contas Regulares": "I",
            "Contas Regulares, com Ressalvas": "II",
            "Contas Irregulares": "III",
        }
        template = templates["dispositivo"]["tomada_de_contas"]["texto_base"]
        for conclusao in ordem:
            grupo = [item for item in responsaveis if item.get("conclusao") == conclusao]
            if not grupo:
                continue
            gestores, _, nomes_negrito = formatar_lista_responsaveis_contas(grupo)
            conclusao_capitalizada = conclusao.capitalize()
            paragrafo = template.format(
                conclusao_capitalizada=conclusao_capitalizada,
                gestores_formatados=gestores,
                inciso_ritce=fundamentos[conclusao],
            )
            if tipo == "TOMADA DE CONTAS ESPECIAL":
                paragrafo = paragrafo.replace("86-D", "90")
            resultado.paragrafos.append(paragrafo)
            resultado.negritos.extend([conclusao_capitalizada, *nomes_negrito])
    return resultado


def construir_nucleo_dispositivo(
    tipo_processo: str,
    responsaveis: list[dict],
    apontamentos: list[dict],
    templates: dict,
    *,
    artigo_orgao: str,
    orgao_completo: str,
    ano_exercicio: str,
) -> ResultadoConclusao:
    """Constrói, na ordem atual, multa, débito e conclusões individuais."""
    resultado = construir_paragrafos_multa_e_debito(
        tipo_processo,
        responsaveis,
        apontamentos,
        templates,
    )
    resultado.incorporar(
        construir_dispositivo_por_conclusoes(
            tipo_processo,
            responsaveis,
            templates,
            artigo_orgao=artigo_orgao,
            orgao_completo=orgao_completo,
            ano_exercicio=ano_exercicio,
        )
    )
    return resultado


def _gestores_associados(
    responsaveis: list[dict],
    apontamentos: list[dict],
    natureza: str,
) -> list[dict]:
    """Retorna a união ordenada dos responsáveis vinculados à natureza."""
    mapa_por_nome = {
        str(responsavel.get("nome", "")).strip(): responsavel
        for responsavel in responsaveis
        if str(responsavel.get("nome", "")).strip()
    }
    nomes = []
    for apontamento in apontamentos:
        if str(apontamento.get(natureza, "")).strip() != "Sim":
            continue
        for nome in obter_responsaveis_apontamento(apontamento, natureza):
            if nome not in nomes:
                nomes.append(nome)
    return [mapa_por_nome[nome] for nome in nomes if nome in mapa_por_nome]


def _numeracoes_com_repercussao(apontamentos: list[dict]) -> list[str]:
    """Reproduz a leitura do primeiro token dos apontamentos com repercussão."""
    numeros = []
    for apontamento in apontamentos:
        if str(apontamento.get("repercussao", "")).strip() != "Sim":
            continue
        texto = str(
            apontamento.get("irregularidade")
            or apontamento.get("item")
            or ""
        ).strip()
        if texto:
            numeros.append(texto.split(maxsplit=1)[0])
    return numeros


def _formatar_gestores_contas_anuais(
    responsaveis: list[dict],
    *,
    incluir_cargo: bool = True,
) -> str:
    """Formata gestores como no parágrafo explicativo de Contas Anuais."""
    masculinos = []
    femininos = []
    empresas = []
    for responsavel in responsaveis:
        nome = str(responsavel.get("nome", "")).strip()
        cargo = str(responsavel.get("cargo", "")).strip()
        texto = f"{nome} ({cargo})" if incluir_cargo else nome
        if cargo.upper().startswith("CNPJ"):
            empresas.append(f"empresa {texto}")
        elif str(responsavel.get("sexo", "")).strip().upper() == "M":
            masculinos.append(texto)
        else:
            femininos.append(texto)

    grupos = []
    if masculinos:
        grupos.append(_formatar_grupo(masculinos, "do Sr.", "dos Srs."))
    if femininos:
        grupos.append(_formatar_grupo(femininos, "da Sra.", "das Sras."))
    if empresas:
        grupos.append(_formatar_grupo(empresas, "da", "das"))
    return " e ".join(grupos)


def _inteiro_nao_negativo(valor) -> int:
    texto = str(valor or "").strip()
    return int(texto) if texto.isdigit() else 0


def construir_fundamentacao_fernanda_contas_ordinarias(
    responsaveis: list[dict],
    apontamentos: list[dict],
    templates: dict | None = None,
    *,
    quantidade_com_responsabilidade=0,
    quantidade_sem_responsabilidade=0,
    falhas_sem_responsabilidade: str = "",
) -> ResultadoFundamentacao:
    """Constrói o parágrafo especial de Fernanda para Contas Ordinárias."""
    gestores_contas_irregulares = [
        responsavel
        for responsavel in responsaveis
        if responsavel.get("conclusao") == "Contas Irregulares"
    ]
    contas_irregulares = bool(gestores_contas_irregulares)
    gestores_multa = _gestores_associados(responsaveis, apontamentos, "multa")
    gestores_repercussao = _gestores_associados(
        responsaveis,
        apontamentos,
        "repercussao",
    )
    gestores_debito = _gestores_associados(responsaveis, apontamentos, "debito")

    gestores_validos = []
    for grupo in (gestores_repercussao, gestores_multa, gestores_debito):
        for gestor in grupo:
            if gestor not in gestores_validos:
                gestores_validos.append(gestor)
    if not gestores_validos:
        sexo_padrao = responsaveis[0].get("sexo", "M") if responsaveis else "M"
        gestores_validos = [{"sexo": sexo_padrao}]

    if len(gestores_validos) > 1:
        tem_homem = any(gestor.get("sexo") == "M" for gestor in gestores_validos)
        artigo_de = "dos" if tem_homem else "das"
        administrador = "Administradores"
        artigo_responsavel = "aos" if tem_homem else "às"
        responsavel = "Responsáveis"
    else:
        sexo = gestores_validos[0].get("sexo")
        artigo_de = "do" if sexo == "M" else "da"
        administrador = "Administrador" if sexo == "M" else "Administradora"
        artigo_responsavel = "ao" if sexo == "M" else "à"
        responsavel = "Responsável"

    quantidade_com = _inteiro_nao_negativo(quantidade_com_responsabilidade)
    quantidade_sem = _inteiro_nao_negativo(quantidade_sem_responsabilidade)
    falhas_sem = str(falhas_sem_responsabilidade or "").strip()
    itens_repercussao = _numeracoes_com_repercussao(apontamentos)
    itens_repercussao_texto = formatar_numeracoes_apontamentos(itens_repercussao)

    consequencias_sancionatorias = []
    if gestores_multa:
        consequencias_sancionatorias.append("a imposição de multa")
    if gestores_debito:
        consequencias_sancionatorias.append("a fixação de débito")
    consequencia = _formatar_enumeracao_juridica(
        consequencias_sancionatorias or ["a imposição de multa"]
    )

    if quantidade_com == 0 and quantidade_sem > 0:
        plural_s = "s" if quantidade_sem > 1 else ""
        plural_item = "ns" if quantidade_sem > 1 else "m"
        plural_verbo = "m" if quantidade_sem > 1 else ""
        template = _texto_template(
            templates,
            ("fundamentacao_especial", "fernanda_contas_ordinarias", "sem_responsabilidade", "texto"),
            "A{plural_s} irregularidade{plural_s} descrita{plural_s} no{plural_s} "
            "ite{plural_item} {falhas_sem} desvela{plural_verbo} a transgressão "
            "a dispositivos constitucionais e a normas de administração financeira "
            "e orçamentária. Entretanto, apesar de opinar pela manutenção desse{plural_s} "
            "aponte{plural_s} para fins da adoção de medidas corretivas, este Órgão "
            "Ministerial manifesta-se pela não responsabilização {artigo_de} "
            "{administrador} relativamente à imposição de multa nessa{plural_s} "
            "inconformidade{plural_s}, por não vislumbrar a presença de dolo ou erro "
            "grosseiro, nos termos do disposto no art. 28 da Lei de Introdução às "
            "Normas do Direito Brasileiro.",
        )
        texto = template.format(
            plural_s=plural_s,
            plural_item=plural_item,
            falhas_sem=falhas_sem,
            plural_verbo=plural_verbo,
            artigo_de=artigo_de,
            administrador=administrador,
        )
        return ResultadoFundamentacao(texto=texto)

    if quantidade_com > 0 and not contas_irregulares:
        plural_s = "s" if quantidade_com > 1 else ""
        plural_verbo = "m" if quantidade_com > 1 else ""
        plural_levar = "em" if quantidade_com > 1 else ""
        template = _texto_template(
            templates,
            ("fundamentacao_especial", "fernanda_contas_ordinarias", "sem_irregularidade", "texto"),
            "O contexto descrito nos autos revela a prática de atos contrários às "
            "normas de administração financeira e orçamentária, ensejando "
            "{consequencia} {artigo_responsavel} {responsavel}. Não obstante, "
            "entende-se que a{plural_s} irregularidade{plural_s} destacada{plural_s} "
            "pela Área Técnica não se reveste{plural_verbo} de relevância bastante "
            "para levar{plural_levar} à emissão de parecer pela irregularidade das "
            "contas {artigo_de} {administrador}.",
        )
        texto = template.format(
            consequencia=consequencia,
            artigo_responsavel=artigo_responsavel,
            responsavel=responsavel,
            plural_s=plural_s,
            plural_verbo=plural_verbo,
            plural_levar=plural_levar,
            artigo_de=artigo_de,
            administrador=administrador,
        )
        return ResultadoFundamentacao(texto=texto)

    if quantidade_com > 0 and contas_irregulares:
        if itens_repercussao:
            plural_s = "s" if len(itens_repercussao) > 1 else ""
            plural_item = "ns" if len(itens_repercussao) > 1 else "m"
            inicio = (
                f"O conjunto das falhas antes descritas, sobretudo o{plural_s} "
                f"ite{plural_item} {itens_repercussao_texto}, revela"
            )
        else:
            inicio = "O conjunto das falhas antes descritas revela"
        gestores_irregulares, _, _ = formatar_lista_responsaveis_contas(
            gestores_contas_irregulares
        )
        julgamento = (
            "o julgamento pela irregularidade das contas ordinárias "
            f"{gestores_irregulares}"
        )
        consequencias = _formatar_enumeracao_juridica(
            [*consequencias_sancionatorias, julgamento]
        )
        template = _texto_template(
            templates,
            ("fundamentacao_especial", "fernanda_contas_ordinarias", "irregularidade", "texto"),
            "{inicio} a prática de atos contrários às normas de administração "
            "financeira e orçamentária que justificam {consequencias}.",
        )
        texto = template.format(inicio=inicio, consequencias=consequencias)
        return ResultadoFundamentacao(
            texto=texto,
            destaques=[
                DestaqueFundamentacao(item) for item in itens_repercussao
            ],
        )

    return ResultadoFundamentacao()


PROCURADORES_PARAGRAFO_RESSALVAS = {
    "FERNANDA ISMAEL",
    "ÂNGELO GRÄBIN BORGHETTI",
    "DANIELA WENDT TONIAZZO",
}


def construir_fundamentacao_contas_anuais(
    procurador: str,
    responsaveis: list[dict],
    apontamentos: list[dict],
    templates: dict | None = None,
) -> ResultadoFundamentacao:
    """Constrói os parágrafos explicativos específicos de Contas Anuais."""
    procurador_normalizado = str(procurador or "").strip().upper()
    gestores_repercussao = _gestores_associados(
        responsaveis,
        apontamentos,
        "repercussao",
    )
    nomes_repercussao = {gestor.get("nome") for gestor in gestores_repercussao}
    gestores_desfavoraveis = [
        gestor
        for gestor in responsaveis
        if gestor.get("conclusao") == "Parecer Desfavorável"
        and gestor.get("nome") in nomes_repercussao
    ]
    gestores_ressalvas_conclusao = [
        gestor
        for gestor in responsaveis
        if gestor.get("conclusao") == "Parecer Favorável, com Ressalvas"
    ]
    gestores_ressalvas_repercussao = [
        gestor
        for gestor in gestores_ressalvas_conclusao
        if gestor.get("nome") in nomes_repercussao
    ]
    itens_repercussao = _numeracoes_com_repercussao(apontamentos)
    paragrafos = []
    destaques = []

    if (
        procurador_normalizado in PROCURADORES_PARAGRAFO_RESSALVAS
        and gestores_ressalvas_conclusao
    ):
        gestores = _formatar_gestores_contas_anuais(
            gestores_ressalvas_conclusao,
            incluir_cargo=False,
        )
        template = _texto_template(
            templates,
            ("fundamentacao_especial", "contas_anuais", "ressalvas", "texto"),
            "O contexto descrito nos autos, ainda que revele a ocorrência de "
            "infrações a dispositivos legais e constitucionais e a normas de "
            "administração financeira e orçamentária, não compromete gravemente "
            "as contas anuais {gestores}.",
        )
        paragrafos.append(template.format(gestores=gestores))

    if gestores_desfavoraveis:
        gestores = _formatar_gestores_contas_anuais(gestores_desfavoraveis)
        if len(itens_repercussao) == 1:
            falhas = f"o item {itens_repercussao[0]}"
        elif len(itens_repercussao) > 1:
            falhas = f"os itens {formatar_numeracoes_apontamentos(itens_repercussao)}"
        else:
            falhas = "as falhas apontadas"
        template = _texto_template(
            templates,
            ("fundamentacao_especial", "contas_anuais", "desfavoravel", "texto"),
            "O conjunto das falhas antes descritas, sobretudo {falhas}, revela a "
            "prática de atos contrários às normas de administração financeira e "
            "orçamentária e se reveste de relevância bastante para ensejar a emissão "
            "de parecer desfavorável à aprovação das contas {gestores}.",
        )
        paragrafos.append(template.format(falhas=falhas, gestores=gestores))
        destaques.extend(DestaqueFundamentacao(item) for item in itens_repercussao)
    elif (
        gestores_ressalvas_repercussao
        and procurador_normalizado not in PROCURADORES_PARAGRAFO_RESSALVAS
    ):
        gestores = _formatar_gestores_contas_anuais(
            gestores_ressalvas_repercussao
        )
        template = _texto_template(
            templates,
            ("fundamentacao_especial", "contas_anuais", "ressalvas", "texto"),
            "O contexto descrito nos autos, ainda que revele a ocorrência de "
            "infrações a dispositivos legais e constitucionais e a normas de "
            "administração financeira e orçamentária, não compromete gravemente "
            "as contas anuais {gestores}.",
        )
        paragrafos.append(template.format(gestores=gestores))

    return ResultadoFundamentacao(
        texto="\r".join(paragrafos),
        destaques=destaques,
    )


def construir_fundamentacao_padrao(
    tipo_processo: str,
    responsaveis: list[dict],
    apontamentos: list[dict],
    templates: dict,
) -> ResultadoFundamentacao:
    """Constrói o parágrafo geral baseado nas falhas com repercussão."""
    tipo = str(tipo_processo or "").strip().upper()
    gestores_repercussao = _gestores_associados(
        responsaveis,
        apontamentos,
        "repercussao",
    )
    nomes_repercussao = {gestor.get("nome") for gestor in gestores_repercussao}
    gestores = [
        gestor
        for gestor in responsaveis
        if gestor.get("nome") in nomes_repercussao
        and gestor.get("conclusao")
        in {"Parecer Desfavorável", "Contas Irregulares"}
    ]
    itens = _numeracoes_com_repercussao(apontamentos)
    if not gestores or not itens:
        return ResultadoFundamentacao()

    itens_texto = formatar_numeracoes_apontamentos(itens)
    if len(itens) == 1:
        texto_itens = f"o item {itens_texto} reveste-se"
    else:
        texto_itens = f"os itens {itens_texto} revestem-se"

    partes_gestores = []
    nomes_com_cargo = []
    for gestor in gestores:
        nome_cargo = _nome_com_cargo(gestor)
        nomes_com_cargo.append(nome_cargo)
        cargo = str(gestor.get("cargo", "")).strip().upper()
        sexo = str(gestor.get("sexo", "")).strip().upper()
        if cargo.startswith("CNPJ"):
            partes_gestores.append(f"da {nome_cargo}")
        elif sexo == "F":
            partes_gestores.append(f"da Sra. {nome_cargo}")
        else:
            partes_gestores.append(f"do Sr. {nome_cargo}")
    texto_gestores = (
        partes_gestores[0]
        if len(partes_gestores) == 1
        else ", ".join(partes_gestores[:-1]) + f" e {partes_gestores[-1]}"
    )

    processos_de_julgamento = {
        "CONTAS ORDINÁRIAS",
        "PROCESSO DE CONTAS ESPECIAIS",
        "TOMADA DE CONTAS ESPECIAL",
    }
    if tipo in processos_de_julgamento:
        conclusao_final = f"pela irregularidade das contas {texto_gestores}"
    else:
        conclusao_final = (
            f"desfavorável à aprovação das contas {texto_gestores}"
        )
    template = templates["conclusao_principal"]["texto_base"]["texto"]
    texto = template.format(
        texto_itens=texto_itens,
        conclusao_final=conclusao_final,
    )
    destaques = [
        DestaqueFundamentacao(
            itens_texto,
            desnegritar_conjuncao=len(itens) > 1,
        ),
        *(DestaqueFundamentacao(nome) for nome in nomes_com_cargo),
    ]
    return ResultadoFundamentacao(texto=texto, destaques=destaques)


def construir_fundamentacao_pre_dispositivo(
    tipo_processo: str,
    procurador: str,
    responsaveis: list[dict],
    apontamentos: list[dict],
    templates: dict,
    *,
    quantidade_com_responsabilidade=0,
    quantidade_sem_responsabilidade=0,
    falhas_sem_responsabilidade: str = "",
) -> ResultadoFundamentacao:
    """Seleciona o construtor compatível com processo e procurador."""
    tipo = str(tipo_processo or "").strip().upper()
    procurador_normalizado = str(procurador or "").strip().upper()
    if tipo == "CONTAS ORDINÁRIAS" and procurador_normalizado == "FERNANDA ISMAEL":
        return construir_fundamentacao_fernanda_contas_ordinarias(
            responsaveis,
            apontamentos,
            templates,
            quantidade_com_responsabilidade=quantidade_com_responsabilidade,
            quantidade_sem_responsabilidade=quantidade_sem_responsabilidade,
            falhas_sem_responsabilidade=falhas_sem_responsabilidade,
        )
    if tipo == "CONTAS ANUAIS" and procurador_normalizado != "GERALDO COSTA DA CAMINO":
        return construir_fundamentacao_contas_anuais(
            procurador_normalizado,
            responsaveis,
            apontamentos,
            templates,
        )
    return construir_fundamentacao_padrao(
        tipo,
        responsaveis,
        apontamentos,
        templates,
    )
