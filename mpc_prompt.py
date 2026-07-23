"""Construção local e testável de prompts jurídicos a partir da GUI."""

from __future__ import annotations

import ntpath
import re
from typing import Any, Iterable

from mpc_regras import obter_responsaveis_apontamento


MARCADORES_SEM_ARQUIVO = {
    "",
    "N/C",
    "NÃO CONSTA",
    "NÃO INFORMADO",
    "RESPONSÁVEL NÃO INTIMADO",
    "NÃO APRESENTOU ESCLARECIMENTOS",
    "NÃO APRESENTOU DEFESA",
}


def _texto(dados: dict[str, Any], chave: str) -> str:
    return str(dados.get(chave, "") or "").strip()


def _nome_arquivo(valor: Any) -> str:
    texto = str(valor or "").strip().strip('"')
    return ntpath.basename(texto) or texto


def _arquivo_informado(valor: Any) -> bool:
    texto = str(valor or "").strip()
    return bool(texto and texto.upper() not in MARCADORES_SEM_ARQUIVO)


def _formatar_lista(itens: Iterable[str]) -> str:
    valores = [str(item).strip() for item in itens if str(item).strip()]
    if not valores:
        return ""
    if len(valores) == 1:
        return valores[0]
    return ", ".join(valores[:-1]) + " e " + valores[-1]


def listar_apontamentos_prompt(dados: dict[str, Any]) -> list[dict[str, Any]]:
    """Normaliza número, descrição e associações dos apontamentos da GUI."""
    resultado = []
    for indice, item in enumerate(
        dados.get("apontamentos_detalhado", []),
        start=1,
    ):
        if not isinstance(item, dict):
            continue
        texto = str(
            item.get("irregularidade") or item.get("item") or ""
        ).strip()
        if not texto:
            continue
        correspondencia = re.match(
            r"^(\d+(?:\.\d+)*)(?:\s*[-–—:]?\s*)(.*)$",
            texto,
        )
        if correspondencia:
            numero = correspondencia.group(1)
            descricao = correspondencia.group(2).strip()
        else:
            numero = f"linha {indice}"
            descricao = texto
        rotulo = numero + (f" - {descricao}" if descricao else "")
        resultado.append(
            {
                "id": f"apontamento_{indice}",
                "numero": numero,
                "descricao": descricao,
                "rotulo": rotulo,
                "dados": item,
            }
        )
    return resultado


def listar_fontes_prompt(dados: dict[str, Any]) -> list[dict[str, Any]]:
    """Reúne os documentos efetivamente informados nas diversas abas."""
    fontes: list[dict[str, Any]] = []

    def adicionar(
        identificador: str,
        categoria: str,
        rotulo: str,
        valor: Any,
        *,
        selecionado_padrao: bool = True,
        responsavel: str = "",
        paginas: str = "",
        peca: str = "",
    ) -> None:
        if not _arquivo_informado(valor):
            return
        fontes.append(
            {
                "id": identificador,
                "categoria": categoria,
                "rotulo": rotulo,
                "nome": _nome_arquivo(valor),
                "caminho": str(valor).strip(),
                "selecionado_padrao": selecionado_padrao,
                "responsavel": responsavel,
                "paginas": str(paginas or "").strip(),
                "peca": str(peca or "").strip(),
            }
        )

    adicionar(
        "relatorio_auditoria",
        "relatorio_auditoria",
        "Relatório de Auditoria",
        dados.get("rag"),
        paginas=_texto(dados, "paginas_rag"),
        peca=_texto(dados, "peca"),
    )
    adicionar(
        "analise_esclarecimentos",
        "analise_esclarecimentos",
        "Análise de Esclarecimentos",
        dados.get("arq_anal_escl"),
        paginas=(
            _texto(dados, "paginas_ae")
            or _texto(dados, "paginas_escl")
        ),
        peca=_texto(dados, "ae_peca"),
    )

    caminhos_individuais = set()
    for indice, responsavel in enumerate(dados.get("responsaveis", []), start=1):
        if not isinstance(responsavel, dict):
            continue
        caminho = _texto(responsavel, "arquivo_esclarecimentos")
        if not _arquivo_informado(caminho):
            continue
        chave_caminho = caminho.casefold()
        if chave_caminho in caminhos_individuais:
            continue
        caminhos_individuais.add(chave_caminho)
        nome = _texto(responsavel, "nome") or f"Administrador {indice}"
        adicionar(
            f"esclarecimentos_{indice}",
            "esclarecimentos_administrador",
            f"Esclarecimentos de {nome}",
            caminho,
            responsavel=nome,
        )

    # Compatibilidade com processos antigos que possuíam somente o campo geral.
    esclarecimentos_gerais = _texto(dados, "esclarecimentos")
    if (
        not caminhos_individuais
        and _arquivo_informado(esclarecimentos_gerais)
        and re.search(r"\.(?:pdf|docx?)\b", esclarecimentos_gerais, re.I)
    ):
        adicionar(
            "esclarecimentos_gerais",
            "esclarecimentos_administrador",
            "Esclarecimentos dos Administradores",
            esclarecimentos_gerais,
            paginas=_texto(dados, "paginas_escl"),
            peca=_texto(dados, "peca_esclarecimentos"),
        )

    adicionar(
        "relatorio_voto",
        "relatorio_voto",
        "Relatório e Voto",
        dados.get("voto"),
        paginas=_texto(dados, "paginas_voto"),
        peca=_texto(dados, "peca_voto"),
    )
    adicionar(
        "parecer_existente",
        "parecer_existente",
        "e-Parecer existente",
        dados.get("arquivo_parecer"),
        selecionado_padrao=False,
    )
    return fontes


def listar_responsaveis_prompt(dados: dict[str, Any]) -> list[dict[str, Any]]:
    """Lista os administradores que podem integrar o contexto do prompt."""
    resultado = []
    for indice, responsavel in enumerate(dados.get("responsaveis", []), start=1):
        if not isinstance(responsavel, dict):
            continue
        nome = _texto(responsavel, "nome")
        if not nome:
            continue
        resultado.append(
            {
                "id": f"responsavel_{indice}",
                "nome": nome,
                "cargo": _texto(responsavel, "cargo"),
                "intimacao": _texto(responsavel, "intimacao"),
                "esclarecimentos": _texto(responsavel, "esclarecimentos"),
                "conclusao": _texto(responsavel, "conclusao"),
            }
        )
    return resultado


def _referencia_fonte(fonte: dict[str, Any]) -> str:
    referencia = f"“{fonte['nome']}”"
    metadados = []
    if str(fonte.get("peca", "")).strip():
        metadados.append(f"peça {fonte['peca']}")
    if str(fonte.get("paginas", "")).strip():
        metadados.append(f"págs. {fonte['paginas']}")
    if metadados:
        referencia += f" ({'; '.join(metadados)})"
    return referencia


def _fontes_categoria(
    fontes: list[dict[str, Any]],
    categoria: str,
) -> list[dict[str, Any]]:
    return [fonte for fonte in fontes if fonte.get("categoria") == categoria]


def _descrever_responsaveis(
    responsaveis: list[dict[str, Any]],
) -> str:
    linhas = []
    for responsavel in responsaveis:
        identificacao = responsavel["nome"]
        if responsavel.get("cargo"):
            identificacao += f" ({responsavel['cargo']})"
        dados = []
        if responsavel.get("intimacao"):
            dados.append(f"intimação: {responsavel['intimacao']}")
        if responsavel.get("esclarecimentos"):
            dados.append(
                f"esclarecimentos: {responsavel['esclarecimentos']}"
            )
        if responsavel.get("conclusao"):
            dados.append(f"conclusão registrada: {responsavel['conclusao']}")
        linhas.append(
            f"- {identificacao}" + (f" — {'; '.join(dados)}" if dados else "")
        )
    return "\n".join(linhas)


def _descrever_associacoes(
    apontamentos: list[dict[str, Any]],
    nomes_permitidos: set[str],
) -> str:
    linhas = []
    for apontamento in apontamentos:
        dados = apontamento["dados"]
        partes = []
        for rotulo, natureza in (
            ("falha", "falha"),
            ("multa", "multa"),
            ("repercussão", "repercussao"),
            ("débito", "debito"),
        ):
            nomes = obter_responsaveis_apontamento(dados, natureza)
            nomes = [nome for nome in nomes if nome in nomes_permitidos]
            if nomes:
                partes.append(f"{rotulo}: {', '.join(nomes)}")
        if str(dados.get("debito", "")).strip() == "Sim" and str(
            dados.get("valor_debito", "")
        ).strip():
            partes.append(f"valor do débito: {dados['valor_debito']}")
        conclusao = str(dados.get("conclusao", "") or "").strip()
        if conclusao:
            partes.insert(0, f"conclusão do item: {conclusao}")
        if partes:
            linhas.append(f"- {apontamento['rotulo']} — {'; '.join(partes)}")
    return "\n".join(linhas)


def construir_prompt(
    dados: dict[str, Any],
    apontamentos: list[dict[str, Any]],
    fontes: list[dict[str, Any]],
    responsaveis: list[dict[str, Any]],
    *,
    incluir_contexto: bool = True,
    incluir_associacoes: bool = True,
    exigir_referencias: bool = True,
    separar_por_item: bool = True,
    orientacoes_adicionais: str = "",
) -> str:
    """Monta a instrução final com concordância e fontes selecionadas."""
    if not apontamentos:
        raise ValueError("Selecione ao menos um apontamento.")
    if not fontes:
        raise ValueError("Selecione ao menos um arquivo de referência.")

    referencias = [_referencia_fonte(fonte) for fonte in fontes]
    arquivos_texto = _formatar_lista(referencias)
    abertura_arquivo = (
        f"Com base no arquivo anexado {arquivos_texto}"
        if len(fontes) == 1
        else f"Com base nos arquivos anexados {arquivos_texto}"
    )
    itens_texto = _formatar_lista(
        [f"{item['rotulo']}" for item in apontamentos]
    )
    objeto = (
        f"o item {itens_texto}"
        if len(apontamentos) == 1
        else f"os itens {itens_texto}"
    )
    linhas = [
        f"{abertura_arquivo}, analise {objeto} do Relatório de Auditoria.",
    ]

    if incluir_contexto:
        contexto = []
        for chave, rotulo in (
            ("processo", "Processo"),
            ("exercicio", "Exercício"),
            ("tipo", "Tipo de processo"),
            ("orgao", "Órgão"),
        ):
            valor = _texto(dados, chave)
            if valor:
                contexto.append(f"{rotulo}: {valor}")
        if contexto:
            linhas.extend(["", "CONTEXTO DO PROCESSO", " | ".join(contexto)])
        if responsaveis:
            linhas.extend(
                [
                    "",
                    "ADMINISTRADORES CONSIDERADOS",
                    _descrever_responsaveis(responsaveis),
                ]
            )

    fontes_rag = _fontes_categoria(fontes, "relatorio_auditoria")
    fontes_analise = _fontes_categoria(fontes, "analise_esclarecimentos")
    fontes_defesa = _fontes_categoria(
        fontes,
        "esclarecimentos_administrador",
    )
    fontes_voto = _fontes_categoria(fontes, "relatorio_voto")
    fontes_parecer = _fontes_categoria(fontes, "parecer_existente")

    referencia_rag = _formatar_lista(
        [_referencia_fonte(fonte) for fonte in fontes_rag]
    ) or "documentação de auditoria selecionada"
    referencia_analise = _formatar_lista(
        [_referencia_fonte(fonte) for fonte in fontes_analise]
    ) or "documentação do Órgão Instrutivo selecionada"

    linhas.extend(
        [
            "",
            "ESTRUTURA OBRIGATÓRIA DA MINUTA",
            "",
            "1. RELATÓRIO DE AUDITORIA",
            "Descreva objetivamente cada falha apontada pela equipe de "
            f"auditoria, conforme consta de {referencia_rag}. Identifique o "
            "critério utilizado, a situação encontrada e os possíveis efeitos, "
            "sem antecipar a opinião do MPC.",
            "",
            "2. ESCLARECIMENTOS DOS ADMINISTRADORES",
        ]
    )
    if fontes_defesa:
        referencia_defesas = _formatar_lista(
            [_referencia_fonte(fonte) for fonte in fontes_defesa]
        )
        linhas.append(
            "Descreva separadamente os argumentos e documentos apresentados "
            "pelos Administradores, constantes "
            + ("do arquivo " if len(fontes_defesa) == 1 else "dos arquivos ")
            + f"{referencia_defesas}. Quando a defesa também estiver resumida "
            f"em {referencia_analise}, confronte as duas fontes e indique "
            "eventual divergência."
        )
    else:
        linhas.append(
            "Registre, com base nos arquivos selecionados e nos dados da GUI, "
            "os esclarecimentos eventualmente apresentados e identifique "
            "expressamente os Administradores que não apresentaram defesa ou "
            "não foram intimados."
        )

    linhas.extend(
        [
            "",
            "3. CONCLUSÕES DO ÓRGÃO INSTRUTIVO",
            "Exponha de maneira autônoma as conclusões do Órgão Instrutivo, "
            f"constantes de {referencia_analise}. Informe se os argumentos "
            "foram acolhidos, rejeitados ou acolhidos parcialmente e não "
            "confunda essa conclusão técnica com a futura opinião do MPC.",
            "",
            "4. ANÁLISE E MINUTA DE PARECER DO MPC",
            "Analise criticamente cada achado, confronte auditoria, defesa e "
            "conclusão instrutiva e apresente fundamentação própria acerca da "
            "procedência, improcedência ou procedência parcial da falha. "
            "Indique, quando cabível, responsabilidade, multa, repercussão nas "
            "contas, débito, alerta ou recomendação.",
        ]
    )

    if fontes_voto:
        linhas.append(
            "Utilize também o Relatório e Voto "
            f"({_formatar_lista([_referencia_fonte(f) for f in fontes_voto])}) "
            "como elemento de confronto na quarta parte, preservando a "
            "independência da opinião do MPC."
        )
    if fontes_parecer:
        linhas.append(
            "O e-Parecer existente "
            f"({_formatar_lista([_referencia_fonte(f) for f in fontes_parecer])}) "
            "serve apenas como referência auxiliar; não reproduza conclusões "
            "sem confrontá-las com os documentos primários."
        )

    if incluir_associacoes:
        nomes = {responsavel["nome"] for responsavel in responsaveis}
        associacoes = _descrever_associacoes(apontamentos, nomes)
        if associacoes:
            linhas.extend(
                [
                    "",
                    "DADOS DE RESPONSABILIZAÇÃO REGISTRADOS NA GUI",
                    associacoes,
                    "Use essas associações como referência de trabalho e "
                    "confirme-as criticamente nos documentos selecionados.",
                ]
            )

    linhas.extend(["", "ORIENTAÇÕES DE QUALIDADE"])
    if separar_por_item and len(apontamentos) > 1:
        linhas.append(
            "- Dentro de cada uma das quatro partes, trate os itens "
            "separadamente e preserve a numeração original."
        )
    linhas.extend(
        [
            "- Não invente fatos, fundamentos, valores, páginas ou argumentos "
            "que não estejam nos arquivos selecionados.",
            "- Identifique expressamente divergências entre auditoria, defesa, "
            "Órgão Instrutivo e MPC.",
            "- Empregue linguagem técnico-jurídica clara, objetiva e impessoal.",
        ]
    )
    if exigir_referencias:
        linhas.append(
            "- Sempre que a informação estiver disponível, indique entre "
            "parênteses o nome do arquivo e a página ou peça que sustenta a "
            "afirmação."
        )
    orientacoes = str(orientacoes_adicionais or "").strip()
    if orientacoes:
        linhas.extend(["", "ORIENTAÇÕES ADICIONAIS DO USUÁRIO", orientacoes])

    return "\n".join(linhas).strip()
