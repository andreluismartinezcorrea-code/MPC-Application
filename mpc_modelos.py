"""Modelo central e migração compatível dos dados do MPC Parecer."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any, Iterable


SCHEMA_VERSION = 1
RESPONSAVEL_NAO_INTIMADO = "Responsável Não Intimado"
NAO_APRESENTOU_ESCLARECIMENTOS = "Não Apresentou Esclarecimentos"
NAO_APRESENTOU_DEFESA_LEGADO = "Não Apresentou Defesa"
ESCLARECIMENTOS_ESPONTANEOS_DESCONSIDERADOS = (
    "Esclarecimentos Espontâneos Desconsiderados"
)


def _texto(valor: Any, padrao: str = "") -> str:
    if valor is None:
        return padrao
    return str(valor)


def _nomes(valor: Any) -> list[str]:
    if isinstance(valor, (list, tuple, set)):
        origem: Iterable[Any] = valor
    else:
        origem = _texto(valor).split(" | ")
    return [str(nome).strip() for nome in origem if str(nome).strip()]


def _arquivo_real(valor: Any) -> bool:
    texto = _texto(valor).strip()
    return bool(
        texto
        and texto
        not in {
            RESPONSAVEL_NAO_INTIMADO,
            NAO_APRESENTOU_ESCLARECIMENTOS,
            NAO_APRESENTOU_DEFESA_LEGADO,
        }
    )


@dataclass(slots=True)
class DadosProcesso:
    exercicio: str = ""
    processo: str = ""
    tipo: str = ""
    orgao: str = ""
    servico: str = ""
    rag: str = ""
    peca: str = ""
    apontes: str = ""


@dataclass(slots=True)
class DadosParecer:
    tipo_parecer: str = ""
    num_parecer: str = ""
    ano_parecer: str = ""
    relator: str = ""
    num_proc: str = ""
    tipo_proc: str = ""
    ano_exercicio: str = ""
    orgao1: str = ""
    procurador: str = ""
    arquivo_parecer: str = ""


@dataclass(slots=True)
class DadosDocumentos:
    arq_anal_escl: str = ""
    pasta: str = ""
    ae_peca: str = ""
    esclarecimentos: str = ""
    peca_esclarecimentos: str = ""
    municipio: str = ""
    doc_probatoria: str = "Sim"
    tramitacao: str = ""
    responsavel_tramitacao: str = "Sem Registro"
    tram_tipo1: str = ""
    tram_num1: str = ""
    tram_tipo2: str = ""
    tram_num2: str = ""


@dataclass(slots=True)
class DadosAnalise:
    apontamento_selecionado: str = ""
    paginas_rag: str = ""
    paginas_escl: str = ""
    paginas_ae: str = ""
    voto: str = ""
    paginas_voto: str = ""
    peca_voto: str = ""
    aux_1: str = ""
    aux_2: str = ""
    aux_3: str = ""
    aux_4: str = ""
    aux_5: str = ""


@dataclass(slots=True)
class ControleFalhas:
    qtd_apontamentos: str = "0"
    falhas_com_resp: str = ""
    qtd_com_resp: str = "0"
    falhas_sem_resp: str = ""
    qtd_sem_resp: str = "0"
    falhas_sugestao_rec: str = ""
    qtd_sugestao_rec: str = "0"


@dataclass(slots=True)
class RegistroProducao:
    registro_id: str = ""
    registro_data: str = ""


@dataclass(slots=True)
class Responsavel:
    nome: str = ""
    cargo: str = ""
    sexo: str = "M"
    intimacao: str = "Sim"
    esclarecimentos: str = NAO_APRESENTOU_ESCLARECIMENTOS
    arquivo_esclarecimentos: str = ""
    regularidade: str = "Sim"
    falhas: str = "Sim"
    multa: str = "Não"
    debito: str = "Não"
    conclusao: str = ""

    @classmethod
    def from_dict(cls, dados: dict[str, Any]) -> "Responsavel":
        responsavel = cls(
            nome=_texto(dados.get("nome")),
            cargo=_texto(dados.get("cargo")),
            sexo=_texto(dados.get("sexo"), "M") or "M",
            intimacao=_texto(dados.get("intimacao"), "Sim") or "Sim",
            esclarecimentos=_texto(
                dados.get("esclarecimentos"),
                NAO_APRESENTOU_ESCLARECIMENTOS,
            ),
            arquivo_esclarecimentos=_texto(
                dados.get("arquivo_esclarecimentos")
            ),
            regularidade=_texto(dados.get("regularidade"), "Sim") or "Sim",
            falhas=_texto(dados.get("falhas"), "Sim") or "Sim",
            multa=_texto(dados.get("multa"), "Não") or "Não",
            debito=_texto(dados.get("debito"), "Não") or "Não",
            conclusao=_texto(dados.get("conclusao")),
        )
        responsavel.normalizar_compatibilidade()
        return responsavel

    def normalizar_compatibilidade(self) -> None:
        """Converte marcações antigas sem alterar decisões jurídicas atuais."""
        if self.esclarecimentos == NAO_APRESENTOU_DEFESA_LEGADO:
            self.esclarecimentos = (
                RESPONSAVEL_NAO_INTIMADO
                if self.intimacao == "Não"
                else NAO_APRESENTOU_ESCLARECIMENTOS
            )
        if self.intimacao == "Não":
            if _arquivo_real(self.arquivo_esclarecimentos):
                self.esclarecimentos = (
                    ESCLARECIMENTOS_ESPONTANEOS_DESCONSIDERADOS
                )
            elif (
                self.esclarecimentos
                != ESCLARECIMENTOS_ESPONTANEOS_DESCONSIDERADOS
            ):
                self.esclarecimentos = RESPONSAVEL_NAO_INTIMADO
                self.arquivo_esclarecimentos = RESPONSAVEL_NAO_INTIMADO

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class Apontamento:
    irregularidade: str = ""
    conclusao: str = ""
    multa: str = "Não"
    debito: str = "Não"
    valor_debito: str = ""
    repercussao: str = "Não"
    responsaveis: list[str] = field(default_factory=list)
    responsaveis_multa: list[str] = field(default_factory=list)
    responsaveis_repercussao: list[str] = field(default_factory=list)
    responsaveis_debito: list[str] = field(default_factory=list)

    @classmethod
    def from_dict(cls, dados: dict[str, Any]) -> "Apontamento":
        responsaveis = _nomes(dados.get("responsaveis"))

        def associacao_compativel(chave: str, natureza: str) -> list[str]:
            if chave in dados:
                return _nomes(dados.get(chave))
            if natureza == "falha" or _texto(dados.get(natureza)).strip() == "Sim":
                return list(responsaveis)
            return []

        return cls(
            irregularidade=_texto(
                dados.get("irregularidade") or dados.get("item")
            ),
            conclusao=_texto(dados.get("conclusao")),
            multa=_texto(dados.get("multa"), "Não") or "Não",
            debito=_texto(dados.get("debito"), "Não") or "Não",
            valor_debito=_texto(dados.get("valor_debito")),
            repercussao=_texto(dados.get("repercussao"), "Não") or "Não",
            responsaveis=responsaveis,
            responsaveis_multa=associacao_compativel(
                "responsaveis_multa", "multa"
            ),
            responsaveis_repercussao=associacao_compativel(
                "responsaveis_repercussao", "repercussao"
            ),
            responsaveis_debito=associacao_compativel(
                "responsaveis_debito", "debito"
            ),
        )

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class ProcessoMPC:
    processo: DadosProcesso = field(default_factory=DadosProcesso)
    parecer: DadosParecer = field(default_factory=DadosParecer)
    documentos: DadosDocumentos = field(default_factory=DadosDocumentos)
    analise: DadosAnalise = field(default_factory=DadosAnalise)
    controle: ControleFalhas = field(default_factory=ControleFalhas)
    registro: RegistroProducao = field(default_factory=RegistroProducao)
    responsaveis: list[Responsavel] = field(default_factory=list)
    apontamentos: list[Apontamento] = field(default_factory=list)
    extras: dict[str, Any] = field(default_factory=dict)
    schema_version: int = SCHEMA_VERSION

    @classmethod
    def from_legacy_dict(cls, dados: dict[str, Any]) -> "ProcessoMPC":
        conhecidos = {
            "schema_version",
            *DadosProcesso.__dataclass_fields__,
            *DadosParecer.__dataclass_fields__,
            *DadosDocumentos.__dataclass_fields__,
            *DadosAnalise.__dataclass_fields__,
            *ControleFalhas.__dataclass_fields__,
            *RegistroProducao.__dataclass_fields__,
            "responsaveis",
            "apontamentos_detalhado",
            "apontamentos_lista",
        }

        def construir(tipo):
            return tipo(
                **{
                    campo: _texto(dados.get(campo), definicao.default)
                    for campo, definicao in tipo.__dataclass_fields__.items()
                }
            )

        apontamentos_brutos = dados.get("apontamentos_detalhado") or [
            {"irregularidade": item}
            for item in dados.get("apontamentos_lista", [])
        ]
        return cls(
            processo=construir(DadosProcesso),
            parecer=construir(DadosParecer),
            documentos=construir(DadosDocumentos),
            analise=construir(DadosAnalise),
            controle=construir(ControleFalhas),
            registro=construir(RegistroProducao),
            responsaveis=[
                Responsavel.from_dict(item)
                for item in dados.get("responsaveis", [])
                if isinstance(item, dict)
            ],
            apontamentos=[
                Apontamento.from_dict(item)
                for item in apontamentos_brutos
                if isinstance(item, dict)
            ],
            extras={
                chave: valor
                for chave, valor in dados.items()
                if chave not in conhecidos
            },
            schema_version=SCHEMA_VERSION,
        )

    def to_legacy_dict(self) -> dict[str, Any]:
        """Mantém o formato já aceito pelas versões anteriores da aplicação."""
        dados: dict[str, Any] = {}
        dados.update(asdict(self.processo))
        dados.update(asdict(self.parecer))
        dados.update(asdict(self.documentos))
        dados.update(asdict(self.analise))
        dados.update(asdict(self.controle))
        dados.update(asdict(self.registro))
        dados.update(self.extras)
        dados["schema_version"] = self.schema_version
        dados["responsaveis"] = [item.to_dict() for item in self.responsaveis]
        dados["apontamentos_detalhado"] = [
            item.to_dict() for item in self.apontamentos
        ]
        dados["apontamentos_lista"] = [
            item.irregularidade for item in self.apontamentos
        ]
        return dados


def normalizar_dados_persistidos(dados: dict[str, Any]) -> dict[str, Any]:
    """Migra qualquer JSON legado para o formato plano atual, sem perdas."""
    if not isinstance(dados, dict):
        raise ValueError("Os dados salvos devem formar um objeto JSON.")
    return ProcessoMPC.from_legacy_dict(dados).to_legacy_dict()
