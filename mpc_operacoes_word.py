"""Controladores do ciclo seguro das operações que escrevem no Word.

As regras de certificação, confirmação, backup, execução e histórico ficam
centralizadas aqui. A GUI fornece apenas as funções concretas de diálogo,
backup, obtenção do documento e apresentação de erros.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable

from mpc_certificacao import (
    OPERACOES_COM_WORD_ATIVO,
    OPERACOES_WORD,
    certificar_dados_pre_word,
    certificar_estrutura_documento,
    possivel_divergencia_documento_processo,
)


@dataclass(frozen=True, slots=True)
class CertificacaoOperacaoWord:
    operacao: str
    nome_documento: str
    total_responsaveis: int
    total_apontamentos: int
    erros: tuple[str, ...]
    avisos: tuple[str, ...]
    verificacoes: tuple[str, ...]
    inclui_backup_word: bool

    @property
    def aprovada(self) -> bool:
        return not self.erros


@dataclass(frozen=True, slots=True)
class ResultadoCicloOperacao:
    operacao: str
    status: str
    erro: Exception | None = None
    certificacao: CertificacaoOperacaoWord | None = None

    @property
    def concluida(self) -> bool:
        return self.status == "concluída"


def certificar_operacao_word(
    operacao: str,
    dados: dict[str, Any],
    obter_documento: Callable[[], Any] | None = None,
) -> CertificacaoOperacaoWord:
    """Certifica dados e, quando necessário, a estrutura do Word ativo."""
    resultado = certificar_dados_pre_word(operacao, dados)
    erros = list(resultado.erros)
    avisos = list(resultado.avisos)
    verificacoes = list(resultado.verificacoes)
    nome_documento = "novo documento gerado pela rotina"
    inclui_word = operacao in OPERACOES_COM_WORD_ATIVO

    if inclui_word:
        if obter_documento is None:
            erros.append("O controlador não recebeu acesso ao documento Word ativo.")
        else:
            try:
                documento = obter_documento()
                nome_documento = str(getattr(documento, "Name", "") or "").strip()
                if not nome_documento:
                    erros.append(
                        "O Word não informou o nome do documento ativo. Ative o "
                        "documento correto e tente novamente."
                    )
                if bool(getattr(documento, "ReadOnly", False)):
                    erros.append(
                        f"O documento '{nome_documento}' está aberto somente para "
                        "leitura e não pode ser alterado."
                    )
                texto = str(documento.Content.Text or "")
                titulos = [
                    str(getattr(controle, "Title", "") or "")
                    for controle in documento.ContentControls
                ]
                erros_estrutura, avisos_estrutura, verificacoes_estrutura = (
                    certificar_estrutura_documento(operacao, texto, titulos)
                )
                erros.extend(erros_estrutura)
                avisos.extend(avisos_estrutura)
                verificacoes.extend(verificacoes_estrutura)
                numero = str(dados.get("processo", "") or "").strip()
                if possivel_divergencia_documento_processo(nome_documento, numero):
                    avisos.append(
                        f"O número do processo '{numero}' não foi identificado no "
                        f"nome do documento ativo ('{nome_documento}'). Confirme se "
                        "este é o Word correto."
                    )
            except Exception as erro_word:
                erros.append(
                    "Não foi possível certificar o documento Word ativo. Abra e "
                    "ative o documento que deseja alterar e confirme que ele está "
                    f"acessível. Detalhe: {erro_word}"
                )

    return CertificacaoOperacaoWord(
        operacao=operacao,
        nome_documento=nome_documento,
        total_responsaveis=resultado.total_responsaveis,
        total_apontamentos=resultado.total_apontamentos,
        erros=tuple(dict.fromkeys(erros)),
        avisos=tuple(dict.fromkeys(avisos)),
        verificacoes=tuple(dict.fromkeys(verificacoes)),
        inclui_backup_word=inclui_word,
    )


def construir_mensagem_confirmacao(certificacao: CertificacaoOperacaoWord) -> str:
    """Monta a confirmação padronizada exibida antes da alteração do Word."""
    itens = "\n".join(
        f"✓ {item.capitalize()}" for item in certificacao.verificacoes
    )
    avisos = ""
    if certificacao.avisos:
        avisos = "\n\nATENÇÃO:\n" + "\n".join(
            f"• {aviso}" for aviso in certificacao.avisos
        )
    texto_backup = (
        " e do estado atual do documento.\n\n"
        if certificacao.inclui_backup_word
        else ". A rotina criará um novo documento Word.\n\n"
    )
    return (
        f"A rotina “{certificacao.operacao}” está pronta para prosseguir.\n\n"
        f"Documento: {certificacao.nome_documento}\n"
        f"Responsáveis considerados: {certificacao.total_responsaveis}\n"
        f"Apontamentos considerados: {certificacao.total_apontamentos}\n\n"
        f"{itens}{avisos}\n\n"
        "Antes de continuar, o programa criará um backup automático dos dados"
        f"{texto_backup}Deseja continuar?"
    )


def executar_ciclo_operacao(
    operacao: str,
    comando: Callable[[], Any],
    *,
    dados: dict[str, Any],
    obter_documento: Callable[[], Any] | None,
    confirmar: Callable[[CertificacaoOperacaoWord], bool],
    criar_backup: Callable[[bool], Any],
    confirmar_sem_backup: Callable[[Exception], bool],
    registrar: Callable[[str, str, Any], Any],
    registrar_historico: bool = True,
) -> ResultadoCicloOperacao:
    """Executa o ciclo completo e devolve um resultado explícito à GUI."""
    certificacao = None
    if operacao in OPERACOES_WORD:
        certificacao = certificar_operacao_word(operacao, dados, obter_documento)
        if not certificacao.aprovada:
            if registrar_historico:
                registrar(
                    operacao,
                    "bloqueada",
                    f"{len(certificacao.erros)} pendência(s) na certificação pré-Word",
                )
            return ResultadoCicloOperacao(operacao, "bloqueada", certificacao=certificacao)
        if not confirmar(certificacao):
            if registrar_historico:
                registrar(operacao, "cancelada", "cancelada pelo usuário")
            return ResultadoCicloOperacao(operacao, "cancelada", certificacao=certificacao)
        try:
            criar_backup(certificacao.inclui_backup_word)
        except Exception as erro_backup:
            if not confirmar_sem_backup(erro_backup):
                if registrar_historico:
                    registrar(operacao, "cancelada", "backup não concluído")
                return ResultadoCicloOperacao(
                    operacao, "cancelada", erro_backup, certificacao
                )

    try:
        if registrar_historico:
            registrar(operacao, "iniciada", "")
        comando()
        if registrar_historico:
            registrar(operacao, "concluída", "")
        return ResultadoCicloOperacao(operacao, "concluída", certificacao=certificacao)
    except Exception as erro:
        if registrar_historico:
            registrar(operacao, "erro", erro)
        return ResultadoCicloOperacao(operacao, "erro", erro, certificacao)
