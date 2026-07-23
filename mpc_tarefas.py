"""Execução segura de tarefas em segundo plano, independente do Tkinter."""

from __future__ import annotations

import queue
import threading
import traceback
from dataclasses import dataclass, field
from typing import Any, Callable


Agendador = Callable[[int, Callable[[], None]], Any]


@dataclass(frozen=True, slots=True)
class ResultadoTarefa:
    valor: Any = None
    erro: BaseException | None = None
    detalhes_erro: str = ""

    @property
    def sucesso(self) -> bool:
        return self.erro is None


@dataclass(slots=True)
class ControleTarefa:
    """Permite consultar e cancelar a entrega tardia de uma tarefa."""

    _cancelamento: threading.Event = field(default_factory=threading.Event)
    _conclusao: threading.Event = field(default_factory=threading.Event)
    thread: threading.Thread | None = None

    def cancelar(self) -> None:
        self._cancelamento.set()

    @property
    def cancelada(self) -> bool:
        return self._cancelamento.is_set()

    @property
    def concluida(self) -> bool:
        return self._conclusao.is_set()


class ExecutorTarefas:
    """Executa no worker e entrega callbacks somente no agendador informado."""

    def __init__(
        self,
        agendar: Agendador,
        *,
        esta_ativo: Callable[[], bool] | None = None,
        intervalo_ms: int = 100,
    ):
        if not callable(agendar):
            raise TypeError("Informe uma função de agendamento válida.")
        self._agendar = agendar
        self._esta_ativo = esta_ativo or (lambda: True)
        self._intervalo_ms = max(1, int(intervalo_ms))

    def _ativo(self) -> bool:
        try:
            return bool(self._esta_ativo())
        except Exception:
            return False

    def executar(
        self,
        tarefa: Callable[..., Any],
        *args: Any,
        ao_sucesso: Callable[[Any], None] | None = None,
        ao_erro: Callable[[BaseException, str], None] | None = None,
        ao_finalizar: Callable[[], None] | None = None,
        preparar_worker: Callable[[], None] | None = None,
        finalizar_worker: Callable[[], None] | None = None,
        daemon: bool = True,
        nome: str | None = None,
    ) -> ControleTarefa:
        if not callable(tarefa):
            raise TypeError("A tarefa deve ser executável.")

        resultados: queue.Queue[ResultadoTarefa] = queue.Queue(maxsize=1)
        controle = ControleTarefa()

        def worker() -> None:
            preparado = False
            try:
                if preparar_worker is not None:
                    preparar_worker()
                    preparado = True
                valor = tarefa(*args)
                resultado = ResultadoTarefa(valor=valor)
            except BaseException as erro:
                resultado = ResultadoTarefa(
                    erro=erro,
                    detalhes_erro=traceback.format_exc(),
                )
            finally:
                if preparado and finalizar_worker is not None:
                    try:
                        finalizar_worker()
                    except BaseException as erro_finalizacao:
                        if "resultado" not in locals() or resultado.sucesso:
                            resultado = ResultadoTarefa(
                                erro=erro_finalizacao,
                                detalhes_erro=traceback.format_exc(),
                            )
            resultados.put(resultado)

        def finalizar_interface() -> None:
            if ao_finalizar is not None:
                try:
                    ao_finalizar()
                except Exception:
                    pass

        def relatar_erro(erro: BaseException, detalhes: str) -> None:
            if ao_erro is not None:
                try:
                    ao_erro(erro, detalhes)
                except Exception:
                    pass

        def consultar() -> None:
            if not self._ativo():
                controle.cancelar()
                return
            try:
                resultado = resultados.get_nowait()
            except queue.Empty:
                try:
                    self._agendar(self._intervalo_ms, consultar)
                except Exception:
                    controle.cancelar()
                return

            controle._conclusao.set()
            if controle.cancelada:
                finalizar_interface()
                return
            try:
                if resultado.sucesso:
                    if ao_sucesso is not None:
                        ao_sucesso(resultado.valor)
                else:
                    relatar_erro(resultado.erro, resultado.detalhes_erro)
            except BaseException as erro_callback:
                relatar_erro(erro_callback, traceback.format_exc())
            finally:
                finalizar_interface()

        thread = threading.Thread(
            target=worker,
            name=nome,
            daemon=daemon,
        )
        controle.thread = thread
        thread.start()
        try:
            self._agendar(self._intervalo_ms, consultar)
        except Exception:
            controle.cancelar()
        return controle


def iniciar_tarefa_isolada(
    tarefa: Callable[[], Any],
    *,
    preparar_worker: Callable[[], None] | None = None,
    finalizar_worker: Callable[[], None] | None = None,
    daemon: bool = True,
    nome: str | None = None,
) -> threading.Thread:
    """Inicia tarefa sem retorno à GUI, com ciclo opcional de preparação."""

    def runner() -> None:
        preparado = False
        try:
            if preparar_worker is not None:
                preparar_worker()
                preparado = True
            tarefa()
        finally:
            if preparado and finalizar_worker is not None:
                finalizar_worker()

    thread = threading.Thread(target=runner, daemon=daemon, name=nome)
    thread.start()
    return thread
