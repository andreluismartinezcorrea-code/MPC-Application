"""Sistema visual e auxiliares de apresentação do MPC Parecer."""

from __future__ import annotations

import re


def calcular_colunas_acoes(
    largura_disponivel: int,
    *,
    largura_minima: int = 230,
    maximo_colunas: int = 3,
) -> int:
    """Calcula uma grade de ações legível para a largura disponível.

    A função é deliberadamente independente do Tkinter para que a regra
    responsiva possa ser testada sem abrir a interface gráfica.
    """
    try:
        largura = max(0, int(largura_disponivel))
        minima = max(1, int(largura_minima))
        limite = max(1, int(maximo_colunas))
    except (TypeError, ValueError):
        return 1
    return max(1, min(limite, largura // minima))


def posicao_acao_grade(indice: int, quantidade_colunas: int) -> tuple[int, int]:
    """Retorna linha e coluna de um botao em uma grade responsiva."""
    indice_seguro = max(0, int(indice))
    colunas_seguras = max(1, int(quantidade_colunas))
    return divmod(indice_seguro, colunas_seguras)


def extrair_progresso_fluxo(resumo: str) -> tuple[int, int, int]:
    """Extrai ``atual``, ``total`` e percentual do resumo textual do fluxo."""
    texto = str(resumo or "")
    correspondencia = re.search(
        r"PROGRESSO:\s*(\d+)\s*/\s*(\d+)\s*\((\d+)%\)",
        texto,
        flags=re.IGNORECASE,
    )
    if not correspondencia:
        return 0, 0, 0
    atual, total, percentual = map(int, correspondencia.groups())
    return atual, total, max(0, min(100, percentual))


def configurar_estilos_aplicacao(estilo) -> None:
    """Aplica a tipografia, espaçamento e hierarquia visual da interface."""
    estilo.configure("App.TFrame")
    estilo.configure("AppHeader.TFrame")

    # Text and Heading Fonts - adjusted sizes for modern look
    estilo.configure(
        "AppTitle.TLabel",
        font=("Segoe UI", 22, "bold"),
    )
    estilo.configure(
        "AppEyebrow.TLabel",
        font=("Segoe UI Semibold", 9),
    )
    estilo.configure(
        "AppProcess.TLabel",
        font=("Segoe UI Semibold", 13),
    )
    estilo.configure(
        "AppMeta.TLabel",
        font=("Segoe UI", 10),
    )
    estilo.configure(
        "AppStatus.TLabel",
        font=("Segoe UI Semibold", 11),
    )
    estilo.configure(
        "Title.TLabel",
        font=("Segoe UI Semibold", 20),
    )
    estilo.configure("Subtitle.TLabel", font=("Segoe UI", 11))

    # Frame styling
    estilo.configure(
        "Section.TLabelframe",
        padding=15,
    )
    estilo.configure(
        "Section.TLabelframe.Label",
        font=("Segoe UI Semibold", 12),
        padding=(0, 0, 0, 5),
    )
    estilo.configure("Header.TLabel", font=("Segoe UI Semibold", 10))

    # Input/Entry padding
    estilo.configure("TEntry", padding=(6, 4))
    estilo.configure("TCombobox", padding=(6, 4))

    # Notebook Tabs - more padding for touch/modern feel
    estilo.configure(
        "TNotebook.Tab",
        font=("Segoe UI Semibold", 11),
        padding=(20, 10),
    )

    # Treeview spacing
    estilo.configure("Treeview", rowheight=32, font=("Segoe UI", 10))
    estilo.configure(
        "Treeview.Heading",
        font=("Segoe UI Semibold", 10),
        padding=(5, 5),
    )

    # Button styling
    estilo.configure("TButton", padding=(12, 8), font=("Segoe UI", 10, "bold"))
    estilo.configure("Primary.TButton", padding=(12, 8), font=("Segoe UI Semibold", 10))


def configurar_widgets_tk_legados(janela, cores) -> None:
    """Harmoniza widgets ``tk`` antigos com o tema atual do ttkbootstrap."""
    fundo = getattr(cores, "bg", "#222222")
    texto = getattr(cores, "fg", "#f5f5f5")
    fundo_campo = getattr(cores, "inputbg", fundo)
    texto_campo = getattr(cores, "inputfg", texto)
    selecao = getattr(cores, "selectbg", getattr(cores, "primary", "#375a7f"))
    selecao_texto = getattr(cores, "selectfg", "#ffffff")

    opcoes = {
        "*Font": ("Segoe UI", 10),
        "*Label.background": fundo,
        "*Label.foreground": texto,
        "*Frame.background": fundo,
        "*Canvas.background": fundo,
        "*Entry.background": fundo_campo,
        "*Entry.foreground": texto_campo,
        "*Entry.insertBackground": texto_campo,
        "*Text.background": fundo_campo,
        "*Text.foreground": texto_campo,
        "*Text.insertBackground": texto_campo,
        "*Listbox.background": fundo_campo,
        "*Listbox.foreground": texto_campo,
        "*Listbox.selectBackground": selecao,
        "*Listbox.selectForeground": selecao_texto,
    }
    for chave, valor in opcoes.items():
        janela.option_add(chave, valor)
