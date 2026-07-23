"""Painel Tkinter da Biblioteca Jurídica Local."""

from __future__ import annotations

import os
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Any, Callable

import pyperclip
import ttkbootstrap as ttk

from mpc_biblioteca import BibliotecaLocal, MODOS_PESQUISA, ResultadoIndexacao


class PainelBibliotecaLocal(ttk.Frame):
    def __init__(
        self,
        master,
        *,
        biblioteca: BibliotecaLocal,
        executar_tarefa: Callable[..., Any],
        inserir_no_word: Callable[[str], Any],
        logger=None,
    ):
        super().__init__(master, padding=12)
        self.biblioteca = biblioteca
        self.executar_tarefa = executar_tarefa
        self.inserir_no_word = inserir_no_word
        self.logger = logger
        self.documento_atual: dict[str, Any] | None = None
        self.status_var = tk.StringVar(value="Biblioteca pronta para configuração.")
        self.categoria_busca_var = tk.StringVar(value="Todos")
        self.modo_busca_var = tk.StringVar(value="Todas as palavras")
        self._construir()
        self.atualizar_acervos()

    def _construir(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        fontes = ttk.LabelFrame(
            self, text="Pastas indexadas", padding=10, style="Section.TLabelframe"
        )
        fontes.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        fontes.columnconfigure(0, weight=1)

        botoes = ttk.Frame(fontes)
        botoes.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(
            botoes, text="+ ADICIONAR PASTA", command=self.adicionar_pasta,
            bootstyle="success",
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            botoes, text="GERENCIAR CATEGORIAS", command=self.gerenciar_categorias,
            bootstyle="secondary-outline",
        ).pack(side="left", padx=6)
        ttk.Button(
            botoes, text="ALTERAR LOCAL", command=self.alterar_pasta,
            bootstyle="info-outline",
        ).pack(side="left", padx=6)
        ttk.Button(
            botoes, text="ALTERAR CATEGORIA", command=self.alterar_categoria,
            bootstyle="info-outline",
        ).pack(side="left", padx=6)
        ttk.Button(
            botoes, text="REMOVER DA BIBLIOTECA", command=self.remover_pasta,
            bootstyle="danger-outline",
        ).pack(side="left", padx=6)
        ttk.Button(
            botoes, text="INDEXAR SELECIONADA", command=self.indexar_selecionada,
            bootstyle="primary-outline",
        ).pack(side="right", padx=(6, 0))
        ttk.Button(
            botoes, text="ATUALIZAR TODOS OS ÍNDICES", command=self.indexar_todas,
            bootstyle="primary",
        ).pack(side="right", padx=6)

        self.tree_acervos = ttk.Treeview(
            fontes,
            columns=("categoria", "pasta", "documentos", "status", "ultima"),
            show="headings", height=4, selectmode="browse",
        )
        for chave, titulo in (
            ("categoria", "Tipo de acervo"), ("pasta", "Pasta"),
            ("documentos", "Documentos"), ("status", "Situação"),
            ("ultima", "Última atualização"),
        ):
            self.tree_acervos.heading(chave, text=titulo)
        self.tree_acervos.column("categoria", width=170, stretch=False)
        self.tree_acervos.column("pasta", width=600)
        self.tree_acervos.column("documentos", width=95, anchor="center", stretch=False)
        self.tree_acervos.column("status", width=160, anchor="center", stretch=False)
        self.tree_acervos.column("ultima", width=155, anchor="center", stretch=False)
        self.tree_acervos.grid(row=1, column=0, sticky="ew")

        busca = ttk.LabelFrame(
            self, text="Pesquisa no acervo local", padding=10,
            style="Section.TLabelframe",
        )
        busca.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        busca.columnconfigure(1, weight=1)
        ttk.Label(busca, text="Palavras ou expressão:").grid(row=0, column=0, padx=(0, 6))
        self.entry_busca = ttk.Entry(busca)
        self.entry_busca.grid(row=0, column=1, sticky="ew", padx=6)
        self.entry_busca.bind("<Return>", lambda _evento: self.pesquisar())
        ttk.Label(busca, text="Modo:").grid(row=0, column=2, padx=(12, 5))
        ttk.Combobox(
            busca,
            textvariable=self.modo_busca_var,
            values=MODOS_PESQUISA,
            state="readonly",
            width=18,
        ).grid(row=0, column=3, padx=5)
        ttk.Label(busca, text="Acervo:").grid(row=0, column=4, padx=(12, 5))
        self.combo_categoria_busca = ttk.Combobox(
            busca, textvariable=self.categoria_busca_var,
            values=("Todos", *self.biblioteca.listar_categorias()),
            state="readonly", width=24,
        )
        self.combo_categoria_busca.grid(row=0, column=5, padx=5)
        ttk.Button(
            busca, text="PESQUISAR", command=self.pesquisar, bootstyle="primary",
        ).grid(row=0, column=6, padx=(8, 0))

        painel = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        painel.grid(row=2, column=0, sticky="nsew")
        esquerda = ttk.Frame(painel, padding=(0, 0, 6, 0))
        direita = ttk.Frame(painel, padding=(6, 0, 0, 0))
        painel.add(esquerda, weight=2)
        painel.add(direita, weight=3)
        esquerda.rowconfigure(1, weight=1)
        esquerda.columnconfigure(0, weight=1)
        ttk.Label(esquerda, text="Resultados encontrados").grid(row=0, column=0, sticky="w")
        self.tree_resultados = ttk.Treeview(
            esquerda,
            columns=("categoria", "arquivo", "tipo", "trecho"),
            show="headings", selectmode="browse",
        )
        self.tree_resultados.heading("categoria", text="Acervo")
        self.tree_resultados.heading("arquivo", text="Arquivo")
        self.tree_resultados.heading("tipo", text="Tipo")
        self.tree_resultados.heading("trecho", text="Trecho localizado")
        self.tree_resultados.column("categoria", width=135, stretch=False)
        self.tree_resultados.column("arquivo", width=260)
        self.tree_resultados.column("tipo", width=55, anchor="center", stretch=False)
        self.tree_resultados.column("trecho", width=420)
        self.tree_resultados.grid(row=1, column=0, sticky="nsew", pady=(5, 0))
        barra_resultados = ttk.Scrollbar(
            esquerda, orient="vertical", command=self.tree_resultados.yview
        )
        barra_resultados.grid(row=1, column=1, sticky="ns", pady=(5, 0))
        self.tree_resultados.configure(yscrollcommand=barra_resultados.set)
        self.tree_resultados.bind("<<TreeviewSelect>>", self._mostrar_resultado)

        direita.rowconfigure(1, weight=1)
        direita.columnconfigure(0, weight=1)
        ttk.Label(
            direita,
            text="Texto do documento — selecione exatamente o trecho que deseja utilizar",
        ).grid(row=0, column=0, sticky="w")
        quadro_texto = ttk.Frame(direita)
        quadro_texto.grid(row=1, column=0, sticky="nsew", pady=(5, 8))
        quadro_texto.rowconfigure(0, weight=1)
        quadro_texto.columnconfigure(0, weight=1)
        self.texto = tk.Text(quadro_texto, wrap="word", undo=False, font=("Segoe UI", 10))
        self.texto.grid(row=0, column=0, sticky="nsew")
        self.texto.tag_configure(
            "resultado_busca",
            background="#ffe066",
            foreground="#111111",
        )
        barra_texto = ttk.Scrollbar(
            quadro_texto, orient="vertical", command=self.texto.yview
        )
        barra_texto.grid(row=0, column=1, sticky="ns")
        self.texto.configure(yscrollcommand=barra_texto.set)
        acoes = ttk.Frame(direita)
        acoes.grid(row=2, column=0, sticky="ew")
        ttk.Button(
            acoes, text="ABRIR DOCUMENTO", command=self.abrir_documento,
            bootstyle="info-outline",
        ).pack(side="left")
        ttk.Button(
            acoes, text="COPIAR TRECHO", command=self.copiar_trecho,
            bootstyle="secondary-outline",
        ).pack(side="left", padx=6)
        ttk.Button(
            acoes, text="INSERIR TRECHO NO WORD", command=self.inserir_trecho,
            bootstyle="success",
        ).pack(side="left", padx=6)

        ttk.Label(
            self, textvariable=self.status_var, style="AppMeta.TLabel",
            bootstyle="secondary", anchor="w",
        ).grid(row=3, column=0, sticky="ew", pady=(8, 0))

    def _id_acervo_selecionado(self) -> int | None:
        selecao = self.tree_acervos.selection()
        return int(selecao[0]) if selecao else None

    def atualizar_acervos(self) -> None:
        self._atualizar_opcoes_categorias()
        for item in self.tree_acervos.get_children():
            self.tree_acervos.delete(item)
        for acervo in self.biblioteca.listar_acervos():
            ultima = str(acervo.get("ultima_indexacao") or "").replace("T", " ")
            self.tree_acervos.insert(
                "", "end", iid=str(acervo["id"]),
                values=(
                    acervo["categoria"], acervo["caminho"], acervo["documentos"],
                    acervo["status"], ultima,
                ),
            )
        estat = self.biblioteca.estatisticas()
        self.status_var.set(
            f"{estat['acervos']} pasta(s) cadastrada(s) • "
            f"{estat['indexados']} documento(s) pesquisável(is) • "
            f"{estat['sem_texto']} sem texto • {estat['erros']} com erro"
        )

    def _atualizar_opcoes_categorias(self) -> None:
        categorias = self.biblioteca.listar_categorias()
        self.combo_categoria_busca.configure(values=("Todos", *categorias))
        if self.categoria_busca_var.get() not in {"Todos", *categorias}:
            self.categoria_busca_var.set("Todos")

    def _escolher_categoria(self, titulo: str, atual: str = "Pareceres") -> str | None:
        janela = tk.Toplevel(self)
        janela.title(titulo)
        janela.transient(self.winfo_toplevel())
        janela.grab_set()
        janela.resizable(False, False)
        resultado = {"valor": None}
        variavel = tk.StringVar(value=atual)
        ttk.Label(
            janela, text="Como esta pasta deve ser classificada?", padding=(18, 16, 18, 6)
        ).pack(anchor="w")
        ttk.Combobox(
            janela, textvariable=variavel,
            values=self.biblioteca.listar_categorias(),
            state="readonly", width=34,
        ).pack(fill="x", padx=18, pady=6)
        botoes = ttk.Frame(janela, padding=18)
        botoes.pack(fill="x")

        def confirmar():
            resultado["valor"] = variavel.get()
            janela.destroy()

        ttk.Button(botoes, text="CANCELAR", command=janela.destroy).pack(side="right")
        ttk.Button(
            botoes, text="CONFIRMAR", command=confirmar, bootstyle="success"
        ).pack(side="right", padx=8)
        janela.wait_window()
        return resultado["valor"]

    def gerenciar_categorias(self) -> None:
        janela = tk.Toplevel(self)
        janela.title("Gerenciar categorias de pastas")
        janela.transient(self.winfo_toplevel())
        janela.grab_set()
        janela.geometry("560x430")
        janela.minsize(500, 380)
        quadro = ttk.Frame(janela, padding=16)
        quadro.pack(fill="both", expand=True)
        quadro.columnconfigure(0, weight=1)
        quadro.rowconfigure(1, weight=1)
        ttk.Label(
            quadro,
            text=(
                "Crie categorias próprias para organizar seus acervos. A categoria "
                "'Outros' é reservada pelo programa."
            ),
            wraplength=510,
            justify="left",
        ).grid(row=0, column=0, sticky="ew", pady=(0, 10))
        lista = tk.Listbox(quadro, exportselection=False, font=("Segoe UI", 10))
        lista.grid(row=1, column=0, sticky="nsew")
        barra = ttk.Scrollbar(quadro, orient="vertical", command=lista.yview)
        barra.grid(row=1, column=1, sticky="ns")
        lista.configure(yscrollcommand=barra.set)
        nome_var = tk.StringVar()
        entrada = ttk.Entry(quadro, textvariable=nome_var)
        entrada.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(12, 8))
        botoes = ttk.Frame(quadro)
        botoes.grid(row=3, column=0, columnspan=2, sticky="ew")

        def atualizar_lista(selecionar=None):
            lista.delete(0, tk.END)
            categorias = self.biblioteca.listar_categorias()
            for categoria in categorias:
                lista.insert(tk.END, categoria)
            if selecionar in categorias:
                indice = categorias.index(selecionar)
                lista.selection_set(indice)
                lista.see(indice)
            self._atualizar_opcoes_categorias()
            self.atualizar_acervos()

        def categoria_selecionada():
            selecao = lista.curselection()
            return lista.get(selecao[0]) if selecao else ""

        def carregar_nome(_evento=None):
            nome_var.set(categoria_selecionada())
            entrada.icursor(tk.END)

        def adicionar():
            try:
                criada = self.biblioteca.adicionar_categoria(nome_var.get())
                nome_var.set("")
                atualizar_lista(criada)
            except Exception as erro:
                messagebox.showerror("Categorias", str(erro), parent=janela)

        def renomear():
            atual = categoria_selecionada()
            if not atual:
                messagebox.showwarning(
                    "Categorias", "Selecione a categoria que deseja renomear.",
                    parent=janela,
                )
                return
            try:
                novo = self.biblioteca.renomear_categoria(atual, nome_var.get())
                atualizar_lista(novo)
            except Exception as erro:
                messagebox.showerror("Categorias", str(erro), parent=janela)

        def excluir():
            atual = categoria_selecionada()
            if not atual:
                messagebox.showwarning(
                    "Categorias", "Selecione a categoria que deseja excluir.",
                    parent=janela,
                )
                return
            if not messagebox.askyesno(
                "Excluir categoria",
                f"Deseja excluir a categoria '{atual}'?",
                parent=janela,
            ):
                return
            try:
                self.biblioteca.remover_categoria(atual)
                nome_var.set("")
                atualizar_lista()
            except Exception as erro:
                messagebox.showerror("Categorias", str(erro), parent=janela)

        lista.bind("<<ListboxSelect>>", carregar_nome)
        ttk.Button(
            botoes, text="ADICIONAR", command=adicionar, bootstyle="success"
        ).pack(side="left")
        ttk.Button(
            botoes, text="RENOMEAR", command=renomear, bootstyle="info-outline"
        ).pack(side="left", padx=6)
        ttk.Button(
            botoes, text="EXCLUIR", command=excluir, bootstyle="danger-outline"
        ).pack(side="left")
        ttk.Button(botoes, text="FECHAR", command=janela.destroy).pack(side="right")
        atualizar_lista()
        entrada.focus_set()

    def adicionar_pasta(self) -> None:
        caminho = filedialog.askdirectory(title="Selecione a pasta do acervo local")
        if not caminho:
            return
        categoria = self._escolher_categoria("Classificar pasta")
        if not categoria:
            return
        try:
            acervo_id = self.biblioteca.adicionar_acervo(caminho, categoria)
            self.atualizar_acervos()
            self.tree_acervos.selection_set(str(acervo_id))
            self.status_var.set("Pasta cadastrada. Clique em INDEXAR SELECIONADA.")
        except Exception as erro:
            messagebox.showerror("Biblioteca Local", str(erro))

    def alterar_pasta(self) -> None:
        acervo_id = self._id_acervo_selecionado()
        if acervo_id is None:
            messagebox.showwarning("Biblioteca Local", "Selecione uma pasta cadastrada.")
            return
        atual = self.biblioteca.obter_acervo(acervo_id)
        caminho = filedialog.askdirectory(
            title="Selecione o novo local da pasta",
            initialdir=atual["caminho"] if atual and Path(atual["caminho"]).is_dir() else None,
        )
        if not caminho:
            return
        try:
            self.biblioteca.alterar_acervo(acervo_id, caminho=caminho)
            self.atualizar_acervos()
            self.status_var.set("Local alterado. Atualize o índice desta pasta.")
        except Exception as erro:
            messagebox.showerror("Biblioteca Local", str(erro))

    def alterar_categoria(self) -> None:
        acervo_id = self._id_acervo_selecionado()
        if acervo_id is None:
            messagebox.showwarning("Biblioteca Local", "Selecione uma pasta cadastrada.")
            return
        atual = self.biblioteca.obter_acervo(acervo_id)
        categoria = self._escolher_categoria(
            "Alterar categoria",
            atual["categoria"] if atual else "Outros",
        )
        if not categoria:
            return
        try:
            self.biblioteca.alterar_acervo(acervo_id, categoria=categoria)
            self.atualizar_acervos()
        except Exception as erro:
            messagebox.showerror("Biblioteca Local", str(erro))

    def remover_pasta(self) -> None:
        acervo_id = self._id_acervo_selecionado()
        if acervo_id is None:
            messagebox.showwarning("Biblioteca Local", "Selecione uma pasta cadastrada.")
            return
        if not messagebox.askyesno(
            "Remover pasta da Biblioteca",
            "A pasta e seus arquivos originais NÃO serão apagados. Somente o índice "
            "de pesquisa desta pasta será removido. Deseja continuar?",
        ):
            return
        self.biblioteca.remover_acervo(acervo_id)
        self.atualizar_acervos()

    def _apos_indexacao(self, resultado: ResultadoIndexacao) -> None:
        self.atualizar_acervos()
        self.status_var.set(resultado.resumo())
        complemento = ""
        if resultado.indisponiveis:
            complemento += f"\n\nPastas indisponíveis: {resultado.indisponiveis}."
        if resultado.sem_texto:
            complemento += (
                "\n\nPDFs sem texto selecionável foram apenas sinalizados; "
                "nenhum OCR foi executado."
            )
        messagebox.showinfo(
            "Indexação concluída", resultado.resumo() + complemento
        )

    def indexar_selecionada(self) -> None:
        acervo_id = self._id_acervo_selecionado()
        if acervo_id is None:
            messagebox.showwarning("Biblioteca Local", "Selecione uma pasta cadastrada.")
            return
        self.status_var.set("Indexando a pasta selecionada em segundo plano...")
        self.executar_tarefa(
            self.biblioteca.indexar_acervo, self._apos_indexacao, acervo_id
        )

    def indexar_todas(self) -> None:
        if not self.biblioteca.listar_acervos():
            messagebox.showwarning(
                "Biblioteca Local", "Adicione ao menos uma pasta antes de indexar."
            )
            return
        self.status_var.set("Atualizando todos os índices em segundo plano...")
        self.executar_tarefa(self.biblioteca.indexar_todos, self._apos_indexacao)

    def pesquisar(self) -> None:
        termo = self.entry_busca.get().strip()
        try:
            resultados = self.biblioteca.pesquisar(
                termo,
                categoria=self.categoria_busca_var.get(),
                modo=self.modo_busca_var.get(),
            )
        except Exception as erro:
            messagebox.showwarning("Pesquisa Local", str(erro))
            return
        for item in self.tree_resultados.get_children():
            self.tree_resultados.delete(item)
        self.documento_atual = None
        self.texto.delete("1.0", tk.END)
        for resultado in resultados:
            trecho = str(resultado["trecho"] or "").replace("\n", " ")
            self.tree_resultados.insert(
                "", "end", iid=str(resultado["id"]),
                values=(
                    resultado["categoria"], resultado["nome"],
                    resultado["extensao"].lstrip(".").upper(), trecho,
                ),
            )
        self.status_var.set(f"Pesquisa concluída: {len(resultados)} resultado(s).")

    def _mostrar_resultado(self, _evento=None) -> None:
        selecao = self.tree_resultados.selection()
        if not selecao:
            return
        self.documento_atual = self.biblioteca.obter_documento(int(selecao[0]))
        self.texto.delete("1.0", tk.END)
        if self.documento_atual:
            self.texto.insert("1.0", self.documento_atual.get("texto", ""))
            self._realcar_resultados()

    def _realcar_resultados(self) -> None:
        self.texto.tag_remove("resultado_busca", "1.0", tk.END)
        termo_original = self.entry_busca.get().strip()
        if not termo_original:
            return
        if self.modo_busca_var.get() == "Expressão exata":
            expressoes = [termo_original.strip('"').strip()]
        else:
            try:
                expressoes = self.biblioteca.termos_pesquisa(termo_original)
            except ValueError:
                return

        primeira_posicao = None

        def marcar(expressao):
            nonlocal primeira_posicao
            inicio = "1.0"
            encontrou = False
            while True:
                posicao = self.texto.search(
                    expressao,
                    inicio,
                    nocase=True,
                    stopindex=tk.END,
                )
                if not posicao:
                    break
                fim = f"{posicao}+{len(expressao)}c"
                self.texto.tag_add("resultado_busca", posicao, fim)
                primeira_posicao = primeira_posicao or posicao
                inicio = fim
                encontrou = True
            return encontrou

        algum_realce = False
        for expressao in (item for item in expressoes if item):
            algum_realce = marcar(expressao) or algum_realce

        # A busca exata do FTS ignora pontuação e quebras simples. Se a
        # expressão literal não estiver idêntica no Text, realça ao menos suas
        # palavras para que o fundamento seja localizado visualmente.
        if not algum_realce and self.modo_busca_var.get() == "Expressão exata":
            try:
                for palavra in self.biblioteca.termos_pesquisa(termo_original):
                    marcar(palavra)
            except ValueError:
                pass
        if primeira_posicao:
            self.texto.see(primeira_posicao)

    def abrir_documento(self) -> None:
        if not self.documento_atual:
            messagebox.showwarning("Biblioteca Local", "Selecione um resultado.")
            return
        caminho = self.documento_atual["caminho"]
        if not Path(caminho).is_file():
            messagebox.showerror(
                "Arquivo indisponível",
                "O arquivo foi movido ou removido. Atualize o índice da pasta.",
            )
            return
        os.startfile(caminho)

    def _trecho_selecionado(self) -> str:
        try:
            return self.texto.get(tk.SEL_FIRST, tk.SEL_LAST).strip()
        except tk.TclError:
            return ""

    def copiar_trecho(self) -> None:
        trecho = self._trecho_selecionado()
        if not trecho:
            messagebox.showwarning(
                "Selecionar trecho",
                "Selecione com o mouse o trecho exato que deseja copiar.",
            )
            return
        pyperclip.copy(trecho)
        self.status_var.set("Trecho copiado para a área de transferência.")

    def inserir_trecho(self) -> None:
        trecho = self._trecho_selecionado()
        if not trecho:
            messagebox.showwarning(
                "Selecionar trecho",
                "Selecione com o mouse o trecho exato que deseja inserir no Word.",
            )
            return
        self.inserir_no_word(trecho)
