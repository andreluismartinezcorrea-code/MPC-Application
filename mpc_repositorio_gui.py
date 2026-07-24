import tkinter as tk
from tkinter import ttk, messagebox, simpledialog



def configurar_aba_repositorio(aba_repositorio, dados_persistidos, script_dir, salvar_repositorio_func, carregar_repositorio_func):
    repositorio = carregar_repositorio_func(script_dir)

    corpo = ttk.Frame(aba_repositorio, padding=(0, 8, 0, 8))
    corpo.pack(fill="both", expand=True)
    corpo.rowconfigure(0, weight=1)
    corpo.columnconfigure(0, weight=1, minsize=250)
    corpo.columnconfigure(1, weight=3, minsize=500)

    # Painel da Esquerda: Lista de Prompts
    painel_lista = ttk.LabelFrame(corpo, text="Meus Prompts", padding=10)
    painel_lista.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
    painel_lista.rowconfigure(0, weight=1)
    painel_lista.columnconfigure(0, weight=1)

    lista_var = tk.StringVar(value=[p["nome"] for p in repositorio])
    lista_prompts = tk.Listbox(painel_lista, listvariable=lista_var, exportselection=False)
    lista_prompts.grid(row=0, column=0, sticky="nsew")
    barra_lista = ttk.Scrollbar(painel_lista, orient="vertical", command=lista_prompts.yview)
    lista_prompts.configure(yscrollcommand=barra_lista.set)
    barra_lista.grid(row=0, column=1, sticky="ns")

    botoes_lista = ttk.Frame(painel_lista)
    botoes_lista.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))

    def adicionar_prompt():
        nome = simpledialog.askstring("Novo Prompt", "Nome do novo prompt:")
        if nome:
            repositorio.append({"nome": nome, "texto": ""})
            atualizar_lista()
            lista_prompts.selection_clear(0, tk.END)
            lista_prompts.selection_set(tk.END)
            selecionar_prompt(None)
            salvar_repositorio_func(script_dir, repositorio)

    def remover_prompt():
        selecao = lista_prompts.curselection()
        if not selecao:
            return
        idx = selecao[0]
        if messagebox.askyesno("Remover", f"Deseja remover '{repositorio[idx]['nome']}'?"):
            repositorio.pop(idx)
            atualizar_lista()
            texto_editor.delete("1.0", tk.END)
            salvar_repositorio_func(script_dir, repositorio)

    ttk.Button(botoes_lista, text="Adicionar", command=adicionar_prompt, bootstyle="success").pack(side="left", expand=True, fill="x", padx=(0, 5))
    ttk.Button(botoes_lista, text="Remover", command=remover_prompt, bootstyle="danger").pack(side="left", expand=True, fill="x")

    def atualizar_lista():
        lista_prompts.delete(0, tk.END)
        for p in repositorio:
            lista_prompts.insert(tk.END, p["nome"])

    # Painel da Direita: Editor
    painel_editor = ttk.LabelFrame(corpo, text="Editor de Prompt", padding=10)
    painel_editor.grid(row=0, column=1, sticky="nsew")
    painel_editor.rowconfigure(1, weight=1)
    painel_editor.columnconfigure(0, weight=1)

    # Barra de ferramentas para inserção
    barra_ferramentas = ttk.Frame(painel_editor)
    barra_ferramentas.grid(row=0, column=0, sticky="ew", pady=(0, 8))

    def inserir_texto(texto):
        texto_editor.insert(tk.INSERT, texto)
        salvar_prompt_atual()

    btn_pdfs = ttk.Menubutton(barra_ferramentas, text="Inserir Arquivos PDF", bootstyle="info-outline")
    menu_pdfs = tk.Menu(btn_pdfs, tearoff=0)
    btn_pdfs["menu"] = menu_pdfs

    # Adicionar itens ao menu PDF dinamicamente
    # Baseado em "fontes" que estariam em dados_persistidos, mas para simplificar
    # vamos usar os apontamentos e responaveis da GUI

    btn_admins = ttk.Menubutton(barra_ferramentas, text="Inserir Administrador", bootstyle="info-outline")
    menu_admins = tk.Menu(btn_admins, tearoff=0)
    btn_admins["menu"] = menu_admins

    btn_falhas = ttk.Menubutton(barra_ferramentas, text="Inserir Falha", bootstyle="info-outline")
    menu_falhas = tk.Menu(btn_falhas, tearoff=0)
    btn_falhas["menu"] = menu_falhas

    btn_pdfs.pack(side="left", padx=(0, 5))
    btn_admins.pack(side="left", padx=(0, 5))
    btn_falhas.pack(side="left")

    texto_editor = tk.Text(painel_editor, wrap="word", undo=True, font=("Segoe UI", 10))
    barra_editor = ttk.Scrollbar(painel_editor, orient="vertical", command=texto_editor.yview)
    texto_editor.configure(yscrollcommand=barra_editor.set)
    texto_editor.grid(row=1, column=0, sticky="nsew")
    barra_editor.grid(row=1, column=1, sticky="ns")

    def salvar_prompt_atual(*_):
        selecao = lista_prompts.curselection()
        if selecao:
            idx = selecao[0]
            repositorio[idx]["texto"] = texto_editor.get("1.0", "end-1c")
            salvar_repositorio_func(script_dir, repositorio)

    texto_editor.bind("<KeyRelease>", salvar_prompt_atual)

    def selecionar_prompt(event):
        selecao = lista_prompts.curselection()
        if selecao:
            idx = selecao[0]
            texto_editor.delete("1.0", tk.END)
            texto_editor.insert(tk.END, repositorio[idx]["texto"])

    lista_prompts.bind("<<ListboxSelect>>", selecionar_prompt)

    return menu_pdfs, menu_admins, menu_falhas, texto_editor, inserir_texto, salvar_prompt_atual
