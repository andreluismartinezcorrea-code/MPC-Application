import json
from pathlib import Path

# Caminho do arquivo onde salvaremos os prompts
REPOSITORIO_PROMPTS_FILE = "repositorio_prompts.json"

PROMPTS_PADRAO = [
    {
        "nome": "Minuta Final (Concisa)",
        "texto": "Vamos elaborar a minuta final do parecer de maneira concisa e objetiva. É essencial que concentremos nossa atenção nos argumentos mais relevantes, garantindo que cada ponto significativo seja apresentado de forma clara e direta. Utilize uma escrita fluida, evitando enumerações, e assegure-se de que a estrutura do texto seja lógica, com parágrafos bem organizados que conduzam o leitor de maneira eficaz. Ao final, revise o texto para garantir a coerência e a clareza das ideias apresentadas."
    }
]



def carregar_repositorio(pasta_inicial):
    caminho = Path(pasta_inicial) / REPOSITORIO_PROMPTS_FILE
    if not caminho.exists():
        return PROMPTS_PADRAO.copy()
    try:
        with open(caminho, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return PROMPTS_PADRAO.copy()



def salvar_repositorio(pasta_inicial, repositorio):
    caminho = Path(pasta_inicial) / REPOSITORIO_PROMPTS_FILE
    try:
        with open(caminho, "w", encoding="utf-8") as f:
            json.dump(repositorio, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"Erro ao salvar repositório: {e}")
