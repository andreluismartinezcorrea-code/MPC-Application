import google.generativeai as genai
import os
import PyPDF2
import re
import json
from tkinter import messagebox

try:
    api_key = os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        api_key = os.getenv("YOUR_API_KEY")
        if not api_key:
            raise ValueError("A variável de ambiente GOOGLE_API_KEY ou YOUR_API_KEY não está definida.")
    genai.configure(api_key=api_key)
except ValueError as e:
    print(f"Erro na configuração da API: {e}")
    exit()

try:
    model = genai.GenerativeModel('gemini-2.5-flash')
except Exception as e:
    print(f"Erro ao carregar o modelo: {e}")
    exit()

def obter_resposta(prompt):
    """Generates a response from the Gemini API based on a given prompt."""
    try:
        model = genai.GenerativeModel('gemini-2.5-flash')
        response = model.generate_content(prompt)
        return response.text
    except genai.APIError as e:
        return f"Erro na API Gemini: {e.message}"
    except Exception as e:
        return f"Erro genérico na geração: {e}"

def extrair_dados_rag_pdf_gemini(caminho_arquivo):
    """Extracts structured data from a PDF audit report using the Gemini API."""
    try:
        with open(caminho_arquivo, 'rb') as arquivo_pdf:
            leitor_pdf = PyPDF2.PdfReader(arquivo_pdf)
            if not leitor_pdf.pages:
                raise IndexError("O arquivo PDF está vazio.")

            texto_completo = ""
            for pagina in leitor_pdf.pages:
                texto_completo += pagina.extract_text()

        nome_arquivo = os.path.basename(caminho_arquivo)
        peca = nome_arquivo[5:13]

        prompt_restante = f"""
        Examine minuciosamente o arquivo "pdf" anexado, buscando extrair as seguintes informações do texto e retornando-as em formato JSON. Todas as strings devem ser limpas, sem caracteres especiais (exceto a máscara do Processo), asteriscos, marcadores de lista ou espaços em excesso. Se uma informação não for encontrada, o valor correspondente deve ser uma string vazia ("") ou uma lista vazia ([]) conforme o tipo de dado.

        ```json
        {{
          "Processo": "string",
          "Orgao": "string",
          "Tipo": "string",
          "ServicoAuditoria": "string",
          "Apontamentos": ["string"],
          "QuantidadeApontamentos": "integer",
          "Gestores": [
            {{
              "nome": "string",
              "cargo": "string"
            }}
          ]
        }}
        ```

        Detalhes e regras para cada campo:

        * **Processo:** O número do processo, formatado como "000000-0200/00-0". Garanta que tenha 6 dígitos antes do primeiro hífen, preenchendo com zeros à esquerda se necessário (ex: "4037-0200/19-4" deve ser "004037-0200/19-4"). Pode estar no cabeçalho ou no corpo do texto.
        * **Orgao:** O nome do órgão, sempre em MAIÚSCULA. Se o nome contiver "Prefeitura", altere para "EXECUTIVO", mantendo o restante do nome (ex: "PM de Eldorado" -> "EXECUTIVO MUNICIPAL DE ELDORADO"). Se contiver "Câmara", altere para "LEGISLATIVO", mantendo o restante do nome (ex: "CM de Eldorado" -> "LEGISLATIVO MUNICIPAL DE ELDORADO"). O órgão NÃO É "TRIBUNAL DE CONTAS DO ESTADO DO RIO GRANDE DO SUL".
        * **Tipo:** O tipo de relatório. Retorne apenas "CONTAS ORDINÁRIAS" ou "CONTAS ANUAIS". Se o documento mencionar "Contas de Governo" ou "Contas de Gestão", categorize como "CONTAS ANUAIS". É IMPORTANTE OBSERVAR A ACENTUAÇÃO CORRETA DAS PALAVRAS. Se for encontrado. Por exemplo: "CONTAS ORDINARIAS" deve ser grafado corretamente, ou seja, deve ser grafado "CONTAS ORDINÁRIAS". O mesmo para "RELATORIO SEM FALHAS", que deve observar a pontuação correta ("Relatório Sem Falhas")
        * **ServicoAuditoria:** O nome do serviço de auditoria, capitalizado (primeira letra de cada palavra em maiúscula, exceto preposições como "de", "do", "da" que devem permanecer em minúscula). Observe a acentuação e grafia correta das palavras e a capitalização.
        * **Apontamentos:** Uma lista (array) de strings com os itens numéricos (ex: "1.11", "1.34", "2.5") encontrados na última tabela do PDF (com colunas "Cargo", "Nome", "Item"). Se não houver apontamentos, ou se a auditoria sugerir que os gestores NÃO sejam intimados, retorne a string `"Relatório Sem Falhas"` para este campo, e o campo `QuantidadeApontamentos` deve ser `0`. Os itens devem ser listados como strings individuais na lista, por exemplo: `["6.4.1", "8.2.1"]`.
        * **QuantidadeApontamentos:** O número total de apontamentos encontrados.
        * **Gestores:** Uma lista (array) de objetos, onde cada objeto contém "nome" e "cargo" do gestor. Se o cargo for "Prefeito Municipal", simplifique para "Prefeito". Se o cargo for "Vice-Prefeito Municipal", simplifique para "Vice-Prefeito".

        Texto:
        ```
        {texto_completo}
        ```
        """
        resposta_gemini_text = obter_resposta(prompt_restante)
        
        try:
            if resposta_gemini_text.startswith("```json") and resposta_gemini_text.endswith("```"):
                resposta_gemini_text = resposta_gemini_text[len("```json"):-len("```")].strip()
            
            dados_extraidos = json.loads(resposta_gemini_text)

            processo = dados_extraidos.get("Processo", "")
            orgao = dados_extraidos.get("Orgao", "")
            tipo = dados_extraidos.get("Tipo", "")
            servico = dados_extraidos.get("ServicoAuditoria", "")
            apontamentos_list = dados_extraidos.get("Apontamentos", [])
            quantidade_de_apontamentos = dados_extraidos.get("QuantidadeApontamentos", 0)
            gestores_json_list = dados_extraidos.get("Gestores", [])

            if processo and processo != "Não encontrado" and re.match(r"\d+-\d{3}/\d{2}-\d", processo):
                parts = processo.split('-')
                first_part = parts[0].zfill(6)
                processo = f"{first_part}-{'-'.join(parts[1:])}"

            apontes_str = apontamentos_list
            if isinstance(apontamentos_list, list) and apontamentos_list:
                if len(apontamentos_list) > 1:
                    apontes_str = ", ".join(apontamentos_list[:-1]) + " e " + apontamentos_list[-1]
                else:
                    apontes_str = apontamentos_list[0]
            elif apontamentos_list == "Relatório Sem Falhas":
                 apontes_str = "Relatório Sem Falhas"
            else:
                 apontes_str = "Não encontrado"

            return {
                "processo": processo,
                "orgao": orgao,
                "tipo": tipo,
                "servico": servico,
                "nome_arquivo": nome_arquivo,
                "peca": peca,
                "apontes": apontes_str,
                "quantidade_de_apontamentos": str(quantidade_de_apontamentos),
                "gestores_cargos": gestores_json_list,
            }

        except json.JSONDecodeError as e:
            messagebox.showerror("Erro", f"Erro ao decodificar JSON da resposta do Gemini: {e}\\nResposta bruta: {resposta_gemini_text[:500]}...")
            return None
        except Exception as e:
            messagebox.showerror("Erro", f"Erro no processamento da resposta do Gemini: {e}\\nResposta bruta: {resposta_gemini_text[:500]}...")
            return None

    except FileNotFoundError:
        messagebox.showerror("Erro", "Arquivo não encontrado.")
        return None
    except PyPDF2.errors.PdfReadError:
        messagebox.showerror("Erro", "Erro ao ler o PDF. O arquivo pode estar corrompido ou protegido.")
        return None
    except IndexError:
        messagebox.showerror("Erro", "O arquivo PDF está vazio ou corrompido.")
        return None
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro na extração do PDF: {e}")
        return None
