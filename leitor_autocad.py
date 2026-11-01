# -*- coding: utf-8 -*-
import os
import sys
import re
import math
import json
import unicodedata
from difflib import get_close_matches
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import warnings

import ezdxf
import openpyxl
from openpyxl.styles import Font
from PIL import ImageGrab, Image
import google.generativeai as genai

# =========================================================================
# CONFIGURAÇÃO DA IA E AMBIENTE
# =========================================================================
warnings.filterwarnings("ignore")
os.environ["GRPC_PYTHON_LOG_LEVEL"] = "error"

# SUA CHAVE DE API (Gemini 1.5 Flash - Free Tier)
# ==============================================================================
# CONFIGURACAO DE AUTENTICACAO - API GEMINI
# ATENCAO: A CHAVE ABAIXO EH UM TOKEN FICTICIO / INVENTADO PARA DEMONSTRACAO!
# Para executar, obtenha uma chave em https://aistudio.google.com/ e configure
# a variavel de ambiente GEMINI_API_KEY ou substitua a variavel abaixo:
# ==============================================================================
CHAVE_INVENTADA_DEMO = "AIzaSyFakeKey_INVENTADO_SUBSTITUA_PELA_SUA_CHAVE_AQUI"
API_KEY = os.environ.get("GEMINI_API_KEY", CHAVE_INVENTADA_DEMO)
genai.configure(api_key=API_KEY)



BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CAMINHO_EXCEL_PADRAO = os.path.join(BASE_DIR, "exemplos", "tabela_tags_referencia.xlsx")
CAMINHO_DXF_PADRAO = os.path.join(BASE_DIR, "exemplos", "planta_rede_exemplo.dxf")
TEMP_FOLDER = os.path.join(BASE_DIR, "temp_processamento")

os.makedirs(TEMP_FOLDER, exist_ok=True)
JSON_PATH = os.path.join(TEMP_FOLDER, "dados_email.json")
JSON_SAIDA_PATH = os.path.join(TEMP_FOLDER, "tags_geradas.json")

LAYERS_IGNORADAS = ["CARIMBO", "FORMATO", "FOLHA", "COTA", "DEFPOINT", "EIXO"]

# =========================================================================
# FUNÇÕES CORE (LIMPEZA E CAD)
# =========================================================================
def normalizar_texto(texto):
    txt = str(texto).strip().upper()
    txt = unicodedata.normalize("NFKD", txt).encode("ASCII", "ignore").decode("ASCII")
    txt = re.sub(r"[^A-Z0-9\s]", " ", txt)
    txt = re.sub(r"\s+", " ", txt).strip()
    txt = re.sub(r"^(RUA|AVENIDA|AV|R|TRAVESSA|ALAMEDA|TV|RODOVIA|ESTRADA|PCA|PRACA)\s+", "", txt)
    return txt.strip()

def classificar_texto(texto):
    txt_limpo = str(texto).strip().upper()
    if not txt_limpo: return None
    if re.search(r"^(RUA|AV|AVENIDA|R\.|TRAVESSA|ALAMEDA|TV\.|TV|RODOVIA|ESTRADA)\b", txt_limpo): return "RUA"
    if len(txt_limpo) > 3: return "RUA"
    return None

def limpar_nome_cliente(txt):
    txt = unicodedata.normalize('NFKD', txt).encode('ASCII', 'ignore').decode('ASCII')
    txt = re.sub(r'(?i)\bcliente\s*:\s*', '', txt)
    return txt.upper().strip()

def get_todas_entidades(msp):
    for e in msp:
        if e.dxftype() == 'INSERT':
            try:
                for sub_e in e.virtual_entities(): yield sub_e
            except Exception: pass
        else: yield e

def carregar_dados_dxf(caminho_dxf):
    doc = ezdxf.readfile(caminho_dxf)
    ruas, codigos = [], []
    for e in get_todas_entidades(doc.modelspace()):
        tipo = e.dxftype()
        if tipo not in ("TEXT", "MTEXT", "ARCALIGNEDTEXT"): continue
        try: layer = str(e.dxf.layer).upper()
        except Exception: continue
        if any(ign in layer for ign in LAYERS_IGNORADAS): continue
        try:
            texto = (e.plain_text() if tipo == "MTEXT" else e.dxf.text).strip()
            if not texto: continue
            ponto = e.dxf.insert
            if "SETORIZA" in layer:
                codigos.append((texto, (ponto[0], ponto[1])))
            elif classificar_texto(texto) == "RUA":
                ruas.append((texto, normalizar_texto(texto), (ponto[0], ponto[1])))
        except Exception: pass
    return ruas, codigos

def buscar_rua_no_dxf(nome_digitado, ruas):
    alvo = normalizar_texto(nome_digitado)
    if len(alvo) < 2: return []
    exatos = [r for r in ruas if r[1] == alvo]
    if exatos: return exatos
    contendo = [r for r in ruas if alvo in r[1]]
    if contendo: return contendo
    palavras = alvo.split()
    match_palavras = [r for r in ruas if all(p in r[1] for p in palavras)]
    if match_palavras: return match_palavras
    nomes_unicos = list(set(r[1] for r in ruas))
    proximos = get_close_matches(alvo, nomes_unicos, n=5, cutoff=0.5)
    if proximos: return [r for r in ruas if r[1] in proximos]
    return []

def achar_setor_cercado(listas_de_ruas, codigos):
    if not codigos or not listas_de_ruas: return "01"
    melhor_setor = "01"
    menor_distancia_maxima = float('inf')
    for texto_codigo, (cx, cy) in codigos:
        try:
            distancias_ate_ruas = [min(math.hypot(cx - px, cy - py) for _, _, (px, py) in p_rua) for p_rua in listas_de_ruas]
            pontuacao = max(distancias_ate_ruas) + (sum(distancias_ate_ruas) * 0.05)
            if pontuacao < menor_distancia_maxima:
                menor_distancia_maxima = pontuacao
                melhor_setor = texto_codigo
        except Exception: pass
    return melhor_setor

def achar_setor_dxf_simples(codigos_dxf):
    setor_encontrado = "01"
    for texto_codigo, (cx, cy) in codigos_dxf:
        if texto_codigo.isdigit():
            setor_encontrado = texto_codigo.zfill(2)
            break
    return setor_encontrado

# =========================================================================
# FUNÇÕES DE GERAÇÃO E ATUALIZAÇÃO DO EXCEL
# =========================================================================
def gerar_sigla_rua_nova(wb, nome_rua):
    aba_ruas = wb["TAG's para Ruas"]
    siglas_usadas = set()
    for row in aba_ruas.iter_rows(values_only=True):
        if len(row) >= 3 and row[1] is not None:
            siglas_usadas.add(str(row[1]).strip().upper())
            
    limpo = limpar_nome_cliente(nome_rua)
    palavras = [p for p in re.split(r'\W+', limpo) if p not in ['E', 'DE', 'DA', 'DO', 'DI', 'DAS', 'DOS', 'RUA', 'AV', 'AVENIDA']]
    sigla_def = "GEN"
    if palavras:
        candidato = "".join([p[0] for p in palavras[:3]]).upper()
        if len(candidato) >= 3 and candidato not in siglas_usadas:
            sigla_def = candidato[:3]
        else:
            for i in range(1, 100):
                test = f"R{i:02d}"
                if test not in siglas_usadas:
                    sigla_def = test
                    break
    return sigla_def

def gerar_opcoes_de_sigla(nome):
    limpo = limpar_nome_cliente(nome)
    palavras = [p for p in re.split(r'\W+', limpo) if p not in ['E', 'DE', 'DA', 'DO', 'DI', 'DAS', 'DOS', 'JR', 'JUNIOR', 'RESIDENCIAL', 'CONDOMINIO', 'PIZZARIA', 'LTDA', 'ME']]
    siglas = []
    if len(palavras) >= 3: siglas.append(palavras[0][0] + palavras[1][0] + palavras[2][0])
    elif len(palavras) == 2: siglas.append((palavras[0][:2] + palavras[1][0]).upper())
    elif len(palavras) == 1: siglas.append(palavras[0][:3])
    
    if not siglas or len(siglas[0]) < 3:
        siglas.extend(["ART", "ATH"])
    
    extras = ["FOP", "ART", "ATH", "ABC", "XYZ", "PRD", "SOL", "TOP", "BEM"]
    for ex in extras:
        if ex not in siglas: siglas.append(ex)
    return [s[:3].upper() for s in siglas if len(s) >= 3]

def atualizar_excel_e_gerar_tags(caminho_excel, nome_cliente, endereco_principal, setor_encontrado):
    wb = openpyxl.load_workbook(caminho_excel)
    
    if not endereco_principal:
        endereco_principal = "GENERICO"
        
    # 1. Encontra ou cria a sigla da rua
    sigla_rua = "XXX"
    is_nova_rua = False
    if "TAG's para Ruas" in wb.sheetnames:
        aba_ruas = wb["TAG's para Ruas"]
        for row in aba_ruas.iter_rows(values_only=True):
            if row and len(row) >= 3 and row[2]:
                if normalizar_texto(endereco_principal) in normalizar_texto(str(row[2])):
                    sigla_rua = str(row[1]).strip()
                    break
                    
        if sigla_rua == "XXX":
            sigla_rua = gerar_sigla_rua_nova(wb, endereco_principal)
            is_nova_rua = True
            
            # Insere a nova rua na planilha
            proxima_linha = aba_ruas.max_row + 1
            while proxima_linha > 1 and aba_ruas.cell(row=proxima_linha-1, column=2).value is None:
                proxima_linha -= 1
            aba_ruas.cell(row=proxima_linha, column=2, value=sigla_rua)
            aba_ruas.cell(row=proxima_linha, column=3, value=endereco_principal.upper())
            celula_aviso = aba_ruas.cell(row=proxima_linha, column=4, value="CRIADA AUTOMATICAMENTE")
            celula_aviso.font = Font(name="Segoe UI", size=10, bold=True, color="FF0000")

    # 2. Gera TAG de Calçada
    setor_formatado = str(setor_encontrado).zfill(3)
    prefixo_busca = f"C{sigla_rua}-{setor_formatado}-"
    maior_sequencial = 0
    for nome_aba in wb.sheetnames:
        for row in wb[nome_aba].iter_rows(values_only=True):
            for cell in row:
                if isinstance(cell, str) and cell.startswith(prefixo_busca):
                    try:
                        seq = int(cell.split('-')[2])
                        if seq > maior_sequencial: maior_sequencial = seq
                    except: pass 
    tag_calcada = f"{prefixo_busca}{str(maior_sequencial + 1).zfill(2)}"

    # 3. Gera TAG de Estação
    siglas_usadas = set()
    if "Estações PEQ PEAD" in wb.sheetnames:
        for row in wb["Estações PEQ PEAD"].iter_rows(values_only=True):
            for cell in row:
                if isinstance(cell, str):
                    match = re.search(r'SISTEMA_INTERNO_2E-([A-Z0-9]{3})-001', cell)
                    if match: siglas_usadas.add(match.group(1))
    
    opcoes = gerar_opcoes_de_sigla(nome_cliente)
    sigla_est = next((op for op in opcoes if op not in siglas_usadas), opcoes[0])
    tag_estacao = f"SISTEMA_INTERNO_2E-{sigla_est}-001-FE01"

    # 4. Insere Calçada na aba Principal
    aba_principal = wb.active
    prox_princ = aba_principal.max_row + 1
    while prox_princ > 1 and aba_principal.cell(row=prox_princ-1, column=1).value is None: prox_princ -= 1
    aba_principal.cell(row=prox_princ, column=1, value=nome_cliente.upper())
    aba_principal.cell(row=prox_princ, column=2, value=tag_calcada)

    # 5. Insere Estação na aba PEQ PEAD
    if "Estações PEQ PEAD" in wb.sheetnames:
        aba_est = wb["Estações PEQ PEAD"]
        prox_est = aba_est.max_row + 1
        while prox_est > 1 and aba_est.cell(row=prox_est-1, column=1).value is None: prox_est -= 1
        aba_est.cell(row=prox_est, column=1, value=nome_cliente.upper())
        aba_est.cell(row=prox_est, column=2, value=tag_estacao)

    wb.save(caminho_excel)
    return tag_calcada, tag_estacao, sigla_rua, is_nova_rua

# =========================================================================
# MODO AUTOMÁTICO (RODA PELO POWER AUTOMATE SEM ABRIR JANELA)
# =========================================================================
def executar_modo_automatico():
    try:
        if not os.path.exists(JSON_PATH):
            sys.exit(0)
        
        with open(JSON_PATH, "r", encoding="utf-8-sig") as f:
            dados = json.loads(f.read(), strict=False)
            
        nome_cliente = limpar_nome_cliente(dados.get("cliente", ""))
        endereco_email = dados.get("endereco", "").strip()
        
        # Puxa o setor do mapa
        _, codigos_dxf = carregar_dados_dxf(CAMINHO_DXF_PADRAO)
        setor_encontrado = achar_setor_dxf_simples(codigos_dxf)

        # Atualiza a planilha, gera tags e avisa se rua é nova
        tag_calcada, tag_estacao, sigla_rua, is_nova_rua = atualizar_excel_e_gerar_tags(
            CAMINHO_EXCEL_PADRAO, nome_cliente, endereco_email, setor_encontrado
        )
        
        dados_saida = {
            "cliente": nome_cliente,
            "setor": str(setor_encontrado),
            "tag_estacao": tag_estacao,
            "tag_calcada": tag_calcada
        }
        
        # Salva o JSON final para o Power Automate montar o email
        with open(JSON_SAIDA_PATH, "w", encoding="utf-8") as f:
            json.dump(dados_saida, f, ensure_ascii=False, indent=4)
            
        sys.exit(0)
    except Exception as e:
        import traceback
        sys.stderr.write(f"ERRO REAL NO PYTHON:\n{str(e)}\n")
        sys.stderr.write(traceback.format_exc())
        sys.exit(1)

# =========================================================================
# FERRAMENTA DE CAPTURA DE TELA (GUI)
# =========================================================================
class SnippingWidget:
    def __init__(self, parent):
        self.top = tk.Toplevel(parent)
        self.top.attributes('-fullscreen', True)
        self.top.attributes('-alpha', 0.3)
        self.top.configure(cursor="cross")
        self.top.attributes('-topmost', True)
        self.canvas = tk.Canvas(self.top, bg="black", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.start_x = None
        self.start_y = None
        self.rect = None
        self.bbox = None
        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)

    def on_press(self, event):
        self.start_x = event.x
        self.start_y = event.y
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, 1, 1, outline='red', width=3, fill="white")

    def on_drag(self, event):
        self.canvas.coords(self.rect, self.start_x, self.start_y, event.x, event.y)

    def on_release(self, event):
        self.bbox = (min(self.start_x, event.x), min(self.start_y, event.y), max(self.start_x, event.x), max(self.start_y, event.y))
        self.top.destroy()

# =========================================================================
# INTERFACE GRÁFICA (GUI) E IA GEMINI
# =========================================================================
class AppTagGeneratorVision:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador Inteligente de TAGs CAD (Gemini Vision & DXF)")
        self.root.geometry("700x800")
        self.root.configure(padx=20, pady=20)
        
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TButton', font=('Segoe UI', 10, 'bold'), padding=5)
        style.configure('Header.TLabel', font=('Segoe UI', 11, 'bold'))

        self.dxf_cache_path = ""
        self.ruas_dxf = []
        self.codigos_dxf = []

        ttk.Label(root, text="1. Caminho dos Arquivos", style='Header.TLabel').pack(anchor='w', pady=(0,5))
        
        frame_dxf = ttk.Frame(root)
        frame_dxf.pack(fill='x', pady=2)
        ttk.Label(frame_dxf, text="Arquivo DXF:", width=15).pack(side='left')
        self.entry_dxf = ttk.Entry(frame_dxf, width=55)
        self.entry_dxf.insert(0, CAMINHO_DXF_PADRAO)
        self.entry_dxf.pack(side='left', padx=5)
        ttk.Button(frame_dxf, text="Procurar", command=self.buscar_dxf).pack(side='left')

        frame_excel = ttk.Frame(root)
        frame_excel.pack(fill='x', pady=2)
        ttk.Label(frame_excel, text="Planilha Excel:", width=15).pack(side='left')
        self.entry_excel = ttk.Entry(frame_excel, width=55)
        self.entry_excel.insert(0, CAMINHO_EXCEL_PADRAO)
        self.entry_excel.pack(side='left', padx=5)
        ttk.Button(frame_excel, text="Procurar", command=self.buscar_excel).pack(side='left')

        ttk.Label(root, text="2. Dados da Permissão de Trabalho (PT)", style='Header.TLabel').pack(anchor='w', pady=(20,5))
        
        frame_cliente = ttk.Frame(root)
        frame_cliente.pack(fill='x', pady=(0, 10))
        ttk.Label(frame_cliente, text="Nome do Cliente:", width=15, foreground="green").pack(side='left')
        self.entry_cliente = ttk.Entry(frame_cliente, width=55)
        self.entry_cliente.pack(side='left', padx=5)

        frame_scanner = ttk.Frame(root)
        frame_scanner.pack(fill='x', pady=(5, 10))
        ttk.Label(frame_scanner, text="Extração com IA Gemini:").pack(side='left')
        ttk.Button(frame_scanner, text="📸 Recortar PT + Mapa e Extrair", command=self.iniciar_captura).pack(side='left', padx=10)

        ttk.Label(root, text="A primeira rua preenchida será usada para gerar a Sigla principal da Calçada.", foreground="blue").pack(anchor='w', pady=(0,5))
        
        self.entries_ruas = []
        for i in range(1, 6):
            frame_r = ttk.Frame(root)
            frame_r.pack(fill='x', pady=2)
            ttk.Label(frame_r, text=f"{i}ª Rua:", width=15).pack(side='left')
            entry_r = ttk.Entry(frame_r, width=55)
            entry_r.pack(side='left', padx=5)
            self.entries_ruas.append(entry_r)

        self.btn_processar = ttk.Button(root, text="🚀 Processar, Gerar TAGs e Salvar no Excel", command=self.processar)
        self.btn_processar.pack(pady=20, fill='x')

        ttk.Label(root, text="3. Resultado Final", style='Header.TLabel').pack(anchor='w', pady=(0,5))
        self.text_resultado = tk.Text(root, height=13, width=75, font=('Consolas', 11), bg="#f4f4f4")
        self.text_resultado.pack(fill='x')

    def buscar_dxf(self):
        c = filedialog.askopenfilename(filetypes=[("DXF", "*.dxf")])
        if c: self.entry_dxf.delete(0, tk.END); self.entry_dxf.insert(0, c)

    def buscar_excel(self):
        c = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if c: self.entry_excel.delete(0, tk.END); self.entry_excel.insert(0, c)

    def print_msg(self, msg):
        self.text_resultado.delete(1.0, tk.END)
        self.text_resultado.insert(tk.END, msg)
        self.root.update()

    def carregar_dxf_memoria(self, caminho):
        if self.dxf_cache_path == caminho: return True
        try:
            self.print_msg("Carregando dicionário DXF na memória...")
            self.ruas_dxf, self.codigos_dxf = carregar_dados_dxf(caminho)
            self.dxf_cache_path = caminho
            return True
        except Exception as e:
            self.print_msg(f"Erro ao carregar DXF: {e}")
            return False

    def iniciar_captura(self):
        self.root.withdraw() 
        snipper = SnippingWidget(self.root)
        self.root.wait_window(snipper.top) 
        self.root.deiconify() 

        if snipper.bbox:
            x1, y1, x2, y2 = snipper.bbox
            if abs(x2-x1) > 10 and abs(y2-y1) > 10:
                img = ImageGrab.grab(bbox=(x1, y1, x2, y2))
                self.enviar_para_gemini(img)

    def enviar_para_gemini(self, img):
        self.print_msg("🧠 Enviando imagem para o Agente Gemini 1.5 Flash... Aguarde.")
        self.root.update()
        
        try:
            # Usando '-latest' para blindar o erro 404
            model = genai.GenerativeModel('gemini-1.5-flash-latest')
            prompt = """
            Você é um engenheiro sênior extraindo dados cirurgicamente de uma imagem de projeto.
            A imagem contém uma tabela com os dados da obra e possivelmente um mapa com ruas.
            1. Extraia o NOME DO CLIENTE.
            2. Extraia o ENDEREÇO DA OBRA (Rua principal, sem número).
            3. Se houver nomes de outras ruas legíveis ao redor no mapa, liste-as.
            
            Retorne APENAS um JSON válido, sem comentários, no formato exato:
            {
                "cliente": "Nome Extracted",
                "endereco_principal": "Rua Principal",
                "ruas_vizinhas": ["Rua A", "Rua B"]
            }
            """
            response = model.generate_content([prompt, img])
            
            # Limpeza do JSON que a IA devolve
            texto_limpo = response.text.replace('```json', '').replace('```', '').strip()
            dados = json.loads(texto_limpo)
            
            # Preenche Interface
            self.entry_cliente.delete(0, tk.END)
            self.entry_cliente.insert(0, dados.get('cliente', '').upper())
            
            for entry in self.entries_ruas: entry.delete(0, tk.END)
            
            idx = 0
            end_princ = dados.get('endereco_principal', '')
            if end_princ:
                self.entries_ruas[idx].insert(0, end_princ.upper())
                idx += 1
                
            for vizinha in dados.get('ruas_vizinhas', []):
                if idx < 5:
                    self.entries_ruas[idx].insert(0, vizinha.upper())
                    idx += 1
                    
            self.print_msg("✅ IA Gemini extraiu os dados com sucesso! Verifique os campos e clique em Processar.")
            
        except Exception as e:
            self.print_msg(f"❌ Erro na extração da IA: {e}\n(Verifique sua conexão de internet e certifique-se de que rodou pip install --upgrade google-generativeai)")

    def processar(self):
        caminho_dxf = self.entry_dxf.get().strip()
        caminho_excel = self.entry_excel.get().strip()
        nome_cliente = self.entry_cliente.get().strip()
        nomes_ruas = [e.get().strip() for e in self.entries_ruas if e.get().strip()]

        if not caminho_dxf or not caminho_excel:
            messagebox.showwarning("Atenção", "Por favor, selecione os caminhos do DXF e do Excel.")
            return

        self.btn_processar.config(state='disabled', text="Calculando Geometria e Lendo Planilhas...")
        self.print_msg("Processando dados...")
        
        try:
            if not self.carregar_dxf_memoria(caminho_dxf): return
                
            ruas_encontradas = []
            for nome in nomes_ruas:
                enc = buscar_rua_no_dxf(nome, self.ruas_dxf)
                if enc: ruas_encontradas.append(enc)
                
            setor_encontrado = achar_setor_cercado(ruas_encontradas, self.codigos_dxf)
            endereco_principal = nomes_ruas[0] if nomes_ruas else ""
            
            # CHAMA A FUNÇÃO QUE CRIA AS TAGS E GRAVA NO EXCEL AO MESMO TEMPO
            tag_calcada, tag_estacao, sigla_rua, is_nova = atualizar_excel_e_gerar_tags(
                caminho_excel, nome_cliente, endereco_principal, setor_encontrado
            )
            
            aviso_sigla = ""
            if is_nova:
                aviso_sigla = f"\n\n[!] AVISO: A rua '{endereco_principal}' NÃO EXISTIA.\nA sigla '{sigla_rua}' foi gerada e registrada em VERMELHO na planilha!"
            
            resultado = (
                f"🗺️ SETOR IDENTIFICADO: {setor_encontrado}\n"
                f"{'-'*60}\n"
                f"🟩 TAG DE ESTAÇÃO: {tag_estacao}\n"
                f"🟦 TAG DE CALÇADA: {tag_calcada}\n"
                f"{aviso_sigla}\n\n"
                f"✅ Tudo foi salvo e atualizado diretamente na sua planilha Excel!"
            )
            self.print_msg(resultado)
            
        except Exception as e:
            import traceback
            self.print_msg(f"❌ ERRO NO PROCESSO:\n{str(e)}\n{traceback.format_exc()}")
        finally:
            self.btn_processar.config(state='normal', text="🚀 Processar, Gerar TAGs e Salvar no Excel")

if __name__ == "__main__":
    # Remove o arquivo temporário se ele existir para garantir que a GUI abra quando você testar manualmente
    if '--auto' in sys.argv or os.path.exists(JSON_PATH):
        executar_modo_automatico()
    else:
        root = tk.Tk()
        app = AppTagGeneratorVision(root)
        root.mainloop()


