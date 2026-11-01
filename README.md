# CAD Tag Vision - Motor IA Híbrido (Gemini Vision + CAD)
> Ferramenta desktop com interface gráfica integrada ao Google Gemini Vision e análise vetorial DXF para anotação e identificação técnica de componentes por setor.

![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)
![AI](https://img.shields.io/badge/AI-Google%20Gemini%20Vision-purple.svg)
![CAD](https://img.shields.io/badge/CAD-ezdxf-brightgreen.svg)
![GUI](https://img.shields.io/badge/GUI-Tkinter-yellow.svg)

---

### ⏱️ Impacto de Negócio & Horas Economizadas (Pé no chão)
- **Tempo de Catalogação:** Redução de **~8 a 10 minutos** por projeto (recorte manual, leitura visual e preenchimento de campos) para **~5 segundos** com extração multimodal via IA.
- **Economia Operacional Real:** Economia estimada de **~15 a 20 horas mensais** para quem realiza triagem técnica e catalogação de ativos.
- **Robustez com IA Multimodal:** Identifica textos manuscritos, plantas antigas ou imagens com baixa resolução onde OCRs tradicionais falham.

---

## 📌 Visão Geral
Aplicação desktop com captura de tela interativa que permite ao engenheiro/operador selecionar uma região de um projeto técnico ou ordem de serviço. A imagem recortada é enviada ao modelo multimodal **Google Gemini (Vision)** para extração semântica dos dados textuais, enquanto o motor analítico local cruza as informações com o mapa vetorial CAD (.DXF) para gerar as TAGs correspondentes ao setor operacional.

## 🚀 Funcionalidades
- **Captura Interativa com Snipping Tool nativo:** O operador recorta a área de interesse diretamente da tela.
- **Extração Semântica com IA:** O modelo Gemini identifica números de processo, equipamentos e descrições técnicas mesmo em imagens ruidosas ou com baixa resolução.
- **Interseção Espacial com CAD:** O motor lê o arquivo `.dxf` e localiza automaticamente o setor correspondente.
- **Autenticação Segura:** Chave de API gerenciada através de variável de ambiente (`GEMINI_API_KEY`).

## 🛠️ Instalação e Configuração

### 1. Dependências
```bash
pip install google-generativeai ezdxf openpyxl pillow
```

### 2. Configurar a Chave da API Gemini
Obtenha uma chave gratuita no [Google AI Studio](https://aistudio.google.com/) e configure no seu ambiente:

No Windows (PowerShell):
```powershell
$env:GEMINI_API_KEY = "sua_chave_real_aqui"
```

No Linux / macOS:
```bash
export GEMINI_API_KEY="sua_chave_real_aqui"
```

### 3. Executar a Aplicação
```bash
python leitor_autocad.py
```
