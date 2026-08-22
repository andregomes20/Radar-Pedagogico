import os
import random
import json
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import customtkinter as ctk
import unicodedata

# Importar os gráficos do módulo graficos.py
import graficos
import relatorios_pdf
import tkinter.simpledialog as simpledialog

def resource_path(relative_path):
    try:
        import sys
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# Configurações do CustomTkinter
ctk.set_appearance_mode("Dark")  # Modos: "System", "Dark", "Light"
ctk.set_default_color_theme("blue")  # Temas: "blue", "green", "dark-blue"

# Lista padrão de disciplinas do Técnico em DS (conforme imagens anexadas)
SUBJECT_LIST = [
    "BIOLOGIA", 
    "CARREIRA E COMPETENCIAS", 
    "EDUCACAO FISICA", 
    "FISICA", 
    "GEOGRAFIA", 
    "HISTORIA", 
    "LINGUA INGLESA", 
    "LINGUA PORTUGUESA", 
    "LOGICA E LINGUAGEM DE PROGRAMACAO", 
    "MATEMATICA", 
    "PROCESSOS DE DESENVOLVIMENTO", 
    "QUIMICA", 
    "REDES", 
    "SOCIOLOGIA"
]

def normalizar_texto(texto):
    """Remove acentos, converte para maiúsculas e remove espaços sobressalentes."""
    if pd.isna(texto) or not isinstance(texto, str):
        return ""
    texto = texto.strip().upper()
    return "".join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

def normalizar_nome_turma(raw_name):
    if not raw_name:
        return ""
    import re
    norm = normalizar_texto(raw_name).strip()
    match = re.search(r'([1-3])\s*[º°ªa]?\s*SERIE\s*([A-E])', norm, re.IGNORECASE)
    if match:
        return f"{match.group(1)}º{match.group(2).upper()}"
    match2 = re.search(r'([1-3])\s*[º°ªa]?\s*([A-E])', norm, re.IGNORECASE)
    if match2:
        return f"{match2.group(1)}º{match2.group(2).upper()}"
    return norm

class SpreadsheetParser:
    """Classe responsável pelo parser inteligente e dinâmico das planilhas de bimestres."""
    
    @staticmethod
    def carregar_arquivo(caminho_arquivo):
        """Detecta o formato (CSV ou Excel) e carrega como DataFrame pandas."""
        ext = os.path.splitext(caminho_arquivo)[1].lower()
        if ext in ['.xlsx', '.xls']:
            return pd.read_excel(caminho_arquivo, header=None)
        elif ext == '.csv':
            try:
                return pd.read_csv(caminho_arquivo, header=None, sep=';', encoding='utf-8')
            except Exception:
                try:
                    return pd.read_csv(caminho_arquivo, header=None, sep=',', encoding='utf-8')
                except Exception:
                    return pd.read_csv(caminho_arquivo, header=None, sep=None, engine='python', encoding='latin1')
        else:
            raise ValueError("Formato de arquivo nao suportado. Use .xlsx, .xls ou .csv")

    @staticmethod
    def analisar_dados(df):
        """
        Analisa o DataFrame e extrai a turma, o bimestre, a lista de disciplinas
        e a lista de alunos com suas notas, faltas, frequência e planos de ação.
        """
        # 1. Detectar Turma (varrer primeiras 10 linhas)
        turma_detectada = ""
        import re
        turma_pattern = re.compile(r'([1-3]\s*[º°]?\s*[A-E]\s*-\s*DS)', re.IGNORECASE)
        for r in range(min(10, len(df))):
            for c in range(len(df.columns)):
                val = str(df.iloc[r, c]).strip()
                if "TURMA:" in val.upper() or "SALA:" in val.upper() or "CLASSE:" in val.upper():
                    if ":" in val:
                        partes = val.split(":", 1)
                        turma_detectada = partes[1].strip()
                    elif c + 1 < len(df.columns):
                        turma_detectada = str(df.iloc[r, c+1]).strip()
                    break
                
                # Alternativa: verificar se a célula em si contém um padrão de turma (Ex: "2ºC - DS")
                match = turma_pattern.search(val)
                if match:
                    turma_detectada = match.group(1).upper()
                    break
            if turma_detectada:
                break
                
        # 2. Detectar Bimestre (varrer primeiras 10 linhas)
        bimestre_detectado = None
        for r in range(min(10, len(df))):
            for c in range(len(df.columns)):
                val = str(df.iloc[r, c]).strip().upper()
                if "FECHAMENTO" in val or "BIMESTRE" in val or "CONSELHO" in val:
                    if "PRIMEIRO" in val or "1B" in val or "1º" in val or "1O" in val:
                        bimestre_detectado = 1
                    elif "SEGUNDO" in val or "2B" in val or "2º" in val or "2O" in val:
                        bimestre_detectado = 2
                    elif "TERCEIRO" in val or "3B" in val or "3º" in val or "3O" in val:
                        bimestre_detectado = 3
                    elif "QUARTO" in val or "4B" in val or "4º" in val or "4O" in val:
                        bimestre_detectado = 4
                    break
            if bimestre_detectado:
                break

        # 3. Encontrar a linha de cabeçalho (que contém "ALUNO")
        r_header = None
        for idx, row in df.iterrows():
            row_str = [normalizar_texto(val) for val in row]
            if "ALUNO" in row_str:
                r_header = idx
                break
                
        if r_header is None:
            r_header = 0
            
        header_vals = [normalizar_texto(val) for val in df.iloc[r_header]]
        
        # 4. Identificar índices das colunas importantes
        col_aluno = header_vals.index("ALUNO") if "ALUNO" in header_vals else 0
        
        # Identificar se a coluna SITUAÇÃO realmente existe
        if "SITUACAO" in header_vals:
            col_situacao = header_vals.index("SITUACAO")
            start_col = col_situacao + 1
        elif col_aluno + 1 < len(df.columns) and "SITUA" in str(df.iloc[r_header, col_aluno + 1]).upper():
            col_situacao = col_aluno + 1
            start_col = col_situacao + 1
        else:
            col_situacao = None
            start_col = col_aluno + 1
        
        # Encontrar colunas de totais
        col_tf = None
        col_fre = None
        col_ft_ano = None
        col_fre_ano = None
        col_media = None
        
        for idx, h_val in enumerate(header_vals):
            if h_val == "TF" or ("TOTAL" in h_val and "FALTA" in h_val):
                col_tf = idx
            elif "FRE(%)" in h_val or "FRE (%)" in h_val or "FREQUENCIA" in h_val:
                if col_fre is None: col_fre = idx
            elif "FT ANO" in h_val:
                col_ft_ano = idx
            elif "FRE ANO" in h_val:
                col_fre_ano = idx
            elif "MEDIA" in h_val or "MÉDIA" in h_val:
                col_media = idx

        # Se as sub-colunas de TOTAL estão na linha abaixo (cabeçalho de duas linhas)
        if r_header + 1 < len(df):
            sub_header = [normalizar_texto(val) for val in df.iloc[r_header + 1]]
            for idx, s_val in enumerate(sub_header):
                if s_val == "TF":
                    col_tf = idx
                elif "FRE" in s_val:
                    if "%" in s_val or "PERCENT" in s_val or col_fre is None:
                        if idx > 15 and col_fre is not None:
                            col_fre_ano = idx
                        else:
                            col_fre = idx
                elif "FT" in s_val:
                    col_ft_ano = idx
                elif "ANO" in s_val:
                    col_fre_ano = idx
                elif "MEDIA" in s_val:
                    col_media = idx

        # Fallbacks caso não tenha encontrado pelo cabeçalho
        if col_tf is None:
            for idx, h_val in enumerate(header_vals):
                if "TOTAL" in h_val or h_val == "TF":
                    col_tf = idx
                    break
            if col_tf is None:
                col_tf = len(df.columns) - 4 if len(df.columns) > 4 else len(df.columns) - 1
                
        if col_fre is None: col_fre = col_tf + 1 if col_tf + 1 < len(df.columns) else None
        if col_ft_ano is None: col_ft_ano = col_fre + 1 if col_fre and col_fre + 1 < len(df.columns) else None
        if col_fre_ano is None: col_fre_ano = col_ft_ano + 1 if col_ft_ano and col_ft_ano + 1 < len(df.columns) else None
        if col_media is None:
            for idx, h_val in enumerate(header_vals):
                if "MEDIA" in h_val or "MÉDIA" in h_val:
                    col_media = idx
                    break
            if col_media is None:
                col_media = len(df.columns) - 1

        # 5. Mapeamento DINÂMICO de disciplinas
        col_subjects = {}
        subjects_list = []
        
        end_col = col_tf if col_tf is not None else len(df.columns)
        
        for c in range(start_col, end_col):
            h_val = str(df.iloc[r_header, c]).strip()
            if pd.isna(df.iloc[r_header, c]) or h_val == "" or h_val.upper() in ["TF", "FRE(%)", "TOTAL", "SITUAÇÃO", "SITUACAO", "MEDIA", "MÉDIA"]:
                continue
            
            # Limpar quebras de linha e códigos numéricos do nome da disciplina na importação
            import re
            h_val_limpo = h_val.split('\n')[0].split('\r')[0]
            h_val_limpo = re.sub(r'\(\d+\)', '', h_val_limpo)
            h_val_limpo = re.sub(r'\b\d+\b', '', h_val_limpo)
            nome_sub = h_val_limpo.strip().upper()
            count = 1
            original_nome = nome_sub
            while nome_sub in col_subjects:
                count += 1
                romanos = {1: "", 2: " II", 3: " III", 4: " IV", 5: " V", 6: " VI"}
                sufixo = romanos.get(count, f" ({count})")
                nome_sub = f"{original_nome}{sufixo}"
                
            col_subjects[nome_sub] = c
            subjects_list.append(nome_sub)

        alunos = []
        aluno_atual = None
        
        # 6. Iterar pelas linhas de dados (após os cabeçalhos)
        r_start = r_header + 1
        if r_start < len(df):
            row_str_next = [str(val).strip().upper() for val in df.iloc[r_start]]
            if "TF" in row_str_next or "MEDIAS" in row_str_next or "M" in row_str_next or "F" in row_str_next:
                r_start += 1
                
        for r in range(r_start, len(df)):
            cell_aluno = df.iloc[r, col_aluno]
            if pd.isna(cell_aluno) or str(cell_aluno).strip() == "":
                continue
                
            cell_aluno_str = str(cell_aluno).strip()
            cell_aluno_norm = normalizar_texto(cell_aluno_str)
            
            # Condição de parada: Fim da tabela de alunos
            if "MEDIA GERAL" in cell_aluno_norm or "ALUNOS DESTAQUE" in cell_aluno_norm or "DISCIPLINAS" in cell_aluno_norm or "LEGENDA" in cell_aluno_norm:
                break
                
            # Verificar se é a linha de "Encaminhamento/Plano de Ação" do aluno anterior
            if "ENCAMINHAMENTO" in cell_aluno_norm or "PLANO DE ACAO" in cell_aluno_norm:
                if aluno_atual:
                    partes = cell_aluno_str.split(":", 1)
                    texto_plano = partes[1].strip() if len(partes) > 1 else ""
                    if texto_plano == "":
                        for c in range(col_aluno + 1, len(df.columns)):
                            val_adj = df.iloc[r, c]
                            if not pd.isna(val_adj) and str(val_adj).strip() != "":
                                texto_plano = str(val_adj).strip()
                                break
                    aluno_atual['plano_acao'] = texto_plano
                continue
            
            # É um novo aluno
            nome = cell_aluno_str.upper()
            situacao = str(df.iloc[r, col_situacao]).strip() if col_situacao is not None else "Ativa"
            if col_situacao is not None and pd.isna(df.iloc[r, col_situacao]):
                situacao = "Ativa"
                
            # Filtrar alunos inativos (Baixa, Transferência, Remanejamento, Remanescentes, etc.)
            situacao_upper = situacao.upper()
            nome_upper = nome.upper()
            is_inativo = False
            for kw in ["BAIXA", "TRANSF", "REMANEJ", "REMANES", "REMANESC"]:
                if kw in situacao_upper or kw in nome_upper:
                    is_inativo = True
                    break
            if is_inativo:
                continue # Ignorar o aluno e não importar!
            
            # Extrair notas das disciplinas
            notas = {}
            for sub, col_idx in col_subjects.items():
                if col_idx < len(df.columns):
                    val = df.iloc[r, col_idx]
                    try:
                        if pd.isna(val) or str(val).strip() in ["", "-", "_", "*"]:
                            notas[sub] = None
                        else:
                            val_str = str(val).replace(",", ".").strip()
                            if val_str.upper() in ["I", "R", "B", "MB"]:
                                mencao_map = {"I": 4.0, "R": 6.0, "B": 8.0, "MB": 10.0}
                                notas[sub] = mencao_map[val_str.upper()]
                            else:
                                notas[sub] = float(val_str)
                    except ValueError:
                        notas[sub] = None
                else:
                    notas[sub] = None

            # Calcular notas baixas (< 5.0) e média local
            notas_validas = [v for v in notas.values() if v is not None]
            calc_notas_baixas = sum(1 for v in notas_validas if v < 5.0)
            calc_media = round(sum(notas_validas) / len(notas_validas), 1) if notas_validas else 0.0

            # Extrair TF (Faltas do bimestre)
            tf = 0
            if col_tf is not None and col_tf < len(df.columns):
                val = df.iloc[r, col_tf]
                try:
                    tf = int(float(str(val).replace(",", ".").strip())) if not pd.isna(val) else 0
                except ValueError:
                    tf = 0
            
            # Extrair Fre% (Presença do bimestre)
            fre = 100.0
            if col_fre is not None and col_fre < len(df.columns):
                val = df.iloc[r, col_fre]
                try:
                    if not pd.isna(val):
                        val_str = str(val).replace("%", "").replace(",", ".").strip()
                        fre = float(val_str)
                        if fre <= 1.0 and fre > 0:
                            fre = fre * 100
                except ValueError:
                    fre = 100.0
                    
            # Extrair FT Ano (Faltas totais do ano)
            ft_ano = tf
            if col_ft_ano is not None and col_ft_ano < len(df.columns):
                val = df.iloc[r, col_ft_ano]
                try:
                    ft_ano = int(float(str(val).replace(",", ".").strip())) if not pd.isna(val) else tf
                except ValueError:
                    ft_ano = tf

            # Extrair Fre Ano (Frequência do ano)
            fre_ano = fre
            if col_fre_ano is not None and col_fre_ano < len(df.columns):
                val = df.iloc[r, col_fre_ano]
                try:
                    if not pd.isna(val):
                        val_str = str(val).replace("%", "").replace(",", ".").strip()
                        fre_ano = float(val_str)
                        if fre_ano <= 1.0 and fre_ano > 0:
                            fre_ano = fre_ano * 100
                except ValueError:
                    fre_ano = fre

            # Média de notas informada na planilha
            planilha_media = calc_media
            if col_media is not None and col_media < len(df.columns):
                val = df.iloc[r, col_media]
                try:
                    planilha_media = float(str(val).replace(",", ".").strip()) if not pd.isna(val) else calc_media
                except ValueError:
                    planilha_media = calc_media

            aluno_atual = {
                'nome': nome,
                'situacao': "Ativa",
                'notas': notas,
                'faltas': tf,
                'frequencia': fre,
                'faltas_ano': ft_ano,
                'frequencia_ano': fre_ano,
                'notas_baixas': calc_notas_baixas,
                'media': planilha_media,
                'plano_acao': ''
            }
            alunos.append(aluno_atual)
            
        return {
            "turma": turma_detectada if turma_detectada else "",
            "bimestre": bimestre_detectado,
            "disciplinas": subjects_list,
            "alunos": alunos
        }

class SchoolAnalyzerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # 0. INICIALIZAR E VERIFICAR LICENÇA
        from licenca import LicenseManager
        self.license_manager = LicenseManager()
        
        # Mostrar a tela inicial de licenciamento/boas-vindas
        self.mostrar_tela_inicial_licenca()

    def inicializar_app_normal(self):
        # Configurar Janela Principal
        self.title("Analisador de Planilhas Escolares — Radar Pedagógico")
        self.geometry("1280x800")
        self.minsize(1100, 700)
        self.resizable(True, True)
        
        # Estado do Sistema (Suporte a múltiplas turmas)
        self.turmas_dados = {}  # { "Turma A": {1: [alunos], 2: [alunos], 3: ..., 4: ...} }
        self.analises_sala = {} # { "Turma A": {1: "Texto da análise...", 2: "..."} }
        self.turma_ativa = ctk.StringVar(value="")
        self.bimestre_ativo = ctk.IntVar(value=1)
        self.filtro_disciplina = ctk.StringVar(value="Todas")
        self.filtro_busca_nome = ctk.StringVar(value="")
        self.filtro_situacao = ctk.StringVar(value="Todos")  # "Todos", "Notas Baixas vs Frequência", "Notas Baixas", "Frequencia Critica", "Destaques", "Com Plano"
        
        # Variável para o aluno atualmente selecionado para edição e filtro comparativo
        self.aluno_selecionado = None
        self.comparativo_aluno_filtro = ctk.StringVar(value="Todos (Sala)")
        self.comparativo_disciplina_filtro = ctk.StringVar(value="Todas (Média Geral)")
        self.detalhes_disciplina_filtro = ctk.StringVar(value="Todas (Média Geral)")
        
        # Arquivo de persistência local das observações
        import sys
        if getattr(sys, 'frozen', False):
            self.persistence_path = os.path.join(os.path.dirname(sys.executable), "dados_salvos.json")
        else:
            self.persistence_path = "dados_salvos.json"
        
        # Carregar dados salvos anteriormente, se houver
        self.carregar_persistência()
        
        # Configurar grid de layout principal (Sidebar + Content)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=0) # Sidebar fixada
        
        # 1. CRIAR SIDEBAR E PÁGINAS
        self.criar_sidebar()
        self.criar_container_paginas()
        
        # Definir a página inicial
        self.mudar_pagina("arquivos")

    def mostrar_tela_inicial_licenca(self):
        from datetime import datetime
        # Configurar Janela para o formato Startup/Bloqueio
        self.title("Radar Pedagógico — Inicialização")
        self.geometry("580x480")
        self.resizable(False, False)
        
        # Configurar grid
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=0)
        
        # Container
        self.lock_frame = ctk.CTkFrame(self, fg_color="#0f172a", corner_radius=0)
        self.lock_frame.grid(row=0, column=0, sticky="nsew")
        
        logo_path = resource_path("logo.png")
        loaded_logo = False
        if os.path.exists(logo_path):
            try:
                from PIL import Image
                pil_img = Image.open(logo_path)
                w, h = pil_img.size
                new_w = 200
                new_h = int(new_w * h / w)
                ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(new_w, new_h))
                self.lbl_lock_logo = ctk.CTkLabel(self.lock_frame, image=ctk_img, text="")
                self.lbl_lock_logo.pack(pady=(20, 10))
                self.lbl_lock_logo.bind("<Double-Button-1>", self.ativar_backdoor_click)
                loaded_logo = True
            except Exception as e:
                print(f"Erro ao carregar logo na tela de licença: {e}")
                
        if not loaded_logo:
            self.lbl_lock_logo = ctk.CTkLabel(
                self.lock_frame, 
                text="🔑 RADAR PEDAGÓGICO", 
                font=ctk.CTkFont(family="Outfit", size=24, weight="bold"),
                text_color="#38bdf8"
            )
            self.lbl_lock_logo.pack(pady=(20, 10))
            self.lbl_lock_logo.bind("<Double-Button-1>", self.ativar_backdoor_click)
        
        # Verificar o status da licença
        is_valid, reason = self.license_manager.verificar_licenca()
        
        # Exibir ID da Máquina
        lbl_machine_id = ctk.CTkLabel(
            self.lock_frame,
            text=f"ID DA MÁQUINA: {self.license_manager.machine_id}",
            font=ctk.CTkFont(family="Courier New", size=16, weight="bold"),
            text_color="#f8fafc",
            fg_color="#1e293b",
            padx=15,
            pady=8,
            corner_radius=6
        )
        lbl_machine_id.pack(pady=(0, 15))
        
        # Calcular status textual da licença
        status = self.license_manager.data.get("status")
        hoje = datetime.now()
        
        if status == "admin_vitalicio":
            status_texto = "Status: Administrador Vitalício Liberado"
            status_color = "#10b981"
        elif status == "ativo":
            status_texto = "Status: Licença Vitalícia Ativa"
            status_color = "#10b981"
        elif status == "ativo_1ano":
            exp_str = self.license_manager.data.get("expira_em")
            if exp_str:
                exp_date = datetime.strptime(exp_str, "%Y-%m-%d")
                dias_restantes = (exp_date - hoje).days
                if dias_restantes > 0:
                    status_texto = f"Status: Licença Temporária (Restam {dias_restantes} dias)"
                    status_color = "#10b981"
                else:
                    status_texto = "Status: Licença de 1 Ano Expirada"
                    status_color = "#f43f5e"
            else:
                status_texto = "Status: Licença de 1 Ano Expirada"
                status_color = "#f43f5e"
        else:
            # Período experimental (8 meses = 244 dias)
            data_ini_str = self.license_manager.data.get("data_inicio")
            data_ini = datetime.strptime(data_ini_str, "%Y-%m-%d")
            dias_uso = (hoje - data_ini).days
            dias_restantes = max(0, 244 - dias_uso)
            
            if is_valid:
                status_texto = f"Status: Período de Avaliação (Restam {dias_restantes} dias)"
                status_color = "#eab308"
            else:
                status_texto = "Status: Período de Avaliação Expirado"
                status_color = "#f43f5e"
                
        # Se houver outro erro de relógio
        if not is_valid and reason == "relogio_alterado":
            status_texto = "Erro: Alteração no relógio do sistema detectada"
            status_color = "#f43f5e"
            
        lbl_status = ctk.CTkLabel(
            self.lock_frame,
            text=status_texto,
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=status_color
        )
        lbl_status.pack(pady=(0, 15))
        self.lbl_status_licenca = lbl_status
        
        # Campo de ativação
        self.entry_chave = ctk.CTkEntry(
            self.lock_frame,
            placeholder_text="Insira a chave de ativação para atualizar",
            width=380,
            height=32,
            fg_color="#1e293b",
            border_color="#334155"
        )
        self.entry_chave.pack(pady=(0, 10))
        
        # Botões de Ação
        btn_frame = ctk.CTkFrame(self.lock_frame, fg_color="transparent")
        btn_frame.pack(pady=10)
        
        self.btn_ativar = ctk.CTkButton(
            btn_frame,
            text="Ativar Licença",
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#0ea5e9",
            hover_color="#0284c7",
            width=150,
            height=34,
            command=self.tentar_ativar_click
        )
        self.btn_ativar.grid(row=0, column=0, padx=5)
        
        self.btn_entrar = ctk.CTkButton(
            btn_frame,
            text="Entrar no Sistema",
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#10b981" if is_valid else "#475569",
            hover_color="#059669" if is_valid else "#475569",
            state="normal" if is_valid else "disabled",
            width=150,
            height=34,
            command=self.entrar_sistema_click
        )
        self.btn_entrar.grid(row=0, column=1, padx=5)
        
        # Suporte a e-mail
        lbl_email_info = ctk.CTkLabel(
            self.lock_frame,
            text="Contato de Suporte: andre.gomes2004@gmail.com",
            font=ctk.CTkFont(size=11),
            text_color="#64748b"
        )
        lbl_email_info.pack(pady=(15, 5))
        
        # Label de erros / feedback
        self.lbl_status_bloqueio = ctk.CTkLabel(
            self.lock_frame,
            text="",
            font=ctk.CTkFont(size=11)
        )
        self.lbl_status_bloqueio.pack()
        
        # Inscrição discreta de direitos autorais no rodapé
        lbl_footer = ctk.CTkLabel(
            self.lock_frame,
            text="Desenvolvido por Prof. André Eduardo Gomes © 2026. Todos os direitos reservados.",
            font=ctk.CTkFont(family="Segoe UI", size=9),
            text_color="#475569"
        )
        lbl_footer.pack(side="bottom", pady=(15, 10))
        
        # Atalho Backdoor Oculto (Ctrl + Shift + Alt + D)
        self.bind_all("<Control-Shift-Alt-KeyPress-d>", self.ativar_backdoor_click)
        self.bind_all("<Control-Shift-Alt-KeyPress-D>", self.ativar_backdoor_click)
        self.bind_all("<Control-Shift-Alt-d>", self.ativar_backdoor_click)
        self.bind_all("<Control-Shift-Alt-D>", self.ativar_backdoor_click)

    def tentar_ativar_click(self):
        chave = self.entry_chave.get()
        sucesso, msg = self.license_manager.tentar_ativar(chave)
        if sucesso:
            messagebox.showinfo("Sucesso", msg)
            
            # Recarregar status da licença na tela
            is_valid, _ = self.license_manager.verificar_licenca()
            self.lbl_status_licenca.configure(
                text="Status: Licença Ativa" if self.license_manager.data.get("status") == "ativo" else "Status: Licença Temporária Ativa",
                text_color="#10b981"
            )
            self.btn_entrar.configure(
                state="normal",
                fg_color="#10b981",
                hover_color="#059669"
            )
            self.lbl_status_bloqueio.configure(text=msg, text_color="#10b981")
            
            # Entrar no sistema automaticamente após ativação bem-sucedida
            self.entrar_sistema_click()
        else:
            self.lbl_status_bloqueio.configure(text=msg, text_color="#f43f5e")

    def ativar_backdoor_click(self, event=None):
        self.license_manager.liberar_backdoor()
        messagebox.showinfo(
            "🔑 Modo Administrador", 
            "Acesso de Administrador Ativado!\n\n"
            "O gerador de chaves para licenças de 1 ano ou vitalícias foi desbloqueado com sucesso.\n"
            "Vá na aba 'Sobre o Sistema' no menu lateral para visualizar o painel do gerador."
        )
        
        # Atualizar tela de licença
        self.lbl_status_licenca.configure(
            text="Status: Administrador Vitalício Liberado",
            text_color="#10b981"
        )
        self.btn_entrar.configure(
            state="normal",
            fg_color="#10b981",
            hover_color="#059669"
        )
        
        # Entrar no sistema automaticamente após backdoor
        self.entrar_sistema_click()

    def entrar_sistema_click(self):
        # Desvincular atalhos do backdoor
        self.unbind_all("<Control-Shift-Alt-KeyPress-d>")
        self.unbind_all("<Control-Shift-Alt-KeyPress-D>")
        self.unbind_all("<Control-Shift-Alt-d>")
        self.unbind_all("<Control-Shift-Alt-D>")
        
        # Destruir o container da tela de licença
        if hasattr(self, 'lock_frame'):
            self.lock_frame.destroy()
            
        # Inicializar a janela principal do app
        self.inicializar_app_normal()


    @property
    def bimestres_dados(self):
        """Propriedade que redireciona acessos legados de bimestres_dados para a turma ativa."""
        t = self.turma_ativa.get()
        if not t:
            return {1: None, 2: None, 3: None, 4: None}
        if t not in self.turmas_dados:
            self.turmas_dados[t] = {1: None, 2: None, 3: None, 4: None}
        return self.turmas_dados[t]
        
    @bimestres_dados.setter
    def bimestres_dados(self, value):
        t = self.turma_ativa.get()
        if t:
            self.turmas_dados[t] = value

    def get_subjects_active_turma(self, b_num=None):
        """Retorna a lista de disciplinas da turma ativa com base nos dados carregados."""
        t = self.turma_ativa.get()
        if not t or t not in self.turmas_dados:
            return []
        
        subjects = set()
        bimestres = [b_num] if b_num else [1, 2, 3, 4]
        for b in bimestres:
            dados = self.turmas_dados[t].get(b)
            if dados:
                for aluno in dados:
                    if 'notas' in aluno:
                        subjects.update([k for k, v in aluno['notas'].items() if k])
        return sorted(list(subjects))
        
    def garantir_bimestre_valido(self):
        """Garante que o bimestre ativo possua dados. Caso contrário, altera para o último carregado."""
        t = self.turma_ativa.get()
        if not t or t not in self.turmas_dados:
            return
            
        b_atual = self.bimestre_ativo.get()
        if self.turmas_dados[t].get(b_atual) is not None:
            return
            
        bimestres_com_dados = [b for b in [1, 2, 3, 4] if self.turmas_dados[t].get(b) is not None]
        if bimestres_com_dados:
            novo_b = max(bimestres_com_dados)
            self.bimestre_ativo.set(novo_b)
            self.combo_bimestre.set(f"{novo_b}º Bimestre")

    def carregar_persistência(self):
        """Carrega dados e caminhos de arquivos salvos na última sessão."""
        if os.path.exists(self.persistence_path):
            try:
                with open(self.persistence_path, 'r', encoding='utf-8') as f:
                    dados = json.load(f)
                    self.turmas_dados = {}
                    
                    if "turmas" in dados:
                        for t_nome, t_data in dados["turmas"].items():
                            t_norm = normalizar_nome_turma(t_nome)
                            if t_norm not in self.turmas_dados:
                                self.turmas_dados[t_norm] = {1: None, 2: None, 3: None, 4: None}
                            for b_str, b_data in t_data.items():
                                b_data_normalized = []
                                if b_data:
                                    for al in b_data:
                                        al_sit = str(al.get("situacao", "Ativa")).strip().upper()
                                        al_nome = str(al.get("nome", "")).strip().upper()
                                        
                                        is_inativo = False
                                        for kw in ["BAIXA", "TRANSF", "REMANEJ", "REMANES", "REMANESC"]:
                                            if kw in al_sit or kw in al_nome:
                                                is_inativo = True
                                                break
                                                
                                        if not is_inativo:
                                            al["situacao"] = "Ativa"
                                            b_data_normalized.append(al)
                                self.turmas_dados[t_norm][int(b_str)] = b_data_normalized if b_data else None
                                
                    self.analises_sala = dados.get("analises_sala", {})
                    
                    active = dados.get("active_turma", "")
                    if active:
                        active_norm = normalizar_nome_turma(active)
                        if active_norm in self.turmas_dados:
                            self.turma_ativa.set(active_norm)
                        elif self.turmas_dados:
                            self.turma_ativa.set(list(self.turmas_dados.keys())[0])
                        else:
                            self.turma_ativa.set("")
                    elif self.turmas_dados:
                        self.turma_ativa.set(list(self.turmas_dados.keys())[0])
                    else:
                        self.turma_ativa.set("")
                        
                    self.garantir_bimestre_valido()
            except Exception as e:
                print(f"Erro ao carregar persistencia: {e}")

    def salvar_persistência(self):
        """Salva o estado atual e as observações modificadas em arquivo JSON local."""
        try:
            turmas_json = {}
            for t_nome, t_data in self.turmas_dados.items():
                turmas_json[t_nome] = {}
                for b_num, b_data in t_data.items():
                    turmas_json[t_nome][str(b_num)] = b_data
            
            with open(self.persistence_path, 'w', encoding='utf-8') as f:
                json.dump({
                    "active_turma": self.turma_ativa.get(),
                    "turmas": turmas_json,
                    "analises_sala": self.analises_sala
                }, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"Erro ao salvar persistencia: {e}")

    def criar_sidebar(self):
        """Cria o menu lateral de navegação com seletor de turmas e bimestres."""
        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0, fg_color="#0f172a")
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(13, weight=1) # Espaçador inferior (linha 13)
        
        logo_path = resource_path("logo.png")
        loaded_logo = False
        if os.path.exists(logo_path):
            try:
                from PIL import Image
                pil_img = Image.open(logo_path)
                w, h = pil_img.size
                new_w = 180
                new_h = int(new_w * h / w)
                ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(new_w, new_h))
                self.lbl_logo = ctk.CTkLabel(
                    self.sidebar, 
                    image=ctk_img,
                    text=""
                )
                self.lbl_logo.grid(row=0, column=0, padx=20, pady=(25, 10))
                loaded_logo = True
            except Exception as e:
                print(f"Erro ao carregar logo.png: {e}")
                
        if not loaded_logo:
            self.lbl_logo = ctk.CTkLabel(
                self.sidebar, 
                text="📊 RADAR PEDAGÓGICO", 
                font=ctk.CTkFont(family="Outfit", size=18, weight="bold"),
                text_color="#38bdf8"
            )
            self.lbl_logo.grid(row=0, column=0, padx=20, pady=(25, 2))
            
        sublogo_text = "Analisador de Bimestres"
        sublogo_color = "#94a3b8"
        if self.license_manager.data.get("status") == "admin_vitalicio":
            sublogo_text = "🔑 Painel do Dono (Admin)"
            sublogo_color = "#10b981"
            
        self.lbl_sublogo = ctk.CTkLabel(
            self.sidebar, 
            text=sublogo_text, 
            font=ctk.CTkFont(family="Outfit", size=11, weight="bold" if sublogo_text != "Analisador de Bimestres" else "normal"),
            text_color=sublogo_color
        )
        self.lbl_sublogo.grid(row=1, column=0, padx=20, pady=(0, 20))
        
        # Seletor de Turma Ativa
        self.lbl_t_select = ctk.CTkLabel(
            self.sidebar, 
            text="Turma Ativa:", 
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#94a3b8"
        )
        self.lbl_t_select.grid(row=2, column=0, padx=20, pady=(10, 2), sticky="w")
        
        t_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        t_frame.grid(row=3, column=0, padx=20, pady=(0, 15), sticky="ew")
        t_frame.grid_columnconfigure(0, weight=1)
        t_frame.grid_columnconfigure(1, weight=0)
        t_frame.grid_columnconfigure(2, weight=0)
        
        turmas_list = list(self.turmas_dados.keys()) if self.turmas_dados else ["Nenhuma Turma"]
        self.combo_turma = ctk.CTkOptionMenu(
            t_frame,
            values=turmas_list,
            command=self.on_turma_alterada,
            fg_color="#1e293b",
            button_color="#334155",
            button_hover_color="#475569"
        )
        self.combo_turma.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        if self.turma_ativa.get():
            self.combo_turma.set(self.turma_ativa.get())
        else:
            self.combo_turma.set("Nenhuma Turma")
            
        self.btn_add_turma = ctk.CTkButton(
            t_frame,
            text="+",
            width=28,
            fg_color="#0d9488",
            hover_color="#0f766e",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.adicionar_turma_click
        )
        self.btn_add_turma.grid(row=0, column=1, padx=2)
        
        self.btn_del_turma = ctk.CTkButton(
            t_frame,
            text="🗑️",
            width=28,
            fg_color="#ef4444",
            hover_color="#dc2626",
            font=ctk.CTkFont(size=12),
            command=self.remover_turma_click
        )
        self.btn_del_turma.grid(row=0, column=2, padx=2)
        
        # Seletor de Bimestre Ativo
        self.lbl_b_select = ctk.CTkLabel(
            self.sidebar, 
            text="Bimestre Ativo:", 
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#94a3b8"
        )
        self.lbl_b_select.grid(row=4, column=0, padx=20, pady=(10, 2), sticky="w")
        
        self.combo_bimestre = ctk.CTkOptionMenu(
            self.sidebar,
            values=["1º Bimestre", "2º Bimestre", "3º Bimestre", "4º Bimestre"],
            command=self.on_bimestre_alterado,
            fg_color="#1e293b",
            button_color="#334155",
            button_hover_color="#475569"
        )
        self.combo_bimestre.grid(row=5, column=0, padx=20, pady=(0, 20), sticky="ew")
        self.combo_bimestre.set(f"{self.bimestre_ativo.get()}º Bimestre")

        # Botões de Navegação
        self.nav_buttons = {}
        paginas = [
            ("arquivos", "📁 Importar Planilhas"),
            ("dashboard", "📈 Dashboard Geral"),
            ("tabela", "📋 Tabela de Alunos"),
            ("graficos", "📊 Graficos da Sala"),
            ("comparativo", "🔄 Comparar Bimestres"),
            ("detalhes", "👤 Detalhes do Aluno"),
            ("sobre", "ℹ️ Sobre o Sistema")
        ]
        
        row_idx = 6
        for pag_id, label in paginas:
            btn = ctk.CTkButton(
                self.sidebar,
                text=label,
                anchor="w",
                font=ctk.CTkFont(size=13, weight="bold"),
                fg_color="transparent",
                text_color="#94a3b8",
                hover_color="#1e293b",
                corner_radius=8,
                command=lambda p=pag_id: self.mudar_pagina(p)
            )
            btn.grid(row=row_idx, column=0, padx=15, pady=4, sticky="ew")
            self.nav_buttons[pag_id] = btn
            row_idx += 1
            
        # Alternador de Tema Escuro/Claro no rodapé (griddado na linha 14, abaixo do espaçador)
        self.theme_switch = ctk.CTkSwitch(
            self.sidebar,
            text="Tema Escuro",
            command=self.toggle_theme,
            font=ctk.CTkFont(size=11),
            text_color="#94a3b8"
        )
        self.theme_switch.select()
        self.theme_switch.grid(row=14, column=0, padx=20, pady=20, sticky="s")

    def toggle_theme(self):
        if self.theme_switch.get() == 1:
            ctk.set_appearance_mode("Dark")
            self.atualizar_estilos_treeview('dark')
            self.desenhar_graficos_ativos()
        else:
            ctk.set_appearance_mode("Light")
            self.atualizar_estilos_treeview('light')
            self.desenhar_graficos_ativos()

    def atualizar_estilos_treeview(self, mode='dark'):
        style = ttk.Style()
        if mode == 'dark':
            style.configure(
                "Treeview",
                background="#1e293b",
                foreground="#f8fafc",
                fieldbackground="#1e293b",
                font=("Segoe UI", 10)
            )
            style.configure(
                "Treeview.Heading",
                background="#0f172a",
                foreground="#f8fafc",
                font=("Segoe UI", 10, "bold")
            )
        else:
            style.configure(
                "Treeview",
                background="#ffffff",
                foreground="#0f172a",
                fieldbackground="#ffffff",
                font=("Segoe UI", 10)
            )
            style.configure(
                "Treeview.Heading",
                background="#e2e8f0",
                foreground="#0f172a",
                font=("Segoe UI", 10, "bold")
            )

    def criar_container_paginas(self):
        """Cria o container central onde as sub-páginas serão renderizadas."""
        self.content_container = ctk.CTkFrame(self, fg_color="transparent")
        self.content_container.grid(row=0, column=1, sticky="nsew", padx=25, pady=25)
        self.content_container.grid_rowconfigure(0, weight=1)
        self.content_container.grid_columnconfigure(0, weight=1)
        
        # Dicionário de construtores de páginas
        self.paginas_builders = {
            "arquivos": self.criar_tela_arquivos,
            "dashboard": self.criar_tela_dashboard,
            "tabela": self.criar_tela_tabela,
            "graficos": self.criar_tela_graficos,
            "comparativo": self.criar_tela_comparativo,
            "detalhes": self.criar_tela_detalhes,
            "sobre": self.criar_tela_sobre
        }
        self.paginas_cache = {}

    def mudar_pagina(self, pagina_id):
        """Esconde a página atual e desenha a nova página selecionada."""
        # Resetar destaque dos botões
        for k, btn in self.nav_buttons.items():
            if k == pagina_id:
                btn.configure(fg_color="#1e293b", text_color="#38bdf8")
            else:
                btn.configure(fg_color="transparent", text_color="#94a3b8")
                
        # Destruir a página antiga
        for child in self.content_container.winfo_children():
            child.destroy()
            
        # Construir e carregar a nova página
        if pagina_id in self.paginas_builders:
            frame = self.paginas_builders[pagina_id]()
            frame.grid(row=0, column=0, sticky="nsew")
            
            # Gatilhos de carregamento de dados específicos por página
            if pagina_id == "dashboard":
                self.carregar_dados_dashboard()
            elif pagina_id == "tabela":
                self.popula_tabela()
            elif pagina_id == "graficos":
                self.desenhar_graficos_ativos()
            elif pagina_id == "comparativo":
                self.atualizar_combo_comp_alunos()
                self.desenhar_comparativo_bimestres()
            elif pagina_id == "detalhes":
                self.preencher_detalhes_aluno()

    def on_bimestre_alterado(self, value):
        num = int(value.split("º")[0])
        self.bimestre_ativo.set(num)
        
        # Atualizar a página ativa
        active_page = None
        for k, btn in self.nav_buttons.items():
            if btn.cget("fg_color") != "transparent":
                active_page = k
                break
                
        if active_page:
            self.mudar_pagina(active_page)

    def on_turma_alterada(self, valor):
        if valor == "Nenhuma Turma" or not valor:
            self.turma_ativa.set("")
            self.combo_turma.set("Nenhuma Turma")
            return
            
        self.turma_ativa.set(valor)
        self.combo_turma.set(valor)
        self.salvar_persistência()
        self.garantir_bimestre_valido()
        
        # Atualizar tela ativa
        active_page = None
        for k, btn in self.nav_buttons.items():
            if btn.cget("fg_color") != "transparent":
                active_page = k
                break
        if active_page:
            self.mudar_pagina(active_page)

    def adicionar_turma_click(self):
        nome = simpledialog.askstring("Nova Turma", "Digite o nome da nova turma:")
        if nome:
            nome = normalizar_nome_turma(nome.strip())
            if not nome:
                return
            if nome in self.turmas_dados:
                messagebox.showinfo("Informação", "Esta turma já existe.")
                self.on_turma_alterada(nome)
                return
            
            self.turmas_dados[nome] = {1: None, 2: None, 3: None, 4: None}
            self.combo_turma.configure(values=list(self.turmas_dados.keys()))
            self.on_turma_alterada(nome)
            messagebox.showinfo("Sucesso", f"Turma '{nome}' criada com sucesso!")

    def remover_turma_click(self):
        t = self.turma_ativa.get()
        if not t:
            messagebox.showwarning("Aviso", "Nenhuma turma selecionada para excluir.")
            return
            
        if messagebox.askyesno("Excluir Turma", f"Deseja realmente excluir a turma '{t}'? Todos os dados associados serão apagados permanentemente."):
            del self.turmas_dados[t]
            if t in self.analises_sala:
                del self.analises_sala[t]
                
            self.salvar_persistência()
            
            turmas_list = list(self.turmas_dados.keys())
            if turmas_list:
                self.combo_turma.configure(values=turmas_list)
                self.on_turma_alterada(turmas_list[0])
            else:
                self.combo_turma.configure(values=["Nenhuma Turma"])
                self.on_turma_alterada("Nenhuma Turma")
                
            messagebox.showinfo("Sucesso", f"Turma '{t}' excluída com sucesso!")

    # ==========================================================================
    # 1. TELA DE ARQUIVOS (IMPORTAÇÃO)
    # ==========================================================================
    def criar_tela_arquivos(self):
        frame = ctk.CTkFrame(self.content_container, fg_color="transparent")
        frame.grid_rowconfigure(2, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        
        # Cabeçalho
        header = ctk.CTkFrame(frame, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        
        lbl_title = ctk.CTkLabel(
            header, 
            text="📁 Importar Planilhas de Bimestres", 
            font=ctk.CTkFont(size=22, weight="bold")
        )
        lbl_title.pack(anchor="w")
        lbl_subtitle = ctk.CTkLabel(
            header, 
            text="Suba os arquivos Excel (.xlsx) ou CSV correspondentes a cada bimestre para analisar.", 
            text_color="#94a3b8"
        )
        lbl_subtitle.pack(anchor="w", pady=(2, 0))
        
        # Grid dos Bimestres (Cartões de Upload)
        grid_cartoes = ctk.CTkFrame(frame, fg_color="transparent")
        grid_cartoes.grid(row=1, column=0, sticky="ew", pady=10)
        for i in range(4):
            grid_cartoes.grid_columnconfigure(i, weight=1)
            
        self.cartoes_status = {}
        for b in range(1, 5):
            card = ctk.CTkFrame(grid_cartoes, fg_color="#1e293b", corner_radius=12, border_width=1, border_color="#334155")
            card.grid(row=0, column=b-1, padx=10, pady=10, sticky="nsew")
            card.grid_columnconfigure(0, weight=1)
            
            # Título do Cartão
            ctk.CTkLabel(
                card, 
                text=f"{b}º Bimestre", 
                font=ctk.CTkFont(size=16, weight="bold"),
                text_color="#f8fafc"
            ).grid(row=0, column=0, padx=15, pady=(15, 5))
            
            # Indicador de Status
            lbl_status = ctk.CTkLabel(
                card, 
                text="Pendente", 
                text_color="#94a3b8",
                font=ctk.CTkFont(size=12, weight="bold")
            )
            lbl_status.grid(row=1, column=0, padx=15, pady=5)
            
            # Botão Importar
            btn_importar = ctk.CTkButton(
                card, 
                text="Selecionar Arquivo", 
                fg_color="#0d9488", 
                hover_color="#0f766e",
                command=lambda bimestre=b: self.importar_planilha_click(bimestre)
            )
            btn_importar.grid(row=2, column=0, padx=15, pady=(15, 5))
            
            # Botão Remover
            btn_remover = ctk.CTkButton(
                card, 
                text="Limpar", 
                fg_color="#475569", 
                hover_color="#334155",
                command=lambda bimestre=b: self.limpar_dados_bimestre(bimestre)
            )
            btn_remover.grid(row=3, column=0, padx=15, pady=(5, 15))
            
            self.cartoes_status[b] = {
                'lbl_status': lbl_status,
                'btn_importar': btn_importar,
                'btn_remover': btn_remover,
                'frame': card
            }
            
        self.atualizar_cartoes_visual()
        
        # Botões de Ação Inferiores
        actions_frame = ctk.CTkFrame(frame, fg_color="transparent")
        actions_frame.grid(row=2, column=0, sticky="s", pady=30)
        
        self.btn_demo = ctk.CTkButton(
            actions_frame,
            text="⚡ Carregar Dados de Demonstracao (Mock)",
            fg_color="#7c3aed",
            hover_color="#6d28d9",
            font=ctk.CTkFont(size=13, weight="bold"),
            height=40,
            command=self.carregar_dados_demo
        )
        self.btn_demo.pack(side="left", padx=10)
        
        return frame

    def atualizar_cartoes_visual(self):
        """Atualiza a cor e o texto do status de cada cartão de bimestre."""
        for b, widgets in self.cartoes_status.items():
            dados = self.bimestres_dados[b]
            if dados:
                widgets['lbl_status'].configure(text=f"✅ Carregado\n({len(dados)} alunos)", text_color="#10b981")
                widgets['frame'].configure(border_color="#10b981")
                widgets['btn_importar'].configure(text="Substituir Arquivo", fg_color="#3b82f6", hover_color="#2563eb")
            else:
                widgets['lbl_status'].configure(text="Pendente", text_color="#94a3b8")
                widgets['frame'].configure(border_color="#334155")
                widgets['btn_importar'].configure(text="Selecionar Arquivo", fg_color="#0d9488", hover_color="#0f766e")

    def importar_planilha_click(self, bimestre_num):
        file_path = filedialog.askopenfilename(
            title=f"Importar Planilha - {bimestre_num}º Bimestre",
            filetypes=[("Arquivos Compativeis", "*.xlsx *.xls *.csv"), ("Planilhas Excel", "*.xlsx *.xls"), ("Arquivos CSV", "*.csv")]
        )
        if not file_path:
            return
            
        try:
            df = SpreadsheetParser.carregar_arquivo(file_path)
            parsed = SpreadsheetParser.analisar_dados(df)
            if not parsed or not parsed["alunos"]:
                messagebox.showerror("Erro de Parser", "Não foi possível extrair dados de alunos desta planilha. Verifique a estrutura.")
                return
                
            turma_detectada = normalizar_nome_turma(parsed["turma"])
            bimestre_detectado = parsed["bimestre"] or bimestre_num
            alunos = parsed["alunos"]
            
            # Confirmar ou adicionar turma vinculada (solicitação obrigatória ao usuário)
            if turma_detectada:
                if not messagebox.askyesno("Confirmar Importação", f"A planilha detectou a sala '{turma_detectada}' para o {bimestre_detectado}º Bimestre.\nDeseja prosseguir com a importação para esta sala?"):
                    return # Abortar para evitar importação errada
            else:
                escolha = simpledialog.askstring(
                    "Confirmar Sala/Turma",
                    "Não detectamos a sala no arquivo.\nDigite o nome da sala para importação:"
                )
                if not escolha:
                    return
                turma_detectada = normalizar_nome_turma(escolha.strip())
                
            if not turma_detectada:
                messagebox.showwarning("Aviso", "O nome da sala não pode ser vazio.")
                return
                
            if turma_detectada not in self.turmas_dados:
                self.turmas_dados[turma_detectada] = {1: None, 2: None, 3: None, 4: None}
                
            self.turma_ativa.set(turma_detectada)
            self.combo_turma.configure(values=list(self.turmas_dados.keys()))
            self.combo_turma.set(turma_detectada)
            
            # Atribuir dados
            self.bimestres_dados[bimestre_detectado] = alunos
            self.salvar_persistência()
            self.atualizar_cartoes_visual()
            
            # Se mudou de bimestre/turma, atualizar página e carregar o dashboard primeiro
            self.combo_bimestre.set(f"{bimestre_detectado}º Bimestre")
            self.bimestre_ativo.set(bimestre_detectado)
            self.garantir_bimestre_valido()
            self.mudar_pagina("dashboard")
            
            messagebox.showinfo("Sucesso!", f"Planilha de '{turma_detectada}' para o {bimestre_detectado}º Bimestre importada com sucesso!\n{len(alunos)} alunos carregados.")
        except Exception as e:
            messagebox.showerror("Erro de Leitura", f"Erro ao processar o arquivo: {str(e)}")

    def limpar_dados_bimestre(self, bimestre_num):
        if self.bimestres_dados[bimestre_num] is None:
            return
        if messagebox.askyesno("Limpar Dados", f"Tem certeza que deseja apagar os dados do {bimestre_num}º Bimestre da turma ativa?"):
            self.bimestres_dados[bimestre_num] = None
            self.salvar_persistência()
            self.atualizar_cartoes_visual()

    def carregar_dados_demo(self):
        """Gera dados de demonstração idênticos àqueles mostrados na planilha da imagem 1."""
        if not messagebox.askyesno("Carregar Demo", "Deseja carregar os dados ficticios de demonstracao para todos os bimestres?"):
            return
            
        turma_demo = "2º D - DS"
        if turma_demo not in self.turmas_dados:
            self.turmas_dados[turma_demo] = {1: None, 2: None, 3: None, 4: None}
            
        self.turma_ativa.set(turma_demo)
        self.combo_turma.configure(values=list(self.turmas_dados.keys()))
        self.combo_turma.set(turma_demo)
        
        # Lista padrão de disciplinas da demonstração
        subjects_demo = [
            "BIOLOGIA", "CARREIRA E COMPETENCIAS", "EDUCACAO FISICA", "FISICA", "GEOGRAFIA", 
            "HISTORIA", "LINGUA INGLESA", "LINGUA PORTUGUESA", "LOGICA E LINGUAGEM DE PROGRAMACAO", 
            "MATEMATICA", "PROCESSOS DE DESENVOLVIMENTO", "QUIMICA", "REDES", "SOCIOLOGIA"
        ]
        
        estudantes_img = [
            ("ALEXANDRE AKILA DE VASCONCELOS SILVA", [5, 6, 5, 7, 10, 9, 2, 9, 8, 9, 9, 5, 8, 0], 31, 98, 56, 91),
            ("DAVI DE OLIVEIRA LIMA", [6, 7, 10, 7, 10, 0, 2, 7, 7, 8, 7, 6, 7, 6], 22, 93, 49, 92),
            ("DAVI HENRIQUE JACINTO", [4, 4, 3, 5, 5, 7, 2, 4, 3, 6, 4, 2, 3, 6], 80, 75, 151, 76),
            ("EDUARDO FRANCISCO DE OLIVEIRA", [8, 8, 10, 7, 9, 9, 5, 6, 6, 8, 9, 6, 8, 6], 11, 97, 25, 96),
            ("ORION NOGUEIRA DE JESUS", [5, 9, 5, 7, 10, 9, 2, 9, 8, 8, 8, 8, 8, 6], 11, 97, 21, 97),
            ("GABRIEL GUIMARAES DOS SANTOS BEZERRA", [7, 7, 10, 6, 10, 9, 1, 8, 6, 7, 5, 5, 3, 3], 45, 86, 86, 86),
            ("GUSTAVO LOPES OLIVEIRA ROLIM", [3, 3, 5, 5, 2, 5, 1, 1, 3, 6, 3, 3, 3, 1], 116, 64, 193, 69),
            ("HALIFFER WILLIAM FERREIRA DA SILVA", [3, 5, 10, 6, 7, 9, 1, 7, 4, 7, 4, 5, 3, 3], 31, 90, 56, 91),
            ("HENZO GABRIEL BATISTA VIEIRA DA SILVA", [5, 6, 5, 5, 8, 8, 2, 7, 8, 9, 9, 3, 7, 6], 34, 89, 68, 89),
            ("JEAN RENAN MARQUES FIORINDO", [4, 4, 5, 5, 3, 7, 8, 1, 5, 6, 5, 0, 3, 2], 116, 64, 198, 69),
            ("LEONARDO JOSE ROCHA DA SILVA", [3, 4, 3, 5, 6, 6, 1, 3, 3, 6, 3, 2, 4, 5], 73, 77, 162, 74),
            ("NICOLAS JULIO PAZ DE ANDRADE", [5, 5, 3, 6, 1, 5, 1, 1, 5, 6, 5, 1, 4, 1], 104, 67, 132, 79),
            ("NICCOLLAS ANDRE FESTINO DA SILVA", [3, 4, 5, 5, 6, 5, 2, 2, 3, 6, 3, 1, 4, 5], 88, 72, 134, 79),
            ("PAULO HENRIQUE ROCHA ARAUJO", [3, 8, 10, 7, 8, 9, 2, 4, 6, 6, 7, 5, 4, 5], 33, 90, 67, 89),
            ("PIETRO AUGUSTO ARAUJO DOS SANTOS", [4, 9, 5, 7, 10, 5, 2, 7, 8, 7, 8, 6, 6, 5], 18, 94, 36, 94),
            ("RAFAELLA DA SILVA SALDANHA", [5, 7, 10, 7, 7, 6, 2, 8, 7, 7, 6, 5, 4, 8], 49, 85, 83, 87),
            ("REGINALDO DA SILVA DIAS JUNIOR", [6, 4, 5, 7, 7, 9, 3, 4, 5, 6, 5, 4, 4, 1], 65, 80, 109, 82),
            ("VICTOR HUGO ACEVILLI", [3, 4, 10, 7, 7, 2, 2, 5, 4, 7, 5, 3, 3, 6], 49, 85, 96, 85),
            ("VICTORIA DA SILVA SOARES", [4, 9, 3, 7, 5, 2, 5, 1, 9, 7, 7, 0, 5, 5], 7, 98, 62, 90),
            ("CARLOS EDUARDO DA SILVA CONCEICAO", [3, 3, 5, 5, 1, 0, 2, 0, 3, 6, 3, 3, 3, 5], 115, 64, 186, 70),
            ("HELLOYSE GABRIELE BARBOSA DA SILVA", [9, 8, 9, 7, 10, 9, 2, 9, 9, 9, 8, 7, 8, 8], 12, 96, 27, 96),
            ("ANA LIVIA CONCEICAO DA SILVA", [10, 8, 10, 7, 10, 9, 2, 9, 9, 9, 9, 8, 8, 10], 12, 96, 24, 96),
            ("MATHEUS HENRIQUE BARRETO DE ARAUJO", [7, 9, 10, 7, 10, 5, 2, 9, 8, 7, 8, 6, 4, 1], 30, 91, 55, 91),
            ("DAVID RYAN MOURA", [4, 4, 5, 5, 6, 6, 2, 1, 3, 7, 3, 9, 3, 8], 81, 75, 141, 78),
            ("DOUGLAS FABIO DE SOUZA SANTOS", [4, 4, 10, 4, 3, 9, 1, 3, 4, 7, 4, 0, 0, 9], 79, 75, 91, 88),
            ("PABLO SANTOS FERREIRA DA SILVA", [5, 8, 10, 5, 8, 9, 2, 4, 8, 8, 8, 6, 4, 6], 40, 87, 46, 93)
        ]

        dados_b2 = []
        for nome, notas_list, tf, fre, ft_ano, fre_ano in estudantes_img:
            notes_dict = {}
            for idx, sub in enumerate(subjects_demo):
                notes_dict[sub] = notas_list[idx] if idx < len(notas_list) else 5.0
                
            calc_nb = sum(1 for v in notes_dict.values() if v < 5.0)
            calc_m = round(sum(notes_dict.values()) / len(notes_dict), 1)
            
            aluno = {
                'nome': nome,
                'situacao': "Ativa",
                'notas': notes_dict,
                'faltas': tf,
                'frequencia': fre,
                'faltas_ano': ft_ano,
                'frequencia_ano': fre_ano,
                'notas_baixas': calc_nb,
                'media': calc_m,
                'plano_acao': 'Acompanhamento Pedagogico' if calc_nb > 5 or fre < 75 else ''
            }
            dados_b2.append(aluno)
            
        self.bimestres_dados[2] = dados_b2
        
        # Gerar variações para 1º, 3º e 4º bimestres
        for b in [1, 3, 4]:
            dados_b = []
            for aluno_b2 in dados_b2:
                novas_notas = {}
                for sub, nota in aluno_b2['notas'].items():
                    if b == 1:
                        mod = random.choice([-1.0, 0.0, 0.5])
                    else:
                        mod = random.choice([0.0, 0.5, 1.0, 1.5 if aluno_b2['nome'] in ["DAVI HENRIQUE JACINTO", "GUSTAVO LOPES OLIVEIRA ROLIM"] else 0.0])
                    
                    nova_nota = min(10.0, max(0.0, nota + mod))
                    novas_notas[sub] = round(nova_nota, 1)
                    
                calc_nb = sum(1 for v in novas_notas.values() if v < 5.0)
                calc_m = round(sum(novas_notas.values()) / len(novas_notas), 1)
                
                if b == 1:
                    faltas = max(1, int(aluno_b2['faltas'] * 0.5))
                    freq = min(100.0, 100 - (faltas * 0.5))
                elif b == 3:
                    faltas = int(aluno_b2['faltas'] * 1.4)
                    freq = max(50.0, 100 - (faltas * 0.4))
                else:
                    faltas = int(aluno_b2['faltas'] * 1.8)
                    freq = max(50.0, 100 - (faltas * 0.35))
                
                aluno = {
                    'nome': aluno_b2['nome'],
                    'situacao': "Ativa",
                    'notas': novas_notas,
                    'faltas': max(0, int(faltas - (b * 2))),
                    'frequencia': round(freq, 1),
                    'faltas_ano': int(faltas),
                    'frequencia_ano': round(freq, 1),
                    'notas_baixas': calc_nb,
                    'media': calc_m,
                    'plano_acao': 'Recuperacao recomendada' if calc_nb > 3 else aluno_b2['plano_acao']
                }
                dados_b.append(aluno)
            self.bimestres_dados[b] = dados_b
            
        self.salvar_persistência()
        self.atualizar_cartoes_visual()
        messagebox.showinfo("Dados de Demonstracao", "Os 4 bimestres foram carregados com dados simulados completos para teste!")
        self.combo_bimestre.set("2º Bimestre")
        self.bimestre_ativo.set(2)
        self.mudar_pagina("dashboard")

    # ==========================================================================
    # 2. TELA DE DASHBOARD (MÉTRICAS GERAIS)
    # ==========================================================================
    def criar_tela_dashboard(self):
        frame = ctk.CTkFrame(self.content_container, fg_color="transparent")
        frame.grid_rowconfigure(2, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        
        # Cabeçalho com botão de PDF
        self.dash_header = ctk.CTkFrame(frame, fg_color="transparent")
        self.dash_header.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        
        lbl_title = ctk.CTkLabel(
            self.dash_header, 
            text="📈 Visao Geral — 2º Bimestre", 
            font=ctk.CTkFont(size=22, weight="bold")
        )
        lbl_title.pack(anchor="w", side="left")
        self.lbl_dash_title = lbl_title
        
        self.btn_pdf_sala = ctk.CTkButton(
            self.dash_header,
            text="Exportar Relatório PDF da Sala 📄",
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#0d9488",
            hover_color="#0f766e",
            command=self.exportar_pdf_sala
        )
        self.btn_pdf_sala.pack(anchor="e", side="right", padx=(10, 0))
        
        # 4 Cards de Métricas Gerais
        metrics_panel = ctk.CTkFrame(frame, fg_color="transparent")
        metrics_panel.grid(row=1, column=0, sticky="ew", pady=10)
        for i in range(4):
            metrics_panel.grid_columnconfigure(i, weight=1)
            
        self.cards = {}
        cards_info = [
            ("Alunos Ativos", "0", "#1e293b", "👨‍🎓"),
            ("Média Geral da Sala", "0.0", "#1e293b", "📊"),
            ("Frequência Média", "0.0%", "#1e293b", "📅"),
            ("Alunos Críticos (<75% Freq)", "0", "#1e293b", "🚨")
        ]
        
        for idx, (title, val, bg, icon) in enumerate(cards_info):
            card = ctk.CTkFrame(metrics_panel, fg_color=bg, corner_radius=10, border_width=1, border_color="#334155")
            card.grid(row=0, column=idx, padx=10, pady=5, sticky="ew")
            card.grid_columnconfigure(0, weight=1)
            
            lbl_title = ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=11, weight="bold"), text_color="#94a3b8")
            lbl_title.grid(row=0, column=0, padx=15, pady=(12, 2), sticky="w")
            
            v_frame = ctk.CTkFrame(card, fg_color="transparent")
            v_frame.grid(row=1, column=0, padx=15, pady=(0, 12), sticky="ew")
            v_frame.grid_columnconfigure(0, weight=1)
            
            lbl_val = ctk.CTkLabel(v_frame, text=val, font=ctk.CTkFont(size=24, weight="bold"), text_color="#f8fafc")
            lbl_val.grid(row=0, column=0, sticky="w")
            
            lbl_icon = ctk.CTkLabel(v_frame, text=icon, font=ctk.CTkFont(size=24))
            lbl_icon.grid(row=0, column=1, sticky="e")
            
            self.cards[idx] = lbl_val

        # Área Inferior
        bot_frame = ctk.CTkFrame(frame, fg_color="transparent")
        bot_frame.grid(row=2, column=0, sticky="nsew", pady=15)
        bot_frame.grid_rowconfigure(0, weight=1)
        bot_frame.grid_columnconfigure(0, weight=1) # Destaques/Disciplinas/BuscaAtiva
        bot_frame.grid_columnconfigure(1, weight=1) # Bloco de Análise
        
        left_sub = ctk.CTkFrame(bot_frame, fg_color="transparent")
        left_sub.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left_sub.grid_rowconfigure(0, weight=1)
        left_sub.grid_rowconfigure(1, weight=1)
        left_sub.grid_rowconfigure(2, weight=1)
        left_sub.grid_columnconfigure(0, weight=1)
        
        card_destaques = ctk.CTkFrame(left_sub, fg_color="#1e293b", corner_radius=10, border_width=1, border_color="#334155")
        card_destaques.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
        card_destaques.grid_rowconfigure(1, weight=1)
        card_destaques.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(
            card_destaques, 
            text="⭐ Alunos Destaque (Maiores Medias)", 
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#f59e0b"
        ).grid(row=0, column=0, padx=15, pady=(10, 5), sticky="w")
        
        self.txt_destaques = ctk.CTkTextbox(
            card_destaques, 
            fg_color="#1e293b", 
            text_color="#f8fafc", 
            font=ctk.CTkFont(family="Segoe UI", size=11),
            border_width=0
        )
        self.txt_destaques.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        
        card_disciplinas = ctk.CTkFrame(left_sub, fg_color="#1e293b", corner_radius=10, border_width=1, border_color="#334155")
        card_disciplinas.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        card_disciplinas.grid_rowconfigure(1, weight=1)
        card_disciplinas.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(
            card_disciplinas, 
            text="📉 Disciplinas com Mais Alunos Abaixo de 5,0", 
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#f43f5e"
        ).grid(row=0, column=0, padx=15, pady=(10, 5), sticky="w")
        
        self.txt_disciplinas = ctk.CTkTextbox(
            card_disciplinas, 
            fg_color="#1e293b", 
            text_color="#f8fafc", 
            font=ctk.CTkFont(family="Segoe UI", size=11),
            border_width=0
        )
        self.txt_disciplinas.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)

        card_busca_ativa = ctk.CTkFrame(left_sub, fg_color="#1e293b", corner_radius=10, border_width=1, border_color="#334155")
        card_busca_ativa.grid(row=2, column=0, sticky="nsew")
        card_busca_ativa.grid_rowconfigure(1, weight=1)
        card_busca_ativa.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(
            card_busca_ativa, 
            text="🚨 Alertas de Busca Ativa (Evasão/Faltas)", 
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#f43f5e"
        ).grid(row=0, column=0, padx=15, pady=(10, 5), sticky="w")
        
        self.txt_busca_ativa = ctk.CTkTextbox(
            card_busca_ativa, 
            fg_color="#1e293b", 
            text_color="#f8fafc", 
            font=ctk.CTkFont(family="Segoe UI", size=11),
            border_width=0
        )
        self.txt_busca_ativa.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)

        right_sub = ctk.CTkFrame(bot_frame, fg_color="#1e293b", corner_radius=10, border_width=1, border_color="#334155")
        right_sub.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        right_sub.grid_rowconfigure(2, weight=1)
        right_sub.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(
            right_sub, 
            text="📝 Analise de Desempenho e Assiduidade", 
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#38bdf8"
        ).grid(row=0, column=0, padx=20, pady=(15, 2), sticky="w")
        
        ctk.CTkLabel(
            right_sub, 
            text="Esta analise e gerada dinamicamente com base nas planilhas e e editavel.", 
            font=ctk.CTkFont(size=11),
            text_color="#94a3b8"
        ).grid(row=1, column=0, padx=20, pady=(0, 10), sticky="w")
        
        self.txt_analise = ctk.CTkTextbox(
            right_sub, 
            fg_color="#0f172a", 
            text_color="#e2e8f0", 
            
            font=ctk.CTkFont(family="Segoe UI", size=13),
            border_width=0,
            wrap="word"
        )
        self.txt_analise.grid(row=2, column=0, sticky="nsew", padx=20, pady=(0, 15))
        
        self.btn_save_analysis = ctk.CTkButton(
            right_sub,
            text="Salvar Observacoes da Turma 💾",
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#0ea5e9",
            hover_color="#0284c7",
            command=self.salvar_observacoes_sala
        )
        self.btn_save_analysis.grid(row=3, column=0, padx=20, pady=(0, 20), sticky="ew")
        
        return frame

    def carregar_dados_dashboard(self):
        # Obter o bimestre ativo selecionado na sidebar
        b_ativo = self.bimestre_ativo.get()
        # Pegar apenas os bimestres carregados que sejam menores ou iguais ao bimestre ativo
        bimestres_carregados = [b for b in range(1, b_ativo + 1) if self.bimestres_dados.get(b) is not None]
        
        # Se nenhum bimestre aplicável estiver carregado, limpar tudo e exibir aviso
        if not bimestres_carregados:
            self.lbl_dash_title.configure(text=f"📈 Visão Geral — Sem Dados no {b_ativo}º Bimestre")
            for i in range(4):
                self.cards[i].configure(text="-")
            self.txt_destaques.delete("1.0", "end")
            self.txt_destaques.insert("1.0", "Carregue planilhas na aba 'Importar Planilhas' para começar.")
            self.txt_disciplinas.delete("1.0", "end")
            self.txt_disciplinas.insert("1.0", "Carregue planilhas na aba 'Importar Planilhas' para começar.")
            self.txt_busca_ativa.delete("1.0", "end")
            self.txt_busca_ativa.insert("1.0", "Carregue planilhas na aba 'Importar Planilhas' para começar.")
            self.txt_analise.delete("1.0", "end")
            self.txt_analise.insert("1.0", f"Sem dados carregados até o {b_ativo}º bimestre nesta turma.")
            return

        # Definir título consolidado dinâmico
        if len(bimestres_carregados) == 1:
            self.lbl_dash_title.configure(text=f"📈 Visão Geral — {bimestres_carregados[0]}º Bimestre")
        else:
            b_ativos_str = ", ".join(f"{b}º" for b in bimestres_carregados)
            self.lbl_dash_title.configure(text=f"📈 Visão Geral Consolidada — Bimestres: {b_ativos_str}")
        
        # Consolidar dados de todos os alunos ativos da turma selecionada
        nomes_alunos = set()
        for b in bimestres_carregados:
            for al in self.bimestres_dados[b]:
                nomes_alunos.add(al['nome'])
                
        consolidado = []
        for nome in sorted(nomes_alunos):
            medias = []
            frequencias = []
            notas_bimestres = []
            situacao = "Ativa"
            
            for b in bimestres_carregados:
                al_reg = next((x for x in self.bimestres_dados[b] if x['nome'] == nome), None)
                if al_reg:
                    medias.append(al_reg['media'])
                    frequencias.append(al_reg['frequencia'])
                    notas_bimestres.append(al_reg['notas'])
                    situacao = al_reg.get('situacao', 'Ativa')
                    
            if not medias:
                continue
                
            # Média das notas por disciplina
            notas_consolidadas = {}
            all_subs = set(k for notas in notas_bimestres for k in notas.keys())
            for sub in all_subs:
                vals = [notas.get(sub) for notas in notas_bimestres if notas.get(sub) is not None]
                if vals:
                    notas_consolidadas[sub] = sum(vals) / len(vals)
                    
            consolidado.append({
                'nome': nome,
                'media': sum(medias) / len(medias),
                'frequencia': sum(frequencias) / len(frequencias),
                'notas': notas_consolidadas,
                'situacao': situacao
            })
            
        alunos_ativos = [al for al in consolidado if al.get('situacao', 'Ativa') == 'Ativa']
        total_alunos = len(alunos_ativos)
        
        if total_alunos == 0:
            for i in range(4):
                self.cards[i].configure(text="-")
            self.txt_destaques.delete("1.0", "end")
            self.txt_destaques.insert("1.0", "Sem alunos ativos nesta turma.")
            self.txt_disciplinas.delete("1.0", "end")
            self.txt_disciplinas.insert("1.0", "Sem dados de disciplinas.")
            self.txt_busca_ativa.delete("1.0", "end")
            self.txt_busca_ativa.insert("1.0", "Sem alertas de busca ativa.")
            return
            
        # Métricas Consolidadas
        media_sala = round(sum(aluno['media'] for aluno in alunos_ativos) / total_alunos, 1)
        freq_media = round(sum(aluno['frequencia'] for aluno in alunos_ativos) / total_alunos, 1)
        
        # Alunos Críticos (<75% frequência acumulada)
        alunos_criticos = sum(1 for aluno in alunos_ativos if aluno['frequencia'] < 75.0)
        
        # Preencher cartões com o valor consolidado único (visão limpa e clara)
        self.cards[0].configure(text=str(total_alunos))
        self.cards[1].configure(text=f"{media_sala:.1f}")
        self.cards[2].configure(text=f"{freq_media:.1f}%")
        self.cards[3].configure(text=str(alunos_criticos))
        if alunos_criticos > 0:
            self.cards[3].configure(text_color="#f43f5e")
        else:
            self.cards[3].configure(text_color="#f8fafc")
            
        # Destaque Absoluto Consolidado (Nenhuma média de disciplina < 5.0 acumulada)
        candidatos_destaque = []
        for aluno in alunos_ativos:
            sem_vermelha = True
            for sub, nota in aluno['notas'].items():
                if nota is not None and nota < 5.0:
                    sem_vermelha = False
                    break
            if sem_vermelha and len(aluno['notas']) > 0:
                candidatos_destaque.append(aluno)
                
        aluno_dest_abs = None
        if candidatos_destaque:
            candidatos_destaque = sorted(candidatos_destaque, key=lambda x: x['media'], reverse=True)
            aluno_dest_abs = candidatos_destaque[0]
            
        # Maiores Médias Gerais Consolidadas
        destaques = sorted(alunos_ativos, key=lambda x: x['media'], reverse=True)[:5]
        self.txt_destaques.delete("1.0", "end")
        
        if aluno_dest_abs:
            self.txt_destaques.insert("end", "⭐ DESTAQUE ABSOLUTO ACUMULADO ⭐\n")
            self.txt_destaques.insert("end", f" Aluno: {aluno_dest_abs['nome']}\n")
            self.txt_destaques.insert("end", f" Média Acumulada: {aluno_dest_abs['media']:.1f} | Frequência: {aluno_dest_abs['frequencia']:.1f}%\n")
            self.txt_destaques.insert("end", " (Nenhuma média de disciplina abaixo de 5.0)\n")
            self.txt_destaques.insert("end", "-"*45 + "\n\n")
            
        self.txt_destaques.insert("end", "Maiores Médias Gerais Acumuladas:\n")
        for idx, d in enumerate(destaques):
            self.txt_destaques.insert("end", f" {idx+1}. {d['nome']:.<38} {d['media']:.1f}\n")
            
        # Disciplinas Críticas Consolidadas
        subjects = set(k for al in consolidado for k in al['notas'].keys())
        disciplinas_abaixo = {}
        for sub in subjects:
            alunos_abaixo = sum(1 for aluno in alunos_ativos if aluno['notas'].get(sub) is not None and aluno['notas'][sub] < 5.0)
            if alunos_abaixo > 0:
                disciplinas_abaixo[sub] = alunos_abaixo
                
        disciplinas_criticas_ordenadas = sorted(disciplinas_abaixo.items(), key=lambda x: x[1], reverse=True)
        
        self.txt_disciplinas.delete("1.0", "end")
        if disciplinas_criticas_ordenadas:
            self.txt_disciplinas.insert("end", "Disciplinas Críticas (Média Geral < 5.0):\n")
            for idx, (sub, qtd) in enumerate(disciplinas_criticas_ordenadas[:5]):
                import re
                sub_limpo = sub.split('\n')[0].split('\r')[0]
                sub_limpo = re.sub(r'\s*\d+$', '', sub_limpo).strip()
                self.txt_disciplinas.insert("end", f" {idx+1}. {sub_limpo:.<30} {qtd} alunos < 5.0\n")
        else:
            self.txt_disciplinas.insert("end", "Nenhuma disciplina com média geral abaixo de 5.0.")
            
        # Alertas de Busca Ativa
        self.txt_busca_ativa.delete("1.0", "end")
        alertas = []
        for aluno in alunos_ativos:
            if aluno['frequencia'] < 75.0:
                alertas.append(f"• {aluno['nome']}: Frequência acumulada crítica de {aluno['frequencia']:.1f}%")
                
        # Alerta de quedas bruscas consecutivas
        alertas_queda = []
        for nome in sorted(nomes_alunos):
            for i in range(len(bimestres_carregados) - 1):
                b_curr = bimestres_carregados[i]
                b_next = bimestres_carregados[i+1]
                al_curr = next((x for x in self.bimestres_dados[b_curr] if x['nome'] == nome and x.get('situacao', 'Ativa') == 'Ativa'), None)
                al_next = next((x for x in self.bimestres_dados[b_next] if x['nome'] == nome and x.get('situacao', 'Ativa') == 'Ativa'), None)
                if al_curr and al_next:
                    diff_f = al_curr['frequencia'] - al_next['frequencia']
                    if diff_f > 15.0:
                        alertas_queda.append(f"• {nome}: Queda de presenças de {b_curr}ºB para {b_next}ºB (-{diff_f:.1f}%)")
                        
        if alertas or alertas_queda:
            if alertas:
                self.txt_busca_ativa.insert("end", "Alunos com Baixa Frequência Acumulada (<75%):\n")
                self.txt_busca_ativa.insert("end", "\n".join(alertas) + "\n\n")
            if alertas_queda:
                self.txt_busca_ativa.insert("end", "📈 Quedas Consecutivas de Assiduidade (>15%):\n")
                self.txt_busca_ativa.insert("end", "\n".join(alertas_queda) + "\n")
        else:
            self.txt_busca_ativa.insert("end", "Nenhum aluno em risco crítico de evasão acumulado.")
            
        # Análise Qualitativa Geral da Turma
        chave_analise = str(b_ativo)
        analise_salva = self.recuperar_analise_salva(chave_analise)
        self.txt_analise.delete("1.0", "end")
        if analise_salva:
            self.txt_analise.insert("1.0", analise_salva)
        else:
            if len(bimestres_carregados) == 1:
                txt_generico = self.gerar_analise_automatica(bimestres_carregados[0], total_alunos, media_sala, freq_media, disciplinas_criticas_ordenadas)
            else:
                txt_generico = self.gerar_analise_automatica_consolidada(bimestres_carregados, total_alunos, media_sala, freq_media, disciplinas_criticas_ordenadas)
            self.txt_analise.insert("1.0", txt_generico)

    def gerar_analise_automatica_consolidada(self, bimestres, total, media, freq, criticas):
        import re
        criticas_limpas = []
        for sub, qtd in criticas[:3]:
            sub_limpo = sub.split('\n')[0].split('\r')[0]
            sub_limpo = re.sub(r'\s*\d+$', '', sub_limpo).strip()
            criticas_limpas.append(f"{sub_limpo} ({qtd} alunos)")
        criticas_str = ", ".join(criticas_limpas)
        b_str = ", ".join(f"{b}º" for b in bimestres)
        txt = (
            f"ANÁLISE DE DESEMPENHO E ASSIDUIDADE ACUMULADA: {self.turma_ativa.get()}\n\n"
            f"A análise consolidada dos bimestres {b_str} aponta uma média acumulada da sala de {media:.1f} e uma assiduidade acumulada de {freq:.1f}% "
            f"para um total de {total} alunos matriculados ativos.\n\n"
            f"DIAGNÓSTICO CRÍTICO CONSOLIDADO:\n"
            f"As disciplinas com maior recorrência de alunos abaixo da média nas planilhas importadas são: {criticas_str if criticas else 'Nenhuma disciplina crítica'}.\n\n"
            f"ESTRATÉGIA IMEDIATA E BUSCA ATIVA:\n"
            f"1. Acompanhamento prioritário de estudantes com alertas de evasão ou frequência acumulada abaixo de 75%.\n"
            f"2. Apoio pedagógico e recuperação contínua nas disciplinas de maior dificuldade apontadas no consolidado.\n"
            f"3. Reunião de alinhamento com responsáveis para reverter o quadro de infrequência."
        )
        return txt

    def gerar_analise_automatica(self, b, total, media, freq, criticas):
        import re
        criticas_limpas = []
        for sub, qtd in criticas[:3]:
            sub_limpo = sub.split('\n')[0].split('\r')[0]
            sub_limpo = re.sub(r'\s*\d+$', '', sub_limpo).strip()
            criticas_limpas.append(f"{sub_limpo} ({qtd} alunos)")
        criticas_str = ", ".join(criticas_limpas)
        txt = (
            f"ANÁLISE DE DESEMPENHO E ASSIDUIDADE: {self.turma_ativa.get()}\n\n"
            f"A análise consolidada do {b}º Bimestre aponta uma média geral da sala de {media:.1f} e uma assiduidade média de {freq:.1f}% "
            f"para um total de {total} alunos matriculados ativos.\n\n"
            f"DIAGNÓSTICO CRÍTICO:\n"
            f"Identificamos que as disciplinas com maior número de alunos com notas em recuperação (< 5.0) são: {criticas_str if criticas else 'Nenhuma disciplina crítica'}.\n"
            f"Há uma correlação direta entre o excesso de faltas e o baixo rendimento nas disciplinas da área técnica de programação.\n\n"
            f"ESTRATÉGIA IMEDIATA:\n"
            f"1. Busca ativa para os alunos em risco crítico de frequência escolar (abaixo de 75%).\n"
            f"2. Nivelamento prático de algoritmos e redes nas primeiras aulas do próximo bimestre.\n"
            f"3. Alinhamento de planos de ação (encaminhamento pedagógico) integrados aos pais."
        )
        return txt

    def recuperar_analise_salva(self, b):
        t = self.turma_ativa.get()
        if not t:
            return None
        return self.analises_sala.get(t, {}).get(str(b))

    def salvar_observacoes_sala(self):
        t = self.turma_ativa.get()
        if not t:
            messagebox.showerror("Erro", "Nenhuma turma ativa selecionada.")
            return
            
        b = self.bimestre_ativo.get()
        texto = self.txt_analise.get("1.0", "end-1c")
        
        if t not in self.analises_sala:
            self.analises_sala[t] = {}
            
        self.analises_sala[t][str(b)] = texto
        self.salvar_persistência()
        messagebox.showinfo("Sucesso", "Análise pedagógica da turma salva com sucesso!")

    # ==========================================================================
    # 3. TELA DE TABELA (GRILLA DE ALUNOS COM FILTROS)
    # ==========================================================================
    def criar_tela_tabela(self):
        frame = ctk.CTkFrame(self.content_container, fg_color="transparent")
        frame.grid_rowconfigure(1, weight=1) # Row 1 (tabela) ganha weight para expandir verticalmente
        frame.grid_rowconfigure(2, weight=0) # Row 2 (legenda) estático
        frame.grid_columnconfigure(0, weight=1)
        
        filter_bar = ctk.CTkFrame(frame, fg_color="#1e293b", corner_radius=10, border_width=1, border_color="#334155")
        filter_bar.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        
        filter_bar.grid_columnconfigure(0, weight=0)
        filter_bar.grid_columnconfigure(1, weight=1)
        filter_bar.grid_columnconfigure(2, weight=0)
        filter_bar.grid_columnconfigure(3, weight=0)
        filter_bar.grid_columnconfigure(4, weight=0)
        filter_bar.grid_columnconfigure(5, weight=0)
        
        ctk.CTkLabel(
            filter_bar, 
            text="Buscar:", 
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#94a3b8"
        ).grid(row=0, column=0, padx=(15, 5), pady=12, sticky="w")
        
        self.entry_busca = ctk.CTkEntry(
            filter_bar, 
            placeholder_text="Nome do aluno...", 
            textvariable=self.filtro_busca_nome,
            fg_color="#0f172a",
            border_color="#334155",
            width=200
        )
        self.entry_busca.grid(row=0, column=1, padx=5, pady=12, sticky="w")
        self.entry_busca.bind("<KeyRelease>", lambda e: self.popula_tabela())
        
        ctk.CTkLabel(
            filter_bar, 
            text="Disciplina:", 
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#94a3b8"
        ).grid(row=0, column=2, padx=(15, 5), pady=12, sticky="w")
        
        subjects = self.get_subjects_active_turma()
        self.combo_disciplina = ctk.CTkOptionMenu(
            filter_bar,
            values=["Todas"] + subjects,
            variable=self.filtro_disciplina,
            command=lambda v: self.popula_tabela(),
            fg_color="#0f172a",
            button_color="#334155",
            width=150
        )
        self.combo_disciplina.grid(row=0, column=3, padx=5, pady=12, sticky="w")

        self.combo_situacao = ctk.CTkOptionMenu(
            filter_bar,
            values=["Filtro Rápido: Todos", "Notas Baixas vs Frequência", "Notas Baixas (< 5.0)", "Frequência Crítica (< 75%)", "Alunos Destaque", "Com Plano de Ação"],
            command=self.on_filtro_rapido_alterado,
            fg_color="#0f172a",
            button_color="#334155",
            width=185
        )
        self.combo_situacao.grid(row=0, column=4, padx=5, pady=12, sticky="w")
        self.combo_situacao.set("Filtro Rápido: Todos")
        
        self.btn_exportar = ctk.CTkButton(
            filter_bar,
            text="Exportar Planilha 📄",
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color="#10b981",
            hover_color="#059669",
            width=130,
            command=self.exportar_excel
        )
        self.btn_exportar.grid(row=0, column=5, padx=(10, 15), pady=12, sticky="e")

        table_container = ctk.CTkFrame(frame, fg_color="#1e293b", corner_radius=10, border_width=1, border_color="#334155")
        table_container.grid(row=1, column=0, sticky="nsew", pady=5)
        table_container.grid_rowconfigure(0, weight=1)
        table_container.grid_columnconfigure(0, weight=1)
        
        ysb = ttk.Scrollbar(table_container, orient="vertical")
        xsb = ttk.Scrollbar(table_container, orient="horizontal")
        
        self.style = ttk.Style()
        self.style.theme_use("default")
        self.atualizar_estilos_treeview('dark')
        
        self.tree = ttk.Treeview(
            table_container,
            yscrollcommand=ysb.set,
            xscrollcommand=xsb.set,
            selectmode="browse"
        )
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        
        ysb.config(command=self.tree.yview)
        xsb.config(command=self.tree.xview)
        
        self.tree.tag_configure("critical", foreground="#f43f5e", background="#3f2225")
        self.tree.tag_configure("warning", foreground="#fb923c", background="#3a2e21")
        self.tree.tag_configure("destaque", foreground="#34d399", background="#1b2e27")
        
        self.tree.bind("<Double-1>", self.on_tabela_double_click)
        
        lbl_info = ctk.CTkLabel(
            frame, 
            text="* Clique duas vezes no aluno para ver o perfil completo e editar o Plano de Ação.", 
            text_color="#94a3b8",
            font=ctk.CTkFont(size=11, slant="italic")
        )
        lbl_info.grid(row=2, column=0, sticky="w", pady=5)

        return frame

    def on_filtro_rapido_alterado(self, value):
        mapeamento = {
            "Filtro Rápido: Todos": "Todos",
            "Notas Baixas vs Frequência": "Notas Baixas vs Frequencia",
            "Notas Baixas (< 5.0)": "Notas Baixas",
            "Frequência Crítica (< 75%)": "Frequencia Critica",
            "Alunos Destaque": "Destaques",
            "Com Plano de Ação": "Com Plano"
        }
        self.filtro_situacao.set(mapeamento.get(value, "Todos"))
        self.popula_tabela()

    def popula_tabela(self):
        b = self.bimestre_ativo.get()
        dados = self.bimestres_dados[b]
        
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        if not dados:
            return
            
        disciplina_selecionada = self.filtro_disciplina.get()
        subjects = self.get_subjects_active_turma(b)
        
        current_combo_values = ["Todas"] + subjects
        self.combo_disciplina.configure(values=current_combo_values)
        if self.filtro_disciplina.get() not in current_combo_values:
            self.filtro_disciplina.set("Todas")
            disciplina_selecionada = "Todas"
            
        situacao = self.filtro_situacao.get()
        
        if situacao == "Notas Baixas vs Frequencia":
            colunas_exibicao = ["aluno", "situacao", "frequencia", "baixas", "media", "plano_acao"]
            cabecalhos = ["Nome do Aluno", "Situação", "Frequência", "N. Vermelhas", "Média Geral", "Disciplinas em Recuperação"]
        elif disciplina_selecionada == "Todas":
            colunas_exibicao = ["aluno", "situacao"] + subjects + ["tf", "fre", "baixas", "media", "plano_acao"]
            cabecalhos = ["Nome do Aluno", "Situação"] + [sub.split()[0] for sub in subjects] + ["TF", "Fre%", "N. Baixas", "Média", "Plano de Ação"]
        else:
            colunas_exibicao = ["aluno", "situacao", "nota_filtrada", "tf", "fre", "baixas", "media", "plano_acao"]
            cabecalhos = ["Nome do Aluno", "Situação", f"Nota: {disciplina_selecionada}", "TF", "Fre%", "Total Baixas", "Média Geral", "Plano de Ação"]
            
        self.tree["columns"] = colunas_exibicao
        self.tree["show"] = "headings"
        
        for idx, col in enumerate(colunas_exibicao):
            self.tree.heading(col, text=cabecalhos[idx], anchor="w" if col in ["aluno", "plano_acao"] else "center")
            if col == "aluno":
                self.tree.column(col, width=220, minwidth=180, stretch=True)
            elif col == "plano_acao":
                self.tree.column(col, width=280, minwidth=150, stretch=True)
            else:
                self.tree.column(col, width=70, minwidth=50, stretch=False, anchor="center")

        busca = normalizar_texto(self.filtro_busca_nome.get())
        
        for aluno in dados:
            if busca and busca not in normalizar_texto(aluno['nome']):
                continue
                
            nota_disc = aluno['notas'].get(disciplina_selecionada) if disciplina_selecionada != "Todas" else None
            
            if situacao == "Notas Baixas":
                if disciplina_selecionada == "Todas":
                    if aluno['notas_baixas'] == 0: continue
                else:
                    if nota_disc is None or nota_disc >= 5.0: continue
            elif situacao == "Frequencia Critica":
                if aluno['frequencia'] >= 75.0: continue
            elif situacao == "Destaques":
                sem_vermelha = True
                for s, n in aluno['notas'].items():
                    if n is not None and n < 5.0:
                        sem_vermelha = False
                        break
                if not sem_vermelha or aluno['media'] < 7.5:
                    continue
            elif situacao == "Com Plano":
                if not aluno['plano_acao'] or aluno['plano_acao'].strip() == "": continue
            elif situacao == "Notas Baixas vs Frequencia":
                if aluno.get('situacao', 'Ativa') != 'Ativa': continue
                if aluno['notas_baixas'] == 0 and aluno['frequencia'] >= 75.0:
                    continue

            tag = ""
            if aluno['frequencia'] < 75.0 and aluno.get('situacao', 'Ativa') == 'Ativa':
                tag = "critical"
            elif aluno['media'] < 5.0 and aluno.get('situacao', 'Ativa') == 'Ativa':
                tag = "warning"
            else:
                sem_vermelha = True
                for s, n in aluno['notas'].items():
                    if n is not None and n < 5.0:
                        sem_vermelha = False
                        break
                if sem_vermelha and aluno['media'] >= 7.5:
                    tag = "destaque"

            if situacao == "Notas Baixas vs Frequencia":
                disc_vermelhas = [sub for sub, nota in aluno['notas'].items() if nota is not None and nota < 5.0]
                disc_str = ", ".join(disc_vermelhas) if disc_vermelhas else "Nenhuma"
                
                valores_linha = [
                    aluno['nome'],
                    aluno.get('situacao', 'Ativa'),
                    f"{aluno['frequencia']:.1f}%",
                    aluno['notas_baixas'],
                    f"{aluno['media']:.1f}",
                    disc_str
                ]
            elif disciplina_selecionada == "Todas":
                valores_linha = [aluno['nome'], aluno.get('situacao', 'Ativa')]
                for sub in subjects:
                    val = aluno['notas'].get(sub)
                    valores_linha.append(f"{val:.1f}" if val is not None else "-")
                valores_linha += [
                    aluno['faltas'],
                    f"{aluno['frequencia']:.1f}%",
                    aluno['notas_baixas'],
                    f"{aluno['media']:.1f}",
                    aluno['plano_acao']
                ]
            else:
                valores_linha = [
                    aluno['nome'],
                    aluno.get('situacao', 'Ativa'),
                    f"{nota_disc:.1f}" if nota_disc is not None else "-",
                    aluno['faltas'],
                    f"{aluno['frequencia']:.1f}%",
                    aluno['notas_baixas'],
                    f"{aluno['media']:.1f}",
                    aluno['plano_acao']
                ]
                
            self.tree.insert("", "end", values=valores_linha, tags=(tag,))

    def on_tabela_double_click(self, event):
        item = self.tree.selection()
        if not item:
            return
        aluno_nome = self.tree.item(item, "values")[0]
        self.aluno_selecionado = aluno_nome
        self.mudar_pagina("detalhes")

    def exportar_excel(self):
        t = self.turma_ativa.get()
        if not t:
            messagebox.showwarning("Aviso", "Nenhuma turma ativa selecionada para exportar.")
            return
            
        b = self.bimestre_ativo.get()
        dados = self.bimestres_dados[b]
        if not dados:
            messagebox.showwarning("Sem Dados", "Não há dados para exportar neste bimestre.")
            return
            
        # Obter filtros ativos da tela
        busca = normalizar_texto(self.filtro_busca_nome.get())
        disciplina_selecionada = self.filtro_disciplina.get()
        situacao = self.filtro_situacao.get()
        
        # 1. Filtrar a lista de alunos de acordo com os filtros de busca e visualização ativos
        dados_filtrados = []
        for aluno in dados:
            if busca and busca not in normalizar_texto(aluno['nome']):
                continue
                
            nota_disc = aluno['notas'].get(disciplina_selecionada) if disciplina_selecionada != "Todas" else None
            
            if situacao == "Notas Baixas":
                if disciplina_selecionada == "Todas":
                    if aluno['notas_baixas'] == 0: continue
                else:
                    if nota_disc is None or nota_disc >= 5.0: continue
            elif situacao == "Frequencia Critica":
                if aluno['frequencia'] >= 75.0: continue
            elif situacao == "Destaques":
                sem_vermelha = True
                for s, n in aluno['notas'].items():
                    if n is not None and n < 5.0:
                        sem_vermelha = False
                        break
                if not sem_vermelha or aluno['media'] < 7.5:
                    continue
            elif situacao == "Com Plano":
                if not aluno['plano_acao'] or aluno['plano_acao'].strip() == "": continue
            elif situacao == "Notas Baixas vs Frequencia":
                if aluno.get('situacao', 'Ativa') not in ['Ativa', 'Ativo', 'ATIVA', 'ATIVO']: continue
                if aluno['notas_baixas'] == 0 and aluno['frequencia'] >= 75.0:
                    continue
            
            dados_filtrados.append(aluno)
            
        if not dados_filtrados:
            messagebox.showwarning("Sem Dados", "Não há alunos correspondentes aos filtros ativos para exportar.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Planilha Excel", "*.xlsx")],
            title=f"Exportar Planilha Filtrada — Turma {t}"
        )
        if not file_path:
            return
            
        try:
            subjects = self.get_subjects_active_turma(b)
            
            # 2. Definir as colunas e formatar as linhas exatamente como aparecem na tela (conforme o filtro selecionado)
            if situacao == "Notas Baixas vs Frequencia":
                colunas_cabecalho = ["Nome do Aluno", "Situação", "Frequência", "N. Vermelhas", "Média Geral", "Disciplinas em Recuperação"]
                linhas = []
                for al in dados_filtrados:
                    disc_vermelhas = [sub for sub, nota in al['notas'].items() if nota is not None and nota < 5.0]
                    disc_str = ", ".join(disc_vermelhas) if disc_vermelhas else "Nenhuma"
                    linha = [
                        al['nome'],
                        al.get('situacao', 'Ativa'),
                        al['frequencia'] / 100.0,  # Convertido para float para formatar como % no Excel
                        al['notas_baixas'],
                        al['media'],
                        disc_str
                    ]
                    linhas.append(linha)
                    
            elif disciplina_selecionada == "Todas":
                colunas_cabecalho = ["Nome do Aluno", "Situação"] + [sub.split()[0] for sub in subjects] + ["TF", "Fre%", "N. Baixas", "Média", "Plano de Ação"]
                linhas = []
                for al in dados_filtrados:
                    linha = [al['nome'], al.get('situacao', 'Ativa')]
                    for sub in subjects:
                        linha.append(al['notas'].get(sub))
                    linha += [
                        al['faltas'],
                        al['frequencia'] / 100.0,
                        al['notas_baixas'],
                        al['media'],
                        al['plano_acao']
                    ]
                    linhas.append(linha)
                    
            else:
                colunas_cabecalho = ["Nome do Aluno", "Situação", f"Nota: {disciplina_selecionada}", "TF", "Fre%", "Total Baixas", "Média Geral", "Plano de Ação"]
                linhas = []
                for al in dados_filtrados:
                    nota_disc = al['notas'].get(disciplina_selecionada)
                    linha = [
                        al['nome'],
                        al.get('situacao', 'Ativa'),
                        nota_disc,
                        al['faltas'],
                        al['frequencia'] / 100.0,
                        al['notas_baixas'],
                        al['media'],
                        al['plano_acao']
                    ]
                    linhas.append(linha)
                    
            # 3. Criar DataFrame e salvar no Excel formatando as colunas
            df_export = pd.DataFrame(linhas, columns=colunas_cabecalho)
            
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                sheet_name_clean = f"Filtro {situacao[:10]}" if len(situacao) > 10 else f"Filtro {situacao}"
                df_export.to_excel(writer, sheet_name=sheet_name_clean, index=False)
                
                workbook = writer.book
                worksheet = writer.sheets[sheet_name_clean]
                
                # Configurar formatação percentual nas colunas de Frequência
                for col_name in ["Frequência", "Fre%"]:
                    if col_name in colunas_cabecalho:
                        idx_col = colunas_cabecalho.index(col_name) + 1
                        for r in range(2, len(linhas) + 2):
                            cell = worksheet.cell(row=r, column=idx_col)
                            if cell.value is not None and isinstance(cell.value, (int, float)):
                                cell.number_format = '0.0%'
                                
                # Ajustar automaticamente as larguras das colunas para visualização premium
                for col in worksheet.columns:
                    max_len = 0
                    col_letter = col[0].column_letter
                    for cell in col:
                        if cell.value is not None:
                            val_str = str(cell.value)
                            if cell.number_format == '0.0%':
                                try:
                                    val_str = f"{float(cell.value) * 100:.1f}%"
                                except ValueError:
                                    pass
                            max_len = max(max_len, len(val_str))
                    worksheet.column_dimensions[col_letter].width = max(max_len + 3, 10)
                    
            messagebox.showinfo("Exportado", f"Planilha filtrada salva com sucesso em:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Erro ao Exportar", f"Falha ao gerar o arquivo Excel: {e}")
            
    # ==========================================================================
    # 4. TELA DE GRÁFICOS (MATPLOTLIB EM TKINTER)
    # ==========================================================================
    def criar_tela_graficos(self):
        frame = ctk.CTkFrame(self.content_container, fg_color="transparent")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=65) # 65% largura para o gráfico
        frame.grid_columnconfigure(1, weight=35) # 35% largura para os controles e textos
        
        self.chart_container_frame = ctk.CTkFrame(frame, fg_color="#1e293b", corner_radius=12, border_width=1, border_color="#334155")
        self.chart_container_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=10)
        
        lbl_msg = ctk.CTkLabel(
            self.chart_container_frame, 
            text="Selecione um grafico para visualizar.", 
            font=ctk.CTkFont(size=14),
            text_color="#94a3b8"
        )
        lbl_msg.pack(expand=True)
        
        ctrl_panel = ctk.CTkFrame(frame, fg_color="#1e293b", corner_radius=12, border_width=1, border_color="#334155")
        ctrl_panel.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=10)
        ctrl_panel.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(
            ctrl_panel, 
            text="📊 Gráficos Disponíveis", 
            font=ctk.CTkFont(size=15, weight="bold")
        ).grid(row=0, column=0, padx=15, pady=(15, 10), sticky="w")
        
        self.tipo_grafico_ativo = ctk.StringVar(value="medias")
        
        btn_medias = ctk.CTkButton(
            ctrl_panel, 
            text="Médias por Disciplina", 
            anchor="w",
            fg_color="#334155",
            command=lambda: self.alterar_tipo_grafico("medias")
        )
        btn_medias.grid(row=1, column=0, padx=15, pady=6, sticky="ew")
        
        btn_baixas_faltas = ctk.CTkButton(
            ctrl_panel, 
            text="Notas Baixas vs. Faltas", 
            anchor="w",
            fg_color="#334155",
            command=lambda: self.alterar_tipo_grafico("baixas_faltas")
        )
        btn_baixas_faltas.grid(row=2, column=0, padx=15, pady=6, sticky="ew")

        self.btn_graficos_map = {
            "medias": btn_medias,
            "baixas_faltas": btn_baixas_faltas
        }
        
        lbl_desc_title = ctk.CTkLabel(
            ctrl_panel,
            text="📝 Análise e Detalhes do Gráfico",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#38bdf8"
        )
        lbl_desc_title.grid(row=3, column=0, padx=15, pady=(20, 5), sticky="w")
        
        self.txt_graficos_desc = tk.Text(
            ctrl_panel,
            bg="#0f172a",
            fg="#e2e8f0",
            
            relief="flat",
            font=("Segoe UI", 10),
            padx=10,
            pady=10,
            wrap="word",
            height=7
        )
        self.txt_graficos_desc.grid(row=4, column=0, sticky="nsew", padx=15, pady=5)

        lbl_stats_title = ctk.CTkLabel(
            ctrl_panel, 
            text="📊 Estatísticas Rápidas", 
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#94a3b8"
        )
        lbl_stats_title.grid(row=5, column=0, padx=15, pady=(15, 5), sticky="w")
        
        self.txt_quick_stats = tk.Text(
            ctrl_panel, 
            bg="#0f172a", 
            fg="#cbd5e1", 
            bd=0, 
            font=("Segoe UI", 10),
            padx=10,
            pady=10,
            wrap="word",
            height=5
        )
        self.txt_quick_stats.grid(row=6, column=0, padx=15, pady=(0, 15), sticky="nsew")
        
        # Permitir que as caixas de textos se expandam
        ctrl_panel.grid_rowconfigure(4, weight=1)
        ctrl_panel.grid_rowconfigure(6, weight=1)
        
        return frame

    def alterar_tipo_grafico(self, tipo):
        self.tipo_grafico_ativo.set(tipo)
        self.desenhar_graficos_ativos()

    def desenhar_graficos_ativos(self):
        b = self.bimestre_ativo.get()
        dados = self.bimestres_dados[b]
        theme = "dark" if self.theme_switch.get() == 1 else "light"
        
        for k, btn in self.btn_graficos_map.items():
            if k == self.tipo_grafico_ativo.get():
                btn.configure(fg_color="#0d9488")
            else:
                btn.configure(fg_color="#334155")
                
        if not dados:
            graficos.limpar_container(self.chart_container_frame)
            lbl = ctk.CTkLabel(
                self.chart_container_frame, 
                text="Por favor, carregue planilhas ou dados de demonstracao\nna aba 'Importar Planilhas' primeiro.",
                text_color="#94a3b8"
            )
            lbl.pack(expand=True)
            self.txt_quick_stats.delete("1.0", "end")
            self.txt_graficos_desc.delete("1.0", "end")
            return
            
        tipo = self.tipo_grafico_ativo.get()
        alunos_ativos = [al for al in dados if al.get('situacao', 'Ativa') == 'Ativa']
        subjects = self.get_subjects_active_turma(b)
        
        if tipo == "medias":
            medias = {}
            for sub in subjects:
                notas_validas = [aluno['notas'][sub] for aluno in alunos_ativos if aluno['notas'].get(sub) is not None]
                if notas_validas:
                    medias[sub] = sum(notas_validas) / len(notas_validas)
            graficos.plotar_medias_disciplinas(self.chart_container_frame, medias, theme=theme)
            
            self.txt_graficos_desc.delete("1.0", "end")
            self.txt_graficos_desc.insert("1.0", 
                "Análise Geral das Médias por Disciplina:\n\n"
                "Este gráfico de barras horizontal apresenta a média coletiva da sala de aula em cada matéria. "
                "As disciplinas estão ordenadas do menor rendimento para o maior, facilitando "
                "o diagnóstico de eixos pedagógicos que necessitam de intervenção ou revisão curricular."
            )
            self.txt_quick_stats.delete("1.0", "end")
            if medias:
                melhor_disc = max(medias.items(), key=lambda x: x[1])
                pior_disc = min(medias.items(), key=lambda x: x[1])
                self.txt_quick_stats.insert("end", 
                     f"Melhor Desempenho:\n{melhor_disc[0].split()[0]}: {melhor_disc[1]:.1f}\n\n"
                     f"Maior Dificuldade:\n{pior_disc[0].split()[0]}: {pior_disc[1]:.1f}\n\n"
                     f"Disciplinas abaixo de 6.0:\n" + 
                     "\n".join([f"- {k.split()[0]}: {v:.1f}" for k,v in medias.items() if v < 6.0])
                )
                
        elif tipo == "baixas_faltas":
            dados_comparativos = []
            for aluno in alunos_ativos:
                dados_comparativos.append({
                    'nome': aluno['nome'],
                    'notas_baixas': aluno['notas_baixas'],
                    'faltas': aluno['faltas'],
                    'frequencia': aluno['frequencia']
                })
            graficos.plotar_notas_baixas_vs_faltas(self.chart_container_frame, dados_comparativos, theme=theme)
            
            self.txt_graficos_desc.delete("1.0", "end")
            self.txt_graficos_desc.insert("1.0", 
                "Top Alunos: Relação Notas Baixas vs. Faltas:\n\n"
                "Este gráfico correlaciona a quantidade de notas vermelhas (abaixo de 5.0) com o total de faltas acumuladas (aulas perdidas). "
                "Permite analisar a relação direta de causa e efeito entre a infrequência escolar de cada estudante e o seu rendimento acadêmico."
            )
            self.txt_quick_stats.delete("1.0", "end")
            alunos_cr_faltas = sum(1 for x in dados_comparativos if x['frequencia'] < 75)
            correlacao = "Forte" if any(x['notas_baixas'] > 3 and x['frequencia'] < 75 for x in dados_comparativos) else "Moderada"
            self.txt_quick_stats.insert("end",
                f"Alunos c/ Freq < 75%:\n {alunos_cr_faltas} alunos\n\n"
                f"Correlacao Faltas x Rendimento:\n {correlacao}\n\n"
                f"Alunos em Alerta Pedagógico:\n" + 
                "\n".join([f"- {x['nome'].split()[0]}: {x['notas_baixas']} nb / {x['faltas']} f" for x in sorted(dados_comparativos, key=lambda x: x['notas_baixas'], reverse=True)[:3]])
            )

    # ==========================================================================
    # 5. TELA DE COMPARATIVO DE BIMESTRES (HISTÓRICO DA SALA E ALUNOS)
    # ==========================================================================
    def criar_tela_comparativo(self):
        frame = ctk.CTkFrame(self.content_container, fg_color="transparent")
        frame.grid_rowconfigure(1, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        
        header = ctk.CTkFrame(frame, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        
        lbl_title = ctk.CTkLabel(
            header, 
            text="🔄 Comparativo entre Bimestres", 
            font=ctk.CTkFont(size=22, weight="bold")
        )
        lbl_title.pack(anchor="w", side="left")
        
        self.combo_comp_aluno = ctk.CTkOptionMenu(
            header,
            values=["Todos (Sala)"],
            variable=self.comparativo_aluno_filtro,
            command=lambda v: self.desenhar_comparativo_bimestres(),
            fg_color="#1e293b",
            button_color="#334155",
            width=200
        )
        self.combo_comp_aluno.pack(anchor="e", side="right")
        
        lbl_aluno_filter = ctk.CTkLabel(
            header,
            text="Filtrar por Aluno:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#94a3b8"
        )
        lbl_aluno_filter.pack(anchor="e", side="right", padx=(10, 5))

        self.combo_comp_disciplina = ctk.CTkOptionMenu(
            header,
            values=["Todas (Média Geral)"],
            variable=self.comparativo_disciplina_filtro,
            command=lambda v: self.desenhar_comparativo_bimestres(),
            fg_color="#1e293b",
            button_color="#334155",
            width=180
        )
        self.combo_comp_disciplina.pack(anchor="e", side="right", padx=(10, 0))
        
        lbl_disc_filter = ctk.CTkLabel(
            header,
            text="Disciplina:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#94a3b8"
        )
        lbl_disc_filter.pack(anchor="e", side="right", padx=(0, 5))
        
        split_frame = ctk.CTkFrame(frame, fg_color="transparent")
        split_frame.grid(row=1, column=0, sticky="nsew")
        split_frame.grid_rowconfigure(0, weight=1)
        split_frame.grid_columnconfigure(0, weight=40) # 40% largura para o gráfico (reduzido)
        split_frame.grid_columnconfigure(1, weight=60) # 60% largura para a tabela (aumentado)
        
        self.comp_chart_frame = ctk.CTkFrame(split_frame, fg_color="#1e293b", corner_radius=12, border_width=1, border_color="#334155")
        self.comp_chart_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        self.comp_table_frame = ctk.CTkFrame(split_frame, fg_color="#1e293b", corner_radius=12, border_width=1, border_color="#334155")
        self.comp_table_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        self.comp_table_frame.grid_rowconfigure(1, weight=1)
        self.comp_table_frame.grid_columnconfigure(0, weight=1)
        
        self.lbl_comp_table_title = ctk.CTkLabel(
            self.comp_table_frame,
            text="Médias Comparativas das Disciplinas",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#38bdf8"
        )
        self.lbl_comp_table_title.grid(row=0, column=0, padx=15, pady=10, sticky="w")
        
        self.tree_comp_medias = ttk.Treeview(self.comp_table_frame, columns=["disciplina", "b1", "b2", "b3", "b4", "evolucao"], show="headings")
        self.tree_comp_medias.grid(row=1, column=0, sticky="nsew", padx=15, pady=(0, 15))
        
        self.tree_comp_medias.heading("disciplina", text="Disciplina", anchor="w")
        self.tree_comp_medias.heading("b1", text="1º Bim")
        self.tree_comp_medias.heading("b2", text="2º Bim")
        self.tree_comp_medias.heading("b3", text="3º Bim")
        self.tree_comp_medias.heading("b4", text="4º Bim")
        self.tree_comp_medias.heading("evolucao", text="Evolucao")
        
        self.tree_comp_medias.column("disciplina", width=140, anchor="w")
        self.tree_comp_medias.column("b1", width=55, anchor="center")
        self.tree_comp_medias.column("b2", width=55, anchor="center")
        self.tree_comp_medias.column("b3", width=55, anchor="center")
        self.tree_comp_medias.column("b4", width=55, anchor="center")
        self.tree_comp_medias.column("evolucao", width=70, anchor="center")
        
        return frame

    def atualizar_combo_comp_alunos(self):
        t = self.turma_ativa.get()
        if not t or t not in self.turmas_dados:
            self.combo_comp_aluno.configure(values=["Todos (Sala)"])
            self.combo_comp_aluno.set("Todos (Sala)")
            self.combo_comp_disciplina.configure(values=["Todas (Média Geral)"])
            self.combo_comp_disciplina.set("Todas (Média Geral)")
            return
            
        nomes = set()
        for b in [1, 2, 3, 4]:
            dados = self.turmas_dados[t].get(b)
            if dados:
                nomes.update([al['nome'] for al in dados if al.get('situacao', 'Ativa') == 'Ativa'])
                
        valid_nomes = ["Todos (Sala)"] + sorted(list(nomes))
        self.combo_comp_aluno.configure(values=valid_nomes)
        
        if self.comparativo_aluno_filtro.get() not in valid_nomes:
            self.comparativo_aluno_filtro.set("Todos (Sala)")
        self.combo_comp_aluno.set(self.comparativo_aluno_filtro.get())
        
        subjects = self.get_subjects_active_turma()
        self.combo_comp_disciplina.configure(values=["Todas (Média Geral)"] + subjects)
        if self.comparativo_disciplina_filtro.get() not in ["Todas (Média Geral)"] + subjects:
            self.comparativo_disciplina_filtro.set("Todas (Média Geral)")
            self.combo_comp_disciplina.set("Todas (Média Geral)")

    def desenhar_comparativo_bimestres(self):
        bimestres_carregados = [b for b, dados in self.bimestres_dados.items() if dados is not None]
        theme = "dark" if self.theme_switch.get() == 1 else "light"
        
        for item in self.tree_comp_medias.get_children():
            self.tree_comp_medias.delete(item)
            
        if not bimestres_carregados:
            graficos.limpar_container(self.comp_chart_frame)
            lbl = ctk.CTkLabel(
                self.comp_chart_frame, 
                text="Por favor, carregue pelo menos um bimestre na aba 'Importar Planilhas' para comparar.",
                text_color="#94a3b8"
            )
            lbl.pack(expand=True)
            return

        aluno_filtro = self.comparativo_aluno_filtro.get()
        disc_selecionada = self.comparativo_disciplina_filtro.get()
        subjects = self.get_subjects_active_turma()
        
        if aluno_filtro == "Todos (Sala)":
            if disc_selecionada == "Todas (Média Geral)":
                self.lbl_comp_table_title.configure(text="Médias Comparativas das Disciplinas")
            else:
                self.lbl_comp_table_title.configure(text=f"Médias da Sala em {disc_selecionada}")
                
            historico_turma = {}
            for b in bimestres_carregados:
                dados = self.bimestres_dados[b]
                alunos_ativos = [al for al in dados if al.get('situacao', 'Ativa') == 'Ativa']
                if alunos_ativos:
                    if disc_selecionada == "Todas (Média Geral)":
                        avg_m = sum(aluno['media'] for aluno in alunos_ativos) / len(alunos_ativos)
                    else:
                        notas_validas = [aluno['notas'][disc_selecionada] for aluno in alunos_ativos if aluno['notas'].get(disc_selecionada) is not None]
                        avg_m = sum(notas_validas) / len(notas_validas) if notas_validas else 0.0
                    avg_f = sum(aluno.get('faltas', 0) for aluno in alunos_ativos) / len(alunos_ativos)
                    historico_turma[b] = {'media': avg_m, 'presenca': avg_f}
                
            graficos.plotar_evolucao_turma(self.comp_chart_frame, historico_turma, theme=theme, disciplina=disc_selecionada)
            
            medias_por_b = {b: {} for b in [1, 2, 3, 4]}
            for b in bimestres_carregados:
                dados = self.bimestres_dados[b]
                alunos_ativos = [al for al in dados if al.get('situacao', 'Ativa') == 'Ativa']
                for sub in subjects:
                    notas_validas = [aluno['notas'][sub] for aluno in alunos_ativos if aluno['notas'].get(sub) is not None]
                    medias_por_b[b][sub] = sum(notas_validas) / len(notas_validas) if notas_validas else None
            
            for sub in subjects:
                b1 = medias_por_b[1].get(sub)
                b2 = medias_por_b[2].get(sub)
                b3 = medias_por_b[3].get(sub)
                b4 = medias_por_b[4].get(sub)
                
                valores_validos = [v for v in [b1, b2, b3, b4] if v is not None]
                if len(valores_validos) >= 2:
                    diff = valores_validos[-1] - valores_validos[0]
                    seta = "▲" if diff > 0 else "▼" if diff < 0 else "="
                    evolucao = f"{seta} {diff:+.1f}"
                else:
                    evolucao = "-"
                    
                valores_linha = [
                    sub,
                    f"{b1:.1f}" if b1 is not None else "-",
                    f"{b2:.1f}" if b2 is not None else "-",
                    f"{b3:.1f}" if b3 is not None else "-",
                    f"{b4:.1f}" if b4 is not None else "-",
                    evolucao
                ]
                self.tree_comp_medias.insert("", "end", values=valores_linha)
        else:
            if disc_selecionada == "Todas (Média Geral)":
                self.lbl_comp_table_title.configure(text=f"Boletim: {aluno_filtro}")
            else:
                self.lbl_comp_table_title.configure(text=f"Histórico em {disc_selecionada}: {aluno_filtro}")
            
            historico_aluno = {}
            notes_bimestrais = {b: {} for b in [1, 2, 3, 4]}
            
            for b in bimestres_carregados:
                dados = self.bimestres_dados[b]
                al_reg = next((x for x in dados if x['nome'] == aluno_filtro), None)
                if al_reg:
                    if disc_selecionada == "Todas (Média Geral)":
                        m_val = al_reg['media']
                    else:
                        m_val = al_reg['notas'].get(disc_selecionada)
                        if m_val is None: m_val = 0.0
                    historico_aluno[b] = {'media': m_val, 'presenca': al_reg['frequencia'], 'faltas': al_reg.get('faltas', 0)}
                    notes_bimestrais[b] = al_reg['notas']
            
            graficos.plotar_evolucao_aluno(self.comp_chart_frame, historico_aluno, aluno_filtro, theme=theme, disciplina=disc_selecionada)
            
            for sub in subjects:
                b1 = notes_bimestrais[1].get(sub)
                b2 = notes_bimestrais[2].get(sub)
                b3 = notes_bimestrais[3].get(sub)
                b4 = notes_bimestrais[4].get(sub)
                
                valores_validos = [v for v in [b1, b2, b3, b4] if v is not None]
                if len(valores_validos) >= 2:
                    diff = valores_validos[-1] - valores_validos[0]
                    seta = "▲" if diff > 0 else "▼" if diff < 0 else "="
                    evolucao = f"{seta} {diff:+.1f}"
                else:
                    evolucao = "-"
                    
                valores_linha = [
                    sub,
                    f"{b1:.1f}" if b1 is not None else "-",
                    f"{b2:.1f}" if b2 is not None else "-",
                    f"{b3:.1f}" if b3 is not None else "-",
                    f"{b4:.1f}" if b4 is not None else "-",
                    evolucao
                ]
                self.tree_comp_medias.insert("", "end", values=valores_linha)


    # ==========================================================================
    # 6. TELA DE DETALHES DO ALUNO (EDITAR PLANO DE AÇÃO E HISTÓRICO INDIVIDUAL)
    # ==========================================================================
    def criar_tela_detalhes(self):
        frame = ctk.CTkFrame(self.content_container, fg_color="transparent")
        frame.grid_rowconfigure(1, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        
        header = ctk.CTkFrame(frame, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        
        lbl_title = ctk.CTkLabel(
            header, 
            text="👤 Perfil e Acompanhamento do Aluno", 
            font=ctk.CTkFont(size=22, weight="bold")
        )
        lbl_title.pack(anchor="w", side="left")
        
        self.btn_pdf_aluno = ctk.CTkButton(
            header,
            text="Exportar Relatório PDF 📄",
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#0d9488",
            hover_color="#0f766e",
            command=self.exportar_pdf_aluno
        )
        self.btn_pdf_aluno.pack(anchor="e", side="right", padx=(10, 0))
        
        self.combo_alunos = ctk.CTkOptionMenu(
            header,
            values=["Nenhum aluno carregado"],
            command=self.on_aluno_selecionado_changed,
            fg_color="#1e293b",
            button_color="#334155",
            width=260
        )
        self.combo_alunos.pack(anchor="e", side="right")

        self.combo_detalhes_disciplina = ctk.CTkOptionMenu(
            header,
            values=["Todas (Média Geral)"],
            variable=self.detalhes_disciplina_filtro,
            command=lambda v: self.preencher_detalhes_aluno(),
            fg_color="#1e293b",
            button_color="#334155",
            width=180
        )
        self.combo_detalhes_disciplina.pack(anchor="e", side="right", padx=(10, 10))
        
        lbl_detalhes_disc = ctk.CTkLabel(
            header,
            text="Disciplina:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#94a3b8"
        )
        lbl_detalhes_disc.pack(anchor="e", side="right", padx=(0, 5))
        
        split_frame = ctk.CTkFrame(frame, fg_color="transparent")
        split_frame.grid(row=1, column=0, sticky="nsew")
        split_frame.grid_rowconfigure(0, weight=28)   # 28% de altura para boletim e plano de ação (reduzido)
        split_frame.grid_rowconfigure(1, weight=72)   # 72% de altura para dar foco total aos gráficos (aumentado)
        split_frame.grid_columnconfigure(0, weight=75) # 75% de largura para tabela e gráfico linha
        split_frame.grid_columnconfigure(1, weight=25) # 25% de largura para plano de ação e radar
        
        # Tabela de Notas colocada diretamente no split_frame (coluna 0, linha 0)
        self.card_historico_notas = ctk.CTkFrame(split_frame, fg_color="#1e293b", corner_radius=12, border_width=1, border_color="#334155")
        self.card_historico_notas.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=(0, 10))
        self.card_historico_notas.grid_rowconfigure(1, weight=1)
        self.card_historico_notas.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(
            self.card_historico_notas,
            text="Notas Detalhadas por Bimestre",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#38bdf8"
        ).grid(row=0, column=0, padx=15, pady=10, sticky="w")
        
        self.tree_detalhes_notas = ttk.Treeview(self.card_historico_notas, columns=["disciplina", "b1", "b2", "b3", "b4", "situacao"], show="headings", height=5)
        self.tree_detalhes_notas.grid(row=1, column=0, sticky="nsew", padx=(15, 0), pady=(0, 15))
        
        # Barra de rolagem para o boletim compacto
        self.card_historico_notas.grid_columnconfigure(1, weight=0)
        tree_scroll = ttk.Scrollbar(self.card_historico_notas, orient="vertical", command=self.tree_detalhes_notas.yview)
        self.tree_detalhes_notas.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.grid(row=1, column=1, sticky="ns", pady=(0, 15), padx=(0, 15))
        
        self.tree_detalhes_notas.heading("disciplina", text="Disciplina", anchor="w")
        self.tree_detalhes_notas.heading("b1", text="1º Bim")
        self.tree_detalhes_notas.heading("b2", text="2º Bim")
        self.tree_detalhes_notas.heading("b3", text="3º Bim")
        self.tree_detalhes_notas.heading("b4", text="4º Bim")
        self.tree_detalhes_notas.heading("situacao", text="Situacao")
        
        self.tree_detalhes_notas.column("disciplina", width=140, anchor="w")
        self.tree_detalhes_notas.column("b1", width=55, anchor="center")
        self.tree_detalhes_notas.column("b2", width=55, anchor="center")
        self.tree_detalhes_notas.column("b3", width=55, anchor="center")
        self.tree_detalhes_notas.column("b4", width=55, anchor="center")
        self.tree_detalhes_notas.column("situacao", width=85, anchor="center")
        
        # Gráfico Linha colocado diretamente no split_frame (coluna 0, linha 1)
        self.card_grafico_aluno_linha = ctk.CTkFrame(split_frame, fg_color="#0f172a", corner_radius=12, border_width=2, border_color="#38bdf8")
        self.card_grafico_aluno_linha.grid(row=1, column=0, sticky="nsew", padx=(0, 10), pady=(10, 0))
        
        # Gráfico Radar colocado diretamente no split_frame (coluna 1, linha 1)
        self.card_grafico_aluno_radar = ctk.CTkFrame(split_frame, fg_color="#0f172a", corner_radius=12, border_width=2, border_color="#38bdf8")
        self.card_grafico_aluno_radar.grid(row=1, column=1, sticky="nsew", padx=(10, 0), pady=(10, 0))
        
        # Plano de Ação colocado diretamente no split_frame (coluna 1, linha 0)
        right_panel = ctk.CTkFrame(split_frame, fg_color="#1e293b", corner_radius=12, border_width=1, border_color="#334155")
        right_panel.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=(0, 10))
        right_panel.grid_rowconfigure(2, weight=1)
        right_panel.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(
            right_panel,
            text="Encaminhamento / Plano de Ação",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#f59e0b"
        ).grid(row=0, column=0, padx=20, pady=(15, 2), sticky="w")
        
        ctk.CTkLabel(
            right_panel,
            text="Escreva a estrategia e orientacoes pedagogicas para o aluno.",
            font=ctk.CTkFont(size=11),
            text_color="#94a3b8"
        ).grid(row=1, column=0, padx=20, pady=(0, 15), sticky="w")
        
        self.txt_plano_acao_aluno = tk.Text(
            right_panel,
            bg="#0f172a",
            fg="#e2e8f0",
            
            relief="flat",
            font=("Segoe UI", 11),
            height=5, # Limitar a altura do campo de texto
            padx=15,
            pady=15,
            wrap="word"
        )
        self.txt_plano_acao_aluno.grid(row=2, column=0, sticky="nsew", padx=20, pady=(0, 15))
        
        self.btn_salvar_plano = ctk.CTkButton(
            right_panel,
            text="Salvar Plano de Acao do Aluno 💾",
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#0ea5e9",
            hover_color="#0284c7",
            command=self.salvar_plano_aluno
        )
        self.btn_salvar_plano.grid(row=3, column=0, padx=20, pady=(0, 20), sticky="ew")
        
        return frame

    def on_aluno_selecionado_changed(self, value):
        self.aluno_selecionado = value
        self.preencher_detalhes_aluno()

    def preencher_detalhes_aluno(self):
        b_ativo = self.bimestre_ativo.get()
        
        # Obter todos os nomes de alunos ativos da turma em qualquer bimestre carregado
        nomes_alunos = set()
        for b in [1, 2, 3, 4]:
            dados_b = self.bimestres_dados.get(b)
            if dados_b:
                nomes_alunos.update([al['nome'] for al in dados_b if al.get('situacao', 'Ativa') == 'Ativa'])
        nomes_alunos = sorted(list(nomes_alunos))
        
        if not nomes_alunos:
            self.combo_alunos.configure(values=["Nenhum aluno carregado"])
            self.combo_alunos.set("Nenhum aluno carregado")
            self.combo_detalhes_disciplina.configure(values=["Todas (Média Geral)"])
            self.combo_detalhes_disciplina.set("Todas (Média Geral)")
            self.txt_plano_acao_aluno.delete("1.0", "end")
            for item in self.tree_detalhes_notas.get_children():
                self.tree_detalhes_notas.delete(item)
            graficos.limpar_container(self.card_grafico_aluno_linha)
            graficos.limpar_container(self.card_grafico_aluno_radar)
            return
            
        self.combo_alunos.configure(values=nomes_alunos)
        
        if self.aluno_selecionado not in nomes_alunos:
            self.aluno_selecionado = nomes_alunos[0]
            
        self.combo_alunos.set(self.aluno_selecionado)
        
        subjects = self.get_subjects_active_turma()
        self.combo_detalhes_disciplina.configure(values=["Todas (Média Geral)"] + subjects)
        if self.detalhes_disciplina_filtro.get() not in ["Todas (Média Geral)"] + subjects:
            self.detalhes_disciplina_filtro.set("Todas (Média Geral)")
            self.combo_detalhes_disciplina.set("Todas (Média Geral)")
            
        disc_selecionada = self.detalhes_disciplina_filtro.get()
        
        historico_aluno = {}
        notas_bimestrais = {b: {} for b in [1, 2, 3, 4]}
        observacao_ativa = ""
        
        for b in [1, 2, 3, 4]:
            dados_b = self.bimestres_dados[b]
            if dados_b:
                al_reg = next((x for x in dados_b if x['nome'] == self.aluno_selecionado), None)
                if al_reg:
                    if disc_selecionada == "Todas (Média Geral)":
                        m_val = al_reg['media']
                    else:
                        m_val = al_reg['notas'].get(disc_selecionada)
                        if m_val is None: m_val = 0.0
                    historico_aluno[b] = {'media': m_val, 'presenca': al_reg['frequencia'], 'faltas': al_reg.get('faltas', 0)}
                    notas_bimestrais[b] = al_reg['notas']
                    if b == b_ativo or (not observacao_ativa and al_reg.get('plano_acao')):
                        observacao_ativa = al_reg['plano_acao']

        self.txt_plano_acao_aluno.delete("1.0", "end")
        self.txt_plano_acao_aluno.insert("1.0", observacao_ativa)
        
        for item in self.tree_detalhes_notas.get_children():
            self.tree_detalhes_notas.delete(item)
            
        for sub in subjects:
            n1 = notas_bimestrais[1].get(sub)
            n2 = notas_bimestrais[2].get(sub)
            n3 = notas_bimestrais[3].get(sub)
            n4 = notas_bimestrais[4].get(sub)
            
            notas_validas = [v for v in [n1, n2, n3, n4] if v is not None]
            media_sub = sum(notas_validas) / len(notas_validas) if notas_validas else 0.0
            situacao = "Aprovado" if media_sub >= 5.0 else "Em Risco" if media_sub > 0 else "-"
            
            valores_linha = [
                sub,
                f"{n1:.1f}" if n1 is not None else "-",
                f"{n2:.1f}" if n2 is not None else "-",
                f"{n3:.1f}" if n3 is not None else "-",
                f"{n4:.1f}" if n4 is not None else "-",
                situacao
            ]
            self.tree_detalhes_notas.insert("", "end", values=valores_linha)
            
        theme = "dark" if self.theme_switch.get() == 1 else "light"
        graficos.plotar_evolucao_aluno(self.card_grafico_aluno_linha, historico_aluno, self.aluno_selecionado, theme=theme, disciplina=disc_selecionada)
        
        notas_ativas = notas_bimestrais[b_ativo]
        if not notas_ativas:
            # Fallback para as primeiras notas carregadas disponíveis
            for b in [1, 2, 3, 4]:
                if notas_bimestrais[b]:
                    notas_ativas = notas_bimestrais[b]
                    break
        graficos.plotar_radar_competencias(self.card_grafico_aluno_radar, notas_ativas, theme=theme)

    def salvar_plano_aluno(self):
        if not self.aluno_selecionado:
            return
            
        b_ativo = self.bimestre_ativo.get()
        texto = self.txt_plano_acao_aluno.get("1.0", "end-1c").strip()
        
        dados_b = self.bimestres_dados[b_ativo]
        if dados_b:
            for idx, aluno in enumerate(dados_b):
                if aluno['nome'] == self.aluno_selecionado:
                    self.bimestres_dados[b_ativo][idx]['plano_acao'] = texto
                    break
                    
        self.salvar_persistência()
        messagebox.showinfo("Salvo", f"Plano de Ação de {self.aluno_selecionado.split()[0]} atualizado no {b_ativo}º Bimestre!")

    # ==========================================================================
    # GERADOR DE PDFS
    # ==========================================================================
    def exportar_pdf_aluno(self):
        aluno_nome = self.aluno_selecionado
        if not aluno_nome:
            messagebox.showwarning("Aviso", "Selecione um aluno para exportar.")
            return
            
        turma_nome = self.turma_ativa.get()
        b_ativo = self.bimestre_ativo.get()
        disc_selecionada = self.detalhes_disciplina_filtro.get()
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Documento PDF", "*.pdf")],
            title=f"Exportar Relatório — {aluno_nome.split()[0]}",
            initialfile=f"Relatorio_{aluno_nome.split()[0]}_{turma_nome.replace(' ', '_')}.pdf"
        )
        if not file_path:
            return
            
        notas_bimestrais = {b: {} for b in [1, 2, 3, 4]}
        frequencias_bimestrais = {b: 100.0 for b in [1, 2, 3, 4]}
        plano_acao = ""
        historico_completo = {}
        
        for b in [1, 2, 3, 4]:
            dados_b = self.turmas_dados.get(turma_nome, {}).get(b)
            if dados_b:
                al_reg = next((x for x in dados_b if x['nome'] == aluno_nome), None)
                if al_reg:
                    notas_bimestrais[b] = al_reg['notas']
                    frequencias_bimestrais[b] = al_reg['frequencia']
                    if disc_selecionada == "Todas (Média Geral)":
                        m_val = al_reg['media']
                    else:
                        m_val = al_reg['notas'].get(disc_selecionada)
                        if m_val is None: m_val = 0.0
                    historico_completo[b] = {'media': m_val, 'presenca': al_reg['frequencia']}
                    if b == b_ativo:
                        plano_acao = al_reg['plano_acao']
                        
        temp_png = os.path.join(os.getcwd(), f"temp_chart_{aluno_nome.split()[0]}.png")
        graficos.exportar_evolucao_aluno_png(temp_png, historico_completo, aluno_nome, disc_selecionada)
        
        try:
            relatorios_pdf.gerar_pdf_aluno(
                caminho_salvar=file_path,
                aluno_nome=aluno_nome,
                turma_nome=turma_nome,
                bimester_ativo=b_ativo,
                notas_bimestrais=notas_bimestrais,
                frequencias_bimestrais=frequencias_bimestrais,
                plano_acao=plano_acao,
                historico_completo=historico_completo,
                chart_image_path=temp_png,
                disciplina=disc_selecionada
            )
            messagebox.showinfo("Sucesso", f"Relatório PDF de {aluno_nome.split()[0]} gerado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro ao Gerar PDF", f"Falha na geração do arquivo PDF: {e}")
        finally:
            if os.path.exists(temp_png):
                try:
                    os.remove(temp_png)
                except Exception:
                    pass

    def exportar_pdf_sala(self):
        turma_nome = self.turma_ativa.get()
        if not turma_nome:
            messagebox.showwarning("Aviso", "Nenhuma turma ativa para exportar.")
            return
            
        b_ativo = self.bimestre_ativo.get()
        dados = self.bimestres_dados[b_ativo]
        if not dados:
            messagebox.showwarning("Aviso", "Não há dados carregados para esta turma no bimestre ativo.")
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Documento PDF", "*.pdf")],
            title=f"Exportar Relatório da Turma — {turma_nome}",
            initialfile=f"Relatorio_{turma_nome.replace(' ', '_')}_{b_ativo}Bimestre.pdf"
        )
        if not file_path:
            return
            
        disc_selecionada = self.comparativo_disciplina_filtro.get()
        
        alunos_ativos = [al for al in dados if al.get('situacao', 'Ativa') == 'Ativa']
        total_alunos = len(alunos_ativos)
        media_sala = round(sum(aluno['media'] for aluno in alunos_ativos) / total_alunos, 2) if total_alunos else 0.0
        freq_media = round(sum(aluno['frequencia'] for aluno in alunos_ativos) / total_alunos, 1) if total_alunos else 0.0
        
        candidatos_destaque = []
        for aluno in alunos_ativos:
            sem_vermelha = True
            for sub, nota in aluno['notas'].items():
                if nota is not None and nota < 5.0:
                    sem_vermelha = False
                    break
            if sem_vermelha and len(aluno['notas']) > 0:
                candidatos_destaque.append(aluno)
                
        aluno_dest_abs = None
        if candidatos_destaque:
            candidatos_destaque = sorted(candidatos_destaque, key=lambda x: x['media'], reverse=True)
            aluno_dest_abs = candidatos_destaque[0]
            
        alunos_destaque_list = sorted(alunos_ativos, key=lambda x: x['media'], reverse=True)[:5]
        
        subjects = self.get_subjects_active_turma(b_ativo)
        disciplinas_abaixo = {}
        for sub in subjects:
            alunos_abaixo = sum(1 for aluno in alunos_ativos if aluno['notas'].get(sub) is not None and aluno['notas'][sub] < 5.0)
            if alunos_abaixo > 0:
                disciplinas_abaixo[sub] = alunos_abaixo
        disciplinas_criticas_ordenadas = sorted(disciplinas_abaixo.items(), key=lambda x: x[1], reverse=True)
        
        analise_sala = self.recuperar_analise_salva(b_ativo)
        if not analise_sala:
            analise_sala = self.gerar_analise_automatica(b_ativo, total_alunos, media_sala, freq_media, disciplinas_criticas_ordenadas)
            
        historico_turma = {}
        for b in [1, 2, 3, 4]:
            dados_b = self.turmas_dados[turma_nome].get(b)
            if dados_b:
                ativos_b = [al for al in dados_b if al.get('situacao', 'Ativa') == 'Ativa']
                if ativos_b:
                    if disc_selecionada == "Todas (Média Geral)":
                        avg_m = sum(aluno['media'] for aluno in ativos_b) / len(ativos_b)
                    else:
                        notas_validas = [aluno['notas'][disc_selecionada] for aluno in ativos_b if aluno['notas'].get(disc_selecionada) is not None]
                        avg_m = sum(notas_validas) / len(notas_validas) if notas_validas else 0.0
                    avg_f = sum(aluno['frequencia'] for aluno in ativos_b) / len(ativos_b)
                    historico_turma[b] = {'media': avg_m, 'presenca': avg_f}
                    
        temp_png = os.path.join(os.getcwd(), f"temp_chart_sala_{turma_nome.replace(' ', '_')}.png")
        graficos.exportar_evolucao_turma_png(temp_png, historico_turma, turma_nome, disc_selecionada)
        
        try:
            relatorios_pdf.gerar_pdf_sala(
                caminho_salvar=file_path,
                turma_nome=turma_nome,
                bimester_ativo=b_ativo,
                dados_alunos=dados,
                media_sala=media_sala,
                freq_media=freq_media,
                alunos_destaque=alunos_destaque_list,
                aluno_destaque_absoluto=aluno_dest_abs,
                disciplinas_criticas=disciplinas_criticas_ordenadas,
                analise_sala=analise_sala,
                chart_image_path=temp_png,
                disciplina=disc_selecionada
            )
            messagebox.showinfo("Sucesso", f"Relatório PDF da sala '{turma_nome}' gerado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro ao Gerar PDF", f"Falha na geração do arquivo PDF da sala: {e}")
        finally:
            if os.path.exists(temp_png):
                try:
                    os.remove(temp_png)
                except Exception:
                    pass

    def criar_tela_sobre(self):
        frame = ctk.CTkFrame(self.content_container, fg_color="transparent")
        frame.grid_rowconfigure(1, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        
        # Cabeçalho
        header = ctk.CTkFrame(frame, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        
        lbl_title = ctk.CTkLabel(
            header, 
            text="ℹ️ Sobre o Sistema", 
            font=ctk.CTkFont(size=22, weight="bold")
        )
        lbl_title.pack(anchor="w", side="left")
        
        # Cartão de Informações Principal dentro de um ScrollableFrame
        scroll_frame = ctk.CTkScrollableFrame(frame, fg_color="transparent")
        scroll_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        scroll_frame.grid_columnconfigure(0, weight=1)
        scroll_frame.grid_rowconfigure(0, weight=1)
        
        card = ctk.CTkFrame(scroll_frame, fg_color="#1e293b", corner_radius=12, border_width=1, border_color="#334155")
        card.pack(fill="both", expand=True, padx=10, pady=10)
        card.grid_columnconfigure(0, weight=1)
        
        logo_path = resource_path("logo.png")
        lbl_logo_img = None
        if os.path.exists(logo_path):
            try:
                from PIL import Image
                pil_img = Image.open(logo_path)
                w, h = pil_img.size
                new_w = 260
                new_h = int(new_w * h / w)
                ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(new_w, new_h))
                lbl_logo_img = ctk.CTkLabel(card, image=ctk_img, text="")
                lbl_logo_img.pack(pady=(20, 10))
            except Exception:
                pass
                
        if not lbl_logo_img:
            # Fallback de Título Textual
            lbl_fallback = ctk.CTkLabel(
                card, 
                text="📊 RADAR PEDAGÓGICO", 
                font=ctk.CTkFont(family="Outfit", size=24, weight="bold"),
                text_color="#38bdf8"
            )
            lbl_fallback.pack(pady=(20, 10))
            
        # Nome do Sistema
        lbl_system_name = ctk.CTkLabel(
            card,
            text="Sistema de Análise e Monitoramento Escolar",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#f8fafc"
        )
        lbl_system_name.pack(pady=(0, 10))
        
        # Divisor
        div = ctk.CTkFrame(card, height=2, fg_color="#334155")
        div.pack(fill="x", padx=100, pady=15)
        
        # Info Box (grid interno de informações)
        info_frame = ctk.CTkFrame(card, fg_color="transparent")
        info_frame.pack(pady=10)
        info_frame.grid_columnconfigure(0, weight=0)
        info_frame.grid_columnconfigure(1, weight=1)
        
        infos = [
            ("👨‍🏫 Desenvolvido por:", "Prof. André Eduardo Gomes"),
            ("📅 Ano de Desenvolvimento:", "2026"),
            ("✉️ E-mail de Contato:", "andre.gomes2004@gmail.com"),
            ("🚀 Versão:", "2.0.0 (Estável)")
        ]
        
        for idx, (label, val) in enumerate(infos):
            lbl_l = ctk.CTkLabel(
                info_frame, 
                text=label, 
                font=ctk.CTkFont(size=13, weight="bold"),
                text_color="#38bdf8",
                anchor="e"
            )
            lbl_l.grid(row=idx, column=0, padx=(10, 15), pady=8, sticky="e")
            
            lbl_r = ctk.CTkLabel(
                info_frame, 
                text=val, 
                font=ctk.CTkFont(size=13),
                text_color="#e2e8f0",
                anchor="w"
            )
            lbl_r.grid(row=idx, column=1, padx=(0, 10), pady=8, sticky="w")
            
        # Descrição institucional
        lbl_desc = ctk.CTkLabel(
            card,
            text="O Radar Pedagógico é uma plataforma de Business Intelligence voltada para docentes e gestores escolares.\n"
                 "Sua principal finalidade é simplificar o processo de importação e processamento de planilhas de turmas,\n"
                 "gerando relatórios consolidados de rendimento acadêmico, alertas de busca ativa contra a evasão escolar,\n"
                 "gráficos interativos e simulações avançadas de conselho de classe.",
            font=ctk.CTkFont(size=12),
            text_color="#94a3b8",
            justify="center"
        )
        lbl_desc.pack(pady=(20, 15))
        
        # Painel do Dono (Apenas se for admin)
        if self.license_manager.data.get("status") == "admin_vitalicio":
            admin_card = ctk.CTkFrame(card, fg_color="#0f172a", corner_radius=8, border_width=1, border_color="#334155")
            admin_card.pack(fill="x", padx=40, pady=(10, 20))
            
            lbl_adm_title = ctk.CTkLabel(
                admin_card,
                text="🔑 Painel do Dono (Gerador de Chaves)",
                font=ctk.CTkFont(size=14, weight="bold"),
                text_color="#10b981"
            )
            lbl_adm_title.pack(pady=(10, 5))
            
            lbl_instr = ctk.CTkLabel(
                admin_card,
                text="Gerar Chave de Ativação Universal (Sem Vínculo de Máquina):",
                font=ctk.CTkFont(size=12),
                text_color="#94a3b8"
            )
            lbl_instr.pack(pady=2)
            
            # Campo onde aparece a chave gerada
            entry_key_output = ctk.CTkEntry(
                admin_card,
                placeholder_text="A chave gerada aparecerá aqui",
                width=320,
                height=30,
                fg_color="#1e293b",
                border_color="#334155",
                justify="center"
            )
            entry_key_output.pack(pady=5)
            
            # Frame de botões
            btn_frame = ctk.CTkFrame(admin_card, fg_color="transparent")
            btn_frame.pack(pady=10)
            
            import random
            import string
            
            def gerar_chave_universal(limite_1ano=True):
                rand_str = ''.join(random.choices(string.ascii_uppercase + string.digits, k=5))
                if limite_1ano:
                    chave = f"ID-365-{rand_str}"
                else:
                    chave = f"ID-VITALICIA-{rand_str}"
                entry_key_output.delete(0, "end")
                entry_key_output.insert(0, chave)
                
            def copiar_chave_gerada():
                chave = entry_key_output.get().strip()
                if not chave or "aparecerá" in chave:
                    messagebox.showwarning("Aviso", "Nenhuma chave foi gerada ainda.")
                    return
                self.clipboard_clear()
                self.clipboard_append(chave)
                messagebox.showinfo("Copiado", f"Chave copiada com sucesso:\n{chave}")
                
            btn_gerar_1ano = ctk.CTkButton(
                btn_frame,
                text="📅 Gerar 1 Ano",
                font=ctk.CTkFont(size=11, weight="bold"),
                fg_color="#059669",
                hover_color="#047857",
                width=110,
                height=30,
                command=lambda: gerar_chave_universal(limite_1ano=True)
            )
            btn_gerar_1ano.grid(row=0, column=0, padx=5)
            
            btn_gerar_vital = ctk.CTkButton(
                btn_frame,
                text="♾️ Gerar Vitalícia",
                font=ctk.CTkFont(size=11, weight="bold"),
                fg_color="#0ea5e9",
                hover_color="#0284c7",
                width=110,
                height=30,
                command=lambda: gerar_chave_universal(limite_1ano=False)
            )
            btn_gerar_vital.grid(row=0, column=1, padx=5)
            
            btn_copiar = ctk.CTkButton(
                btn_frame,
                text="📋 Copiar",
                font=ctk.CTkFont(size=11, weight="bold"),
                fg_color="#10b981",
                hover_color="#059669",
                width=80,
                height=30,
                command=copiar_chave_gerada
            )
            btn_copiar.grid(row=0, column=2, padx=5)
            
        return frame

if __name__ == "__main__":
    app = SchoolAnalyzerApp()
    app.mainloop()
