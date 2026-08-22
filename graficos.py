import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Configurações de estilo para manter compatibilidade com o CustomTkinter (Escuro / Claro)
COLOR_CONFIGS = {
    'dark': {
        'bg': '#0f172a',       # Slate 900 (mais contraste)
        'fg': '#f8fafc',       # Slate 50
        'grid': '#334155',     # Slate 700
        'accent': '#38bdf8',   # Sky 400
        'accent2': '#fb7185',  # Rose 400
        'bar': '#0ea5e9',      # Sky 500
        'text': '#e2e8f0'      # Slate 200
    },
    'light': {
        'bg': '#ffffff',
        'fg': '#0f172a',       # Slate 900
        'grid': '#e2e8f0',     # Slate 200
        'accent': '#0284c7',   # Sky 600
        'accent2': '#e11d48',  # Rose 600
        'bar': '#0284c7',
        'text': '#334155'
    }
}

def configurar_estilo(theme='dark'):
    cfg = COLOR_CONFIGS.get(theme, COLOR_CONFIGS['dark'])
    plt.rcParams.update({
        'figure.facecolor': cfg['bg'],
        'axes.facecolor': cfg['bg'],
        'axes.edgecolor': cfg['grid'],
        'axes.labelcolor': cfg['text'],
        'xtick.color': cfg['text'],
        'ytick.color': cfg['text'],
        'text.color': cfg['text'],
        'grid.color': cfg['grid'],
        'grid.linestyle': '--',
        'grid.alpha': 0.6,
        'font.family': 'sans-serif',
        'font.sans-serif': ['Segoe UI', 'DejaVu Sans', 'Arial']
    })
    return cfg

def limpar_container(parent_frame):
    """Remove widgets antigos do frame de exibição do gráfico."""
    for widget in parent_frame.winfo_children():
        widget.destroy()

def plotar_medias_disciplinas(parent_frame, medias_disciplinas, theme='dark'):
    """
    Gera um gráfico de barras horizontal das médias por disciplina,
    ordenado da maior para a menor (como na Imagem 2).
    """
    limpar_container(parent_frame)
    cfg = configurar_estilo(theme)
    
    if not medias_disciplinas:
        fig, ax = plt.subplots(figsize=(6, 4))
        ax.text(0.5, 0.5, "Nenhum dado de disciplina disponivel", 
                ha='center', va='center', fontsize=12)
        ax.axis('off')
        canvas = FigureCanvasTkAgg(fig, master=parent_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)
        return canvas

    # Ordenar as disciplinas pelas médias
    disciplinas_ordenadas = sorted(medias_disciplinas.items(), key=lambda x: x[1])
    nomes_originais = [x[0] for x in disciplinas_ordenadas]
    
    # Limpar códigos e quebras de linha dos nomes das disciplinas
    nomes = []
    import re
    for n in nomes_originais:
        n_limpo = n.split('\n')[0].split('\r')[0]
        n_limpo = re.sub(r'\(\d+\)', '', n_limpo)
        n_limpo = re.sub(r'\b\d+\b', '', n_limpo)
        n_limpo = n_limpo.strip()
        if len(n_limpo) > 22:
            n_limpo = n_limpo[:19] + "..."
        nomes.append(n_limpo)
        
    valores = [x[1] for x in disciplinas_ordenadas]

    fig, ax = plt.subplots(figsize=(5.6, 3.8), dpi=100)
    bars = ax.barh(nomes, valores, color=cfg['bar'], height=0.6, edgecolor=cfg['bg'])
    
    # Adicionar os valores nas pontas das barras
    for bar in bars:
        width = bar.get_width()
        ax.text(width + 0.1, bar.get_y() + bar.get_height()/2, f'{width:.1f}', 
                ha='left', va='center', color=cfg['text'], fontweight='bold', fontsize=9)
        
    ax.set_title("Média Geral da Sala por Disciplina", fontsize=13, fontweight='bold', pad=15)
    ax.set_xlabel("Média", fontsize=10, labelpad=8)
    ax.set_xlim(0, 10.5)
    ax.xaxis.grid(True)
    ax.set_axisbelow(True)
    # Configurar tamanho menor da fonte do eixo Y para garantir que caiba na margem
    ax.tick_params(axis='y', labelsize=8.5, labelcolor=cfg['text'])
    
    # Remover bordas desnecessárias
    for spine in ['top', 'right']:
        ax.spines[spine].set_visible(False)
        
    # Margem esquerda confortável para os nomes longos das disciplinas
    fig.subplots_adjust(left=0.42, right=0.90, top=0.88, bottom=0.15)
    
    canvas = FigureCanvasTkAgg(fig, master=parent_frame)
    canvas.draw()
    canvas.get_tk_widget().pack(fill='both', expand=True)
    return canvas

def plotar_evolucao_aluno(parent_frame, historico_aluno, nome_aluno, theme='dark', disciplina='Média Geral'):
    """
    Gera um gráfico de linha mostrando a evolução do aluno ao longo dos bimestres.
    Eixo Y1: Média Geral (0 a 10)
    Eixo Y2: Faltas (Absoluto)
    Inclui animação de interpolação suave ao trocar de disciplina.
    """
    cfg = COLOR_CONFIGS.get(theme, COLOR_CONFIGS['dark'])
    bimestres = sorted(historico_aluno.keys())
    
    if not bimestres:
        limpar_container(parent_frame)
        fig, ax = plt.subplots(figsize=(6, 4))
        ax.text(0.5, 0.5, "Dados insuficientes para gerar grafico de evolucao", 
                ha='center', va='center', fontsize=12)
        ax.axis('off')
        canvas = FigureCanvasTkAgg(fig, master=parent_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)
        for attr in ['current_fig', 'current_ax1', 'current_ax2', 'current_line1', 'current_line2', 
                     'current_ann_medias', 'current_ann_presencas', 'current_medias', 'current_presencas', 
                     'current_nome', 'current_bimestres']:
            if hasattr(parent_frame, attr):
                delattr(parent_frame, attr)
        return canvas
        
    x_labels = [f'{b}º Bim' for b in bimestres]
    medias = [historico_aluno[b]['media'] for b in bimestres]
    # Usar faltas no lugar da frequência/presença
    presencas = [float(historico_aluno[b].get('faltas', 0)) for b in bimestres]
    
    # Verificar se podemos animar (mesmo aluno e mesmos bimestres já plotados)
    can_animate = False
    if (hasattr(parent_frame, 'current_nome') and parent_frame.current_nome == nome_aluno and
        hasattr(parent_frame, 'current_bimestres') and parent_frame.current_bimestres == bimestres and
        hasattr(parent_frame, 'current_fig') and hasattr(parent_frame, 'current_line1')):
        can_animate = True
        
    if can_animate:
        fig = parent_frame.current_fig
        ax1 = parent_frame.current_ax1
        ax2 = parent_frame.current_ax2
        line1 = parent_frame.current_line1
        line2 = parent_frame.current_line2
        ann_medias = parent_frame.current_ann_medias
        ann_presencas = parent_frame.current_ann_presencas
        
        old_medias = parent_frame.current_medias
        old_presencas = parent_frame.current_presencas
        
        # Obter nome curto para evitar quebra de layout
        parts = nome_aluno.split()
        nome_curto = parts[0] + " " + parts[-1] if len(parts) > 1 else nome_aluno
        ax1.set_title(f"Evolução Escolar — {nome_curto}\n{disciplina}", fontsize=11, fontweight='bold', pad=12)
        
        steps = 12
        delay = 15 # ms
        
        # Limites Y dinâmicos para faltas na animação
        max_f = max(presencas) if presencas else 0
        ax2.set_ylim(-0.5, max(max_f + 2, 10))
        
        def run_step(step_idx):
            if not parent_frame.winfo_exists():
                return
            t = step_idx / steps
            current_medias = [old_m + (new_m - old_m) * t for old_m, new_m in zip(old_medias, medias)]
            current_presencas = [old_p + (new_p - old_p) * t for old_p, new_p in zip(old_presencas, presencas)]
            
            line1[0].set_ydata(current_medias)
            line2[0].set_ydata(current_presencas)
            
            for idx, m_val in enumerate(current_medias):
                if idx < len(ann_medias):
                    ann_medias[idx].xy = (x_labels[idx], m_val)
                    ann_medias[idx].set_text(f'{m_val:.1f}')
                    
            for idx, p_val in enumerate(current_presencas):
                if idx < len(ann_presencas):
                    ann_presencas[idx].xy = (x_labels[idx], p_val)
                    ann_presencas[idx].set_text(f'{p_val:.0f} f')
                    
            fig.canvas.draw_idle()
            
            if step_idx < steps:
                parent_frame.after(delay, lambda: run_step(step_idx + 1))
            else:
                parent_frame.current_medias = medias
                parent_frame.current_presencas = presencas
                
        run_step(1)
        return fig.canvas
    else:
        limpar_container(parent_frame)
        cfg = configurar_estilo(theme)
        
        fig, ax1 = plt.subplots(figsize=(5.2, 3.6), dpi=100)
        
        # Plot da Média
        color = cfg['accent']
        ax1.set_xlabel('Bimestre', labelpad=8)
        ax1.set_ylabel('Média de Notas', color=color, fontweight='bold')
        line1 = ax1.plot(x_labels, medias, marker='o', linewidth=2.5, color=color, label='Média de Notas')
        ax1.tick_params(axis='y', labelcolor=color)
        ax1.set_ylim(0, 10.5)
        ax1.grid(True)
        
        # Segundo eixo Y para Faltas
        ax2 = ax1.twinx()  
        color = cfg['accent2']
        ax2.set_ylabel('Faltas', color=color, fontweight='bold')
        line2 = ax2.plot(x_labels, presencas, marker='s', linewidth=2, color=color, linestyle='--', label='Faltas')
        ax2.tick_params(axis='y', labelcolor=color)
        max_f = max(presencas) if presencas else 0
        ax2.set_ylim(-0.5, max(max_f + 2, 10))
        
        # Adicionar anotações de valores
        ann_medias = []
        for i, txt in enumerate(medias):
            ann = ax1.annotate(f'{txt:.1f}', (x_labels[i], medias[i]), textcoords="offset points", 
                         xytext=(0,10), ha='center', fontweight='bold', fontsize=9, color=cfg['accent'])
            ann_medias.append(ann)
            
        ann_presencas = []
        for i, txt in enumerate(presencas):
            ann = ax2.annotate(f'{txt:.0f} f', (x_labels[i], presencas[i]), textcoords="offset points", 
                         xytext=(0,-15), ha='center', fontweight='bold', fontsize=9, color=cfg['accent2'])
            ann_presencas.append(ann)
            
        # Obter nome curto para evitar quebra de layout
        parts = nome_aluno.split()
        nome_curto = parts[0] + " " + parts[-1] if len(parts) > 1 else nome_aluno
        ax1.set_title(f"Evolução Escolar — {nome_curto}\n{disciplina}", fontsize=11, fontweight='bold', pad=12)
        
        # Remover bordas superiores
        ax1.spines['top'].set_visible(False)
        ax2.spines['top'].set_visible(False)
        
        # Ajustar margens manuais para dar espaço confortável para Média de Notas e Faltas
        fig.subplots_adjust(left=0.16, right=0.84, top=0.82, bottom=0.15)
        
        canvas = FigureCanvasTkAgg(fig, master=parent_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)
        
        # Guardar referências para futura animação
        parent_frame.current_fig = fig
        parent_frame.current_ax1 = ax1
        parent_frame.current_ax2 = ax2
        parent_frame.current_line1 = line1
        parent_frame.current_line2 = line2
        parent_frame.current_ann_medias = ann_medias
        parent_frame.current_ann_presencas = ann_presencas
        parent_frame.current_medias = medias
        parent_frame.current_presencas = presencas
        parent_frame.current_nome = nome_aluno
        parent_frame.current_bimestres = bimestres
        
        return canvas

def plotar_evolucao_turma(parent_frame, historico_turma, theme='dark', disciplina='Média Geral'):
    """
    Plota a evolução da Média Geral da Sala e Faltas Médias da Sala ao longo dos bimestres.
    Inclui animação de interpolação suave ao trocar de disciplina.
    """
    cfg = COLOR_CONFIGS.get(theme, COLOR_CONFIGS['dark'])
    bimestres = sorted(historico_turma.keys())
    
    if not bimestres:
        limpar_container(parent_frame)
        fig, ax = plt.subplots(figsize=(6, 4))
        ax.text(0.5, 0.5, "Carregue os bimestres para visualizar a evolucao da turma", 
                ha='center', va='center', fontsize=12)
        ax.axis('off')
        canvas = FigureCanvasTkAgg(fig, master=parent_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)
        for attr in ['current_fig', 'current_ax1', 'current_ax2', 'current_line1', 'current_line2', 
                     'current_ann_medias', 'current_ann_presencas', 'current_medias', 'current_presencas', 
                     'current_nome', 'current_bimestres']:
            if hasattr(parent_frame, attr):
                delattr(parent_frame, attr)
        return canvas
        
    x_labels = [f'{b}º Bim' for b in bimestres]
    medias = [historico_turma[b]['media'] for b in bimestres]
    presencas = [historico_turma[b]['presenca'] for b in bimestres]
    
    # Verificar se podemos animar
    can_animate = False
    if (hasattr(parent_frame, 'current_nome') and parent_frame.current_nome == "turma" and
        hasattr(parent_frame, 'current_bimestres') and parent_frame.current_bimestres == bimestres and
        hasattr(parent_frame, 'current_fig') and hasattr(parent_frame, 'current_line1')):
        can_animate = True
        
    if can_animate:
        fig = parent_frame.current_fig
        ax1 = parent_frame.current_ax1
        ax2 = parent_frame.current_ax2
        line1 = parent_frame.current_line1
        line2 = parent_frame.current_line2
        ann_medias = parent_frame.current_ann_medias
        ann_presencas = parent_frame.current_ann_presencas
        
        old_medias = parent_frame.current_medias
        old_presencas = parent_frame.current_presencas
        
        import re
        disc_limpa = disciplina.split('\n')[0].split('\r')[0]
        disc_limpa = re.sub(r'\(\d+\)', '', disc_limpa)
        disc_limpa = re.sub(r'\b\d+\b', '', disc_limpa).strip()
        ax1.set_title(f"Evolução Histórica da Turma\n{disc_limpa}", fontsize=13, fontweight='bold', pad=15)
        
        steps = 12
        delay = 15 # ms
        
        # Limites Y dinâmicos para faltas na animação da turma
        max_f = max(presencas) if presencas else 0
        ax2.set_ylim(-0.5, max(max_f + 2, 10))
        
        def run_step(step_idx):
            if not parent_frame.winfo_exists():
                return
            t = step_idx / steps
            current_medias = [old_m + (new_m - old_m) * t for old_m, new_m in zip(old_medias, medias)]
            current_presencas = [old_p + (new_p - old_p) * t for old_p, new_p in zip(old_presencas, presencas)]
            
            line1[0].set_ydata(current_medias)
            line2[0].set_ydata(current_presencas)
            
            for idx, m_val in enumerate(current_medias):
                if idx < len(ann_medias):
                    ann_medias[idx].xy = (x_labels[idx], m_val)
                    ann_medias[idx].set_text(f'{m_val:.2f}')
                    
            for idx, p_val in enumerate(current_presencas):
                if idx < len(ann_presencas):
                    ann_presencas[idx].xy = (x_labels[idx], p_val)
                    ann_presencas[idx].set_text(f'{p_val:.1f} f')
                    
            fig.canvas.draw_idle()
            
            if step_idx < steps:
                parent_frame.after(delay, lambda: run_step(step_idx + 1))
            else:
                parent_frame.current_medias = medias
                parent_frame.current_presencas = presencas
                
        run_step(1)
        return fig.canvas
    else:
        limpar_container(parent_frame)
        cfg = configurar_estilo(theme)
        
        fig, ax1 = plt.subplots(figsize=(5.2, 3.6), dpi=100)
        
        # Plot Média da Sala
        color = cfg['accent']
        ax1.set_xlabel('Bimestre', labelpad=8)
        ax1.set_ylabel('Média Geral da Sala', color=color, fontweight='bold')
        line1 = ax1.plot(x_labels, medias, marker='o', linewidth=3, color=color, label='Média Geral')
        ax1.tick_params(axis='y', labelcolor=color)
        ax1.set_ylim(0, 10.5)
        ax1.grid(True)
        
        # Plot Faltas Médias
        ax2 = ax1.twinx()
        color = cfg['accent2']
        ax2.set_ylabel('Faltas Médias', color=color, fontweight='bold')
        line2 = ax2.plot(x_labels, presencas, marker='^', linewidth=2.5, color=color, linestyle='-.', label='Faltas Médias')
        ax2.tick_params(axis='y', labelcolor=color)
        max_f = max(presencas) if presencas else 0
        ax2.set_ylim(-0.5, max(max_f + 2, 10))
        
        # Adicionar anotações
        ann_medias = []
        for i, txt in enumerate(medias):
            ann = ax1.annotate(f'{txt:.2f}', (x_labels[i], medias[i]), textcoords="offset points", 
                         xytext=(0,10), ha='center', fontweight='bold', fontsize=9, color=cfg['accent'])
            ann_medias.append(ann)
            
        ann_presencas = []
        for i, txt in enumerate(presencas):
            ann = ax2.annotate(f'{txt:.1f} f', (x_labels[i], presencas[i]), textcoords="offset points", 
                         xytext=(0,-15), ha='center', fontweight='bold', fontsize=9, color=cfg['accent2'])
            ann_presencas.append(ann)
            
        import re
        disc_limpa = disciplina.split('\n')[0].split('\r')[0]
        disc_limpa = re.sub(r'\(\d+\)', '', disc_limpa)
        disc_limpa = re.sub(r'\b\d+\b', '', disc_limpa).strip()
        ax1.set_title(f"Evolução Histórica da Turma\n{disc_limpa}", fontsize=13, fontweight='bold', pad=15)
        ax1.spines['top'].set_visible(False)
        ax2.spines['top'].set_visible(False)
        
        # Ajustar margens manuais para dar espaço confortável para Média de Notas e Faltas
        fig.subplots_adjust(left=0.16, right=0.84, top=0.82, bottom=0.15)
        
        canvas = FigureCanvasTkAgg(fig, master=parent_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)
        
        parent_frame.current_fig = fig
        parent_frame.current_ax1 = ax1
        parent_frame.current_ax2 = ax2
        parent_frame.current_line1 = line1
        parent_frame.current_line2 = line2
        parent_frame.current_ann_medias = ann_medias
        parent_frame.current_ann_presencas = ann_presencas
        parent_frame.current_medias = medias
        parent_frame.current_presencas = presencas
        parent_frame.current_nome = "turma"
        parent_frame.current_bimestres = bimestres
        
        return canvas

def plotar_notas_baixas_vs_faltas(parent_frame, dados_alunos, theme='dark'):
    """
    Gera um gráfico comparativo agrupado de Notas Baixas (< 5.0) e Faltas (TF)
    para os alunos com mais notas baixas (como na Imagem 3).
    dados_alunos: lista de dicts contendo {'nome': X, 'notas_baixas': Y, 'faltas': Z, 'frequencia': W}
    """
    limpar_container(parent_frame)
    cfg = configurar_estilo(theme)
    
    if not dados_alunos:
        fig, ax = plt.subplots(figsize=(6, 4))
        ax.text(0.5, 0.5, "Sem dados comparativos disponiveis", 
                ha='center', va='center', fontsize=12)
        ax.axis('off')
        canvas = FigureCanvasTkAgg(fig, master=parent_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)
        return canvas

    # Ordenar pelos que têm mais notas baixas
    dados_ordenados = sorted(dados_alunos, key=lambda x: x['notas_baixas'], reverse=True)[:10] # Top 10
    
    nomes = [x['nome'].split()[0] + ' ' + (x['nome'].split()[1][0] + '.' if len(x['nome'].split()) > 1 else '') for x in dados_ordenados]
    notas_baixas = [x['notas_baixas'] for x in dados_ordenados]
    faltas = [x['faltas'] for x in dados_ordenados]
    
    x = np.arange(len(nomes))
    width = 0.35  # Largura das barras
    
    fig, ax1 = plt.subplots(figsize=(5.6, 3.8), dpi=100)
    
    # Barras de Notas Baixas (Eixo Y1)
    rects1 = ax1.bar(x - width/2, notas_baixas, width, label='Notas Baixas (<5.0)', color=cfg['accent2'])
    ax1.set_ylabel('Quantidade de Notas Baixas', color=cfg['accent2'], fontweight='bold')
    ax1.tick_params(axis='y', labelcolor=cfg['accent2'])
    ax1.set_xticks(x)
    ax1.set_xticklabels(nomes, rotation=25, ha='right', fontsize=9)
    ax1.grid(True)
    
    # Barras de Faltas (Eixo Y2)
    ax2 = ax1.twinx()
    rects2 = ax2.bar(x + width/2, faltas, width, label='Faltas (TF)', color='#f59e0b') # Laranja para faltas
    ax2.set_ylabel('Total de Faltas (Aulas)', color='#f59e0b', fontweight='bold')
    ax2.tick_params(axis='y', labelcolor='#f59e0b')
    
    # Título e Legenda conjunta
    plt.title("Top Alunos: Relação Notas Baixas vs. Faltas", fontsize=13, fontweight='bold', pad=15)
    
    # Adicionar legenda
    lines, labels = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines + lines2, labels + labels2, loc='upper right', framealpha=0.9, facecolor=cfg['bg'], edgecolor=cfg['grid'])
    
    ax1.spines['top'].set_visible(False)
    ax2.spines['top'].set_visible(False)
    
    plt.tight_layout()
    
    canvas = FigureCanvasTkAgg(fig, master=parent_frame)
    canvas.draw()
    canvas.get_tk_widget().pack(fill='both', expand=True)
    return canvas


def exportar_evolucao_aluno_png(filepath, historico_aluno, nome_aluno, disciplina='Média Geral'):
    """Gera o gráfico de evolução do aluno e salva como imagem PNG."""
    import matplotlib
    matplotlib.use('Agg') # Usar backend não-GUI para evitar conflitos de threads
    import matplotlib.pyplot as plt
    
    cfg = configurar_estilo('light') # Relatórios em PDF costumam usar estilo claro (fundo branco)
    bimestres = sorted(historico_aluno.keys())
    if not bimestres:
        return False
        
    x_labels = [f'{b}º Bim' for b in bimestres]
    medias = [historico_aluno[b]['media'] for b in bimestres]
    presencas = [historico_aluno[b]['presenca'] for b in bimestres]
    
    fig, ax1 = plt.subplots(figsize=(7, 3.5), dpi=150)
    
    # Plot Média
    color = cfg['accent']
    ax1.set_xlabel('Bimestre', labelpad=8)
    ax1.set_ylabel('Média de Notas', color=color, fontweight='bold')
    ax1.plot(x_labels, medias, marker='o', linewidth=2.5, color=color)
    ax1.tick_params(axis='y', labelcolor=color)
    ax1.set_ylim(0, 10.5)
    ax1.grid(True)
    
    # Plot Frequência
    ax2 = ax1.twinx()
    color = cfg['accent2']
    ax2.set_ylabel('Frequência (%)', color=color, fontweight='bold')
    ax2.plot(x_labels, presencas, marker='s', linewidth=2, color=color, linestyle='--')
    ax2.tick_params(axis='y', labelcolor=color)
    ax2.set_ylim(0, 105)
    
    # Anotações
    for i, txt in enumerate(medias):
        ax1.annotate(f'{txt:.1f}', (x_labels[i], medias[i]), textcoords="offset points", 
                     xytext=(0,8), ha='center', fontweight='bold', fontsize=8, color=cfg['accent'])
    for i, txt in enumerate(presencas):
        ax2.annotate(f'{txt:.0f}%', (x_labels[i], presencas[i]), textcoords="offset points", 
                     xytext=(0,-12), ha='center', fontweight='bold', fontsize=8, color=cfg['accent2'])
                     
    plt.title(f"Evolução Escolar — {nome_aluno}\n({disciplina})", fontsize=11, fontweight='bold', pad=12)
    ax1.spines['top'].set_visible(False)
    ax2.spines['top'].set_visible(False)
    
    plt.tight_layout()
    fig.savefig(filepath, bbox_inches='tight')
    plt.close(fig)
    return True

def exportar_evolucao_turma_png(filepath, historico_turma, nome_turma, disciplina='Média Geral'):
    """Gera o gráfico de evolução da turma e salva como imagem PNG."""
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    
    cfg = configurar_estilo('light')
    bimestres = sorted(historico_turma.keys())
    if not bimestres:
        return False
        
    x_labels = [f'{b}º Bim' for b in bimestres]
    medias = [historico_turma[b]['media'] for b in bimestres]
    presencas = [historico_turma[b]['presenca'] for b in bimestres]
    
    fig, ax1 = plt.subplots(figsize=(7, 3.5), dpi=150)
    
    # Plot Média da Sala
    color = cfg['accent']
    ax1.set_xlabel('Bimestre', labelpad=8)
    ax1.set_ylabel('Média Geral da Sala', color=color, fontweight='bold')
    ax1.plot(x_labels, medias, marker='o', linewidth=3, color=color)
    ax1.tick_params(axis='y', labelcolor=color)
    ax1.set_ylim(0, 10.5)
    ax1.grid(True)
    
    # Plot Presença Média
    ax2 = ax1.twinx()
    color = cfg['accent2']
    ax2.set_ylabel('Presença Média (%)', color=color, fontweight='bold')
    ax2.plot(x_labels, presencas, marker='^', linewidth=2.5, color=color, linestyle='-.')
    ax2.tick_params(axis='y', labelcolor=color)
    ax2.set_ylim(0, 105)
    
    # Anotações
    for i, txt in enumerate(medias):
        ax1.annotate(f'{txt:.2f}', (x_labels[i], medias[i]), textcoords="offset points", 
                     xytext=(0,8), ha='center', fontweight='bold', fontsize=8, color=cfg['accent'])
    for i, txt in enumerate(presencas):
        ax2.annotate(f'{txt:.1f}%', (x_labels[i], presencas[i]), textcoords="offset points", 
                     xytext=(0,-12), ha='center', fontweight='bold', fontsize=8, color=cfg['accent2'])
                     
    plt.title(f"Evolução Histórica da Turma — {nome_turma}\n({disciplina})", fontsize=11, fontweight='bold', pad=12)
    ax1.spines['top'].set_visible(False)
    ax2.spines['top'].set_visible(False)
    
    plt.tight_layout()
    fig.savefig(filepath, bbox_inches='tight')
    plt.close(fig)
    return True


def plotar_radar_competencias(parent_frame, notas_aluno, theme='dark'):
    """
    Gera um gráfico de radar (polar) agrupando as disciplinas em 5 áreas de competências.
    notas_aluno: dicionário contendo as notas da disciplina no bimestre.
    """
    limpar_container(parent_frame)
    cfg = configurar_estilo(theme)
    
    if not notas_aluno:
        fig, ax = plt.subplots(figsize=(2.7, 2.7))
        fig.patch.set_facecolor(cfg['bg'])
        ax.set_facecolor(cfg['bg'])
        ax.text(0.5, 0.5, "Sem dados para radar", ha='center', va='center', fontsize=10, color=cfg['text'])
        ax.axis('off')
        canvas = FigureCanvasTkAgg(fig, master=parent_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)
        return canvas
        
    categorias = {
        'TÉCNICA': [],
        'LINGUAGENS': [],
        'EXATAS': [],
        'BIOLÓGICAS': [],
        'HUMANAS': []
    }
    
    for sub, nota in notas_aluno.items():
        if nota is None:
            continue
        sub_upper = sub.upper()
        
        # 1. Técnica
        if any(kw in sub_upper for kw in ['LOGICA', 'PROGRAMACAO', 'PROCESSOS', 'REDES', 'COMPUTAÇÃO', 'DESENVOLVIMENTO', 'CARREIRA', 'COMPETENCIAS', 'TECNICA']):
            categorias['TÉCNICA'].append(nota)
        # 2. Linguagens (contendo Educação Física)
        elif any(kw in sub_upper for kw in ['PORTUGUESA', 'PORTUGUES', 'INGLESA', 'INGLES', 'LITERATURA', 'ARTES', 'EDUCACAO FISICA', 'EDUCAÇÃO FÍSICA']):
            categorias['LINGUAGENS'].append(nota)
        # 3. Exatas (contendo Física e Química)
        elif any(kw in sub_upper for kw in ['MATEMATICA', 'FISICA', 'QUIMICA']):
            categorias['EXATAS'].append(nota)
        # 4. Biológicas
        elif any(kw in sub_upper for kw in ['BIOLOGIA', 'CIENCIAS']):
            categorias['BIOLÓGICAS'].append(nota)
        # 5. Humanas
        elif any(kw in sub_upper for kw in ['HISTORIA', 'GEOGRAFIA', 'SOCIOLOGIA', 'FILOSOFIA']):
            categorias['HUMANAS'].append(nota)
        else:
            categorias['LINGUAGENS'].append(nota)
            
    valores = []
    labels = []
    for cat in ['TÉCNICA', 'LINGUAGENS', 'EXATAS', 'BIOLÓGICAS', 'HUMANAS']:
        notas_list = categorias[cat]
        if notas_list:
            valores.append(sum(notas_list) / len(notas_list))
        else:
            valores.append(0.0)
        labels.append(cat)
        
    num_vars = len(labels)
    angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
    
    valores += valores[:1]
    angles += angles[:1]
    
    fig, ax = plt.subplots(figsize=(2.7, 2.7), subplot_kw=dict(projection='polar'), dpi=100)
    fig.patch.set_facecolor(cfg['bg'])
    ax.set_facecolor(cfg['bg'])
    
    # Rotação para colocar o primeiro eixo no topo
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    
    # Desenhar eixos e rótulos
    plt.xticks(angles[:-1], labels, color=cfg['text'], fontsize=6, fontweight='bold')
    
    # Configurar limites do Y (notas 0 a 10)
    ax.set_rlabel_position(0)
    plt.yticks([2.0, 4.0, 6.0, 8.0, 10.0], ["2", "4", "6", "8", "10"], color="#94a3b8", fontsize=6)
    ax.set_ylim(0, 10.5)
    
    # Desenhar dados
    color_accent = cfg['accent']
    ax.plot(angles, valores, color=color_accent, linewidth=2, linestyle='solid')
    ax.fill(angles, valores, color=color_accent, alpha=0.22)
    
    # Configurar cores das linhas de grade polar
    ax.grid(color=cfg['grid'], linestyle='--', alpha=0.6)
    ax.spines['polar'].set_color(cfg['grid'])
    
    plt.title("Radar de Competências", fontsize=11, fontweight='bold', color=cfg['fg'], pad=12)
    
    # IMPORTANTE: subplots_adjust ao invés de tight_layout para dar margem para os labels circulares
    fig.subplots_adjust(left=0.26, right=0.74, top=0.74, bottom=0.26)
    
    canvas = FigureCanvasTkAgg(fig, master=parent_frame)
    canvas.draw()
    canvas.get_tk_widget().pack(fill='both', expand=True)
    return canvas
