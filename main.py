import os
import sqlite3
from tkinter import *
from tkinter import ttk, filedialog, messagebox
from datetime import datetime, date
from docx import Document
import matplotlib.pyplot as plt
from PIL import Image, ImageTk
import calendar
from tkcalendar import DateEntry
import webbrowser
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
# =============================================
# CONFIGURAÇÕES INICIAIS
# =============================================

COR_FUNDO = "#f8f9fa"
COR_PRIMARIA = "#2c3e50"
COR_SECUNDARIA = "#34495e"
COR_DESTAQUE = "#3498db"
COR_TEXTO = "#2c3e50"
COR_CARD = "#ffffff"
COR_ALERTA = "#e74c3c"
COR_SUCESSO = "#2ecc71"

# Configurar diretórios
os.makedirs("comprovantes", exist_ok=True)
os.makedirs("relatorios", exist_ok=True)

# =============================================
# BANCO DE DADOS - ESTRUTURA SIMPLIFICADA
# =============================================

def inicializar_banco_dados():
    conn = sqlite3.connect("sistema.db")
    cursor = conn.cursor()
    
    # Tabela de clientes (proprietários)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS clientes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL,
        telefone TEXT,
        email TEXT,
        endereco TEXT
    );
    """)
    
    # Tabela de imóveis
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS imoveis (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cliente_id INTEGER,
        endereco TEXT NOT NULL,
        quartos INTEGER,
        banheiros INTEGER,
        FOREIGN KEY (cliente_id) REFERENCES clientes (id)
    );
    """)
    
    # Tabela de serviços de limpeza
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS limpezas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        imovel_id INTEGER,
        data DATE NOT NULL,
        hora_inicio TEXT,
        hora_fim TEXT,
        horas_trabalhadas REAL,
        valor_hora REAL DEFAULT 30.0,
        valor_total REAL,
        observacoes TEXT,
        FOREIGN KEY (imovel_id) REFERENCES imoveis (id)
    );
    """)
    
    # Tabela de itens do enxoval
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS tipos_enxoval (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL,
        preco_unitario REAL DEFAULT 0.0
    );
    """)
    
    # Tabela de consumo de enxoval
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS consumo_enxoval (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        imovel_id INTEGER,
        item_id INTEGER,
        data DATE NOT NULL,
        quantidade INTEGER DEFAULT 1,
        FOREIGN KEY (imovel_id) REFERENCES imoveis (id),
        FOREIGN KEY (item_id) REFERENCES tipos_enxoval (id)
    );
    """)
    
    # Tabela de suprimentos
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS suprimentos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL,
        preco_unitario REAL DEFAULT 0.0
    );
    """)
    
    # Tabela de reposição de suprimentos
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS reposicao_suprimentos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        imovel_id INTEGER,
        suprimento_id INTEGER,
        data DATE NOT NULL,
        quantidade INTEGER DEFAULT 1,
        valor_gasto REAL,
        comprovante_path TEXT,
        FOREIGN KEY (imovel_id) REFERENCES imoveis (id),
        FOREIGN KEY (suprimento_id) REFERENCES suprimentos (id)
    );
    """)
    
    # Inserir dados básicos se as tabelas estiverem vazias
    cursor.execute("SELECT COUNT(*) FROM tipos_enxoval")
    if cursor.fetchone()[0] == 0:
        itens_enxoval = [
            ('Lençol Solteiro', 45.0),
            ('Lençol Casal', 65.0),
            ('Lençol Queen', 75.0),
            ('Lençol King', 85.0),
            ('Toalha de Banho', 35.0),
            ('Toalha de Rosto', 25.0),
            ('Toalha de Mesa', 30.0),
            ('Cobertor', 90.0),
            ('Edredom', 120.0)
        ]
        cursor.executemany("INSERT INTO tipos_enxoval (nome, preco_unitario) VALUES (?, ?)", itens_enxoval)
    
    cursor.execute("SELECT COUNT(*) FROM suprimentos")
    if cursor.fetchone()[0] == 0:
        suprimentos = [
            ('Sabonete', 1.5),
            ('Shampoo', 5.0),
            ('Condicionador', 5.0),
            ('Papel Higiênico', 0.5),
            ('Papel Toalha', 2.0),
            ('Café', 15.0),
            ('Açúcar', 8.0),
            ('Sabão em Pó', 12.0),
            ('Amaciante', 10.0)
        ]
        cursor.executemany("INSERT INTO suprimentos (nome, preco_unitario) VALUES (?, ?)", suprimentos)
    
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS itens_servico (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL,
        tipo TEXT NOT NULL CHECK(tipo IN ('enxoval', 'suprimento', 'limpeza')),
        preco_unitario REAL DEFAULT 0.0,
        unidade_medida TEXT
    );
    """)






    conn.commit()
    conn.close()

# =============================================
# FUNÇÕES AUXILIARES
# =============================================

def calcular_horas(hora_inicio, hora_fim):
    """Calcula a diferença entre duas horas no formato HH:MM"""
    try:
        inicio = datetime.strptime(hora_inicio, "%H:%M")
        fim = datetime.strptime(hora_fim, "%H:%M")
        diferenca = fim - inicio
        horas = diferenca.seconds / 3600
        return round(horas, 2)
    except ValueError:
        return 0.0

def selecionar_arquivo(entry_widget, diretorio="comprovantes"):
    """Permite selecionar um arquivo e move para o diretório de comprovantes"""
    caminho = filedialog.askopenfilename()
    if caminho:
        nome_arquivo = os.path.basename(caminho)
        destino = os.path.join(diretorio, nome_arquivo)
        os.replace(caminho, destino)
        entry_widget.delete(0, END)
        entry_widget.insert(0, destino)

def formatar_moeda(valor):
    """Formata um valor float para string monetária"""
    return f"R$ {valor:.2f}".replace(".", ",")

def gerar_relatorio_semanal():
    """Gera um relatório semanal em formato DOCX"""
    conn = sqlite3.connect("sistema.db")
    cursor = conn.cursor()
    
    # Data de início (7 dias atrás)
    data_inicio = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")
    
    # Criar documento
    doc = Document()
    doc.add_heading('Relatório Semanal de Serviços', 0)
    
    # Adicionar data
    doc.add_paragraph(f"Período: {data_inicio} a {datetime.now().strftime('%Y-%m-%d')}")
    
    # Limpezas realizadas
    doc.add_heading('Limpezas Realizadas', level=1)
    cursor.execute("""
        SELECT i.endereco, l.data, l.horas_trabalhadas, l.valor_total 
        FROM limpezas l
        JOIN imoveis i ON l.imovel_id = i.id
        WHERE date(l.data) >= date(?)
        ORDER BY l.data
    """, (data_inicio,))
    
    tabela_limpezas = doc.add_table(rows=1, cols=4)
    tabela_limpezas.style = 'Light Shading'
    cabecalho = tabela_limpezas.rows[0].cells
    cabecalho[0].text = 'Imóvel'
    cabecalho[1].text = 'Data'
    cabecalho[2].text = 'Horas'
    cabecalho[3].text = 'Valor'
    
    total_horas = 0
    total_valor = 0
    
    for row in cursor.fetchall():
        linha = tabela_limpezas.add_row().cells
        linha[0].text = row[0]
        linha[1].text = row[1]
        linha[2].text = str(row[2])
        linha[3].text = formatar_moeda(row[3])
        total_horas += row[2]
        total_valor += row[3]
    
    doc.add_paragraph(f"Total de horas trabalhadas: {total_horas}")
    doc.add_paragraph(f"Total a receber por limpezas: {formatar_moeda(total_valor)}")
    
    # Enxoval utilizado
    doc.add_heading('Enxoval Utilizado', level=1)
    cursor.execute("""
        SELECT i.endereco, t.nome, SUM(c.quantidade), t.preco_unitario
        FROM consumo_enxoval c
        JOIN imoveis i ON c.imovel_id = i.id
        JOIN tipos_enxoval t ON c.item_id = t.id
        WHERE date(c.data) >= date(?)
        GROUP BY i.endereco, t.nome
        ORDER BY i.endereco
    """, (data_inicio,))
    
    tabela_enxoval = doc.add_table(rows=1, cols=4)
    tabela_enxoval.style = 'Light Shading'
    cabecalho = tabela_enxoval.rows[0].cells
    cabecalho[0].text = 'Imóvel'
    cabecalho[1].text = 'Item'
    cabecalho[2].text = 'Quantidade'
    cabecalho[3].text = 'Valor Unitário'
    
    total_itens = 0
    total_valor_enxoval = 0
    
    for row in cursor.fetchall():
        linha = tabela_enxoval.add_row().cells
        linha[0].text = row[0]
        linha[1].text = row[1]
        linha[2].text = str(row[2])
        linha[3].text = formatar_moeda(row[3])
        total_itens += row[2]
        total_valor_enxoval += row[2] * row[3]
    
    doc.add_paragraph(f"Total de itens utilizados: {total_itens}")
    doc.add_paragraph(f"Total a receber por enxoval: {formatar_moeda(total_valor_enxoval)}")
    
    # Suprimentos repostos
    doc.add_heading('Suprimentos Repostos', level=1)
    cursor.execute("""
        SELECT i.endereco, s.nome, SUM(r.quantidade), r.valor_gasto
        FROM reposicao_suprimentos r
        JOIN imoveis i ON r.imovel_id = i.id
        JOIN suprimentos s ON r.suprimento_id = s.id
        WHERE date(r.data) >= date(?)
        GROUP BY i.endereco, s.nome
        ORDER BY i.endereco
    """, (data_inicio,))
    
    tabela_suprimentos = doc.add_table(rows=1, cols=4)
    tabela_suprimentos.style = 'Light Shading'
    cabecalho = tabela_suprimentos.rows[0].cells
    cabecalho[0].text = 'Imóvel'
    cabecalho[1].text = 'Item'
    cabecalho[2].text = 'Quantidade'
    cabecalho[3].text = 'Valor Gasto'
    
    total_suprimentos = 0
    total_valor_suprimentos = 0
    
    for row in cursor.fetchall():
        linha = tabela_suprimentos.add_row().cells
        linha[0].text = row[0]
        linha[1].text = row[1]
        linha[2].text = str(row[2])
        linha[3].text = formatar_moeda(row[3])
        total_suprimentos += row[2]
        total_valor_suprimentos += row[3]
    
    doc.add_paragraph(f"Total de suprimentos repostos: {total_suprimentos}")
    doc.add_paragraph(f"Total gasto com suprimentos: {formatar_moeda(total_valor_suprimentos)}")
    doc.add_paragraph(f"Valor fixo semanal por gestão: {formatar_moeda(50.0 * cursor.execute('SELECT COUNT(DISTINCT imovel_id) FROM limpezas WHERE date(data) >= date(?)', (data_inicio,)).fetchone()[0])}")
    
    # Total geral
    doc.add_heading('Resumo Financeiro', level=1)
    total_geral = total_valor + total_valor_enxoval + (50.0 * cursor.execute('SELECT COUNT(DISTINCT imovel_id) FROM limpezas WHERE date(data) >= date(?)', (data_inicio,)).fetchone()[0])
    doc.add_paragraph(f"Total a receber: {formatar_moeda(total_geral)}", style='Heading 2')
    
    # Salvar documento
    nome_arquivo = f"relatorios/Relatorio_{datetime.now().strftime('%Y%m%d')}.docx"
    doc.save(nome_arquivo)
    conn.close()
    
    return nome_arquivo

# =============================================
# INTERFACE PRINCIPAL
# =============================================

class SistemaGestaoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Gestão de Serviços para Airbnb")
        self.root.geometry("1100x700")
        self.root.configure(bg=COR_FUNDO)
        
        # Configurações de estilo
        self.fonte_titulo = ("Helvetica", 14, "bold")
        self.fonte_normal = ("Helvetica", 11)
        self.fonte_pequena = ("Helvetica", 9)
        
        # Inicializar banco de dados
        inicializar_banco_dados()
        
        # Criar layout principal
        self.criar_cabecalho()
        self.criar_menu_lateral()
        self.criar_area_conteudo()
        
        # Mostrar tela inicial
        self.mostrar_tela("dashboard")
    
    def criar_cabecalho(self):
        """Cria o cabeçalho do sistema"""
        frame = Frame(self.root, bg=COR_PRIMARIA, height=70)
        frame.pack(fill=X)
        
        Label(frame, 
              text="GESTÃO DE SERVIÇOS ", 
              bg=COR_PRIMARIA, 
              fg="white", 
              font=("Helvetica", 16, "bold")).pack(side=LEFT, padx=20)
        
        self.label_data = Label(frame, 
                              text=datetime.now().strftime("%d/%m/%Y %H:%M"),
                              bg=COR_PRIMARIA,
                              fg="white",
                              font=self.fonte_normal)
        self.label_data.pack(side=RIGHT, padx=20)
    
    def criar_menu_lateral(self):
        """Cria o menu de navegação lateral"""
        self.frame_menu = Frame(self.root, bg=COR_SECUNDARIA, width=200)
        self.frame_menu.pack(side=LEFT, fill=Y)
        
        # Menu de módulos
        modulos = [
            ("Dashboard", "dashboard"),
            ("Clientes", "clientes"),
            ("Imóveis", "imoveis"),
            ("Limpeza", "limpeza"),
            ("Enxoval", "enxoval"),
            ("Configurar Itens", "config_itens"),
            ("Suprimentos", "suprimentos"),
            ("Relatórios", "relatorios")
        ]
        
        for texto, comando in modulos:
            btn = Button(self.frame_menu,
                        text=texto,
                        command=lambda c=comando: self.mostrar_tela(c),
                        bg=COR_SECUNDARIA,
                        fg="white",
                        activebackground=COR_DESTAQUE,
                        activeforeground="white",
                        borderwidth=0,
                        font=self.fonte_normal,
                        anchor="w",
                        padx=20,
                        pady=12)
            btn.pack(fill=X)
    
    def criar_area_conteudo(self):
        """Cria a área de conteúdo principal"""
        self.frame_conteudo = Frame(self.root, bg=COR_FUNDO)
        self.frame_conteudo.pack(fill=BOTH, expand=True, padx=20, pady=20)
        
        # Criar todos os frames de conteúdo
        self.criar_dashboard()
        self.criar_clientes()
        self.criar_imoveis()
        self.criar_limpeza()
        self.criar_enxoval()
        self.criar_suprimentos()
        self.criar_relatorios()
        self.criar_config_itens()

    def mostrar_tela(self, tela):
        """Controla qual tela mostrar"""
        for widget in self.frame_conteudo.winfo_children():
            widget.pack_forget()
        
        if tela == "dashboard":
            self.frame_dashboard.pack(fill=BOTH, expand=True)
            self.atualizar_dashboard()
        elif tela == "clientes":
            self.frame_clientes.pack(fill=BOTH, expand=True)
            self.carregar_clientes()
        elif tela == "imoveis":
            self.frame_imoveis.pack(fill=BOTH, expand=True)
            self.carregar_imoveis()
        elif tela == "limpeza":
            self.frame_limpeza.pack(fill=BOTH, expand=True)
            self.carregar_limpezas()
        elif tela == "enxoval":
            self.frame_enxoval.pack(fill=BOTH, expand=True)
            self.carregar_itens_enxoval()
        elif tela == "suprimentos":
            self.frame_suprimentos.pack(fill=BOTH, expand=True)
            self.carregar_suprimentos()
        elif tela == "relatorios":
            self.frame_relatorios.pack(fill=BOTH, expand=True)
        elif tela == "config_itens":
            self.frame_config_itens.pack(fill=BOTH, expand=True)

    # =============================================
    # MÓDULO DASHBOARD
    # =============================================

    def criar_dashboard(self):
        """Cria o dashboard inicial"""
        self.frame_dashboard = Frame(self.frame_conteudo, bg=COR_FUNDO)

        # Frame para os cards de resumo
        frame_cards = Frame(self.frame_dashboard, bg=COR_FUNDO)
        frame_cards.pack(fill=X, pady=10)

        # Cards de resumo
        self.card_limpezas = self.criar_card(frame_cards, "Limpezas Esta Semana", "0", COR_DESTAQUE)
        self.card_enxoval = self.criar_card(frame_cards, "Enxoval Utilizado", "0", COR_SECUNDARIA)
        self.card_suprimentos = self.criar_card(frame_cards, "Suprimentos Repostos", "0", COR_PRIMARIA)
        self.card_receber = self.criar_card(frame_cards, "Total a Receber", "R$ 0,00", COR_SUCESSO)

        # Frame para os gráficos
        frame_graficos = Frame(self.frame_dashboard, bg=COR_FUNDO)
        frame_graficos.pack(fill=BOTH, expand=True)

        # ============================
        # Gráfico de Limpezas por Dia
        # ============================
        frame_grafico1 = Frame(frame_graficos, bg=COR_FUNDO)
        frame_grafico1.pack(side=LEFT, fill=BOTH, expand=True, padx=5, pady=5)
        Label(frame_grafico1, text="Limpezas por Dia", bg=COR_FUNDO, font=self.fonte_titulo).pack()

        fig1 = Figure(figsize=(4, 2.5), dpi=100)
        ax1 = fig1.add_subplot(111)
        dias = ['Seg', 'Ter', 'Qua', 'Qui', 'Sex']
        quantidades = [3, 5, 2, 4, 6]
        ax1.bar(dias, quantidades, color='#4CAF50')  # verde moderno
        ax1.set_facecolor('#FFFFFF')
        fig1.patch.set_facecolor(COR_CARD)
        ax1.tick_params(labelsize=8)
        ax1.spines['top'].set_visible(False)
        ax1.spines['right'].set_visible(False)

        self.canvas_grafico1 = FigureCanvasTkAgg(fig1, master=frame_grafico1)
        self.canvas_grafico1.draw()
        self.canvas_grafico1.get_tk_widget().pack(fill=BOTH, expand=True)

        # ==========================================
        # Gráfico de Itens de Enxoval Mais Utilizados
        # ==========================================
        frame_grafico2 = Frame(frame_graficos, bg=COR_FUNDO)
        frame_grafico2.pack(side=LEFT, fill=BOTH, expand=True, padx=5, pady=5)
        Label(frame_grafico2, text="Itens de Enxoval Mais Utilizados", bg=COR_FUNDO, font=self.fonte_titulo).pack()

        fig2 = Figure(figsize=(4, 2.5), dpi=100)
        ax2 = fig2.add_subplot(111)
        itens = ['Toalhas', 'Lençóis', 'Fronhas', 'Cobertores']
        usos = [40, 30, 20, 10]
        cores = ['#2196F3', '#FFC107', '#FF5722', '#9C27B0']
        ax2.pie(usos, labels=itens, autopct='%1.1f%%', colors=cores, startangle=140)
        fig2.patch.set_facecolor(COR_CARD)

        self.canvas_grafico2 = FigureCanvasTkAgg(fig2, master=frame_grafico2)
        self.canvas_grafico2.draw()
        self.canvas_grafico2.get_tk_widget().pack(fill=BOTH, expand=True)

    # ...existing code...

    def criar_card(self, parent, titulo, valor, cor):
        """Cria um card de resumo para o dashboard"""
        card = Frame(parent, bg=COR_CARD, bd=0, 
                    highlightthickness=1, highlightbackground="#e0e0e0")
        card.pack(side=LEFT, padx=10, fill=BOTH, expand=True)
        
        Label(card, text=titulo, bg=COR_CARD, fg=COR_TEXTO, 
              font=self.fonte_titulo).pack(pady=(10, 5), padx=10, anchor="w")
        
        label_valor = Label(card, text=valor, bg=COR_CARD, fg=cor, 
                          font=("Helvetica", 24, "bold"))
        label_valor.pack(pady=(0, 10))
        
        return label_valor
    
    def atualizar_dashboard(self):
        """Atualiza os dados do dashboard"""
        conn = sqlite3.connect("sistema.db")
        cursor = conn.cursor()
        
        # Limpezas esta semana
        cursor.execute("""
            SELECT COUNT(*), SUM(valor_total) 
            FROM limpezas 
            WHERE date(data) >= date('now', '-7 days')
        """)
        limpezas = cursor.fetchone()
        self.card_limpezas.config(text=f"{limpezas[0]}\n{formatar_moeda(limpezas[1] or 0)}")
        
        # Enxoval utilizado
        cursor.execute("""
            SELECT COUNT(*), SUM(t.preco_unitario * c.quantidade)
            FROM consumo_enxoval c
            JOIN tipos_enxoval t ON c.item_id = t.id
            WHERE date(c.data) >= date('now', '-7 days')
        """)
        enxoval = cursor.fetchone()
        self.card_enxoval.config(text=f"{enxoval[0]}\n{formatar_moeda(enxoval[1] or 0)}")
        
        # Suprimentos repostos
        cursor.execute("""
            SELECT COUNT(*), SUM(valor_gasto)
            FROM reposicao_suprimentos
            WHERE date(data) >= date('now', '-7 days')
        """)
        suprimentos = cursor.fetchone()
        self.card_suprimentos.config(text=f"{suprimentos[0]}\n{formatar_moeda(suprimentos[1] or 0)}")
        
        # Total a receber
        total = (limpezas[1] or 0) + (enxoval[1] or 0) + (50.0 * cursor.execute("SELECT COUNT(DISTINCT imovel_id) FROM limpezas WHERE date(data) >= date('now', '-7 days')").fetchone()[0])
        self.card_receber.config(text=formatar_moeda(total))
        
        # Gerar gráficos
        self.gerar_grafico_limpezas()
        self.gerar_grafico_enxoval()
        
        conn.close()
    
    def gerar_grafico_limpezas(self):
        """Gera gráfico de limpezas por dia"""
        conn = sqlite3.connect("sistema.db")
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT date(data), COUNT(*), SUM(valor_total)
            FROM limpezas
            WHERE date(data) >= date('now', '-30 days')
            GROUP BY date(data)
            ORDER BY date(data)
        """)
        
        datas = []
        quantidades = []
        valores = []
        
        for row in cursor.fetchall():
            datas.append(row[0][5:])  # Mostrar apenas dia/mês
            quantidades.append(row[1])
            valores.append(row[2])
        
        fig, ax = plt.subplots(figsize=(6, 3))
        ax.bar(datas, quantidades, color=COR_DESTAQUE)
        ax.set_title('Limpezas por Dia (últimos 30 dias)')
        ax.set_ylabel('Quantidade')
        plt.xticks(rotation=45)
        plt.tight_layout()
        
        # Salvar e exibir no canvas
        caminho = "grafico_limpezas.png"
        plt.savefig(caminho, dpi=80)
        plt.close()
        
        self.exibir_imagem_no_canvas(caminho, self.canvas_grafico1)
        conn.close()
    
    def gerar_grafico_enxoval(self):
        """Gera gráfico de itens de enxoval mais utilizados"""
        conn = sqlite3.connect("sistema.db")
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT t.nome, SUM(c.quantidade)
            FROM consumo_enxoval c
            JOIN tipos_enxoval t ON c.item_id = t.id
            WHERE date(c.data) >= date('now', '-30 days')
            GROUP BY t.nome
            ORDER BY SUM(c.quantidade) DESC
            LIMIT 5
        """)
        
        itens = []
        quantidades = []
        
        for row in cursor.fetchall():
            itens.append(row[0])
            quantidades.append(row[1])
        
        fig, ax = plt.subplots(figsize=(6, 3))
        ax.barh(itens, quantidades, color=COR_SECUNDARIA)
        ax.set_title('Itens Mais Utilizados (últimos 30 dias)')
        ax.set_xlabel('Quantidade')
        plt.tight_layout()
        
        # Salvar e exibir no canvas
        caminho = "grafico_enxoval.png"
        plt.savefig(caminho, dpi=80)
        plt.close()
        
        self.exibir_imagem_no_canvas(caminho, self.canvas_grafico2)
        conn.close()
    
    def exibir_imagem_no_canvas(self, caminho_imagem, canvas):
        from PIL import Image, ImageTk

        # Acesse o widget Tkinter real dentro do FigureCanvasTkAgg
        tk_widget = canvas.get_tk_widget()

        # Obtenha dimensões reais do widget
        largura = tk_widget.winfo_width()
        altura = tk_widget.winfo_height()

        # Se ainda não estiver visível, espera o layout carregar
        if largura == 1 or altura == 1:
            # Agenda nova tentativa após 100ms
            self.frame_dashboard.after(100, lambda: self.exibir_imagem_no_canvas(caminho_imagem, canvas))
            return

        # Carrega e redimensiona a imagem
        img = Image.open(caminho_imagem)
        img = img.resize((largura, altura), Image.LANCZOS)
        img_tk = ImageTk.PhotoImage(img)

        # Armazena referência da imagem para evitar garbage collection
        canvas._img_ref = img_tk

        # Adiciona a imagem ao canvas
        canvas_widget = tk_widget
        label = Label(canvas_widget, image=img_tk, bg=COR_CARD)
        label.place(relx=0.5, rely=0.5, anchor='center')

    
    # =============================================
    # MÓDULO CLIENTES
    # =============================================
    
    def criar_clientes(self):
        """Cria a interface para gestão de clientes"""
        self.frame_clientes = Frame(self.frame_conteudo, bg=COR_FUNDO)
        
        # Treeview para listar clientes
        frame_tree = Frame(self.frame_clientes, bg=COR_FUNDO)
        frame_tree.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        scrollbar = Scrollbar(frame_tree)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        self.tree_clientes = ttk.Treeview(frame_tree, columns=('id', 'nome', 'telefone', 'email'), 
                                        yscrollcommand=scrollbar.set)
        self.tree_clientes.pack(fill=BOTH, expand=True)
        scrollbar.config(command=self.tree_clientes.yview)
        
        self.tree_clientes.heading('#0', text='ID')
        self.tree_clientes.heading('#1', text='Nome')
        self.tree_clientes.heading('#2', text='Telefone')
        self.tree_clientes.heading('#3', text='Email')
        
        # Formulário para adicionar/editar clientes
        frame_form = Frame(self.frame_clientes, bg=COR_CARD, bd=0, 
                         highlightthickness=1, highlightbackground="#e0e0e0")
        frame_form.pack(fill=X, padx=10, pady=10)
        
        Label(frame_form, text="Cadastro de Clientes", bg=COR_CARD, 
              fg=COR_TEXTO, font=self.fonte_titulo).pack(pady=(10, 5), anchor="w", padx=10)
        
        # Campos do formulário
        campos = [
            ("Nome:", Entry(frame_form)),
            ("Telefone:", Entry(frame_form)),
            ("Email:", Entry(frame_form)),
            ("Endereço:", Entry(frame_form))
        ]
        
        self.entry_cliente_nome = campos[0][1]
        self.entry_cliente_telefone = campos[1][1]
        self.entry_cliente_email = campos[2][1]
        self.entry_cliente_endereco = campos[3][1]
        
        for texto, widget in campos:
            frame = Frame(frame_form, bg=COR_CARD)
            frame.pack(fill=X, padx=10, pady=5)
            Label(frame, text=texto, bg=COR_CARD).pack(side=LEFT, padx=5)
            widget.pack(side=LEFT, expand=True, fill=X)
        
        # Botões
        frame_botoes = Frame(frame_form, bg=COR_CARD)
        frame_botoes.pack(fill=X, padx=10, pady=10)
        
        Button(frame_botoes, text="Adicionar", command=self.adicionar_cliente,
              bg=COR_DESTAQUE, fg="white").pack(side=LEFT, padx=5)
        Button(frame_botoes, text="Limpar", command=self.limpar_form_cliente,
              bg=COR_SECUNDARIA, fg="white").pack(side=LEFT, padx=5)
    
    def carregar_clientes(self):
        """Carrega os clientes no TreeView"""
        conn = sqlite3.connect("sistema.db")
        cursor = conn.cursor()
        cursor.execute("SELECT id, nome, telefone, email FROM clientes")
        
        # Limpar treeview
        for item in self.tree_clientes.get_children():
            self.tree_clientes.delete(item)
        
        # Adicionar novos itens
        for row in cursor.fetchall():
            self.tree_clientes.insert('', 'end', values=row)
        
        conn.close()
    
    def adicionar_cliente(self):
        """Adiciona um novo cliente ao banco de dados"""
        nome = self.entry_cliente_nome.get()
        telefone = self.entry_cliente_telefone.get()
        email = self.entry_cliente_email.get()
        endereco = self.entry_cliente_endereco.get()
        
        if nome:
            conn = sqlite3.connect("sistema.db")
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO clientes (nome, telefone, email, endereco)
                VALUES (?, ?, ?, ?)
            """, (nome, telefone, email, endereco))
            conn.commit()
            conn.close()
            
            messagebox.showinfo("Sucesso", "Cliente cadastrado com sucesso!")
            self.carregar_clientes()
            self.limpar_form_cliente()
        else:
            messagebox.showerror("Erro", "Informe pelo menos o nome do cliente")
    
    def limpar_form_cliente(self):
        """Limpa o formulário de clientes"""
        self.entry_cliente_nome.delete(0, END)
        self.entry_cliente_telefone.delete(0, END)
        self.entry_cliente_email.delete(0, END)
        self.entry_cliente_endereco.delete(0, END)
    
    # =============================================
    # MÓDULO IMÓVEIS
    # =============================================
    
    def criar_imoveis(self):
        """Cria a interface para gestão de imóveis"""
        self.frame_imoveis = Frame(self.frame_conteudo, bg=COR_FUNDO)
        
        # Treeview (similar ao exemplo anterior)
        
        # Formulário compacto
        frame_form = Frame(self.frame_imoveis, bg=COR_CARD, bd=1, relief=SOLID)
        frame_form.pack(fill=X, padx=5, pady=5)
        
        Label(frame_form, text="Cadastro de Imóveis", bg=COR_CARD, 
            fg=COR_TEXTO, font=self.fonte_pequena).grid(row=0, column=0, columnspan=2, sticky='w', pady=(5, 0))
        
        # Cliente
        Label(frame_form, text="Cliente:", bg=COR_CARD, font=self.fonte_pequena).grid(
            row=1, column=0, sticky='e', padx=2, pady=1)
        
        self.combo_cliente_imovel = ttk.Combobox(frame_form, width=28, font=self.fonte_pequena)
        self.combo_cliente_imovel.grid(row=1, column=1, sticky='we', padx=2, pady=1)
        
        # Outros campos
        campos = [
            ("Endereço:", Entry(frame_form, width=30, font=self.fonte_pequena)),
            ("Quartos:", Entry(frame_form, width=5, font=self.fonte_pequena)),
            ("Banheiros:", Entry(frame_form, width=5, font=self.fonte_pequena))
        ]
        
        for i, (texto, widget) in enumerate(campos, start=2):
            Label(frame_form, text=texto, bg=COR_CARD, font=self.fonte_pequena).grid(
                row=i, column=0, sticky='e', padx=2, pady=1)
            widget.grid(row=i, column=1, sticky='w', padx=2, pady=1)
        
        self.entry_imovel_endereco, self.entry_imovel_quartos, self.entry_imovel_banheiros = [c[1] for c in campos]
        
        # Botões
        frame_botoes = Frame(frame_form, bg=COR_CARD)
        frame_botoes.grid(row=len(campos)+2, column=0, columnspan=2, pady=3)
        
        Button(frame_botoes, text="Adicionar", command=self.adicionar_imovel,
            bg=COR_DESTAQUE, fg="white", font=self.fonte_pequena, padx=5).pack(side=LEFT, padx=2)
        Button(frame_botoes, text="Limpar", command=self.limpar_form_imovel,
            bg=COR_SECUNDARIA, fg="white", font=self.fonte_pequena, padx=5).pack(side=LEFT, padx=2)
        
        frame_form.columnconfigure(1, weight=1)



    def carregar_imoveis(self):
        """Carrega os imóveis no TreeView e atualiza o combobox de clientes"""
        conn = sqlite3.connect("sistema.db")
        cursor = conn.cursor()
        
        # Carregar clientes no combobox
        cursor.execute("SELECT id, nome FROM clientes")
        clientes = cursor.fetchall()
        self.combo_cliente_imovel['values'] = [f"{c[0]} - {c[1]}" for c in clientes]
        
        if clientes:
            self.combo_cliente_imovel.current(0)
        
        # Carregar imóveis no treeview
        cursor.execute("""
            SELECT i.id, c.nome, i.endereco, i.quartos, i.banheiros
            FROM imoveis i
            JOIN clientes c ON i.cliente_id = c.id
        """)
        
        # Limpar treeview
        for item in self.tree_imoveis.get_children():
            self.tree_imoveis.delete(item)
        
        # Adicionar novos itens
        for row in cursor.fetchall():
            self.tree_imoveis.insert('', 'end', values=row)
        
        conn.close()
    
    def adicionar_imovel(self):
        """Adiciona um novo imóvel ao banco de dados"""
        cliente_id = self.combo_cliente_imovel.get().split(" - ")[0]
        endereco = self.entry_imovel_endereco.get()
        quartos = self.entry_imovel_quartos.get()
        banheiros = self.entry_imovel_banheiros.get()
        
        if cliente_id and endereco:
            try:
                quartos = int(quartos) if quartos else 0
                banheiros = int(banheiros) if banheiros else 0
                
                conn = sqlite3.connect("sistema.db")
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO imoveis (cliente_id, endereco, quartos, banheiros)
                    VALUES (?, ?, ?, ?)
                """, (cliente_id, endereco, quartos, banheiros))
                conn.commit()
                conn.close()
                
                messagebox.showinfo("Sucesso", "Imóvel cadastrado com sucesso!")
                self.carregar_imoveis()
                self.limpar_form_imovel()
            except ValueError:
                messagebox.showerror("Erro", "Quartos e banheiros devem ser números inteiros")
        else:
            messagebox.showerror("Erro", "Informe pelo menos o cliente e o endereço")
    
    def limpar_form_imovel(self):
        """Limpa o formulário de imóveis"""
        self.entry_imovel_endereco.delete(0, END)
        self.entry_imovel_quartos.delete(0, END)
        self.entry_imovel_banheiros.delete(0, END)
    
    # =============================================
    # MÓDULO LIMPEZA
    # =============================================
    
    def criar_limpeza(self):
        """Cria a interface para gestão de limpezas"""
        self.frame_limpeza = Frame(self.frame_conteudo, bg=COR_FUNDO)
        
        # Treeview para listar limpezas
        frame_tree = Frame(self.frame_limpeza, bg=COR_FUNDO)
        frame_tree.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        scrollbar = Scrollbar(frame_tree)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        self.tree_limpezas = ttk.Treeview(frame_tree, 
                                        columns=('id', 'imovel', 'data', 'horas', 'valor_hora', 'valor_total'),
                                        yscrollcommand=scrollbar.set)
        self.tree_limpezas.pack(fill=BOTH, expand=True)
        scrollbar.config(command=self.tree_limpezas.yview)
        
        self.tree_limpezas.heading('#0', text='ID')
        self.tree_limpezas.heading('#1', text='Imóvel')
        self.tree_limpezas.heading('#2', text='Data')
        self.tree_limpezas.heading('#3', text='Horas')
        self.tree_limpezas.heading('#4', text='Valor/Hora')
        self.tree_limpezas.heading('#5', text='Valor Total')
        
        # Formulário para adicionar limpezas
        frame_form = Frame(self.frame_limpeza, bg=COR_CARD, bd=0, 
                         highlightthickness=1, highlightbackground="#e0e0e0")
        frame_form.pack(fill=X, padx=10, pady=10)
        
        Label(frame_form, text="Registro de Limpeza", bg=COR_CARD, 
              fg=COR_TEXTO, font=self.fonte_titulo).pack(pady=(10, 5), anchor="w", padx=10)
        
        # Combobox para selecionar imóvel
        frame_imovel = Frame(frame_form, bg=COR_CARD)
        frame_imovel.pack(fill=X, padx=10, pady=5)
        Label(frame_imovel, text="Imóvel:", bg=COR_CARD).pack(side=LEFT, padx=5)
        self.combo_imovel_limpeza = ttk.Combobox(frame_imovel)
        self.combo_imovel_limpeza.pack(side=LEFT, expand=True, fill=X)
        
        # Campos de data e horários
        campos = [
            ("Data:", DateEntry(frame_form, date_pattern='dd/mm/yyyy')),
            ("Hora Início (HH:MM):", Entry(frame_form)),
            ("Hora Fim (HH:MM):", Entry(frame_form)),
            ("Valor por Hora (R$):", Entry(frame_form)),
            ("Observações:", Text(frame_form, height=3))
        ]
        
        self.entry_limpeza_data = campos[0][1]
        self.entry_limpeza_hora_inicio = campos[1][1]
        self.entry_limpeza_hora_fim = campos[2][1]
        self.entry_limpeza_valor_hora = campos[3][1]
        self.entry_limpeza_observacoes = campos[4][1]
        
        # Configurar valor padrão para hora
        self.entry_limpeza_valor_hora.insert(0, "30.00")
        
        for texto, widget in campos:
            frame = Frame(frame_form, bg=COR_CARD)
            frame.pack(fill=X, padx=10, pady=5)
            Label(frame, text=texto, bg=COR_CARD).pack(side=LEFT, padx=5)
            widget.pack(side=LEFT, expand=True, fill=X)
        
        # Botões
        frame_botoes = Frame(frame_form, bg=COR_CARD)
        frame_botoes.pack(fill=X, padx=10, pady=10)
        
        Button(frame_botoes, text="Adicionar", command=self.adicionar_limpeza,
              bg=COR_DESTAQUE, fg="white").pack(side=LEFT, padx=5)
        Button(frame_botoes, text="Calcular", command=self.calcular_limpeza,
              bg=COR_SECUNDARIA, fg="white").pack(side=LEFT, padx=5)
        Button(frame_botoes, text="Limpar", command=self.limpar_form_limpeza,
              bg=COR_ALERTA, fg="white").pack(side=LEFT, padx=5)
    
    def carregar_limpezas(self):
        """Carrega as limpezas no TreeView e atualiza o combobox de imóveis"""
        conn = sqlite3.connect("sistema.db")
        cursor = conn.cursor()
        
        # Carregar imóveis no combobox
        cursor.execute("SELECT id, endereco FROM imoveis")
        imoveis = cursor.fetchall()
        self.combo_imovel_limpeza['values'] = [f"{i[0]} - {i[1]}" for i in imoveis]
        
        if imoveis:
            self.combo_imovel_limpeza.current(0)
        
        # Carregar limpezas no treeview
        cursor.execute("""
            SELECT l.id, i.endereco, l.data, l.horas_trabalhadas, l.valor_hora, l.valor_total
            FROM limpezas l
            JOIN imoveis i ON l.imovel_id = i.id
            ORDER BY l.data DESC
        """)
        
        # Limpar treeview
        for item in self.tree_limpezas.get_children():
            self.tree_limpezas.delete(item)
        
        # Adicionar novos itens formatados
        for row in cursor.fetchall():
            self.tree_limpezas.insert('', 'end', values=(
                row[0], row[1], row[2], 
                f"{row[3]:.2f}h", 
                formatar_moeda(row[4]), 
                formatar_moeda(row[5])
            ))
        
        conn.close()
    
    def calcular_limpeza(self):
        """Calcula o valor da limpeza com base nas horas trabalhadas"""
        hora_inicio = self.entry_limpeza_hora_inicio.get()
        hora_fim = self.entry_limpeza_hora_fim.get()
        valor_hora = self.entry_limpeza_valor_hora.get()
        
        try:
            horas = calcular_horas(hora_inicio, hora_fim)
            valor = horas * float(valor_hora)
            messagebox.showinfo("Cálculo", 
                              f"Horas trabalhadas: {horas:.2f}h\n"
                              f"Valor total: {formatar_moeda(valor)}")
        except ValueError:
            messagebox.showerror("Erro", "Informe horários no formato HH:MM e valor por hora numérico")
    
    def adicionar_limpeza(self):
        """Adiciona uma nova limpeza ao banco de dados"""
        imovel_id = self.combo_imovel_limpeza.get().split(" - ")[0]
        data = self.entry_limpeza_data.get_date()
        hora_inicio = self.entry_limpeza_hora_inicio.get()
        hora_fim = self.entry_limpeza_hora_fim.get()
        valor_hora = self.entry_limpeza_valor_hora.get()
        observacoes = self.entry_limpeza_observacoes.get("1.0", END).strip()
        
        if imovel_id and data and hora_inicio and hora_fim and valor_hora:
            try:
                horas = calcular_horas(hora_inicio, hora_fim)
                valor_total = horas * float(valor_hora)
                
                conn = sqlite3.connect("sistema.db")
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO limpezas 
                    (imovel_id, data, hora_inicio, hora_fim, horas_trabalhadas, valor_hora, valor_total, observacoes)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (imovel_id, data, hora_inicio, hora_fim, horas, float(valor_hora), valor_total, observacoes))
                conn.commit()
                conn.close()
                
                messagebox.showinfo("Sucesso", "Limpeza registrada com sucesso!")
                self.carregar_limpezas()
                self.limpar_form_limpeza()
            except ValueError:
                messagebox.showerror("Erro", "Valor por hora deve ser numérico")
        else:
            messagebox.showerror("Erro", "Preencha todos os campos obrigatórios")
    
    def limpar_form_limpeza(self):
        """Limpa o formulário de limpeza"""
        self.entry_limpeza_hora_inicio.delete(0, END)
        self.entry_limpeza_hora_fim.delete(0, END)
        self.entry_limpeza_observacoes.delete("1.0", END)
    
    # =============================================
    # MÓDULO ENXOVAL
    # =============================================
    
    def criar_enxoval(self):
        """Cria a interface para gestão de enxoval"""
        self.frame_enxoval = Frame(self.frame_conteudo, bg=COR_FUNDO)
        
        # Treeview para listar consumo de enxoval
        frame_tree = Frame(self.frame_enxoval, bg=COR_FUNDO)
        frame_tree.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        scrollbar = Scrollbar(frame_tree)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        self.tree_enxoval = ttk.Treeview(frame_tree, 
                                       columns=('id', 'imovel', 'item', 'quantidade', 'data', 'valor_unitario', 'valor_total'),
                                       yscrollcommand=scrollbar.set)
        self.tree_enxoval.pack(fill=BOTH, expand=True)
        scrollbar.config(command=self.tree_enxoval.yview)
        
        self.tree_enxoval.heading('#0', text='ID')
        self.tree_enxoval.heading('#1', text='Imóvel')
        self.tree_enxoval.heading('#2', text='Item')
        self.tree_enxoval.heading('#3', text='Quantidade')
        self.tree_enxoval.heading('#4', text='Data')
        self.tree_enxoval.heading('#5', text='Valor Unitário')
        self.tree_enxoval.heading('#6', text='Valor Total')
        
        # Formulário para registrar consumo de enxoval
        frame_form = Frame(self.frame_enxoval, bg=COR_CARD, bd=0, 
                          highlightthickness=1, highlightbackground="#e0e0e0")
        frame_form.pack(fill=X, padx=10, pady=10)
        
        Label(frame_form, text="Registro de Consumo de Enxoval", bg=COR_CARD, 
              fg=COR_TEXTO, font=self.fonte_titulo).pack(pady=(10, 5), anchor="w", padx=10)
        
        # Combobox para selecionar imóvel
        frame_imovel = Frame(frame_form, bg=COR_CARD)
        frame_imovel.pack(fill=X, padx=10, pady=5)
        Label(frame_imovel, text="Imóvel:", bg=COR_CARD).pack(side=LEFT, padx=5)
        self.combo_imovel_enxoval = ttk.Combobox(frame_imovel)
        self.combo_imovel_enxoval.pack(side=LEFT, expand=True, fill=X)
        
        # Combobox para selecionar item
        frame_item = Frame(frame_form, bg=COR_CARD)
        frame_item.pack(fill=X, padx=10, pady=5)
        Label(frame_item, text="Item:", bg=COR_CARD).pack(side=LEFT, padx=5)
        self.combo_item_enxoval = ttk.Combobox(frame_item)
        self.combo_item_enxoval.pack(side=LEFT, expand=True, fill=X)
        
        # Campos de quantidade e data
        campos = [
            ("Quantidade:", Entry(frame_form)),
            ("Data:", DateEntry(frame_form, date_pattern='dd/mm/yyyy'))
        ]
        
        self.entry_enxoval_quantidade = campos[0][1]
        self.entry_enxoval_data = campos[1][1]
        
        for texto, widget in campos:
            frame = Frame(frame_form, bg=COR_CARD)
            frame.pack(fill=X, padx=10, pady=5)
            Label(frame, text=texto, bg=COR_CARD).pack(side=LEFT, padx=5)
            widget.pack(side=LEFT, expand=True, fill=X)
        
        # Botões
        frame_botoes = Frame(frame_form, bg=COR_CARD)
        frame_botoes.pack(fill=X, padx=10, pady=10)
        
        Button(frame_botoes, text="Adicionar", command=self.adicionar_consumo_enxoval,
              bg=COR_DESTAQUE, fg="white").pack(side=LEFT, padx=5)
        Button(frame_botoes, text="Limpar", command=self.limpar_form_enxoval,
              bg=COR_SECUNDARIA, fg="white").pack(side=LEFT, padx=5)
    
    def carregar_itens_enxoval(self):
        """Carrega os itens de enxoval no TreeView e atualiza os comboboxes"""
        conn = sqlite3.connect("sistema.db")
        cursor = conn.cursor()
        
        # Carregar imóveis no combobox
        cursor.execute("SELECT id, endereco FROM imoveis")
        imoveis = cursor.fetchall()
        self.combo_imovel_enxoval['values'] = [f"{i[0]} - {i[1]}" for i in imoveis]
        
        if imoveis:
            self.combo_imovel_enxoval.current(0)
        
        # Carregar itens de enxoval no combobox
        cursor.execute("SELECT id, nome, preco_unitario FROM tipos_enxoval")
        itens = cursor.fetchall()
        self.combo_item_enxoval['values'] = [f"{i[0]} - {i[1]} (R$ {i[2]:.2f})" for i in itens]
        
        if itens:
            self.combo_item_enxoval.current(0)
        
        # Carregar consumo de enxoval no treeview
        cursor.execute("""
            SELECT c.id, i.endereco, t.nome, c.quantidade, c.data, t.preco_unitario, 
                   (c.quantidade * t.preco_unitario) as valor_total
            FROM consumo_enxoval c
            JOIN imoveis i ON c.imovel_id = i.id
            JOIN tipos_enxoval t ON c.item_id = t.id
            ORDER BY c.data DESC
        """)
        
        # Limpar treeview
        for item in self.tree_enxoval.get_children():
            self.tree_enxoval.delete(item)
        
        # Adicionar novos itens formatados
        for row in cursor.fetchall():
            self.tree_enxoval.insert('', 'end', values=(
                row[0], row[1], row[2], row[3], row[4],
                formatar_moeda(row[5]), 
                formatar_moeda(row[6])
            ))
        
        conn.close()
    
    def criar_config_itens(self):
        """Cria a interface para configurar itens (enxoval vs. suprimentos)"""
        self.frame_config_itens = Frame(self.frame_conteudo, bg=COR_FUNDO)
        
        # Treeview para listar todos os itens
        frame_tree = Frame(self.frame_config_itens, bg=COR_FUNDO)
        frame_tree.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        scrollbar = Scrollbar(frame_tree)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        self.tree_itens = ttk.Treeview(frame_tree, columns=('id', 'nome', 'tipo', 'preco', 'unidade'), 
                                    yscrollcommand=scrollbar.set)
        self.tree_itens.pack(fill=BOTH, expand=True)
        scrollbar.config(command=self.tree_itens.yview)
        
        self.tree_itens.heading('#0', text='ID')
        self.tree_itens.heading('#1', text='Nome')
        self.tree_itens.heading('#2', text='Tipo')
        self.tree_itens.heading('#3', text='Preço Unitário')
        self.tree_itens.heading('#4', text='Unidade')
        
        # Formulário para adicionar/editar itens
        frame_form = Frame(self.frame_config_itens, bg=COR_CARD, bd=0, 
                        highlightthickness=1, highlightbackground="#e0e0e0")
        frame_form.pack(fill=X, padx=10, pady=10)
        
        Label(frame_form, text="Adicionar/Editar Item", bg=COR_CARD, 
            fg=COR_TEXTO, font=self.fonte_titulo).pack(pady=(10, 5), anchor="w", padx=10)
        
        # Campos do formulário
        campos = [
            ("Nome:", Entry(frame_form)),
            ("Tipo:", ttk.Combobox(frame_form, values=['enxoval', 'suprimento', 'limpeza'])),
            ("Preço Unitário (R$):", Entry(frame_form)),
            ("Unidade de Medida:", Entry(frame_form))
        ]
        
        self.entry_item_nome = campos[0][1]
        self.combo_item_tipo = campos[1][1]
        self.entry_item_preco = campos[2][1]
        self.entry_item_unidade = campos[3][1]
        
        for texto, widget in campos:
            frame = Frame(frame_form, bg=COR_CARD)
            frame.pack(fill=X, padx=10, pady=5)
            Label(frame, text=texto, bg=COR_CARD).pack(side=LEFT, padx=5)
            widget.pack(side=LEFT, expand=True, fill=X)
        
        # Botões
        frame_botoes = Frame(frame_form, bg=COR_CARD)
        frame_botoes.pack(fill=X, padx=10, pady=10)
        
        Button(frame_botoes, text="Salvar", command=self.salvar_item_config,
            bg=COR_DESTAQUE, fg="white").pack(side=LEFT, padx=5)
        Button(frame_botoes, text="Remover", command=self.remover_item_config,
            bg=COR_ALERTA, fg="white").pack(side=LEFT, padx=5)
        Button(frame_botoes, text="Limpar", command=self.limpar_form_item,
            bg=COR_SECUNDARIA, fg="white").pack(side=LEFT, padx=5)
        
        # Carregar dados iniciais
        self.carregar_itens_config()

    def carregar_itens_config(self):
        """Carrega todos os itens para configuração"""
        conn = sqlite3.connect("sistema.db")
        cursor = conn.cursor()
        
        # Verificar se a tabela existe (caso o sistema seja atualizado)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS itens_servico (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            tipo TEXT NOT NULL CHECK(tipo IN ('enxoval', 'suprimento', 'limpeza')),
            preco_unitario REAL DEFAULT 0.0,
            unidade_medida TEXT
        );
        """)
        
        cursor.execute("SELECT id, nome, tipo, preco_unitario, unidade_medida FROM itens_servico")
        
        # Limpar treeview
        for item in self.tree_itens.get_children():
            self.tree_itens.delete(item)
        
        # Adicionar itens formatados
        for row in cursor.fetchall():
            preco = f"R$ {row[3]:.2f}" if row[3] else "N/A"
            self.tree_itens.insert('', 'end', values=(row[0], row[1], row[2], preco, row[4] or "-"))
        
        conn.close()





    def adicionar_consumo_enxoval(self):
        """Adiciona um novo consumo de enxoval ao banco de dados"""
        imovel_id = self.combo_imovel_enxoval.get().split(" - ")[0]
        item_id = self.combo_item_enxoval.get().split(" - ")[0]
        quantidade = self.entry_enxoval_quantidade.get()
        data = self.entry_enxoval_data.get_date()
        
        if imovel_id and item_id and quantidade and data:
            try:
                quantidade = int(quantidade)
                
                conn = sqlite3.connect("sistema.db")
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO consumo_enxoval 
                    (imovel_id, item_id, quantidade, data)
                    VALUES (?, ?, ?, ?)
                """, (imovel_id, item_id, quantidade, data))
                conn.commit()
                conn.close()
                
                messagebox.showinfo("Sucesso", "Consumo de enxoval registrado com sucesso!")
                self.carregar_itens_enxoval()
                self.limpar_form_enxoval()
            except ValueError:
                messagebox.showerror("Erro", "Quantidade deve ser um número inteiro")
        else:
            messagebox.showerror("Erro", "Preencha todos os campos obrigatórios")
    
    def limpar_form_enxoval(self):
        """Limpa o formulário de enxoval"""
        self.entry_enxoval_quantidade.delete(0, END)
    

    def salvar_item_config(self):
        nome = self.entry_item_nome.get()
        tipo = self.combo_item_tipo.get()
        preco = self.entry_item_preco.get()
        unidade = self.entry_item_unidade.get()
        
        if not nome or not tipo:
            messagebox.showerror("Erro", "Preencha pelo menos Nome e Tipo!")
            return
        
        try:
            preco = float(preco) if preco else 0.0
        except ValueError:
            messagebox.showerror("Erro", "Preço deve ser um número!")
            return
        
        conn = sqlite3.connect("sistema.db")
        cursor = conn.cursor()
        
        
        if hasattr(self, 'item_selecionado_id') and self.item_selecionado_id:
            cursor.execute("""
                UPDATE itens_servico 
                SET nome=?, tipo=?, preco_unitario=?, unidade_medida=?
                WHERE id=?
            """, (nome, tipo, preco, unidade, self.item_selecionado_id))
        else:
            cursor.execute("""
                INSERT INTO itens_servico (nome, tipo, preco_unitario, unidade_medida)
                VALUES (?, ?, ?, ?)
            """, (nome, tipo, preco, unidade))
        
        # Sincronizar com as tabelas usadas nos módulos
        if tipo == "enxoval":
            cursor.execute("INSERT OR IGNORE INTO tipos_enxoval (nome, preco_unitario) VALUES (?, ?)", (nome, preco))
        elif tipo == "suprimento":
            cursor.execute("INSERT OR IGNORE INTO suprimentos (nome, preco_unitario) VALUES (?, ?)", (nome, preco))
        # Se quiser, pode fazer o mesmo para "limpeza" em outra tabela
        
        conn.commit()
        conn.close()
        
        messagebox.showinfo("Sucesso", "Item salvo com sucesso!")
        self.carregar_itens_config()
        self.limpar_form_item()

    def remover_item_config(self):
        """Remove o item selecionado"""
        if not hasattr(self, 'item_selecionado_id') or not self.item_selecionado_id:
            messagebox.showerror("Erro", "Nenhum item selecionado!")
            return
        
        resposta = messagebox.askyesno("Confirmar", "Tem certeza que deseja remover este item?")
        if resposta:
            conn = sqlite3.connect("sistema.db")
            cursor = conn.cursor()
            cursor.execute("DELETE FROM itens_servico WHERE id=?", (self.item_selecionado_id,))
            conn.commit()
            conn.close()
            
            messagebox.showinfo("Sucesso", "Item removido!")
            self.carregar_itens_config()
            self.limpar_form_item()

    def limpar_form_item(self):
        """Limpa o formulário de itens"""
        self.entry_item_nome.delete(0, END)
        self.combo_item_tipo.set('')
        self.entry_item_preco.delete(0, END)
        self.entry_item_unidade.delete(0, END)
        
        if hasattr(self, 'item_selecionado_id'):
            del self.item_selecionado_id

























    # =============================================
    # MÓDULO SUPRIMENTOS
    # =============================================
    
    def criar_suprimentos(self):
        """Cria a interface para gestão de suprimentos"""
        self.frame_suprimentos = Frame(self.frame_conteudo, bg=COR_FUNDO)
        
        # Treeview para listar reposição de suprimentos
        frame_tree = Frame(self.frame_suprimentos, bg=COR_FUNDO)
        frame_tree.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        scrollbar = Scrollbar(frame_tree)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        self.tree_suprimentos = ttk.Treeview(frame_tree, 
                                           columns=('id', 'imovel', 'item', 'quantidade', 'data', 'valor_gasto', 'comprovante'),
                                           yscrollcommand=scrollbar.set)
        self.tree_suprimentos.pack(fill=BOTH, expand=True)
        scrollbar.config(command=self.tree_suprimentos.yview)
        
        self.tree_suprimentos.heading('#0', text='ID')
        self.tree_suprimentos.heading('#1', text='Imóvel')
        self.tree_suprimentos.heading('#2', text='Item')
        self.tree_suprimentos.heading('#3', text='Quantidade')
        self.tree_suprimentos.heading('#4', text='Data')
        self.tree_suprimentos.heading('#5', text='Valor Gasto')
        self.tree_suprimentos.heading('#6', text='Comprovante')
        
        # Formulário para registrar reposição de suprimentos
        frame_form = Frame(self.frame_suprimentos, bg=COR_CARD, bd=0, 
                         highlightthickness=1, highlightbackground="#e0e0e0")
        frame_form.pack(fill=X, padx=10, pady=10)
        
        Label(frame_form, text="Reposição de Suprimentos", bg=COR_CARD, 
              fg=COR_TEXTO, font=self.fonte_titulo).pack(pady=(10, 5), anchor="w", padx=10)
        
        # Combobox para selecionar imóvel
        frame_imovel = Frame(frame_form, bg=COR_CARD)
        frame_imovel.pack(fill=X, padx=10, pady=5)
        Label(frame_imovel, text="Imóvel:", bg=COR_CARD).pack(side=LEFT, padx=5)
        self.combo_imovel_suprimento = ttk.Combobox(frame_imovel)
        self.combo_imovel_suprimento.pack(side=LEFT, expand=True, fill=X)
        
        # Combobox para selecionar item
        frame_item = Frame(frame_form, bg=COR_CARD)
        frame_item.pack(fill=X, padx=10, pady=5)
        Label(frame_item, text="Item:", bg=COR_CARD).pack(side=LEFT, padx=5)
        self.combo_item_suprimento = ttk.Combobox(frame_item)
        self.combo_item_suprimento.pack(side=LEFT, expand=True, fill=X)
        
        # Campos de quantidade, data e valor
        campos = [
            ("Quantidade:", Entry(frame_form)),
            ("Data:", DateEntry(frame_form, date_pattern='dd/mm/yyyy')),
            ("Valor Gasto (R$):", Entry(frame_form))
        ]
        
        self.entry_suprimento_quantidade = campos[0][1]
        self.entry_suprimento_data = campos[1][1]
        self.entry_suprimento_valor = campos[2][1]
        
        for texto, widget in campos:
            frame = Frame(frame_form, bg=COR_CARD)
            frame.pack(fill=X, padx=10, pady=5)
            Label(frame, text=texto, bg=COR_CARD).pack(side=LEFT, padx=5)
            widget.pack(side=LEFT, expand=True, fill=X)
        
        # Campo para comprovante
        frame_comprovante = Frame(frame_form, bg=COR_CARD)
        frame_comprovante.pack(fill=X, padx=10, pady=5)
        Label(frame_comprovante, text="Comprovante:", bg=COR_CARD).pack(side=LEFT, padx=5)
        self.entry_suprimento_comprovante = Entry(frame_comprovante)
        self.entry_suprimento_comprovante.pack(side=LEFT, expand=True, fill=X)
        Button(frame_comprovante, text="Selecionar", command=lambda: selecionar_arquivo(self.entry_suprimento_comprovante),
              bg=COR_SECUNDARIA, fg="white").pack(side=LEFT, padx=5)
        
        # Botões
        frame_botoes = Frame(frame_form, bg=COR_CARD)
        frame_botoes.pack(fill=X, padx=10, pady=10)
        
        Button(frame_botoes, text="Adicionar", command=self.adicionar_reposicao_suprimento,
              bg=COR_DESTAQUE, fg="white").pack(side=LEFT, padx=5)
        Button(frame_botoes, text="Limpar", command=self.limpar_form_suprimento,
              bg=COR_SECUNDARIA, fg="white").pack(side=LEFT, padx=5)
    
    def carregar_suprimentos(self):
        """Carrega os suprimentos no TreeView e atualiza os comboboxes"""
        conn = sqlite3.connect("sistema.db")
        cursor = conn.cursor()
        
        # Carregar imóveis no combobox
        cursor.execute("SELECT id, endereco FROM imoveis")
        imoveis = cursor.fetchall()
        self.combo_imovel_suprimento['values'] = [f"{i[0]} - {i[1]}" for i in imoveis]
        
        if imoveis:
            self.combo_imovel_suprimento.current(0)
        
        # Carregar itens de suprimento no combobox
        cursor.execute("SELECT id, nome FROM suprimentos")
        itens = cursor.fetchall()
        self.combo_item_suprimento['values'] = [f"{i[0]} - {i[1]}" for i in itens]
        
        if itens:
            self.combo_item_suprimento.current(0)
        
        # Carregar reposição de suprimentos no treeview
        cursor.execute("""
            SELECT r.id, i.endereco, s.nome, r.quantidade, r.data, r.valor_gasto, r.comprovante_path
            FROM reposicao_suprimentos r
            JOIN imoveis i ON r.imovel_id = i.id
            JOIN suprimentos s ON r.suprimento_id = s.id
            ORDER BY r.data DESC
        """)
        
        # Limpar treeview
        for item in self.tree_suprimentos.get_children():
            self.tree_suprimentos.delete(item)
        
        # Adicionar novos itens formatados
        for row in cursor.fetchall():
            comprovante = "Sim" if row[6] else "Não"
            self.tree_suprimentos.insert('', 'end', values=(
                row[0], row[1], row[2], row[3], row[4],
                formatar_moeda(row[5]), 
                comprovante
            ))
        
        conn.close()
    
    def adicionar_reposicao_suprimento(self):
        """Adiciona uma nova reposição de suprimento ao banco de dados"""
        imovel_id = self.combo_imovel_suprimento.get().split(" - ")[0]
        item_id = self.combo_item_suprimento.get().split(" - ")[0]
        quantidade = self.entry_suprimento_quantidade.get()
        data = self.entry_suprimento_data.get_date()
        valor = self.entry_suprimento_valor.get()
        comprovante = self.entry_suprimento_comprovante.get()
        
        if imovel_id and item_id and quantidade and data and valor:
            try:
                quantidade = int(quantidade)
                valor = float(valor)
                
                conn = sqlite3.connect("sistema.db")
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO reposicao_suprimentos 
                    (imovel_id, suprimento_id, quantidade, data, valor_gasto, comprovante_path)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (imovel_id, item_id, quantidade, data, valor, comprovante))
                conn.commit()
                conn.close()
                
                messagebox.showinfo("Sucesso", "Reposição de suprimento registrada com sucesso!")
                self.carregar_suprimentos()
                self.limpar_form_suprimento()
            except ValueError:
                messagebox.showerror("Erro", "Quantidade deve ser inteiro e valor deve ser numérico")
        else:
            messagebox.showerror("Erro", "Preencha todos os campos obrigatórios")
    
    def limpar_form_suprimento(self):
        """Limpa o formulário de suprimentos"""
        self.entry_suprimento_quantidade.delete(0, END)
        self.entry_suprimento_valor.delete(0, END)
        self.entry_suprimento_comprovante.delete(0, END)
    
    # =============================================
    # MÓDULO RELATÓRIOS
    # =============================================
    
    def criar_relatorios(self):
        """Cria a interface para geração de relatórios"""
        self.frame_relatorios = Frame(self.frame_conteudo, bg=COR_FUNDO)
        
        # Frame para os botões de relatório
        frame_botoes = Frame(self.frame_relatorios, bg=COR_FUNDO)
        frame_botoes.pack(fill=X, padx=10, pady=20)
        
        # Botão para gerar relatório semanal
        btn_relatorio_semanal = Button(frame_botoes, 
                                     text="Gerar Relatório Semanal",
                                     command=self.gerar_relatorio_semanal_interface,
                                     bg=COR_DESTAQUE,
                                     fg="white",
                                     font=self.fonte_titulo,
                                     padx=20,
                                     pady=10)
        btn_relatorio_semanal.pack(fill=X)
        
        # Frame para visualização do relatório
        frame_visualizacao = Frame(self.frame_relatorios, bg=COR_FUNDO)
        frame_visualizacao.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        Label(frame_visualizacao, text="Último Relatório Gerado", bg=COR_FUNDO, 
              font=self.fonte_titulo).pack(anchor="w")
        
        self.texto_relatorio = Text(frame_visualizacao, wrap=WORD, bg=COR_CARD)
        self.texto_relatorio.pack(fill=BOTH, expand=True)
        
        # Botão para abrir relatório no Word
        btn_abrir_word = Button(frame_visualizacao,
                              text="Abrir no Word",
                              command=self.abrir_relatorio_word,
                              bg=COR_SECUNDARIA,
                              fg="white")
        btn_abrir_word.pack(side=RIGHT, padx=5, pady=5)
    
    def gerar_relatorio_semanal_interface(self):
        """Gera e exibe o relatório semanal na interface"""
        try:
            caminho = gerar_relatorio_semanal()
            
            # Exibir resumo na interface
            with open(caminho, 'rb') as f:
                doc = Document(f)
                texto = ""
                for para in doc.paragraphs:
                    texto += para.text + "\n"
                
                self.texto_relatorio.delete(1.0, END)
                self.texto_relatorio.insert(1.0, texto)
            
            messagebox.showinfo("Sucesso", f"Relatório gerado com sucesso!\n{caminho}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar relatório: {str(e)}")
    
    def abrir_relatorio_word(self):
        """Abre o último relatório gerado no Word"""
        try:
            # Buscar o arquivo mais recente na pasta de relatórios
            arquivos = [f for f in os.listdir("relatorios") if f.endswith(".docx")]
            if arquivos:
                arquivos.sort(reverse=True)
                caminho = os.path.join("relatorios", arquivos[0])
                webbrowser.open(caminho)
            else:
                messagebox.showwarning("Aviso", "Nenhum relatório encontrado na pasta 'relatorios'")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o relatório: {str(e)}")

# =============================================
# INICIALIZAÇÃO DO SISTEMA
# =============================================

if __name__ == "__main__":
    from datetime import timedelta  # Import necessário para o timedelta
    
    root = Tk()
    app = SistemaGestaoApp(root)
    root.mainloop()
