import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
import sys # Adicionado para lidar com erros de codificação em alguns sistemas

# Lógica de Cálculo
# Esta é uma lógica simplificada baseada em práticas comuns da NBR 5410.
# Não substitui um projeto elétrico profissional.
def calcular_bitolas(tipo_comodo):
    """
    Retorna as bitolas recomendadas com base no tipo de cômodo.
    NBR 5410 define bitolas mínimas: 1.5mm² para iluminação, 2.5mm² para TUGs.
    Circuitos específicos (TUEs) como chuveiros exigem dimensionamento próprio.
    """
    bitolas = {
        'iluminacao': '1.5 mm²',
        'tomadas': '2.5 mm²',
        'especifico': '-'
    }
    if tipo_comodo == 'Banheiro com Chuveiro Elétrico':
        # Chuveiros de 5500W a 7800W em 220V geralmente exigem 4mm² ou 6mm².
        # Usaremos 6.0mm² como uma recomendação segura.
        bitolas['especifico'] = '6.0 mm² (Chuveiro)'
    elif tipo_comodo in ['Cozinha', 'Área de Serviço']:
        # Cozinhas podem ter circuitos de tomadas mais robustos (TUEs).
        # A recomendação mínima para TUGs é 2.5mm², mas um circuito dedicado
        # para torneira elétrica ou forno pode exigir 4.0mm².
        # Para simplificar, mantemos o padrão de TUGs.
        bitolas['tomadas'] = '2.5 mm²'

    return bitolas

# Funções da Interface

def adicionar_comodo():
    """Adiciona um cômodo na tabela da interface."""
    nome = entry_nome.get()
    largura_str = entry_largura.get().replace(',', '.')
    comprimento_str = entry_comprimento.get().replace(',', '.')
    tipo = combo_tipo.get()

    if not all([nome, largura_str, comprimento_str, tipo]):
        messagebox.showerror("Erro", "Todos os campos devem ser preenchidos.")
        return

    try:
        largura = float(largura_str)
        comprimento = float(comprimento_str)
        if largura <= 0 or comprimento <= 0:
            raise ValueError
    except ValueError:
        messagebox.showerror("Erro de Entrada", "Largura e Comprimento devem ser números positivos.")
        return

    area = largura * comprimento
    bitolas = calcular_bitolas(tipo)

    # Insere os dados na tabela (Treeview)
    tree.insert('', 'end', values=(
        nome,
        f"{area:.2f} m²",
        tipo,
        bitolas['iluminacao'],
        bitolas['tomadas'],
        bitolas['especifico']
    ))

    # Limpa os campos de entrada
    entry_nome.delete(0, 'end')
    entry_largura.delete(0, 'end')
    entry_comprimento.delete(0, 'end')
    combo_tipo.set('')


def get_treeview_data():
    """Extrai os dados da tabela (Treeview) para um formato de lista de dicionários."""
    data = []
    columns = [tree.heading(col)['text'] for col in tree['columns']]
    for item_id in tree.get_children():
        values = tree.item(item_id, 'values')
        data.append(dict(zip(columns, values)))
    return data


def exportar_para_excel():
    """Exporta os dados da tabela para um arquivo Excel."""
    data = get_treeview_data()
    if not data:
        messagebox.showwarning("Aviso", "Não há dados para exportar.")
        return

    df = pd.DataFrame(data)
    
    try:
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")],
            title="Salvar como Excel"
        )
        if not filepath:
            return

        df.to_excel(filepath, index=False, engine='openpyxl')
        messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso em:\n{filepath}")
    except Exception as e:
        messagebox.showerror("Erro ao Salvar", f"Ocorreu um erro: {e}")


def exportar_para_pdf():
    """Exporta os dados da tabela para um arquivo PDF."""
    data = get_treeview_data()
    if not data:
        messagebox.showwarning("Aviso", "Não há dados para exportar.")
        return

    try:
        filepath = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")],
            title="Salvar como PDF"
        )
        if not filepath:
            return

        doc = SimpleDocTemplate(filepath, pagesize=landscape(letter))
        elements = []
        
        # Corrigido para lidar com fontes e codificação UTF-8
        try:
            styles = getSampleStyleSheet()
        except Exception:
            # Fallback para caso de problemas com fontes/codificação
            from reportlab.lib.styles import StyleSheet1
            styles = StyleSheet1()
            styles.add(ParagraphStyle(name='Normal', fontName='Helvetica', fontSize=10))
            styles.add(ParagraphStyle(name='h1', fontName='Helvetica-Bold', fontSize=14))

        # --- INÍCIO DA CORREÇÃO ---
        # A linha abaixo (title = ...) deve estar DENTRO do bloco try,
        # e estava faltando o argumento 'styles['h1']' e tinha uma vírgula a mais.

        # Título
        title_text = "Dimensionamento Elétrico Residencial (Estimativa)"
        title = Paragraph(title_text, styles['h1'])
        elements.append(title)
        elements.append(Spacer(1, 12))
        
        # Dados da Tabela
        headers = [tree.heading(col)['text'] for col in tree['columns']]
        table_data = [headers] + [list(item.values()) for item in data]
        
        # Criar tabela no ReportLab
        table = Table(table_data)
        
        # Estilo da Tabela
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ])
        table.setStyle(style)

        elements.append(table)
        elements.append(Spacer(1, 24))

        # Aviso
        disclaimer_text = (
            "<b>AVISO IMPORTANTE:</b> Esta é uma estimativa simplificada e não substitui um projeto elétrico "
            "completo e a análise de um profissional qualificado. Todos os serviços elétricos devem seguir "
            "rigorosamente a norma NBR 5410 e ser executados por um eletricista certificado."
        )
        disclaimer = Paragraph(disclaimer_text, styles['Normal'])
        elements.append(disclaimer)

        doc.build(elements)
        messagebox.showinfo("Sucesso", f"PDF salvo com sucesso em:\n{filepath}")

    except Exception as e:
        messagebox.showerror("Erro ao Salvar", f"Ocorreu um erro ao gerar o PDF: {e}")
        # --- FIM DA CORREÇÃO ---

def limpar_tabela():
    """Limpa todos os dados da tabela."""
    for item in tree.get_children():
        tree.delete(item)

# Configuração da Janela Principal
root = tk.Tk()
root.title("Calculadora de Dimensionamento Elétrico Residencial")
root.geometry("950x550")
root.resizable(True, True)

# Frame da Entrada de Dados
frame_entrada = ttk.LabelFrame(root, text="Dados do Cômodo")
frame_entrada.pack(padx=10, pady=10, fill='x', expand=False)

# Widgets de entrada
ttk.Label(frame_entrada, text="Nome do Cômodo:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
entry_nome = ttk.Entry(frame_entrada, width=20)
entry_nome.grid(row=0, column=1, padx=5, pady=5, sticky='w')

ttk.Label(frame_entrada, text="Largura (m):").grid(row=0, column=2, padx=5, pady=5, sticky='w')
entry_largura = ttk.Entry(frame_entrada, width=10)
entry_largura.grid(row=0, column=3, padx=5, pady=5, sticky='w')

ttk.Label(frame_entrada, text="Comprimento (m):").grid(row=0, column=4, padx=5, pady=5, sticky='w')
entry_comprimento = ttk.Entry(frame_entrada, width=10)
entry_comprimento.grid(row=0, column=5, padx=5, pady=5, sticky='w')

ttk.Label(frame_entrada, text="Tipo de Cômodo:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
tipos_comodo = [
    'Quarto', 'Sala', 'Cozinha', 'Banheiro', 'Banheiro com Chuveiro Elétrico',
    'Área de Serviço', 'Corredor', 'Garagem', 'Área Externa'
]
combo_tipo = ttk.Combobox(frame_entrada, values=tipos_comodo, width=25)
combo_tipo.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky='w')

# Botão de adicionar
btn_adicionar = ttk.Button(frame_entrada, text="Adicionar Cômodo", command=adicionar_comodo)
btn_adicionar.grid(row=1, column=3, columnspan=2, padx=5, pady=10)

# Frame da Tabela de Resultados
frame_tabela = ttk.Frame(root)
frame_tabela.pack(padx=10, pady=5, fill='both', expand=True)

# Definição das colunas
columns = ('nome', 'area', 'tipo', 'bitola_luz', 'bitola_tomada', 'bitola_especifico')
tree = ttk.Treeview(frame_tabela, columns=columns, show='headings')

# Cabeçalhos
tree.heading('nome', text='Cômodo')
tree.heading('area', text='Área (m²)')
tree.heading('tipo', text='Tipo')
tree.heading('bitola_luz', text='Bitola Iluminação')
tree.heading('bitola_tomada', text='Bitola Tomadas (TUG)')
tree.heading('bitola_especifico', text='Bitola Específico (TUE)')

# Largura das colunas
tree.column('nome', width=120)
tree.column('area', width=80, anchor='center')
tree.column('tipo', width=150)
tree.column('bitola_luz', width=120, anchor='center')
tree.column('bitola_tomada', width=130, anchor='center')
tree.column('bitola_especifico', width=150, anchor='center')

tree.pack(side='left', fill='both', expand=True)

# Scrollbar
scrollbar = ttk.Scrollbar(frame_tabela, orient='vertical', command=tree.yview)
tree.configure(yscrollcommand=scrollbar.set)
scrollbar.pack(side='right', fill='y')


# Frame de Ações (EXPORTAR/LIMPAR)
frame_acoes = ttk.Frame(root)
frame_acoes.pack(padx=10, pady=10, fill='x', expand=False)

btn_excel = ttk.Button(frame_acoes, text="Exportar para Excel", command=exportar_para_excel)
btn_excel.pack(side='left', padx=5)

btn_pdf = ttk.Button(frame_acoes, text="Exportar para PDF", command=exportar_para_pdf)
btn_pdf.pack(side='left', padx=5)

btn_limpar = ttk.Button(frame_acoes, text="Limpar Tabela", command=limpar_tabela)
btn_limpar.pack(side='right', padx=5)

# Aviso da Interface
label_aviso = ttk.Label(root, text="⚠️ Lembre-se: Este programa é apenas para estimativas. Consulte sempre um eletricista qualificado.", foreground="red")
label_aviso.pack(pady=(0, 10))

root.mainloop()