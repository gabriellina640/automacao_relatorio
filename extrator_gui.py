import sys
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
from docx.shared import Pt
# --- NOVAS IMPORTAÇÕES PARA ALINHAMENTO ---
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL

# --- 1. Lógica de Extração ---

def limpar_nome_sessao(texto_completo):
    """
    Retorna apenas o número e 'Virtual'. Ex: '8ª Virtual'
    """
    match_num = re.search(r"(\d+[ªº°])", texto_completo)
    numero = match_num.group(1) if match_num else ""
    tipo = "Virtual" if "Virtual" in texto_completo else ""
    resultado = f"{numero} {tipo}".strip()
    return resultado if resultado else "Sessão"

def encontrar_sessao_formatada(doc):
    for i in range(min(15, len(doc.paragraphs))):
        texto = doc.paragraphs[i].text.strip()
        if "Sessão" in texto and "pauta da" in texto.lower():
            return limpar_nome_sessao(texto)
    return "Sessão"

def extrair_dados(caminho_arquivo):
    doc = Document(caminho_arquivo)
    lista_itens = []
    item_atual = {}
    sessao_formatada = encontrar_sessao_formatada(doc)
    
    regex_processo = re.compile(r"Processo\s*n[º\.]?\s*([\d\.\-\/]+)", re.IGNORECASE)

    for para in doc.paragraphs:
        texto = para.text.strip()
        
        match = regex_processo.search(texto)
        if match:
            if item_atual:
                lista_itens.append(item_atual)
            item_atual = {
                "processo": match.group(1),
                "assunto": "",
                "conselheiro": "",
                "sessao": sessao_formatada
            }
            continue
        
        if (texto.startswith("Objeto:") or texto.startswith("Assunto:")) and item_atual:
            item_atual["assunto"] = texto.split(":", 1)[1].strip()
            
        elif texto.startswith("Relator:") and item_atual:
            conteudo_completo = texto.split(":", 1)[1].strip()
            item_atual["conselheiro"] = conteudo_completo

    if item_atual:
        lista_itens.append(item_atual)
    return lista_itens

# --- 2. Geração do Word (Com Alinhamento Centralizado) ---
def gerar_word(dados, caminho_saida):
    doc = Document()
    
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(8)
    
    # Remove espaçamentos globais
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.space_before = Pt(0)

    doc.add_heading('Relatório de Processos', 0)

    colunas = ["Nº Processo", "Assunto", "DISTRIBUIDO P/ CONSELHEIRO(A)", "Sessão", "Data da Assinatura", "Data da Publicação"]
    
    table = doc.add_table(rows=1, cols=len(colunas))
    table.style = 'Table Grid'
    table.autofit = False 

    # --- Configuração do Cabeçalho ---
    hdr_cells = table.rows[0].cells
    for i, col in enumerate(colunas):
        hdr_cells[i].text = col
        
        # Alinhamento Vertical (Centro)
        hdr_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        
        # Alinhamento Horizontal (Centro) + Fonte
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        
        run = p.runs[0]
        run.font.name = 'Times New Roman'
        run.font.size = Pt(8)
        run.font.bold = True

    # --- Preenchimento dos Dados ---
    for item in dados:
        row = table.add_row().cells
        
        valores = [
            item['processo'], 
            item['assunto'], 
            item['conselheiro'], 
            item['sessao'], 
            "", 
            ""
        ]
        
        for i, valor in enumerate(valores):
            cell = row[i]
            cell.text = valor
            
            # 1. Alinhamento Vertical da Célula
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            
            # 2. Formatação do Parágrafo (Horizontal + Fonte)
            for paragraph in cell.paragraphs:
                # Alinhamento Horizontal
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Remove espaços extras
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.line_spacing = 1.0 
                
                for run in paragraph.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(8)
                    
                    # Negrito apenas na coluna Processo (índice 0)
                    if i == 0:
                        run.font.bold = True

    doc.save(caminho_saida)

# --- 3. Interface Gráfica ---
def selecionar_arquivo():
    arquivo_origem = filedialog.askopenfilename(
        title="Selecione a Pauta (.docx)",
        filetypes=[("Arquivos Word", "*.docx")]
    )
    
    if not arquivo_origem:
        return

    try:
        lbl_status.config(text="Processando...", fg="blue")
        root.update()
        
        dados = extrair_dados(arquivo_origem)
        
        if not dados:
            messagebox.showwarning("Aviso", "Nenhum processo encontrado.")
            lbl_status.config(text="Aguardando...", fg="black")
            return

        arquivo_destino = filedialog.asksaveasfilename(
            title="Salvar Tabela Como",
            defaultextension=".docx",
            filetypes=[("Arquivos Word", "*.docx")],
            initialfile="Tabela_Formatada.docx"
        )

        if not arquivo_destino:
            lbl_status.config(text="Cancelado.", fg="black")
            return

        gerar_word(dados, arquivo_destino)
        
        lbl_status.config(text="Concluído!", fg="green")
        sessao_msg = dados[0]['sessao'] if dados else "N/A"
        messagebox.showinfo("Sucesso", f"Tabela gerada!\nAlinhamento: Centralizado\nSessão: {sessao_msg}\nTotal: {len(dados)} processos.")
        
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
        lbl_status.config(text="Erro.", fg="red")

root = tk.Tk()
root.title("Extrator MP - Centralizado")
root.geometry("400x200")

lbl_instrucao = tk.Label(root, text="Selecione a Pauta para extrair:", pady=20)
lbl_instrucao.pack()

btn_processar = tk.Button(root, text="Selecionar Arquivo", command=selecionar_arquivo, bg="#2b2b2b", fg="white", font=("Times New Roman", 12, "bold"), padx=10, pady=5)
btn_processar.pack()

lbl_status = tk.Label(root, text="Aguardando...", pady=20)
lbl_status.pack()

root.mainloop()