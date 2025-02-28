import pandas as pd
import PyPDF2
import re
import tkinter as tk
from tkinter import simpledialog, messagebox, Toplevel, Label, Text, Scrollbar, Button, Frame, StringVar, OptionMenu, filedialog
import os

# Função principal para executar a validação
def executar_validacao():
    # Inicializa o tkinter para as caixas de diálogo
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    
    # Criar janela para selecionar o tipo de equipamento
    dialog_window = Toplevel(root)
    dialog_window.title("Seleção de Equipamento")
    dialog_window.geometry("300x150")
    dialog_window.resizable(False, False)
    
    tipo_equipamento = StringVar(dialog_window)
    tipo_equipamento.set("Selecione")  # valor inicial
    
    Label(dialog_window, text="Selecione o tipo de equipamento:").pack(pady=10)
    
    option_menu = OptionMenu(dialog_window, tipo_equipamento, "Computador", "Monitor")
    option_menu.pack(pady=10)
    
    resultado = {"selecionado": False, "tipo": ""}
    
    def confirmar():
        if tipo_equipamento.get() != "Selecione":
            resultado["selecionado"] = True
            resultado["tipo"] = tipo_equipamento.get()
            dialog_window.destroy()
    
    Button(dialog_window, text="Confirmar", command=confirmar).pack(pady=10)
    
    dialog_window.transient(root)
    dialog_window.grab_set()
    root.wait_window(dialog_window)
    
    if not resultado["selecionado"]:
        messagebox.showinfo("Cancelado", "Seleção cancelada.")
        root.destroy()
        return
    
    # Baseado na seleção, solicitar o arquivo Excel
    tipo_selecionado = resultado["tipo"]
    messagebox.showinfo("Selecionar Excel", f"Por favor, selecione o arquivo Excel para {tipo_selecionado}")
    caminho_excel = filedialog.askopenfilename(
        title=f"Selecione o arquivo Excel para {tipo_selecionado}",
        filetypes=[("Arquivos Excel", "*.xlsx;*.xls")]
    )
    
    if not caminho_excel:
        messagebox.showinfo("Cancelado", "Nenhum arquivo Excel selecionado.")
        root.destroy()
        return
    
    # Solicitar o arquivo PDF
    messagebox.showinfo("Selecionar PDF", f"Por favor, selecione o arquivo PDF para {tipo_selecionado}")
    caminho_pdf = filedialog.askopenfilename(
        title=f"Selecione o arquivo PDF para {tipo_selecionado}",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    
    if not caminho_pdf:
        messagebox.showinfo("Cancelado", "Nenhum arquivo PDF selecionado.")
        root.destroy()
        return
    
    # Extrair o número do arquivo do nome do PDF
    numero = os.path.basename(caminho_pdf).split('.')[0]
    
    print(f"Tipo de equipamento: {tipo_selecionado}")
    print(f"Caminho do Excel: {caminho_excel}")
    print(f"Caminho do PDF: {caminho_pdf}")
    
    # Função para extrair números de patrimônio do PDF
    def extrair_patrimonios_pdf(caminho_pdf):
        patrimonios = set()
        try:
            with open(caminho_pdf, "rb") as file:
                reader = PyPDF2.PdfReader(file)
                for page in reader.pages:
                    text = page.extract_text()
                    encontrados = re.findall(r'\b(\d{9})\b', text)  # Captura números de 9 dígitos
                    patrimonios.update(encontrados)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir o arquivo PDF: {e}")
        return patrimonios
    
    # Função para extrair números de patrimônio do Excel usando a segunda coluna
    def extrair_patrimonios_excel(caminho_excel):
        patrimonios = set()
        try:
            df = pd.read_excel(caminho_excel)
            
            # Verifica se existem pelo menos 2 colunas
            if len(df.columns) < 2:
                raise ValueError("O arquivo Excel não possui pelo menos duas colunas.")
            
            # Usa a segunda coluna (índice 1) como coluna de patrimônios
            coluna_patrimonio = df.columns[1]
            print(f"Usando a coluna '{coluna_patrimonio}' como coluna de patrimônios")
            
            # Converte para string e remove NaN
            patrimonios = set(df[coluna_patrimonio].dropna().astype(str))
            
            # Remove quaisquer valores que não sejam numéricos de 9 dígitos
            patrimonios = {p for p in patrimonios if re.match(r'^\d{9}$', p)}
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar o Excel: {e}")
        return patrimonios
    
    # Carregar dados das fontes
    try:
        patrimonios_pdf = extrair_patrimonios_pdf(caminho_pdf)
        print(f"Encontrados {len(patrimonios_pdf)} patrimônios no PDF")
    except Exception as e:
        print(f"Erro ao processar o PDF: {e}")
        patrimonios_pdf = set()
    
    try:
        patrimonios_excel = extrair_patrimonios_excel(caminho_excel)
        print(f"Encontrados {len(patrimonios_excel)} patrimônios no Excel")
    except Exception as e:
        print(f"Erro ao processar o Excel: {e}")
        patrimonios_excel = set()
    
    # Validar correspondência
    patrimonios_em_ambos = patrimonios_pdf & patrimonios_excel
    patrimonios_somente_no_pdf = patrimonios_pdf - patrimonios_excel
    patrimonios_somente_no_excel = patrimonios_excel - patrimonios_pdf
    
    # Preparar o resultado para exibição
    resumo = f"RESULTADOS DA VALIDAÇÃO:\n\n"
    resumo += f"Tipo de equipamento: {tipo_selecionado}\n"
    resumo += f"Arquivo PDF: {os.path.basename(caminho_pdf)}\n\n"
    resumo += f"Patrimônios em ambas as fontes: {len(patrimonios_em_ambos)}\n"
    resumo += f"Patrimônios apenas no PDF: {len(patrimonios_somente_no_pdf)}\n"
    resumo += f"Patrimônios apenas no Excel: {len(patrimonios_somente_no_excel)}\n\n"
    
    if patrimonios_somente_no_pdf:
        resumo += "Exemplos de patrimônios que estão apenas no PDF:\n"
        for p in list(patrimonios_somente_no_pdf)[:5]:  # Mostra até 5 exemplos
            resumo += f"- {p}\n"
        resumo += "\n"
    
    if patrimonios_somente_no_excel:
        resumo += "Exemplos de patrimônios que estão apenas no Excel:\n"
        for p in list(patrimonios_somente_no_excel)[:5]:  # Mostra até 5 exemplos
            resumo += f"- {p}\n"
        resumo += "\n"
    
    # Criar DataFrames separados para cada lista
    df_ambos = pd.DataFrame({'Patrimônios em Ambos': list(patrimonios_em_ambos)}) if patrimonios_em_ambos else pd.DataFrame({'Patrimônios em Ambos': ['N/A']})
    df_pdf = pd.DataFrame({'Patrimônios Apenas no PDF': list(patrimonios_somente_no_pdf)}) if patrimonios_somente_no_pdf else pd.DataFrame({'Patrimônios Apenas no PDF': ['N/A']})
    df_excel = pd.DataFrame({'Patrimônios Apenas no Excel': list(patrimonios_somente_no_excel)}) if patrimonios_somente_no_excel else pd.DataFrame({'Patrimônios Apenas no Excel': ['N/A']})
    
    # Solicitar local para salvar o resultado
    messagebox.showinfo("Salvar Resultado", "Selecione onde deseja salvar o arquivo de resultados")
    caminho_saida = filedialog.asksaveasfilename(
        title="Salvar arquivo de resultados",
        defaultextension=".xlsx",
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )
    
    if not caminho_saida:
        # Se o usuário cancelar, ainda mostra os resultados mas não salva o arquivo
        messagebox.showinfo("Aviso", "O arquivo de resultados não será salvo.")
    else:
        # Salvar resultados para análise em diferentes planilhas
        with pd.ExcelWriter(caminho_saida) as writer:
            df_ambos.to_excel(writer, sheet_name='Patrimônios em Ambos', index=False)
            df_pdf.to_excel(writer, sheet_name='Apenas no PDF', index=False)
            df_excel.to_excel(writer, sheet_name='Apenas no Excel', index=False)
            
            # Criar uma aba de resumo
            df_resumo = pd.DataFrame({
                'Categoria': ['Patrimônios em Ambos', 'Patrimônios Apenas no PDF', 'Patrimônios Apenas no Excel'],
                'Quantidade': [len(patrimonios_em_ambos), len(patrimonios_somente_no_pdf), len(patrimonios_somente_no_excel)]
            })
            df_resumo.to_excel(writer, sheet_name='Resumo', index=False)
        
        resumo += f"Resultados detalhados salvos em: {caminho_saida}\n"
    
    # Criar janela de resultados
    result_window = Toplevel(root)
    result_window.title("Resultados da Validação")
    result_window.geometry("600x500")
    
    # Frame para conter o texto e a scrollbar
    text_frame = Frame(result_window)
    text_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    scrollbar = Scrollbar(text_frame)
    scrollbar.pack(side="right", fill="y")
    
    result_text = Text(text_frame, yscrollcommand=scrollbar.set, wrap="word")
    result_text.insert("1.0", resumo)
    result_text.config(state="disabled")  # Torna o texto somente leitura
    result_text.pack(side="left", fill="both", expand=True)
    
    scrollbar.config(command=result_text.yview)
    
    def fechar():
        result_window.destroy()
        root.destroy()
    
    Button(result_window, text="Fechar", command=fechar).pack(pady=10)
    
    # Manter a janela de resultados aberta até o usuário fechar
    result_window.transient(root)
    result_window.grab_set()
    root.wait_window(result_window)

# Iniciar o programa
if __name__ == "__main__":
    executar_validacao()