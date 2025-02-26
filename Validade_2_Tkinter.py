import pandas as pd
import PyPDF2
import re
import tkinter as tk
from tkinter import simpledialog
import os

# Inicializa o tkinter para a caixa de diálogo
root = tk.Tk()
root.withdraw()  # Esconde a janela principal

# Caminho do arquivo Excel (fixo)
caminho_excel = 'C:/Users/p0134255/Desktop/Validação/Planilha_Definitiva___Onda_5_Monitor.ods'

# Solicita o número ao usuário
numero = simpledialog.askstring("Entrada", "Digite o número do arquivo:")

# Verifica se o usuário digitou algo
if numero:
    # Construindo o caminho do arquivo PDF com o número inserido
    caminho_pdf = f'C:/Users/p0134255/Desktop/Validação/{numero}.pdf'
    print(f"Caminho do PDF configurado: {caminho_pdf}")
else:
    print("Nenhum número foi digitado. Usando o caminho padrão.")
    # Caminho padrão caso o usuário não digite nada
    caminho_pdf = 'C:/Users/p0134255/Desktop/Validação/939861.pdf'

# Função para extrair números de patrimônio do PDF
def extrair_patrimonios_pdf(caminho_pdf):
    patrimonios = set()
    with open(caminho_pdf, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            text = page.extract_text()
            encontrados = re.findall(r'\b(\d{9})\b', text)  # Captura números de 9 dígitos
            patrimonios.update(encontrados)
    return patrimonios

# Função para extrair números de patrimônio do Excel usando a segunda coluna
def extrair_patrimonios_excel(caminho_excel):
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

# Exibir resultados
print(f"\nRESULTADOS DA VALIDAÇÃO:")
print(f"Patrimônios em ambas as fontes: {len(patrimonios_em_ambos)}")
print(f"Patrimônios apenas no PDF: {len(patrimonios_somente_no_pdf)}")
print(f"Patrimônios apenas no Excel: {len(patrimonios_somente_no_excel)}")

if patrimonios_somente_no_pdf:
    print("\nExemplos de patrimônios que estão apenas no PDF:")
    for p in list(patrimonios_somente_no_pdf)[:5]:  # Mostra até 5 exemplos
        print(f"- {p}")

if patrimonios_somente_no_excel:
    print("\nExemplos de patrimônios que estão apenas no Excel:")
    for p in list(patrimonios_somente_no_excel)[:5]:  # Mostra até 5 exemplos
        print(f"- {p}")

# Criar DataFrames separados para cada lista
df_ambos = pd.DataFrame({'Patrimônios em Ambos': list(patrimonios_em_ambos)}) if patrimonios_em_ambos else pd.DataFrame({'Patrimônios em Ambos': ['N/A']})
df_pdf = pd.DataFrame({'Patrimônios Apenas no PDF': list(patrimonios_somente_no_pdf)}) if patrimonios_somente_no_pdf else pd.DataFrame({'Patrimônios Apenas no PDF': ['N/A']})
df_excel = pd.DataFrame({'Patrimônios Apenas no Excel': list(patrimonios_somente_no_excel)}) if patrimonios_somente_no_excel else pd.DataFrame({'Patrimônios Apenas no Excel': ['N/A']})

# Salvar resultados para análise em diferentes planilhas
caminho_saida = "C:/Users/p0134255/Desktop/Validação/validacao_patrimonios.xlsx"
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

print(f"\nResultados salvos em: {caminho_saida}")
