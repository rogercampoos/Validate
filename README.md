# Validação de Patrimônios

Este programa compara números de patrimônio entre arquivos PDF e Excel, identificando quais números estão presentes em ambos os arquivos, apenas no PDF ou apenas no Excel.

## Requisitos

- Python 3.6 ou superior
- Bibliotecas: pandas, PyPDF2, tkinter (geralmente vem com Python)

## Instalação

Instale as dependências necessárias:

```
pip install -r requirements.txt
```

## Como usar

1. Execute o script:

```
python validacao_patrimonios.py
```

2. Selecione o tipo de equipamento (Computador ou Monitor)
3. Selecione o arquivo Excel quando solicitado
4. Selecione o arquivo PDF quando solicitado
5. Escolha onde salvar o arquivo de resultados

## Formato dos arquivos

- **Excel**: O programa usa a segunda coluna do arquivo Excel para extrair números de patrimônio de 9 dígitos.
- **PDF**: O programa extrai todos os números de 9 dígitos encontrados no texto do PDF.

## Resultados

O programa gera:
- Um resumo na tela mostrando quantos patrimônios foram encontrados em cada fonte
- Um arquivo Excel com quatro abas:
  - Resumo: contagem de patrimônios por categoria
  - Patrimônios em Ambos: números encontrados tanto no PDF quanto no Excel
  - Apenas no PDF: números encontrados somente no PDF
  - Apenas no Excel: números encontrados somente no Excel
