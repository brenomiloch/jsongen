import pandas as pd
import tkinter as tk
from tkinter import filedialog

def xlsx_to_dataframe(xlsx_file_path):
    df = pd.read_excel(xlsx_file_path)
    return df

def convert_and_save_json():
    # Solicitar ao usuário o arquivo XLSX
    xlsx_file_path = filedialog.askopenfilename(title='Selecione o arquivo XLSX', filetypes=[('XLSX files', '*.xlsx')])
    
    if not xlsx_file_path:
        print('Nenhum arquivo selecionado. Operação cancelada.')
        return

    # Ler o arquivo XLSX e converter para DataFrame
    xlsx_dataframe = xlsx_to_dataframe(xlsx_file_path)

    # Solicitar ao usuário o local para salvar o arquivo JSON
    json_file_path = filedialog.asksaveasfilename(title='Salve o arquivo JSON', defaultextension='.json', filetypes=[('JSON files', '*.json')])

    if not json_file_path:
        print('Nenhum local de salvamento selecionado. Operação cancelada.')
        return

    # Converter o DataFrame para JSON e salvar no arquivo
    xlsx_dataframe.to_json(json_file_path, orient='records')

    # Exibir mensagem de conclusão
    print(f'DataFrame convertido para JSON e salvo em: {json_file_path}')

# Criar janela principal
root = tk.Tk()
root.withdraw()  # Ocultar a janela principal

# Chamar a função para converter e salvar
convert_and_save_json()
