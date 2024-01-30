import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Conversor XLSX para JSON")

        # Botão para importar arquivo XLSX
        self.import_button = tk.Button(root, text="Importar XLSX", command=self.import_xlsx)
        self.import_button.pack(pady=20)

        # Botão para converter e salvar como JSON
        self.convert_button = tk.Button(root, text="Converter e Salvar JSON", command=self.convert_and_save_json)
        self.convert_button.pack(pady=20)

    def import_xlsx(self):
        xlsx_file_path = filedialog.askopenfilename(title='Selecione o arquivo XLSX', filetypes=[('XLSX files', '*.xlsx')])
        if xlsx_file_path:
            messagebox.showinfo("Sucesso", f"Arquivo XLSX importado:\n{xlsx_file_path}")

    def convert_and_save_json(self):
        xlsx_file_path = filedialog.askopenfilename(title='Selecione o arquivo XLSX', filetypes=[('XLSX files', '*.xlsx')])
        if not xlsx_file_path:
            messagebox.showwarning("Aviso", "Nenhum arquivo XLSX selecionado. Operação cancelada.")
            return

        xlsx_dataframe = pd.read_excel(xlsx_file_path)

        json_file_path = filedialog.asksaveasfilename(title='Salve o arquivo JSON', defaultextension='.json', filetypes=[('JSON files', '*.json')])
        if not json_file_path:
            messagebox.showwarning("Aviso", "Nenhum local de salvamento selecionado. Operação cancelada.")
            return

        xlsx_dataframe.to_json(json_file_path, orient='records')
        messagebox.showinfo("Sucesso", f'DataFrame convertido para JSON e salvo em:\n{json_file_path}')

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
