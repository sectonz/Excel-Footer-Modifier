import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client
import openpyxl
import threading

def modify_excel_footer(file_path, left_footer_text, center_footer_text, right_footer_text, progress_bar):
    try:
        file_extension = os.path.splitext(file_path)[1].lower()
        
        if file_extension == '.xls':
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False  

            workbook = excel.Workbooks.Open(file_path)

            for sheet in workbook.Sheets:
                sheet.PageSetup.LeftFooter = left_footer_text
                sheet.PageSetup.CenterFooter = center_footer_text
                sheet.PageSetup.RightFooter = right_footer_text

                progress_bar['value'] += 1
            
            workbook.Save()
            workbook.Close()
            excel.Quit()

        elif file_extension == '.xlsx':
            workbook = openpyxl.load_workbook(file_path)

            for sheet in workbook.worksheets:
                sheet.oddFooter.left.text = left_footer_text
                sheet.oddFooter.center.text = center_footer_text
                sheet.oddFooter.right.text = right_footer_text
                sheet.evenFooter.left.text = left_footer_text
                sheet.evenFooter.center.text = center_footer_text
                sheet.evenFooter.right.text = right_footer_text

                progress_bar['value'] += 1

            workbook.save(file_path)

        return True
        
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao modificar o arquivo: {e}")
        return False

def center_window(window):
    """Centraliza a janela na tela."""
    window.update_idletasks()  
    width = window.winfo_width()
    height = window.winfo_height()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    window.geometry(f"{width}x{height}+{x}+{y}")

def process_files(left_text, center_text, right_text):
    progress_window = tk.Toplevel(root)
    progress_window.title("Alterações em Progresso")
    progress_window.geometry("300x100")
    center_window(progress_window)

    progress_window.configure(bg="#fafafa")
    
    style = ttk.Style()
    style.configure("Progress.TLabel", background="#fafafa", foreground="black")

    ttk.Label(progress_window, text="Alterações em progresso...", style="Progress.TLabel").pack(pady=10)    
    progress_bar = ttk.Progressbar(progress_window, orient='horizontal', length=250, mode='determinate')
    progress_bar.pack(pady=10)
    
    progress_bar['value'] = 0  
    progress_bar['maximum'] = len(selected_files) 
    
    successes = 0
    for file_path in selected_files:
        if modify_excel_footer(file_path, left_text, center_text, right_text, progress_bar):
            successes += 1
    
    messagebox.showinfo("Finalizado.", f"{successes} rodapé(s) atualizado(s) com sucesso!")
    
    selected_files.clear()
    update_files_label()
    
    progress_window.destroy()  

def on_submit():
    left_text = left_footer_entry.get().strip()
    center_text = center_footer_entry.get().strip()
    right_text = right_footer_entry.get().strip()
    
    if not (left_text or center_text or right_text):
        messagebox.showinfo("Aviso", "Texto de seção não inserido.")
        return
    
    threading.Thread(target=process_files, args=(left_text, center_text, right_text)).start()

def select_files():
    file_paths = filedialog.askopenfilenames(
        title="Selecione o(s) arquivo(s) Excel",
        filetypes=[("Arquivos Excel", "*.xls;*.xlsx")]
    )
    if file_paths:
        selected_files.clear()  
        selected_files.extend(file_paths)
        update_files_label()

def update_files_label():
    if not selected_files:
        files_label.config(text="Nenhum arquivo selecionado")
    else:
        file_names = [os.path.basename(f) for f in selected_files]
        if len(file_names) > 7:
            display_text = ", ".join(file_names[:7]) + f" + {len(file_names) - 7} outros"
        else:
            display_text = ", ".join(file_names)
        files_label.config(text=display_text)

root = tk.Tk()
root.title("Modificador de Rodapé - by André Luiz Correia Filho")
root.configure(bg='#fafafa')
root.geometry("450x500")
root.eval('tk::PlaceWindow . center')
center_window(root)


style = ttk.Style()
style.theme_use('clam')
style.configure('TLabel', background='#fafafa', font=('Arial', 10))
style.configure('TButton', background='#fafafa', foreground='black', font=('Arial', 10))
style.configure('TEntry', font=('Arial', 10))
style.map('TButton', background=[('active', '#f0f0f0')])
style.configure('TFrame', background='#fafafa')

selected_files = []

main_frame = ttk.Frame(root, padding="20 20 20 20")
main_frame.pack(fill=tk.BOTH, expand=True)
main_frame.columnconfigure(0, weight=1)

select_button = ttk.Button(main_frame, text="Selecionar Arquivos", command=select_files, width=25)
select_button.grid(row=0, column=0, pady=10, sticky='ew')

files_label = ttk.Label(main_frame, text="Nenhum arquivo selecionado", wraplength=400)
files_label.grid(row=1, column=0, pady=10, sticky='ew')

ttk.Label(main_frame, text="Seção da Esquerda:").grid(row=2, column=0, sticky='w', padx=20, pady=(10, 0))
left_footer_entry = ttk.Entry(main_frame, width=50)
left_footer_entry.grid(row=3, column=0, pady=5,sticky='w')

ttk.Label(main_frame, text="Seção do Centro:").grid(row=4, column=0, sticky='w', padx=20, pady=(10, 0))
center_footer_entry = ttk.Entry(main_frame, width=50)
center_footer_entry.grid(row=5, column=0, pady=5,sticky='w')

ttk.Label(main_frame, text="Seção da Direita:").grid(row=6, column=0, sticky='w', padx=20, pady=(10, 0))
right_footer_entry = ttk.Entry(main_frame, width=50)
right_footer_entry.grid(row=7, column=0, pady=5,sticky='w')

bottom_frame = ttk.Frame(root)
bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=20, pady=20)

submit_button = ttk.Button(bottom_frame, text="Confirmar Alterações", command=on_submit, width=25)
submit_button.pack(anchor='s')

root.mainloop()
