#Bibliotecas 

import tkinter as tk
from tkinter import messagebox, ttk
import pyodbc
import pandas as pd
import os

#________________________________________________________________________________________________________________________________________
# Váriaveis do banco de dados:

server = '000.000.0.000'
database = 'Nome_do_banco'
username = 'Nome_do_usuario_do_Banco'
password = 'Senha_do_usuario'
conn_string = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'

#________________________________________________________________________________________________________________________________________
#classe Principal:

class RegistroApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Registro de Equipamentos")
        self.root.geometry("800x600")
        self.root.configure(bg="#f0f0f0")  # Cor de fundo da janela
        self.registros = []

        # Estilos
        self.bg_color = "#f0f0f0"  # Cinza muito claro
        self.frame_bg = "#ffffff"  # Branco
        self.label_color = "#333333"  # Cinza escuro
        self.entry_bg = "#e0e0e0"  # Cinza claro
        self.button_bg = "#007bff"  # Azul
        self.button_fg = "#ffffff"  # Branco
        self.alt_button_bg = "#28a745"  # Verde
        self.alt_button_fg = "#ffffff"  # Branco

        # Cabeçalho
        self.header_frame = tk.Frame(root, bg="#004080", pady=10)
        self.header_frame.pack(fill=tk.X)
        tk.Label(self.header_frame, text="Registro de Equipamentos", bg="#004080", fg="#ffffff", font=('Arial', 16, 'bold')).pack()

        # Painel de Formulário
        self.form_frame = tk.Frame(root, bg=self.frame_bg, padx=20, pady=20)
        self.form_frame.pack(fill=tk.X, padx=20, pady=(10, 0))

        tk.Label(self.form_frame, text="Nome:", bg=self.frame_bg, fg=self.label_color, font=('Arial', 12, 'bold')).grid(row=0, column=0, sticky="w", pady=5)
        self.nome_entry = tk.Entry(self.form_frame, width=50, bg=self.entry_bg, font=('Arial', 10))
        self.nome_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(self.form_frame, text="Equipamento:", bg=self.frame_bg, fg=self.label_color, font=('Arial', 12, 'bold')).grid(row=1, column=0, sticky="w", pady=5)
        self.equipamento_entry = tk.Entry(self.form_frame, width=50, bg=self.entry_bg, font=('Arial', 10))
        self.equipamento_entry.grid(row=1, column=1, padx=5, pady=5)

        # Container para campos adicionais de equipamento
        self.extra_equipamentos_frame = tk.Frame(self.form_frame, bg=self.frame_bg)
        self.extra_equipamentos_frame.grid(row=2, column=1, columnspan=1, pady=10)

        # Botões para adicionar/remover campos de equipamento
        self.button_frame = tk.Frame(self.form_frame, bg=self.frame_bg)
        self.button_frame.grid(row=3, column=1, columnspan=2, pady=5)

        # Botão para exportar para Excel
        self.export_button = tk.Button(self.button_frame, text="Exportar para Excel", command=self.exportar_para_excel, bg=self.button_bg, fg=self.alt_button_fg, font=('Arial', 12, 'bold'))
        self.export_button.pack(side=tk.LEFT, padx=5)

        # Botão para salvar registro
        self.save_button = tk.Button(self.button_frame, text="Salvar Registro", command=self.salvar_registro, bg=self.button_bg, fg=self.button_fg, font=('Arial', 12, 'bold'))
        self.save_button.pack(side=tk.LEFT, padx=5)

        # Botão para adicionar registros
        self.add_button = tk.Button(self.button_frame, text="+", width=2, height=1, command=self.adicionar_campo_equipamento, bg=self.alt_button_bg, fg=self.alt_button_fg, font=('Arial', 14, 'bold'))
        self.add_button.pack(side=tk.LEFT, padx=5)

        # Botão para remover registro
        self.remove_button = tk.Button(self.button_frame, text="-", width=2, height=1, command=self.remover_campo_equipamento, bg=self.alt_button_bg, fg=self.alt_button_fg, font=('Arial', 14, 'bold'))
        self.remove_button.pack(side=tk.LEFT, padx=5)

        # Painel de Exibição dos registros
        self.registros_frame = tk.Frame(root, bg=self.frame_bg)
        self.registros_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(10, 20))

        tk.Label(self.registros_frame, text="Registros Salvos:", bg=self.frame_bg, fg=self.label_color, font=('Arial', 14, 'bold')).pack(anchor="w", padx=10)

        # Barra de rolagem
        self.registros_canvas = tk.Canvas(self.registros_frame, bg=self.bg_color, highlightthickness=0)
        self.registros_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scrollbar = ttk.Scrollbar(self.registros_frame, orient="vertical", command=self.registros_canvas.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.registros_canvas.configure(yscrollcommand=self.scrollbar.set)

        self.registros_frame_inner = tk.Frame(self.registros_canvas, bg=self.bg_color)
        self.registros_canvas.create_window((0, 0), window=self.registros_frame_inner, anchor="nw")

        self.registros_frame.bind("<Configure>", self.on_frame_configure)
        self.registros_canvas.bind_all("<MouseWheel>", self.on_mouse_wheel)

        self.carregar_registros()
        self.contador = 0  
        self.extra_entries = [] 

#________________________________________________________________________________________________________________________________________
#Funções:

    def adicionar_campo_equipamento(self):
       
        row = self.contador * 2
        self.contador += 1

        new_entry = tk.Entry(self.extra_equipamentos_frame, width=50, bg=self.entry_bg, font=('Arial', 10))
        new_entry.grid(row=row, column=1, padx=5, pady=5)

        self.extra_entries.append(new_entry)

    def remover_campo_equipamento(self):
        if self.extra_entries:
        
            entry = self.extra_entries.pop()
            entry.destroy()

            self.contador -= 1

    def salvar_registro(self):
        nome = self.nome_entry.get()
        equipamentos = [self.equipamento_entry.get()]
        equipamentos += [entry.get() for entry in self.extra_entries if entry.get().strip()]

        if not nome or not equipamentos:
            messagebox.showwarning("Erro", "Nome e Equipamento(s) são obrigatórios!")
            return

        try:
            conn = pyodbc.connect(conn_string)
            cursor = conn.cursor()

            insert_query = """
            INSERT INTO registros (nome_equipamento, retidara, devolucao, statu, nome_pessoa)
            VALUES (?, GETDATE(), NULL, 'Pendente', ?)
            """
            for equipamento in equipamentos:
                cursor.execute(insert_query, equipamento, nome)

            conn.commit()
            messagebox.showinfo("Sucesso", "Registro salvo com sucesso!")
            self.carregar_registros()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar registro: {e}")
        finally:
            cursor.close()
            conn.close()

    def excluir_registro(self, id):
        confirm = messagebox.askyesno("Confirmar Exclusão", "Tem certeza de que deseja excluir este registro?")
        if confirm:
            try:
                conn = pyodbc.connect(conn_string)
                cursor = conn.cursor()

                delete_query = "DELETE FROM registros WHERE id = ?"
                cursor.execute(delete_query, id)
                conn.commit()
                messagebox.showinfo("Sucesso", "Registro excluído com sucesso!")
                self.carregar_registros()
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao excluir registro: {e}")
            finally:
                cursor.close()
                conn.close()

    def carregar_registros(self):

        for widget in self.registros_frame_inner.winfo_children():
            widget.destroy() 

        try:
            conn = pyodbc.connect(conn_string)
            cursor = conn.cursor()

            query = """
            SELECT id, nome_equipamento, retidara, devolucao, statu, nome_pessoa 
            FROM registros 
            WHERE retidara >= DATEADD(DAY, -14, GETDATE());
            """
            cursor.execute(query)
            registros = cursor.fetchall()

            for registro in registros:
                frame_registro = tk.Frame(self.registros_frame_inner, padx=10, pady=10, bg=self.bg_color)
                frame_registro.pack(fill=tk.X, pady=5)

                equipamento_str = f"Equipamento(s): {registro.nome_equipamento}"
                retirada_str = f"Retirada: {registro.retidara.strftime('%d/%m/%Y %H:%M:%S')}"
                devolucao_str = f"Devolução: {registro.devolucao.strftime('%d/%m/%Y %H:%M:%S') if registro.devolucao else 'Pendente'}"
                status_str = f"Status: {registro.statu}"
                nome_pessoa_str = f"Nome: {registro.nome_pessoa}"
                
                tk.Label(frame_registro, text=f"Registro:", bg=self.bg_color, fg=self.label_color, font=('Arial', 10, 'bold')).pack(anchor="w")
                tk.Label(frame_registro, text=nome_pessoa_str, bg=self.bg_color, fg=self.label_color, font=('Arial', 10)).pack(anchor="w")
                tk.Label(frame_registro, text=equipamento_str, bg=self.bg_color, fg=self.label_color, font=('Arial', 10)).pack(anchor="w")
                tk.Label(frame_registro, text=retirada_str, bg=self.bg_color, fg=self.label_color, font=('Arial', 10)).pack(anchor="w")
                tk.Label(frame_registro, text=devolucao_str, bg=self.bg_color, fg=self.label_color, font=('Arial', 10)).pack(anchor="w")
                tk.Label(frame_registro, text=status_str, bg=self.bg_color, fg=self.label_color, font=('Arial', 10)).pack(anchor="w")

                if registro.statu == "Devolvido":
                    button_bg = "#28a745" 
                    button_text = "Devolvido"
                else:
                    button_bg = "#dc3545" 
                    button_text = "Pendente"

                toggle_button = tk.Button(frame_registro, text=button_text, command=lambda id=registro.id: self.alternar_status(id), bg=button_bg, fg=self.alt_button_fg, font=('Arial', 10, 'bold'))
                toggle_button.pack(side=tk.RIGHT, padx=10)

                delete_button = tk.Button(frame_registro, text="Excluir", command=lambda id=registro.id: self.excluir_registro(id), bg=self.button_bg, fg="#ffffff", font=('Arial', 10, 'bold'))
                delete_button.pack(side=tk.RIGHT, padx=10)

            self.registros_canvas.update_idletasks()
            self.registros_canvas.config(scrollregion=self.registros_canvas.bbox("all"))

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar registros: {e}")
        finally:
            cursor.close()
            conn.close()

    def alternar_status(self, id):
        try:
            conn = pyodbc.connect(conn_string)
            cursor = conn.cursor()

            cursor.execute("SELECT statu FROM registros WHERE id = ?", id)
            status_atual = cursor.fetchone()[0]

            if status_atual == "Devolvido":
                update_query = """
                UPDATE registros
                SET devolucao = NULL, statu = 'Pendente'
                WHERE id = ?
                """
            else:
                update_query = """
                UPDATE registros
                SET devolucao = GETDATE(), statu = 'Devolvido'
                WHERE id = ?
                """

            cursor.execute(update_query, id)
            conn.commit()
            messagebox.showinfo("Sucesso", f"Status alterado com sucesso!")
            self.carregar_registros()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao alterar status: {e}")
        finally:
            cursor.close()
            conn.close()

    def exportar_para_excel(self):
        try:
            conn = pyodbc.connect(conn_string)
            cursor = conn.cursor()

            query = """
            SELECT id, nome_equipamento, retidara, devolucao, statu, nome_pessoa 
            FROM registros 
            WHERE retidara >= DATEADD(DAY, -31, GETDATE());
            """
            cursor.execute(query)
            registros = cursor.fetchall()

            if not registros:
                messagebox.showinfo("Aviso", "Nenhum registro encontrado para exportar.")
                return

            dados = []
            for registro in registros:
                dados.append({
                    "ID": registro.id,
                    "Nome Equipamento": registro.nome_equipamento,
                    "Retirada": registro.retidara,
                    "Devolução": registro.devolucao.strftime('%d/%m/%Y %H:%M:%S') if registro.devolucao else 'Pendente',
                    "Status": registro.statu,
                    "Nome Pessoa": registro.nome_pessoa
                })

            df = pd.DataFrame(dados)

            pasta_documentos = os.path.expanduser("~/Documents")
            arquivo_excel = os.path.join(pasta_documentos, "registros.xlsx")

            df.to_excel(arquivo_excel, index=False, engine='openpyxl')

            messagebox.showinfo("Sucesso", f"Registros exportados para '{arquivo_excel}' com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar registros: {e}")
        finally:
            cursor.close()
            conn.close()

    def on_mouse_wheel(self, event):
        self.registros_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def on_frame_configure(self, event):
        self.registros_canvas.configure(scrollregion=self.registros_canvas.bbox("all"))

if __name__ == "__main__":
    root = tk.Tk()
    app = RegistroApp(root)
    root.mainloop()
