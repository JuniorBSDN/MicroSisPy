import tkinter as tk
from tkinter import ttk, messagebox
import datetime
import os
import pandas as pd
from escpos.printer import Usb  # Pode ser Usb, Serial ou Network dependendo da sua impressora
import sys


class ControleFinanceiroApp:
    def __init__(self, master):
        self.master = master
        master.title("MicroSis - CAIXA REGISTRADORA CLASSICA")
        master.geometry("1000x680")
        master.resizable(False, False)  # Impede que a janela seja redimensionada

        self.valores = []
        self.comanda_counter = self.get_next_comanda_number()
        self.arquivo_excel = "transacoes.xlsx"  # Nome do arquivo Excel

        #------IMAGENS DOS BOTOES------
        self.add = tk.PhotoImage(file="add.png")
        self.remove = tk.PhotoImage(file="remove.png")
        self.comanda = tk.PhotoImage(file="comanda.png")
        self.resumo = tk.PhotoImage(file="resumo.png")
        # --- Cores para o Tema Escuro Básico ---
        self.bg_dark = "#222222"  # Um tom mais escuro de azul-cinza para o fundo
        self.fg_light = "#ecf0f1"  # Cor do texto claro (quase branco)
        self.frame_bg = "#111111"  # Um tom ligeiramente mais claro para os frames
        self.entry_bg = "#222222"  # Fundo de campos de entrada
        self.entry_fg = "#00BFFF"  # Texto em campos de entrada
        self.button_bg = "silver"  # Azul vibrante para botões
        self.button_fg = "#111111"  # Texto branco nos botões
        self.tree_header_bg = "#444444"  # Fundo do cabeçalho da tabela
        self.tree_bg = "silver"  # Fundo da tabela (um pouco mais escuro)
        self.tree_fg = "#111111"  # Texto da tabela
        self.tree_select_bg = "#444444"  # Fundo da seleção na tabela
        self.total_fg = "#00BFFF"  # Verde para o total

        master.config(bg=self.bg_dark)  # Define o fundo da janela principal

        # --- Configurações da impressora térmica (AJUSTE AQUI!) ---
        self.impressora_vendor_id = 0x04b8
        self.impressora_product_id = 0x0202
        self.printer = None

        # Cria o arquivo Excel se ele não existir, com os cabeçalhos
        if not os.path.exists(self.arquivo_excel):
            df_vazio = pd.DataFrame(columns=["Numero_Comanda", "Data_Hora", "Valores_Cadastrados", "Total"])
            df_vazio.to_excel(self.arquivo_excel, index=False)

        # --- Widgets com place e cores simples ---

        self.logo = tk.Label(master, bg=self.bg_dark, text="MicroSistem", fg="#00BFFF",font=("consolas 20 bold"), padx=10, pady=10)
        self.logo.place(x=100, y=20, width=400, height=50)
        self.titulo = tk.Label(master, bg=self.bg_dark, text="--CAIXA REGISTRADORA CLASSICA--",font=("system 10 bold"), fg="#00BFFF", padx=10, pady=10)
        self.titulo.place(x=180, y=65, width=250, height=20)

        # Frame para entrada de valores
        self.frame_entrada = tk.Frame(master, bg=self.frame_bg, bd=0, relief="groove")
        self.frame_entrada.place(x=20, y=120, width=610, height=100)

        self.labelInfo = tk.Label(self.frame_entrada,bg=self.frame_bg, text="Inserir valor", font=("system 12 bold"), fg=self.fg_light,)
        self.labelInfo.place(x=20, y=10, width=100, height=20)

        self.label_valor = tk.Label(self.frame_entrada, text="Valor (R$):", bg=self.frame_bg, fg=self.fg_light, font=("system 20 bold"))
        self.label_valor.place(x=20, y=50, width=140, height=20)

        self.entrada_valor = tk.Entry(self.frame_entrada, bg=self.entry_bg, bd=0, fg=self.entry_fg,insertbackground=self.entry_fg, font=("system", 20, "bold"))
        self.entrada_valor.place(x=170, y=35, width=420, height=50)
        self.entrada_valor.bind("<Return>", self.adicionar_valor_event)
        self.entrada_valor.focus_set()

        self.botao_adicionar = tk.Button(master, image=self.add, bg=self.button_bg, fg=self.button_fg,
                                         font=("system", 12, "bold"), relief="flat", command=self.adicionar_valor)
        self.botao_adicionar.place(x=660, y=125, width=320, height=90)  # Ajustei altura para ficar melhor

        # Tabela para exibir valores
        self.frame_tabela = tk.Frame(master, bg=self.frame_bg, bd=0, relief="groove")
        self.frame_tabela.place(x=20, y=240, width=610, height=215)

        # Configurações de estilo para Treeview (ainda usando ttk para a tabela ser moderna)
        style = ttk.Style()
        style.theme_use("clam")  # 'clam' é um bom ponto de partida para personalizar Treeview
        style.configure("Custom.Treeview",background=self.tree_bg,foreground=self.tree_fg,
                        fieldbackground=self.tree_bg,
                        font=("system", 14, "bold"))#tamanho da fonte na tabela
        style.map("Custom.Treeview",
                  background=[("selected", self.tree_select_bg)])
        style.configure("Custom.Treeview.Heading",
                        background=self.tree_header_bg,
                        foreground=self.fg_light,
                        font=("system", 10, "bold"))

        self.tree = ttk.Treeview(self.frame_tabela, columns=("Valor"), show="headings", style="Custom.Treeview")
        self.tree.heading("Valor", text="Lista dos adicionados")
        self.tree.column("Valor", width=600, anchor="nw")
        self.tree.place(x=0, y=0, relwidth=1, relheight=1)

        self.scrollbar = tk.Scrollbar(self.tree, orient="vertical", command=self.tree.yview, bg=self.frame_bg)
        self.scrollbar.place(relx=1, rely=0, relheight=1, anchor="ne")

        self.tree.configure(yscrollcommand=self.scrollbar.set)

        # Rótulo para o Total
        self.label_total = tk.Label(master, text="Total: R$ 0.00", bg="black", fg=self.total_fg,font=("System", 30, "bold"))
        self.label_total.place(x=20, y=480, width=610, height=100)

        # Botões
        self.botao_remover = tk.Button(master, image=self.remove, bg=self.button_bg,fg=self.button_fg, font=("System", 12, "bold"), relief="flat",command=self.remover_valor)
        self.botao_remover.place(x=660, y=240, width=320, height=95)  # Altura ajustada

        self.botao_finalizar = tk.Button(master, image=self.comanda, bg=self.button_bg, fg=self.button_fg, font=("System", 12, "bold"), relief="flat", command=self.finalizar_comanda)
        self.botao_finalizar.place(x=660, y=360, width=320, height=95)  # Altura ajustada

        # NOVO BOTÃO PARA ABRIR NOVA JANELA
        self.botao_abrir_nova_janela = tk.Button(master, image=self.resumo, bg=self.button_bg, fg=self.button_fg,
                                                font=("System", 12, "bold"), relief="flat", command=self.abrir_nova_janela)
        self.botao_abrir_nova_janela.place(x=660, y=480, width=320, height=95) # Posição para o novo botão

        #_____informações de comandos______
        self.labeInfo1 = tk.Label(master, text= "* Press ( ENTER ) - Adiciona o valor na lista",fg="white", bg="#222222", font="System 10 bold")
        self.labeInfo1.place(x=20,y=600,width=290,height=20)

        self.labeInfo2 = tk.Label(master, text="* Press ( F ) - Finaliza a tranzação, emite a comanda e salva", fg="white", bg="#222222",
                                  font="System 10 bold")
        self.labeInfo2.place(x=20, y=620, width=390, height=20)

        # Bind a tecla 'f' para finalizar e imprimir
        master.bind('<f>', self.finalizar_comanda_event)
        master.bind('<F>', self.finalizar_comanda_event)

        self.atualizar_tabela()

    def get_next_comanda_number(self):
        # Percorre apenas os arquivos no diretório atual para evitar problemas de permissão ou caminhos.
        comanda_files = [f for f in os.listdir('.') if f.startswith('comanda_') and f.endswith('.txt')]
        if not comanda_files:
            return 1

        numbers = []
        for f in comanda_files:
            try:
                num_str = f.replace('comanda_', '').replace('.txt', '')
                numbers.append(int(num_str))
            except ValueError:
                continue

        if numbers:
            return max(numbers) + 1
        return 1

    def adicionar_valor_event(self, event):
        self.adicionar_valor()

    def adicionar_valor(self):
        try:
            valor_str = self.entrada_valor.get().replace(",", ".")
            valor = float(valor_str)
            if valor <= 0:
                messagebox.showwarning("Valor Inválido", "Por favor, insira um valor positivo.")
                return
            self.valores.append(valor)
            self.entrada_valor.delete(0, tk.END)
            self.atualizar_tabela()
        except ValueError:
            messagebox.showerror("Erro de Entrada", "Por favor, insira um número válido.")

    def remover_valor(self):
        selecionado = self.tree.selection()
        if not selecionado:
            messagebox.showwarning("Nenhum Item Selecionado", "Por favor, selecione um valor para remover.")
            return

        for item in selecionado:
            index = self.tree.index(item)
            if 0 <= index < len(self.valores):
                self.valores.pop(index)
        self.atualizar_tabela()

    def atualizar_tabela(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for valor in self.valores:
            self.tree.insert("", "end", values=(f"R$ {valor:.2f}",))

        total = sum(self.valores)
        self.label_total.config(text=f"Total: R$ {total:.2f}")

    def finalizar_comanda_event(self, event):
        self.finalizar_comanda()

    def finalizar_comanda(self):
        if not self.valores:
            messagebox.showwarning("Nenhum Valor", "Não há valores para finalizar a comanda.")
            return

        total = sum(self.valores)
        data_hora_obj = datetime.datetime.now()
        data_hora_str = data_hora_obj.strftime("%Y-%m-%d %H:%M:%S")

        numero_comanda_atual = self.comanda_counter
        nome_arquivo_txt = f"comanda_{numero_comanda_atual}.txt"
        self.comanda_counter += 1

        try:
            # Gerar o arquivo TXT (Comanda)
            conteudo_comanda_txt = self.gerar_conteudo_comanda(numero_comanda_atual, data_hora_str, self.valores, total)
            with open(nome_arquivo_txt, "w") as f:
                f.write(conteudo_comanda_txt)
            # Salvar no Excel
            self.salvar_no_excel(numero_comanda_atual, data_hora_str, self.valores, total)

            # --- Escolha a opção de impressão aqui! ---
            # Opção 1: Impressão Térmica (recomendado se for ESC/POS)
            self.imprimir_comanda_termica(conteudo_comanda_txt)

            # Opção 2: Impressão via OS (para impressoras comuns no Windows)
            # Descomente a linha abaixo e comente a linha da impressão térmica se quiser usar esta.
            # self.imprimir_txt_via_os(nome_arquivo_txt)

            # --- Adicionado para apagar o arquivo TXT após a impressão ---
            if os.path.exists(nome_arquivo_txt):
                os.remove(nome_arquivo_txt)
                print(f"Arquivo '{nome_arquivo_txt}' apagado com sucesso.") # Mensagem de depuração
            # -----------------------------------------------------------

            self.valores = []
            self.atualizar_tabela()
            self.entrada_valor.focus_set()

        except Exception as e:
            messagebox.showerror("Erro ao Finalizar Transação", f"Ocorreu um erro: {e}")

    def gerar_conteudo_comanda(self, numero_comanda, data_hora_str_original, valores_cadastrados, total):
        conteudo = ""
        conteudo += f"COMANDA: {numero_comanda}\n"

        # Obter o fuso horário de Vigia, Pará, Brasil (GMT-3)
        #data_hora_local = datetime.datetime.now() - datetime.timedelta(hours=3)  # GMT-3
        #conteudo += f"Data/Hora: {data_hora_local.strftime('%d/%m/%Y %H:%M:%S')}\n"

        conteudo += "-" * 30 + "\n"
        conteudo += "Itens:\n"
        for i, valor in enumerate(valores_cadastrados):
            conteudo += f"{i + 1}. R$ {valor:.2f}\n"
        conteudo += "-" * 30 + "\n"
        conteudo += f"Total pago: R$ {total:.2f}\n"
        conteudo += "----------------------------\n"
        return conteudo

    def salvar_no_excel(self, numero_comanda, data_hora, valores_cadastrados, total):
        try:
            df = pd.read_excel(self.arquivo_excel)
            valores_str = ", ".join([f"R$ {v:.2f}" for v in valores_cadastrados])

            nova_linha = pd.DataFrame([{
                "Numero_Comanda": numero_comanda,
                "Data_Hora": data_hora,
                "Valores_Cadastrados": valores_str,
                "Total": total
            }])

            df = pd.concat([df, nova_linha], ignore_index=True)
            df.to_excel(self.arquivo_excel, index=False)

        except Exception as e:
            messagebox.showerror("Erro ao Salvar no Excel", f"Não foi possível salvar os dados no Excel: {e}")

    def imprimir_comanda_termica(self, conteudo):
        try:
            self.printer = Usb(self.impressora_vendor_id, self.impressora_product_id, in_ep=0x82, out_ep=0x01)
            self.printer.text(conteudo)
            self.printer.cut()

        except Exception as e:
            messagebox.showwarning("Erro de Impressão Térmica",
                                   f"Não foi possível imprimir na impressora térmica. Verifique a conexão e os IDs. Erro: {e}")
        finally:
            if self.printer:
                self.printer.close()

    def imprimir_txt_via_os(self, nome_arquivo_txt):
        if sys.platform == "win32":
            try:
                os.startfile(nome_arquivo_txt, "print")

            except Exception as e:
                messagebox.showerror("Erro de Impressão via OS",
                                     f"Não foi possível imprimir o arquivo '{nome_arquivo_txt}' via sistema operacional. Erro: {e}")
        else:
            messagebox.showwarning("Impressão via OS",
                                   "Esta função de impressão via OS é otimizada para Windows. Para Linux/macOS, a impressão de TXT diretamente pode exigir ferramentas de linha de comando como 'lpr'.")

    # --- Nova função para abrir uma nova janela ---
    def abrir_nova_janela(self):
        nova_janela = tk.Toplevel(self.master) # Cria uma nova janela top-level
        nova_janela.title("Resumo de entradas")
        nova_janela.geometry("400x300")
        nova_janela.config(bg="black")
        nova_janela.transient(self.master) # Faz a nova janela aparecer sobre a principal
        nova_janela.grab_set() # Bloqueia a interação com a janela principal enquanto a nova está aberta

        # Adicione widgets à nova janela
        label_nova_janela = tk.Label(nova_janela,fg="red",bg="black", text="MicroSis", font=("Consolas", 16))
        label_nova_janela.pack(pady=20)
        labelText = tk.Label(nova_janela,fg="red",bg="black", text=" caixa registradora classica", font=("System", 10))
        labelText.pack(pady=20)


        botao_fechar = tk.Button(nova_janela,fg="red",bg="black", text="Fechar Janela", command=nova_janela.destroy)
        botao_fechar.pack(pady=10)

        self.master.wait_window(nova_janela) # Espera a nova janela ser fechada para interagir com a principal


# Inicia a aplicação
if __name__ == "__main__":
    root = tk.Tk()
    app = ControleFinanceiroApp(root)
    root.mainloop()