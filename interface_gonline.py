import tkinter as tk
from fileinput import filename
from tkinter import filedialog, Frame
import tkinter.scrolledtext as st
import tkinter.font as tkfont
import shutil
import mammoth
from bs4 import BeautifulSoup
import os


# Cria a janela principal
root = tk.Tk()
root.title("Gerador de Estrutura CRCPR Online")


input_path = tk.StringVar()
output_path = tk.StringVar()
content_read = tk.StringVar()

def input_file():
    filetypes = (("Arquivos Word", "*.docx"),
                 ("All files", "*.*"))
    filepath = filedialog.askopenfilename(filetypes=filetypes)
    if filename:
        text_inpath["text"] = filepath
        text_inpath["anchor"] = "e"
        input_path.set(filepath)


def output_file():
    outpath = filedialog.askdirectory()
    outpath += "/boletim-copia.docx"

    if filename:
        text_outpath["text"] = outpath
        text_outpath["anchor"] = "e"
        output_path.set(outpath)


def copy_file():
    try:
        shutil.copy(input_path.get(), output_path.get())
        print("Cópia feita com sucesso!")
    except shutil.SameFileError:
        print("(!) Arquivo de cópia e arquivo final são o mesmo arquivo.")
    except PermissionError:
        print("(!) Permissão negada")
    except Exception as e:
        print(f"(!) Erro: {e}")


def transform_file():
    out_path = output_path.get()
    output_html = out_path.replace("-copia.docx", "-output.html")

    with open(out_path, "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
        text = result.value  # The generated HTML
        try:
            with open(output_html, "w", encoding="utf-8") as html_file:
                html_file.write(text)
            print("Arquivo html criado com sucesso!")
        except Exception as e:
            print(f"(!) Erro: {e}")


def create_structure():
    destino = output_path.get()
    prep_online = destino.replace("boletim-copia.docx", "estrutura-crcpr-online.txt")
    output_html = destino.replace("-copia.docx", "-output.html")
    lista_materias = []
    infos_antes = []
    erro_email = "maito"

    with open(output_html, "r", encoding="utf-8") as file:
        html_content = file.read()

    soup = BeautifulSoup(html_content, "lxml")

    marcador_materia = soup.find(string="Lista de matérias:")

    for texto in marcador_materia.find_all_previous(["p", "li"]):
        if erro_email in texto.get_text():
            corrige_email = texto.get_text().replace("maito", "mailto")
            infos_antes.append(corrige_email)
        else:
            infos_antes.append(texto.get_text())

    infos_antes.reverse()
    infos_antes.pop(0)  # remove o texto extra do topo do arquivo
    infos_antes.pop(0)
    infos_antes.pop(-1)  # remove o "Lista matérias" dessa lista

    for link in marcador_materia.find_all_next("a"):
        noticia_id = link.get("href").replace("https://www3.crcpr.org.br/crcpr/noticias/", "")
        lista_materias.append(noticia_id)

    with open(prep_online, "w", encoding="utf-8") as arquivo:
        for item in infos_antes:
            arquivo.write(item + "\n\n")
        for item in lista_materias:
            arquivo.write(item + ",")


def remove_files():
    arquivo1 = output_path.get()
    arquivo2 = arquivo1.replace("-copia.docx", "-output.html")

    lista_paths = [arquivo1, arquivo2]

    for item in lista_paths:
        if os.path.exists(item):
            os.remove(item)
            print("Arquivos removidos com sucesso!")
        else:
            print("Arquivo não encontrado!")


def open_final():
    destino = output_path.get()
    prep_online = destino.replace("boletim-copia.docx", "estrutura-crcpr-online.txt")

    with open(prep_online, "r", encoding="utf-8") as arquivo:
        conteudo = arquivo.read()
        content_file.delete("1.0", tk.END)
        content_file.insert(tk.END, conteudo)


def gerar_online():
    copy_file()
    transform_file()
    create_structure()
    remove_files()
    open_final()
    btn_run["text"] = "Concluído!"
    btn_run["bg"] = "lightgreen"
    btn_run["fg"] = "darkgreen"


# Estilo
corpo_fonte = tkfont.Font(font="Calibri", size=10)
bg_color = "#f1f5f9"
root.config(bg=bg_color)

main_frame = Frame(root, bg=bg_color, padx=20)
main_frame.pack(side="left")


# Arquivo de entrada

input_frame1 = Frame(main_frame, pady=10, bg=bg_color)
input_frame1.pack(fill="x")

wrapper_inpath = tk.LabelFrame(input_frame1,text="Arquivo do Boletim",font=corpo_fonte,labelanchor="nw",fg="#1e293b",bg=bg_color)
wrapper_inpath.pack(ipady=3, ipadx=15, fill="x")

text_inpath = tk.Label(wrapper_inpath,width=24,text="-- Escolha um arquivo (.docx) --",font="Calibri 11 italic",fg="#475569",anchor="w",bg=bg_color)
text_inpath.pack(side="left", fill="both",padx=3, pady=6)

btn_inpath = tk.Button(wrapper_inpath,text="Procurar",font=corpo_fonte,pady=2,relief="solid",bd=0,bg="#b5cdff",command=input_file)
btn_inpath.pack(side="top", fill="both", padx=5)


# Arquivo de saída

input_frame2 = Frame(main_frame, pady=10, bg=bg_color)
input_frame2.pack(fill="x")

wrapper_outpath = tk.LabelFrame(input_frame2,text="Pasta de Saída",font=corpo_fonte,labelanchor="nw",fg="#1e293b",bg=bg_color)
wrapper_outpath.pack(ipady=3, ipadx=15, fill="x")

text_outpath = tk.Label(wrapper_outpath,width=24,text="-- Escolha uma pasta --",font="Calibri 11 italic",anchor="w",fg="#475569",bg=bg_color)
text_outpath.pack(side="left", fill="both", padx=3, pady=6)

btn_outpath = tk.Button(wrapper_outpath,text="Procurar",font=corpo_fonte,pady=2,relief="solid",bd=0,bg="#b5cdff",command=output_file)
btn_outpath.pack(side="top", fill="both", padx=5)


# Preview do conteúdo

lateral_frame = Frame(root, padx=5, pady=15, bg=bg_color)
lateral_frame.pack(side="right",padx=5)

wrapper_content = tk.LabelFrame(lateral_frame,text="Estrutura CRCPR Online",font=corpo_fonte,labelanchor="nw",bg=bg_color,fg="#1e293b")
wrapper_content.pack(side="top", ipady=5)

content_file = st.ScrolledText(wrapper_content, wrap="word",bg="#f8fafc",foreground="black",borderwidth=0,border=0,width=60,height=15)
content_file.pack(fill="both", padx=5)


# Área do botão

bottom_frame = Frame(main_frame, pady=10, bg=bg_color)
bottom_frame.pack(fill="x")

btn_run = tk.Button(bottom_frame,text="Gerar",bg="#3549ff",fg="white",pady=10,font="Calibri 13 bold",relief="solid",bd=0,command=gerar_online)
btn_run.pack(fill="x") # botão "gerar"



root.mainloop()
