# importaçoes
import PyPDF2
import os

merger = PyPDF2.PdfMerger()

lista_arquivo = os.listdir('arquivos')
lista_arquivo.sort()

for arquivo in lista_arquivo:
    if ".pdf" in arquivo:
        merger.append(f"arquivos/{arquivo}")
merger.write("PDF Final.pdf")
