from tkinter import Tk, filedialog
from tabulate import tabulate
import pdfplumber
import os
from termcolor import colored  # type: ignore

def selecionar_arquivo():
    root = Tk()
    root.withdraw()  # Esconde a janela principal do tkinter
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo",
        filetypes=[
            ("PDF Files", "*.pdf"),
            ("Excel Files", "*.xlsx *.xls"),
            ("All Files", "*.*")
        ]
    )
    root.destroy()  # Fecha a janela do tkinter
    return caminho_arquivo if caminho_arquivo else None

def extrair_tabelas_pdf(caminho_pdf):
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            tabelas_por_pagina = []
            for pagina in pdf.pages:
                tabelas = pagina.extract_tables()
                tabelas_por_pagina.append(tabelas if tabelas else [])
            return tabelas_por_pagina
    except Exception as e:
        print(colored(f"Erro ao ler o PDF: {e}", "red"))
        return None

def limpar_dados(tabela):
    tabela = [linha for linha in tabela if any(cell and str(cell).strip() for cell in linha)]
    tabela = [[str(cell).strip() for cell in linha if cell and str(cell).strip()] for linha in tabela]
    return tabela

def remover_linhas_de_todas(data_por_pagina, headers_por_pagina):
    for pagina_num, data in enumerate(data_por_pagina):
        print(colored(f"\n Caso de teste ", "blue"))
        headers = headers_por_pagina[pagina_num]
        data_com_indices = [[i + 1] + linha for i, linha in enumerate(data)]
        headers_com_indices = ["Linha"] + headers
        print(tabulate(data_com_indices, headers=headers_com_indices, tablefmt="grid"))
    
    while True:
        linhas_para_remover = input("Digite os números das linhas que deseja remover (separados por vírgula): ").strip()
        if not linhas_para_remover:
            print(colored("Nenhuma linha foi removida.", "yellow"))
            return data_por_pagina
        try:
            indices = sorted(set(int(i.strip()) - 1 for i in linhas_para_remover.split(",")), reverse=True)
            for pagina_num in range(len(data_por_pagina)):
                data_por_pagina[pagina_num] = [linha for i, linha in enumerate(data_por_pagina[pagina_num]) if i not in indices]
            print(colored("\nTabelas atualizadas.", "green"))
            return data_por_pagina
        except ValueError:
            print(colored("Erro: Insira números válidos separados por vírgula.", "red"))

def exportar_todas_tabelas(headers_por_pagina, data_por_pagina, nome_arquivo="Tabelas.md"):
    try:
        with open(nome_arquivo, "w", encoding="utf-8") as md_file:
            for pagina_num, (headers, data) in enumerate(zip(headers_por_pagina, data_por_pagina)):
                # Escreve o título da tabela
                md_file.write(f"## Caso de Teste - {pagina_num + 1}\n\n")
                
                # Escreve o cabeçalho
                md_file.write(f"| {' | '.join(headers)} |\n")
                md_file.write(f"|{'|'.join(['---'] * len(headers))}|\n")
                
                # Escreve os dados
                for linha in data:
                    linha_formatada = [str(celula).replace("\n", "<br>") for celula in linha]
                    md_file.write(f"| {' | '.join(linha_formatada)} |\n")
                
                # Adiciona separador entre tabelas
                md_file.write("\n<br>\n\n")
        
        print(colored(f"\nTodas as tabelas foram exportadas para '{nome_arquivo}'.", "green"))
        return True
    except Exception as e:
        print(colored(f"Erro ao exportar tabelas: {e}", "red"))
        return False

def menu_interativo(caminho_arquivo):
    if not caminho_arquivo or not os.path.exists(caminho_arquivo):
        print(colored("Erro: Nenhum arquivo selecionado ou arquivo não encontrado.", "red"))
        return
    
    tabelas_por_pagina = extrair_tabelas_pdf(caminho_arquivo)
    if not tabelas_por_pagina:
        print(colored("Erro: Nenhuma tabela encontrada no PDF.", "red"))
        return

    headers_por_pagina, data_por_pagina = [], []
    for pagina_num, tabelas in enumerate(tabelas_por_pagina):
        for tabela in tabelas:
            tabela_limpa = limpar_dados(tabela)
            if tabela_limpa and len(tabela_limpa) > 0:  # Garante que tem pelo menos os headers
                headers_por_pagina.append(tabela_limpa[0])
                data_por_pagina.append(tabela_limpa[1:] if len(tabela_limpa) > 1 else [])
    
    while True:
        print("\n--- Tablelo ---")
        print("[1] - Selecionar novo arquivo")
        print("[2] - Exibir tabelas")
        print("[3] - Remover linhas")
        print("[4] - Exportar todas as tabelas para Markdown")
        print("[0] - Sair")
        opcao = input("Escolha uma opção: ").strip()
        
        if opcao == "1":
            novo_arquivo = selecionar_arquivo()
            if novo_arquivo:
                caminho_arquivo = novo_arquivo
                tabelas_por_pagina = extrair_tabelas_pdf(caminho_arquivo)
                headers_por_pagina, data_por_pagina = [], []
                for pagina_num, tabelas in enumerate(tabelas_por_pagina):
                    for tabela in tabelas:
                        tabela_limpa = limpar_dados(tabela)
                        if tabela_limpa and len(tabela_limpa) > 0:
                            headers_por_pagina.append(tabela_limpa[0])
                            data_por_pagina.append(tabela_limpa[1:] if len(tabela_limpa) > 1 else [])
        elif opcao == "2":
            for pagina_num, data in enumerate(data_por_pagina):
                print(colored(f"\nTabela da página {pagina_num + 1}:", "blue"))
                print(tabulate(data, headers=headers_por_pagina[pagina_num], tablefmt="grid"))
        elif opcao == "3":
            data_por_pagina = remover_linhas_de_todas(data_por_pagina, headers_por_pagina)
        elif opcao == "4":
            exportar_todas_tabelas(headers_por_pagina, data_por_pagina)
        elif opcao == "0":
            print("Saindo...")
            break
        else:
            print(colored("Opção inválida. Tente novamente.", "red"))

if __name__ == "__main__":
    print("Selecione o arquivo PDF para começar...")
    caminho_arquivo = selecionar_arquivo()
    if caminho_arquivo:
        menu_interativo(caminho_arquivo)
    else:
        print(colored("Nenhum arquivo selecionado. O programa será encerrado.", "red"))
