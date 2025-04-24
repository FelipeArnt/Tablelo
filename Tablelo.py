from tkinter import Tk, filedialog
import pdfplumber # type: ignore
import os
from rich.console import Console # type: ignore
from rich.panel import Panel # type: ignore
from rich.table import Table # type: ignore
from rich import print # type: ignore
from rich.prompt import Prompt # type: ignore

class Extractor:
    def __init__(self):
        self.console = Console()
    
    def selecionar_arquivo(self):
        """Seleciona um arquivo usando interface gráfica"""
        root = Tk()
        root.withdraw()
        caminho_arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo",
            filetypes=[("PDF Files", "*.pdf")]
        )
        root.destroy()
        return caminho_arquivo if caminho_arquivo else None

    def extrair_tabelas(self, caminho_pdf, paginas_selecionadas=None):
        """Extrai tabelas de páginas específicas de um arquivo PDF"""
        try:
            with pdfplumber.open(caminho_pdf) as pdf:
                total_paginas = len(pdf.pages)
                
                # Se nenhuma página for especificada, processa todas
                if paginas_selecionadas is None:
                    return [pagina.extract_tables() for pagina in pdf.pages], total_paginas
                
                # Processa apenas as páginas selecionadas
                tabelas = []
                for num_pagina in paginas_selecionadas:
                    if 1 <= num_pagina <= total_paginas:
                        tabelas.append(pdf.pages[num_pagina-1].extract_tables())
                    else:
                        self.console.print(f"[yellow]Aviso: Página {num_pagina} não existe no documento[/yellow]")
                
                return tabelas, total_paginas
        except Exception as e:
            self.console.print(f"[red]Erro ao ler o PDF: {e}[/red]")
            return None, 0

class Processador:
    def __init__(self):
        self.console = Console()
    
    def limpar_dados(self, tabela):
        """Remove linhas e células vazias da tabela"""
        tabela = [linha for linha in tabela if any(cell and str(cell).strip() for cell in linha)]
        return [[str(cell).strip() for cell in linha if cell and str(cell).strip()] for linha in tabela]
    
    def processar_tabelas(self, tabelas_por_pagina):
        """Processa todas as tabelas extraídas"""
        headers_por_pagina, data_por_pagina = [], []
        for tabelas in tabelas_por_pagina:
            for tabela in tabelas:
                tabela_limpa = self.limpar_dados(tabela)
                if tabela_limpa and len(tabela_limpa) > 0:
                    headers_por_pagina.append(tabela_limpa[0])
                    data_por_pagina.append(tabela_limpa[1:] if len(tabela_limpa) > 1 else [])
        return headers_por_pagina, data_por_pagina
    
    def remover_linhas(self, data_por_pagina, headers_por_pagina):
        """Interface para remoção de linhas das tabelas"""
        for pagina_num, data in enumerate(data_por_pagina):
            self.console.print(Panel.fit(f"Caso de Teste {pagina_num + 1}", style="blue"))
            
            table = Table(show_header=True, header_style="bold magenta")
            table.add_column("Linha")
            for header in headers_por_pagina[pagina_num]:
                table.add_column(header)
            
            for i, linha in enumerate(data, 1):
                table.add_row(str(i), *[str(item) for item in linha])
            
            self.console.print(table)
        
        while True:
            linhas_para_remover = Prompt.ask("Digite os números das linhas que deseja remover (separados por vírgula)")
            if not linhas_para_remover:
                self.console.print("[yellow]Nenhuma linha foi removida.[/yellow]")
                return data_por_pagina
            
            try:
                indices = sorted({int(i.strip()) - 1 for i in linhas_para_remover.split(",")}, reverse=True)
                return [
                    [linha for i, linha in enumerate(pagina) if i not in indices]
                    for pagina in data_por_pagina
                ]
            except ValueError:
                self.console.print("[red]Erro: Insira números válidos separados por vírgula.[/red]")

    def filtrar_tabelas_por_palavra_chave(self, headers_por_pagina, data_por_pagina, palavras_chave):
        """Filtra tabelas que contêm palavras-chave específicas"""
        novos_headers = []
        novos_dados = []
        tabelas_removidas = 0
        
        for headers, data in zip(headers_por_pagina, data_por_pagina):
            # Verifica se a tabela contém alguma palavra-chave
            contem_palavra_chave = any(
                any(any(palavra.lower() in str(cell).lower() 
                    for palavra in palavras_chave)
                    for cell in linha)
                for linha in [headers] + data
            )
            
            if not contem_palavra_chave:
                novos_headers.append(headers)
                novos_dados.append(data)
            else:
                tabelas_removidas += 1
        
        self.console.print(f"[green]Foram removidas {tabelas_removidas} tabelas contendo as palavras-chave.[/green]")
        return novos_headers, novos_dados

class Exportador:
    def __init__(self):
        self.console = Console()
    
    def exportar_markdown(self, headers_por_pagina, data_por_pagina, nome_arquivo="Tabelas.md"):
        """Exporta as tabelas para um arquivo Markdown"""
        try:
            with open(nome_arquivo, "w", encoding="utf-8") as md_file:
                for headers, data in zip(headers_por_pagina, data_por_pagina):
                    md_file.write("| " + " | ".join(headers) + " |\n")
                    md_file.write("|" + "|".join(["---"] * len(headers)) + "|\n")
                    for linha in data:
                        md_file.write("| " + " | ".join(str(cell).replace("\n", "<br>") for cell in linha) + " |\n")
                    md_file.write("\n\n")
            
            self.console.print(f"[green]Tabelas exportadas para '{nome_arquivo}'[/green]")
            return True
        except Exception as e:
            self.console.print(f"[red]Erro ao exportar: {e}[/red]")
            return False

class Tablelo:
    def __init__(self):
        self.console = Console()
        self.extractor = Extractor()
        self.processor = Processador()
        self.exporter = Exportador()
        self.caminho_arquivo = None
        self.headers_por_pagina = []
        self.data_por_pagina = []
        self.total_paginas = 0
    
    def selecionar_paginas(self):
        """Interface para seleção de páginas a serem processadas"""
        self.console.print(f"\nDocumento possui {self.total_paginas} páginas no total")
        while True:
            entrada = Prompt.ask(
                "Digite os números das páginas que deseja processar (ex: 1,3-5,7)\n"
                "Ou deixe em branco para processar todas",
                default=""
            )
            
            if not entrada.strip():
                return None  # Processa todas
            
            try:
                paginas = set()
                for parte in entrada.split(','):
                    parte = parte.strip()
                    if '-' in parte:
                        inicio, fim = map(int, parte.split('-'))
                        paginas.update(range(inicio, fim + 1))
                    else:
                        paginas.add(int(parte))
                
                return sorted(paginas)
            except ValueError:
                self.console.print("[red]Formato inválido. Use números e intervalos como 1,3-5,7[/red]")

    def carregar_arquivo(self, caminho=None):
        """Carrega e processa um novo arquivo"""
        self.caminho_arquivo = caminho or self.extractor.selecionar_arquivo()
        if not self.caminho_arquivo or not os.path.exists(self.caminho_arquivo):
            self.console.print("[red]Arquivo inválido ou não encontrado[/red]")
            return False
        
        # Primeira extração para saber o total de páginas
        _, self.total_paginas = self.extractor.extrair_tabelas(self.caminho_arquivo)
        
        # Seleciona páginas para processar
        paginas = self.selecionar_paginas()
        
        tabelas, _ = self.extractor.extrair_tabelas(self.caminho_arquivo, paginas)
        if not tabelas:
            return False
        
        self.headers_por_pagina, self.data_por_pagina = self.processor.processar_tabelas(tabelas)
        return True
    
    def exibir_tabelas(self):
        """Exibe todas as tabelas processadas"""
        for pagina_num, (headers, data) in enumerate(zip(self.headers_por_pagina, self.data_por_pagina)):
            self.console.print(Panel.fit(f"Tabela {pagina_num + 1}", style="blue"))
            
            table = Table(show_header=True, header_style="bold magenta")
            for header in headers:
                table.add_column(header)
            
            for linha in data:
                table.add_row(*[str(item) for item in linha])
            
            self.console.print(table)
    
    def filtrar_tabelas_indesejadas(self):
        """Filtra tabelas contendo palavras-chave específicas"""
        palavras_padrao = "marca d'água,analisado por,assinatura,elaborado por"
        palavras_chave = Prompt.ask(
            "Digite as palavras-chave para filtrar (separadas por vírgula)",
            default=palavras_padrao
        ).split(",")
        
        palavras_chave = [palavra.strip().lower() for palavra in palavras_chave if palavra.strip()]
        
        if not palavras_chave:
            self.console.print("[yellow]Nenhuma palavra-chave fornecida.[/yellow]")
            return
        
        self.headers_por_pagina, self.data_por_pagina = self.processor.filtrar_tabelas_por_palavra_chave(
            self.headers_por_pagina, self.data_por_pagina, palavras_chave
        )

    def executar(self):
        """Loop principal da aplicação"""
        self.console.print(Panel.fit("Tablelo - Extrator de Tabelas PDF", style="bold blue"))
        
        if not self.carregar_arquivo():
            return
        
        while True:
            self.console.print(Panel.fit("Menu Principal", style="bold blue"))
            self.console.print("[1] Selecionar novo arquivo")
            self.console.print("[2] Exibir tabelas")
            self.console.print("[3] Remover linhas")
            self.console.print("[4] Filtrar tabelas por palavras-chave")
            self.console.print("[5] Exportar para Markdown")
            self.console.print("[0] Sair")
            
            opcao = Prompt.ask("Escolha uma opção", choices=["0", "1", "2", "3", "4", "5"])
            
            if opcao == "1":
                self.carregar_arquivo()
            elif opcao == "2":
                self.exibir_tabelas()
            elif opcao == "3":
                self.data_por_pagina = self.processor.remover_linhas(
                    self.data_por_pagina, self.headers_por_pagina
                )
            elif opcao == "4":
                self.filtrar_tabelas_indesejadas()
            elif opcao == "5":
                self.exporter.exportar_markdown(
                    self.headers_por_pagina, self.data_por_pagina
                )
            elif opcao == "0":
                self.console.print("[green]Até logo![/green]")
                break

if __name__ == "__main__":
    app = Tablelo()
    app.executar()
