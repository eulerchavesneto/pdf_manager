import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
from pathlib import Path
import sys
import subprocess
import importlib.util
import importlib.metadata
import re

# --- Verificação e Instalação de Dependências ---

def check_and_install_dependencies():
    """
    Verifica se as dependências necessárias estão instaladas.
    Se não estiverem, tenta instalá-las usando pip.
    """
    required_packages = {
        'PyPDF2': 'PyPDF2',
        'Pillow': 'Pillow',
        'fitz': 'PyMuPDF',  # fitz é importado de PyMuPDF
        'docx': 'python-docx',
    }
    
    missing_packages = []
    print("Verificando dependências...")
    
    for package_name, install_name in required_packages.items():
        # Tenta encontrar a especificação do módulo. Se não encontrar, ele não está instalado.
        if importlib.util.find_spec(package_name) is None:
            print(f"[FALTA] {package_name} não encontrado. Será instalado como '{install_name}'.")
            missing_packages.append(install_name)
        else:
            try:
                version = importlib.metadata.version(install_name)
                print(f"[OK] {package_name} (versão {version}) está instalado.")
            except importlib.metadata.PackageNotFoundError:
                 print(f"[FALTA] {package_name} não encontrado. Será instalado como '{install_name}'.")
                 missing_packages.append(install_name)


    if not missing_packages:
        print("\nTodas as dependências já estão satisfeitas.")
        return True

    print(f"\nInstalando pacotes necessários: {', '.join(missing_packages)}")
    try:
        # Usa subprocess para chamar o pip e instalar os pacotes faltantes.
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install"] + missing_packages,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
        print("Pacotes instalados com sucesso!")
        return True
    except subprocess.CalledProcessError:
        messagebox.showerror(
            "Erro de Dependência",
            f"Não foi possível instalar os pacotes necessários: {', '.join(missing_packages)}.\n"
            f"Por favor, instale-os manualmente executando:\n"
            f"pip install {' '.join(missing_packages)}"
        )
        return False

# Importa as bibliotecas após a verificação para garantir que elas existem.
# Isso evita erros de importação se as dependências não estiverem presentes.
if check_and_install_dependencies():
    from PyPDF2 import PdfMerger, PdfReader, PdfWriter
    from PyPDF2.generic import Destination
    import PIL.Image
    import fitz  # PyMuPDF
    import docx

def sanitize_filename(name):
    """Remove caracteres inválidos de um nome de arquivo e substitui quebras de linha."""
    name = name.replace('\n', ' ').replace('\r', '') # Substitui quebras de linha por espaço
    return re.sub(r'[\\/*?:"<>|]', "", name).strip()

class SplitDefinitionDialog(simpledialog.Dialog):
    """Caixa de diálogo para adicionar ou editar uma definição de divisão de PDF."""
    def __init__(self, parent, title, initial_values=None):
        self.initial_values = initial_values or {}
        super().__init__(parent, title)

    def body(self, master):
        ttk.Label(master, text="Nome do Arquivo (sem .pdf):").grid(row=0, sticky=tk.W, padx=5, pady=2)
        self.name_entry = ttk.Entry(master, width=40)
        self.name_entry.grid(row=1, sticky=(tk.W, tk.E), padx=5, pady=2)
        self.name_entry.insert(0, self.initial_values.get('name', ''))

        ttk.Label(master, text="Intervalo de Páginas (ex: 1-10):").grid(row=2, sticky=tk.W, padx=5, pady=2)
        self.pages_entry = ttk.Entry(master, width=40)
        self.pages_entry.grid(row=3, sticky=(tk.W, tk.E), padx=5, pady=2)
        self.pages_entry.insert(0, self.initial_values.get('pages', ''))
        
        return self.name_entry # Foco inicial

    def apply(self):
        name = self.name_entry.get()
        self.result = {
            'name': sanitize_filename(name), # Limpa o nome do arquivo
            'pages': self.pages_entry.get().strip()
        }

class PDFManagerApp:
    """
    Aplicação com interface gráfica para manipular arquivos PDF, permitindo juntar,
    cortar, reorganizar, converter e inserir páginas.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("Gerenciador de PDFs Otimizado")
        self.root.geometry("850x650")
        self.root.minsize(700, 500)

        # Diretório padrão do usuário para salvar arquivos
        self.USER_HOME_DIR = Path.home()

        # --- Variáveis de Estado ---
        self.pdf_files_to_join = []
        self.join_output_file = tk.StringVar(value=str(self.USER_HOME_DIR / "PDFs_Combinados.pdf"))
        self.interleave_var = tk.BooleanVar(value=False)

        self.cut_pdf_file = tk.StringVar()
        self.cut_pages = tk.StringVar()
        self.cut_output_file = tk.StringVar(value=str(self.USER_HOME_DIR / "PDF_Cortado.pdf"))

        self.split_pdf_file = tk.StringVar()
        self.split_output_folder = tk.StringVar(value=str(self.USER_HOME_DIR))
        self.split_definitions = []
        self.split_import_level = tk.IntVar(value=0) # 0 para importar todos os níveis

        self.reorder_pdf_file = tk.StringVar()
        self.reorder_output_file = tk.StringVar(value=str(self.USER_HOME_DIR / "PDF_Reorganizado.pdf"))
        self.reorder_page_indices = []

        self.convert_pdf_file = tk.StringVar()
        self.convert_output_folder = tk.StringVar(value=str(self.USER_HOME_DIR))
        self.convert_format = tk.StringVar(value="jpg")
        self.convert_dpi = tk.IntVar(value=300)
        
        self.insert_target_pdf = tk.StringVar()
        self.insert_source_pdf = tk.StringVar()
        self.insert_position = tk.IntVar(value=1)
        self.insert_output_file = tk.StringVar(value=str(self.USER_HOME_DIR / "PDF_com_Paginas_Inseridas.pdf"))
        self.insert_type = tk.StringVar(value="pdf")

        self.status_var = tk.StringVar(value="Pronto")

        self.setup_ui()

    # --- Configuração da Interface Gráfica (UI) ---

    def setup_ui(self):
        """Configura a estrutura principal da UI, incluindo o notebook com abas."""
        style = ttk.Style()
        style.configure("Accent.TButton", font=("Arial", 12, "bold"), padding=10)
        style.configure("TLabel", padding=2)
        style.configure("TButton", padding=5)
        style.configure("TFrame", padding=10)
        style.configure("TLabelframe", padding=10)

        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)

        tabs = {
            "Juntar PDFs": self.setup_join_tab,
            "Cortar PDF": self.setup_cut_tab,
            "Dividir PDF": self.setup_split_tab,
            "Reorganizar PDF": self.setup_reorder_tab,
            "Converter PDF": self.setup_convert_tab,
            "Inserir Páginas": self.setup_insert_tab,
        }

        for tab_name, setup_function in tabs.items():
            frame = ttk.Frame(notebook)
            notebook.add(frame, text=tab_name)
            setup_function(frame)
            
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=5)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def setup_join_tab(self, parent):
        title_label = ttk.Label(parent, text="Juntar PDFs", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        files_frame = ttk.LabelFrame(parent, text="Arquivos PDF")
        files_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        add_button = ttk.Button(files_frame, text="Adicionar PDFs", command=self.add_files_to_join)
        add_button.pack(side=tk.TOP, anchor=tk.W, pady=5, padx=5)
        self.join_listbox = self._create_listbox_with_controls(
            files_frame,
            up_command=lambda: self._move_listbox_item(self.join_listbox, self.pdf_files_to_join, -1),
            down_command=lambda: self._move_listbox_item(self.join_listbox, self.pdf_files_to_join, 1),
            remove_command=self.remove_file_from_join,
            clear_command=self.clear_join_files
        )
        merge_options_frame = ttk.LabelFrame(parent, text="Opções de Mesclagem")
        merge_options_frame.pack(fill=tk.X, pady=5)
        interleave_check = ttk.Checkbutton(merge_options_frame, text="Intercalar Páginas (ex: pág. 1 de cada, depois pág. 2 de cada, etc.)", variable=self.interleave_var)
        interleave_check.pack(anchor=tk.W, pady=2)
        self._create_output_frame(parent, "Arquivo de Saída Combinado", self.join_output_file, self.select_join_output_file)
        join_button = ttk.Button(parent, text="Juntar PDFs", command=self.join_pdfs, style="Accent.TButton")
        join_button.pack(pady=10)

    def setup_cut_tab(self, parent):
        title_label = ttk.Label(parent, text="Cortar PDF", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        self._create_file_input_frame(parent, "Arquivo PDF para Cortar", self.cut_pdf_file, self.select_pdf_to_cut)
        pages_frame = ttk.LabelFrame(parent, text="Páginas a Extrair")
        pages_frame.pack(fill=tk.X, pady=5)
        ttk.Label(pages_frame, text="Digite o intervalo de páginas (ex: 1-5, 7, 9-12)").pack(anchor=tk.W, pady=2)
        ttk.Entry(pages_frame, textvariable=self.cut_pages).pack(fill=tk.X, pady=5)
        self._create_output_frame(parent, "Arquivo de Saída Cortado", self.cut_output_file, self.select_cut_output_file)
        cut_button = ttk.Button(parent, text="Cortar PDF", command=self.cut_pdf, style="Accent.TButton")
        cut_button.pack(pady=10)

    def setup_split_tab(self, parent):
        title_label = ttk.Label(parent, text="Dividir PDF em Vários Arquivos", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        self._create_file_input_frame(parent, "Arquivo PDF para Dividir", self.split_pdf_file, self.select_pdf_to_split)
        definitions_frame = ttk.LabelFrame(parent, text="Definições de Divisão")
        definitions_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        columns = ('name', 'pages')
        self.split_treeview = ttk.Treeview(definitions_frame, columns=columns, show='headings', height=8)
        self.split_treeview.heading('name', text='Nome do Arquivo de Saída')
        self.split_treeview.heading('pages', text='Intervalo de Páginas')
        self.split_treeview.column('name', width=250)
        self.split_treeview.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        scrollbar = ttk.Scrollbar(definitions_frame, orient="vertical", command=self.split_treeview.yview)
        self.split_treeview.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        controls_frame = ttk.Frame(parent)
        controls_frame.pack(fill=tk.X, pady=5)
        ttk.Button(controls_frame, text="Adicionar...", command=self.add_split_definition).pack(side=tk.LEFT, padx=2)
        ttk.Button(controls_frame, text="Editar...", command=self.edit_split_definition).pack(side=tk.LEFT, padx=2)
        ttk.Button(controls_frame, text="Remover", command=self.remove_split_definition).pack(side=tk.LEFT, padx=2)
        
        import_frame = ttk.Frame(controls_frame)
        import_frame.pack(side=tk.LEFT, padx=10)
        ttk.Button(import_frame, text="Importar do Sumário", command=self.import_bookmarks_as_splits).pack(side=tk.LEFT)
        ttk.Label(import_frame, text="Importar até Nível (0=todos):").pack(side=tk.LEFT, padx=(5,2))
        ttk.Spinbox(import_frame, from_=0, to=10, width=4, textvariable=self.split_import_level).pack(side=tk.LEFT)

        self._create_folder_output_frame(parent, "Pasta de Saída", self.split_output_folder, self.select_split_output_folder)
        split_button = ttk.Button(parent, text="Dividir PDF", command=self.split_pdf, style="Accent.TButton")
        split_button.pack(pady=10)

    def setup_reorder_tab(self, parent):
        title_label = ttk.Label(parent, text="Reorganizar PDF", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        self._create_file_input_frame(parent, "Arquivo PDF para Reorganizar", self.reorder_pdf_file, self.select_pdf_to_reorder)
        pages_frame = ttk.LabelFrame(parent, text="Ordem das Páginas")
        pages_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        ttk.Label(pages_frame, text="Arraste as páginas para a posição desejada usando os botões.").pack(anchor=tk.W, pady=2)
        self.reorder_listbox = self._create_listbox_with_controls(
            pages_frame,
            up_command=lambda: self._move_listbox_item(self.reorder_listbox, self.reorder_page_indices, -1),
            down_command=lambda: self._move_listbox_item(self.reorder_listbox, self.reorder_page_indices, 1)
        )
        extra_actions_frame = ttk.Frame(pages_frame)
        extra_actions_frame.pack(fill=tk.X, pady=5)
        ttk.Button(extra_actions_frame, text="Inverter Ordem", command=self.reverse_pages).pack(side=tk.LEFT, padx=2)
        ttk.Button(extra_actions_frame, text="Separar Pares/Ímpares", command=self.even_odd_sort_pages).pack(side=tk.LEFT, padx=2)
        self._create_output_frame(parent, "Arquivo de Saída Reorganizado", self.reorder_output_file, self.select_reorder_output_file)
        reorder_button = ttk.Button(parent, text="Aplicar Reorganização", command=self.reorder_pdf, style="Accent.TButton")
        reorder_button.pack(pady=10)

    def setup_convert_tab(self, parent):
        title_label = ttk.Label(parent, text="Converter PDF", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        self._create_file_input_frame(parent, "Arquivo PDF para Converter", self.convert_pdf_file, self.select_pdf_to_convert)
        options_frame = ttk.LabelFrame(parent, text="Opções de Conversão")
        options_frame.pack(fill=tk.X, pady=5)
        format_frame = ttk.Frame(options_frame)
        format_frame.pack(fill=tk.X, pady=5)
        ttk.Label(format_frame, text="Formato:").pack(side=tk.LEFT, padx=(0, 10))
        for fmt in ['jpg', 'png', 'docx', 'html']:
            ttk.Radiobutton(format_frame, text=fmt.upper(), variable=self.convert_format, value=fmt).pack(side=tk.LEFT, padx=5)
        dpi_frame = ttk.Frame(options_frame)
        dpi_frame.pack(fill=tk.X, pady=5)
        ttk.Label(dpi_frame, text="Resolução (DPI para imagens):").pack(side=tk.LEFT, padx=(0, 10))
        for dpi in [72, 150, 300, 600]:
            ttk.Radiobutton(dpi_frame, text=str(dpi), variable=self.convert_dpi, value=dpi).pack(side=tk.LEFT, padx=5)
        self._create_folder_output_frame(parent, "Pasta de Saída", self.convert_output_folder, self.select_convert_output_folder)
        convert_button = ttk.Button(parent, text="Converter PDF", command=self.convert_pdf, style="Accent.TButton")
        convert_button.pack(pady=10)

    def setup_insert_tab(self, parent):
        title_label = ttk.Label(parent, text="Inserir Páginas", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        self._create_file_input_frame(parent, "PDF de Destino", self.insert_target_pdf, self.select_target_pdf)
        source_frame = ttk.LabelFrame(parent, text="Páginas a Inserir")
        source_frame.pack(fill=tk.X, pady=5)
        pdf_radio = ttk.Radiobutton(source_frame, text="Inserir de outro PDF", variable=self.insert_type, value="pdf", command=self.toggle_insert_source_frame)
        pdf_radio.pack(anchor=tk.W)
        self.pdf_source_frame = ttk.Frame(source_frame)
        self.pdf_source_frame.pack(fill=tk.X, padx=(20, 0))
        ttk.Entry(self.pdf_source_frame, textvariable=self.insert_source_pdf).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(self.pdf_source_frame, text="Selecionar...", command=self.select_source_pdf).pack(side=tk.RIGHT)
        blank_radio = ttk.Radiobutton(source_frame, text="Inserir página em branco", variable=self.insert_type, value="blank", command=self.toggle_insert_source_frame)
        blank_radio.pack(anchor=tk.W)
        position_frame = ttk.LabelFrame(parent, text="Posição de Inserção")
        position_frame.pack(fill=tk.X, pady=5)
        ttk.Label(position_frame, text="Inserir após a página:").pack(side=tk.LEFT, padx=(0, 10))
        self.position_spinbox = ttk.Spinbox(position_frame, from_=0, to=9999, textvariable=self.insert_position, width=5)
        self.position_spinbox.pack(side=tk.LEFT)
        ttk.Label(position_frame, text="(0 = início do documento)").pack(side=tk.LEFT, padx=10)
        self._create_output_frame(parent, "Arquivo de Saída com Páginas Inseridas", self.insert_output_file, self.select_insert_output_file)
        insert_button = ttk.Button(parent, text="Inserir Páginas", command=self.insert_pages, style="Accent.TButton")
        insert_button.pack(pady=10)
        self.toggle_insert_source_frame()

    # --- Métodos Auxiliares da UI (Reutilizáveis) ---

    def _create_file_input_frame(self, parent, label_text, text_variable, command):
        frame = ttk.LabelFrame(parent, text=label_text)
        frame.pack(fill=tk.X, pady=5)
        entry = ttk.Entry(frame, textvariable=text_variable)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        button = ttk.Button(frame, text="Selecionar Arquivo...", command=command)
        button.pack(side=tk.RIGHT)
        return frame

    def _create_output_frame(self, parent, label_text, text_variable, command):
        frame = ttk.LabelFrame(parent, text=label_text)
        frame.pack(fill=tk.X, pady=5)
        entry = ttk.Entry(frame, textvariable=text_variable)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        button = ttk.Button(frame, text="Salvar Como...", command=command)
        button.pack(side=tk.RIGHT)
        return frame
        
    def _create_folder_output_frame(self, parent, label_text, text_variable, command):
        frame = ttk.LabelFrame(parent, text=label_text)
        frame.pack(fill=tk.X, pady=5)
        entry = ttk.Entry(frame, textvariable=text_variable)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        button = ttk.Button(frame, text="Selecionar Pasta...", command=command)
        button.pack(side=tk.RIGHT)
        return frame

    def _create_listbox_with_controls(self, parent, up_command, down_command, remove_command=None, clear_command=None):
        list_frame = ttk.Frame(parent)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        listbox = tk.Listbox(list_frame, selectmode=tk.SINGLE, height=10)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=listbox.yview)
        listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        control_frame = ttk.Frame(parent)
        control_frame.pack(fill=tk.X)
        ttk.Button(control_frame, text="Mover para Cima", command=up_command).pack(side=tk.LEFT, padx=2)
        ttk.Button(control_frame, text="Mover para Baixo", command=down_command).pack(side=tk.LEFT, padx=2)
        if remove_command:
            ttk.Button(control_frame, text="Remover", command=remove_command).pack(side=tk.LEFT, padx=2)
        if clear_command:
            ttk.Button(control_frame, text="Limpar Tudo", command=clear_command).pack(side=tk.LEFT, padx=2)
        return listbox

    def _move_listbox_item(self, listbox, data_list, direction):
        selected_indices = listbox.curselection()
        if not selected_indices: return
        idx = selected_indices[0]
        new_idx = idx + direction
        if 0 <= new_idx < len(data_list):
            data_list[idx], data_list[new_idx] = data_list[new_idx], data_list[idx]
            item_text = listbox.get(idx)
            listbox.delete(idx)
            listbox.insert(new_idx, item_text)
            listbox.selection_set(new_idx)
            listbox.activate(new_idx)

    # --- Manipuladores de Eventos (Seleção de Arquivos) ---

    def _select_file(self, title, filetypes, variable_to_set, on_success=None):
        file = filedialog.askopenfilename(title=title, filetypes=filetypes)
        if file:
            variable_to_set.set(file)
            if on_success:
                on_success(file)
    
    def _select_save_file(self, title, filetypes, initial_file, variable_to_set):
        file = filedialog.asksaveasfilename(title=title, filetypes=filetypes, initialfile=initial_file, defaultextension=".pdf")
        if file:
            variable_to_set.set(file)

    def add_files_to_join(self):
        files = filedialog.askopenfilenames(title="Selecionar arquivos PDF", filetypes=[("Arquivos PDF", "*.pdf")])
        if files:
            for file in files:
                if file not in self.pdf_files_to_join:
                    self.pdf_files_to_join.append(file)
                    self.join_listbox.insert(tk.END, Path(file).name)
            self.status_var.set(f"{len(self.pdf_files_to_join)} arquivos selecionados para junção.")

    def select_pdf_to_cut(self):
        self._select_file("Selecionar PDF para Cortar", [("Arquivos PDF", "*.pdf")], self.cut_pdf_file, self._update_status_with_page_count)

    def select_pdf_to_split(self):
        self._select_file("Selecionar PDF para Dividir", [("Arquivos PDF", "*.pdf")], self.split_pdf_file, self._update_status_with_page_count)

    def select_pdf_to_reorder(self):
        self._select_file("Selecionar PDF para Reorganizar", [("Arquivos PDF", "*.pdf")], self.reorder_pdf_file, self.load_pages_for_reordering)

    def select_pdf_to_convert(self):
        self._select_file("Selecionar PDF para Converter", [("Arquivos PDF", "*.pdf")], self.convert_pdf_file, lambda f: self.status_var.set(f"Arquivo carregado: {Path(f).name}"))

    def select_target_pdf(self):
        self._select_file("Selecionar PDF de Destino", [("Arquivos PDF", "*.pdf")], self.insert_target_pdf, self._update_insert_spinbox)
        
    def select_source_pdf(self):
        self._select_file("Selecionar PDF de Origem", [("Arquivos PDF", "*.pdf")], self.insert_source_pdf, self._update_status_with_page_count)

    def select_join_output_file(self):
        self._select_save_file("Salvar PDF Combinado Como", [("Arquivos PDF", "*.pdf")], "PDFs_Combinados.pdf", self.join_output_file)

    def select_cut_output_file(self):
        self._select_save_file("Salvar PDF Cortado Como", [("Arquivos PDF", "*.pdf")], "PDF_Cortado.pdf", self.cut_output_file)

    def select_split_output_folder(self):
        folder = filedialog.askdirectory(title="Selecionar pasta para salvar arquivos divididos")
        if folder:
            self.split_output_folder.set(folder)

    def select_reorder_output_file(self):
        self._select_save_file("Salvar PDF Reorganizado Como", [("Arquivos PDF", "*.pdf")], "PDF_Reorganizado.pdf", self.reorder_output_file)
        
    def select_insert_output_file(self):
        self._select_save_file("Salvar PDF com Páginas Inseridas Como", [("Arquivos PDF", "*.pdf")], "PDF_com_Paginas_Inseridas.pdf", self.insert_output_file)

    def select_convert_output_folder(self):
        folder = filedialog.askdirectory(title="Selecionar pasta para salvar arquivos convertidos")
        if folder:
            self.convert_output_folder.set(folder)

    # --- Lógica de Manipulação de Listas e Treeview ---

    def remove_file_from_join(self):
        selected_indices = self.join_listbox.curselection()
        if not selected_indices: return
        idx = selected_indices[0]
        self.join_listbox.delete(idx)
        self.pdf_files_to_join.pop(idx)
        self.status_var.set(f"{len(self.pdf_files_to_join)} arquivos restantes.")

    def clear_join_files(self):
        self.join_listbox.delete(0, tk.END)
        self.pdf_files_to_join.clear()
        self.status_var.set("Lista de arquivos para junção limpa.")

    def add_split_definition(self):
        dialog = SplitDefinitionDialog(self.root, "Adicionar Nova Divisão")
        if dialog.result:
            if not dialog.result['name'] or not dialog.result['pages']:
                messagebox.showwarning("Entrada Inválida", "Nome do arquivo e páginas são obrigatórios.")
                return
            self.split_definitions.append(dialog.result)
            self._update_split_treeview()

    def edit_split_definition(self):
        selected_item = self.split_treeview.focus()
        if not selected_item:
            messagebox.showwarning("Nenhuma Seleção", "Por favor, selecione uma divisão para editar.")
            return
        item_index = self.split_treeview.index(selected_item)
        initial_values = self.split_definitions[item_index]
        dialog = SplitDefinitionDialog(self.root, "Editar Divisão", initial_values)
        if dialog.result:
            if not dialog.result['name'] or not dialog.result['pages']:
                messagebox.showwarning("Entrada Inválida", "Nome do arquivo e páginas são obrigatórios.")
                return
            self.split_definitions[item_index] = dialog.result
            self._update_split_treeview()

    def remove_split_definition(self):
        selected_item = self.split_treeview.focus()
        if not selected_item:
            messagebox.showwarning("Nenhuma Seleção", "Por favor, selecione uma divisão para remover.")
            return
        item_index = self.split_treeview.index(selected_item)
        del self.split_definitions[item_index]
        self._update_split_treeview()

    def _update_split_treeview(self):
        self.split_treeview.delete(*self.split_treeview.get_children())
        for definition in self.split_definitions:
            self.split_treeview.insert('', tk.END, values=(definition['name'], definition['pages']))

    def load_pages_for_reordering(self, filepath):
        self.reorder_listbox.delete(0, tk.END)
        self.reorder_page_indices.clear()
        try:
            pdf = PdfReader(filepath)
            num_pages = len(pdf.pages)
            for i in range(num_pages):
                self.reorder_listbox.insert(tk.END, f"Página {i + 1}")
                self.reorder_page_indices.append(i)
            self.status_var.set(f"{num_pages} páginas carregadas para reorganização.")
        except Exception as e:
            self._handle_error("Erro ao ler PDF", e)

    def reverse_pages(self):
        if not self.reorder_page_indices: return
        self.reorder_page_indices.reverse()
        self._update_listbox_from_data(self.reorder_listbox, [f"Página {i+1}" for i in self.reorder_page_indices])
        self.status_var.set("Ordem das páginas invertida.")

    def even_odd_sort_pages(self):
        if not self.reorder_page_indices: return
        odd_indices = self.reorder_page_indices[0::2]
        even_indices = self.reorder_page_indices[1::2]
        self.reorder_page_indices = odd_indices + even_indices
        self._update_listbox_from_data(self.reorder_listbox, [f"Página {i+1}" for i in self.reorder_page_indices])
        self.status_var.set("Páginas reorganizadas: ímpares, depois pares.")
        
    def toggle_insert_source_frame(self):
        state = "normal" if self.insert_type.get() == "pdf" else "disabled"
        for widget in self.pdf_source_frame.winfo_children():
            widget.configure(state=state)

    # --- Lógica Principal (Operações com PDF) ---

    def join_pdfs(self):
        if not self._validate_inputs(self.pdf_files_to_join, self.join_output_file): return
        output_path = self.join_output_file.get()
        try:
            if self.interleave_var.get():
                self._interleave_pdfs(output_path)
            else:
                self._simple_merge_pdfs(output_path)
            self._handle_success(f"PDFs unidos com sucesso em: {output_path}")
        except Exception as e:
            self._handle_error("Erro ao juntar PDFs", e)

    def _simple_merge_pdfs(self, output_path):
        merger = PdfMerger()
        for i, pdf_file in enumerate(self.pdf_files_to_join):
            self._update_progress(f"Processando arquivo {i+1}/{len(self.pdf_files_to_join)}", i, len(self.pdf_files_to_join))
            merger.append(pdf_file)
        merger.write(output_path)
        merger.close()

    def _interleave_pdfs(self, output_path):
        readers = [PdfReader(f) for f in self.pdf_files_to_join]
        writer = PdfWriter()
        max_pages = max(len(r.pages) for r in readers)
        for page_num in range(max_pages):
            self._update_progress(f"Intercalando página {page_num+1}/{max_pages}", page_num, max_pages)
            for reader in readers:
                if page_num < len(reader.pages):
                    writer.add_page(reader.pages[page_num])
        with open(output_path, "wb") as out_file:
            writer.write(out_file)

    def cut_pdf(self):
        if not self._validate_inputs(self.cut_pdf_file, self.cut_output_file, self.cut_pages): return
        input_path = self.cut_pdf_file.get()
        output_path = self.cut_output_file.get()
        try:
            reader = PdfReader(input_path)
            page_indices = self._parse_page_ranges(self.cut_pages.get(), len(reader.pages))
            if page_indices is None: return
            writer = PdfWriter()
            for i, page_index in enumerate(page_indices):
                self._update_progress(f"Extraindo página {i+1}/{len(page_indices)}", i, len(page_indices))
                writer.add_page(reader.pages[page_index])
            with open(output_path, "wb") as out_file:
                writer.write(out_file)
            self._handle_success(f"{len(page_indices)} páginas extraídas com sucesso para: {output_path}")
        except Exception as e:
            self._handle_error("Erro ao cortar PDF", e)

    def split_pdf(self):
        if not self._validate_inputs(self.split_pdf_file, self.split_output_folder, self.split_definitions): return
        input_path = self.split_pdf_file.get()
        output_folder = Path(self.split_output_folder.get())
        try:
            reader = PdfReader(input_path)
            max_pages = len(reader.pages)
            for i, definition in enumerate(self.split_definitions):
                name = definition['name']
                pages_str = definition['pages']
                self._update_progress(f"Processando '{name}' ({i+1}/{len(self.split_definitions)})", i, len(self.split_definitions))
                page_indices = self._parse_page_ranges(pages_str, max_pages)
                if page_indices is None:
                    self.status_var.set(f"Erro no intervalo '{pages_str}' para '{name}'. Pulando.")
                    continue
                writer = PdfWriter()
                for page_index in page_indices:
                    writer.add_page(reader.pages[page_index])
                output_path = output_folder / f"{name}.pdf"
                with open(output_path, "wb") as out_file:
                    writer.write(out_file)
            self._handle_success(f"{len(self.split_definitions)} arquivos criados com sucesso em: {output_folder}")
        except Exception as e:
            self._handle_error("Erro ao dividir o PDF", e)

    def reorder_pdf(self):
        if not self._validate_inputs(self.reorder_pdf_file, self.reorder_output_file, self.reorder_page_indices): return
        input_path = self.reorder_pdf_file.get()
        output_path = self.reorder_output_file.get()
        try:
            reader = PdfReader(input_path)
            writer = PdfWriter()
            for i, page_index in enumerate(self.reorder_page_indices):
                self._update_progress(f"Reorganizando página {i+1}/{len(self.reorder_page_indices)}", i, len(self.reorder_page_indices))
                writer.add_page(reader.pages[page_index])
            with open(output_path, "wb") as out_file:
                writer.write(out_file)
            self._handle_success(f"PDF reorganizado com sucesso em: {output_path}")
        except Exception as e:
            self._handle_error("Erro ao reorganizar PDF", e)

    def convert_pdf(self):
        if not self._validate_inputs(self.convert_pdf_file, self.convert_output_folder): return
        input_path = Path(self.convert_pdf_file.get())
        output_folder = Path(self.convert_output_folder.get())
        base_name = input_path.stem
        fmt = self.convert_format.get()
        try:
            doc = fitz.open(input_path)
            total_pages = len(doc)
            if fmt in ['jpg', 'png']:
                dpi = self.convert_dpi.get()
                matrix = fitz.Matrix(dpi / 72, dpi / 72)
                for i, page in enumerate(doc):
                    self._update_progress(f"Convertendo página {i+1}/{total_pages}", i, total_pages)
                    pix = page.get_pixmap(matrix=matrix)
                    pix.save(output_folder / f"{base_name}_pagina_{i+1}.{fmt}")
                self._handle_success(f"{total_pages} páginas convertidas para {fmt.upper()} em: {output_folder}")
            elif fmt == 'docx':
                word_doc = docx.Document()
                for i, page in enumerate(doc):
                    self._update_progress(f"Convertendo página {i+1}/{total_pages}", i, total_pages)
                    word_doc.add_paragraph(page.get_text())
                    if i < total_pages - 1:
                        word_doc.add_page_break()
                output_path = output_folder / f"{base_name}.docx"
                word_doc.save(output_path)
                self._handle_success(f"PDF convertido para DOCX em: {output_path}")
            elif fmt == 'html':
                output_path = output_folder / f"{base_name}.html"
                html_content = ["<html><head><title>{base_name}</title></head><body>"]
                for i, page in enumerate(doc):
                    self._update_progress(f"Convertendo página {i+1}/{total_pages}", i, total_pages)
                    html_content.append(f"<div><h2>Página {i+1}</h2><p>{page.get_text('html')}</p></div><hr>")
                html_content.append("</body></html>")
                with open(output_path, "w", encoding="utf-8") as f:
                    f.write("\n".join(html_content))
                self._handle_success(f"PDF convertido para HTML em: {output_path}")
        except Exception as e:
            self._handle_error("Erro durante a conversão", e)

    def insert_pages(self):
        if not self._validate_inputs(self.insert_target_pdf, self.insert_output_file): return
        target_path = self.insert_target_pdf.get()
        output_path = self.insert_output_file.get()
        position = self.insert_position.get()
        try:
            target_reader = PdfReader(target_path)
            writer = PdfWriter()
            if not (0 <= position <= len(target_reader.pages)):
                messagebox.showwarning("Aviso", f"Posição de inserção inválida. O PDF tem {len(target_reader.pages)} páginas.")
                return
            writer.append(fileobj=target_reader, pages=(0, position))
            if self.insert_type.get() == "pdf":
                source_path = self.insert_source_pdf.get()
                if not self._validate_inputs(source_path): return
                source_reader = PdfReader(source_path)
                writer.append(fileobj=source_reader)
                msg = f"{len(source_reader.pages)} páginas inseridas"
            else:
                writer.add_blank_page()
                msg = "Página em branco inserida"
            writer.append(fileobj=target_reader, pages=(position, len(target_reader.pages)))
            with open(output_path, "wb") as f:
                writer.write(f)
            self._handle_success(f"{msg} com sucesso em: {output_path}")
        except Exception as e:
            self._handle_error("Erro ao inserir páginas", e)
            
    # --- Funções Auxiliares e de Validação ---
    
    def import_bookmarks_as_splits(self):
        """Lê os bookmarks do PDF selecionado e preenche a lista de divisões."""
        if not self._validate_inputs(self.split_pdf_file): return
        
        try:
            input_path = self.split_pdf_file.get()
            reader = PdfReader(input_path)
            
            if not reader.outline:
                messagebox.showinfo("Sem Sumário", "Este PDF não contém um sumário (bookmarks) para importar.")
                return

            self.split_definitions.clear()
            total_pages = len(reader.pages)
            max_level = self.split_import_level.get()
            
            bookmarks = self._flatten_bookmarks(reader.outline, reader, max_level)
            
            if not bookmarks:
                 messagebox.showinfo("Sem Sumário", "Não foi possível extrair um sumário válido para o nível especificado.")
                 return

            for i, (title, start_page_num) in enumerate(bookmarks):
                if i + 1 < len(bookmarks):
                    end_page_num = bookmarks[i+1][1] - 1
                else:
                    end_page_num = total_pages
                
                if start_page_num > end_page_num:
                    end_page_num = start_page_num
                
                self.split_definitions.append({
                    'name': sanitize_filename(title),
                    'pages': f"{start_page_num}-{end_page_num}"
                })
            
            self._update_split_treeview()
            self.status_var.set(f"{len(self.split_definitions)} divisões importadas do sumário.")

        except Exception as e:
            self._handle_error("Erro ao importar sumário", e)

    def _flatten_bookmarks(self, outline_items, reader, max_level, level=1):
        """Função recursiva para achatar a lista de bookmarks até um nível máximo."""
        bookmarks = []
        if max_level > 0 and level > max_level:
            return []

        for item in outline_items:
            if isinstance(item, list):
                bookmarks.extend(self._flatten_bookmarks(item, reader, max_level, level + 1))
            elif isinstance(item, Destination):
                try:
                    # Método mais moderno e robusto para obter o número da página
                    page_num = reader.get_page_number(item.page) + 1
                except Exception:
                    # Fallback para versões mais antigas ou casos diferentes
                    page_num = reader.get_destination_page_number(item) + 1
                
                if page_num is not None:
                    bookmarks.append((item.title, page_num))
        return bookmarks

    def _update_status_with_page_count(self, filepath):
        try:
            num_pages = len(PdfReader(filepath).pages)
            self.status_var.set(f"Arquivo '{Path(filepath).name}' carregado ({num_pages} páginas).")
        except Exception as e:
            self._handle_error(f"Não foi possível ler o arquivo {Path(filepath).name}", e, show_messagebox=False)

    def _update_insert_spinbox(self, filepath):
        self._update_status_with_page_count(filepath)
        try:
            num_pages = len(PdfReader(filepath).pages)
            self.position_spinbox.config(to=num_pages)
        except Exception:
            pass

    def _update_listbox_from_data(self, listbox, data):
        listbox.delete(0, tk.END)
        for item in data:
            listbox.insert(tk.END, item)

    def _parse_page_ranges(self, pages_input, max_pages):
        indices = set()
        try:
            if not pages_input: raise ValueError("O intervalo de páginas não pode estar vazio.")
            parts = pages_input.replace(" ", "").split(',')
            for part in parts:
                if '-' in part:
                    start, end = map(int, part.split('-'))
                    if start < 1 or end > max_pages or start > end:
                        raise ValueError(f"Intervalo '{start}-{end}' está fora dos limites (1-{max_pages}).")
                    indices.update(range(start - 1, end))
                else:
                    page = int(part)
                    if not (1 <= page <= max_pages):
                        raise ValueError(f"Página '{page}' está fora dos limites (1-{max_pages}).")
                    indices.add(page - 1)
            return sorted(list(indices))
        except ValueError as e:
            messagebox.showerror("Erro de Entrada", f"Formato de página inválido: {e}")
            return None

    def _validate_inputs(self, *args):
        for arg in args:
            value = arg.get() if isinstance(arg, tk.Variable) else arg
            if not value:
                messagebox.showwarning("Entrada Inválida", "Por favor, preencha todos os campos obrigatórios.")
                return False
        return True

    def _update_progress(self, status_text, current, total):
        self.status_var.set(f"{status_text} ({current+1}/{total})")
        self.root.update_idletasks()

    def _handle_success(self, message):
        self.status_var.set("Concluído!")
        messagebox.showinfo("Sucesso", message)

    def _handle_error(self, message, error, show_messagebox=True):
        self.status_var.set(f"Erro: {message}")
        print(f"ERRO DETALHADO: {error}")
        if show_messagebox:
            messagebox.showerror("Erro", f"{message}:\n{error}")

if __name__ == "__main__":
    if 'PdfReader' in globals():
        root = tk.Tk()
        app = PDFManagerApp(root)
        root.mainloop()
    else:
        print("\nA aplicação não pode ser iniciada devido à falta de dependências.")
