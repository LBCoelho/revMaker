import sys
import subprocess
import importlib.util
import os
import unicodedata
import webbrowser



print("Iniciando... Verificando dependências necessárias.")

version = "v1.2"

required_packages = {
    'pywin32': 'win32com',
    'PyPDF2': 'PyPDF2',
    'FreeSimpleGUI': 'FreeSimpleGUI'
}

def check_and_install(): #Instalação de dependencias // Deixar ativo somente se for executar o programa via .bat ou pelo CMD. Caso for compilar o .exe, comentar a função de iniciação
    pacotes_instalados = 0
    for package_name, import_name in required_packages.items():
        spec = importlib.util.find_spec(import_name)
        if spec is None:
            print(f"Dependência '{package_name}' não encontrada. Instalando...")
            try:
                subprocess.run(
                    [sys.executable, "-m", "pip", "install", "--quiet", package_name],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL, 
                    check=True
                )
                print(f" -> '{package_name}' instalado com sucesso.")
                pacotes_instalados += 1
            except subprocess.CalledProcessError:
                print(f"ERRO: Falha ao instalar '{package_name}'.", file=sys.stderr)
                print("Por favor, execute o comando manualmente:", file=sys.stderr)
                print(f"pip install {package_name}", file=sys.stderr)
                sys.exit(1)


def normalizar_caminho(caminho):
    # Normaliza acentuação e converte para formato longo do Windows
    caminho_normalizado = unicodedata.normalize('NFC', caminho)
    if not caminho_normalizado.startswith('\\\\?\\'):
        caminho_normalizado = r'\\?\\' + caminho_normalizado.replace('/', '\\')
    return caminho_normalizado

    
    if pacotes_instalados > 0:
        print("-" * 50)
        print("Instalação de dependências concluída. Reinicie o script.")
        print("-" * 50)
        sys.exit()
    else:
        print("Dependências já estão em dia.")

#check_and_install()

print("Todas as dependências estão prontas. Iniciando o aplicativo...")
print("-" * 50)

import threading
import re
import shutil
from pathlib import Path
import FreeSimpleGUI as sg
import win32com.client
from PyPDF2 import PdfReader, PdfWriter

# CRIAR NOVA PASTA DE REVISÃO

def processar_revisao(diretorio_base_str, aux_files_list, status_callback):

    diretorio_base = Path(diretorio_base_str)
    status_callback(f"Analisando diretório: {diretorio_base}", 10) 
    
    padrao_rev = re.compile(r'^Rev\.(\d{1,2})$')
    pastas_rev_encontradas = []
    
    for item in diretorio_base.iterdir():
        if item.is_dir():
            match = padrao_rev.match(item.name)
            if match:
                numero_rev = int(match.group(1))
                pastas_rev_encontradas.append((numero_rev, item))

    if not pastas_rev_encontradas:
        raise Exception("Erro: Nenhuma pasta no formato 'Rev.X' foi encontrada. Verifique a formatação das pastas ou o diretorio")

    rev_atual_num, pasta_rev_atual_CAMINHO_ANTIGO = max(pastas_rev_encontradas, key=lambda item: item[0])
    status_callback(f"Revisão mais alta encontrada: {pasta_rev_atual_CAMINHO_ANTIGO.name}", 20)
    #
    pdfs_na_pasta_atual = list(pasta_rev_atual_CAMINHO_ANTIGO.glob('*.pdf'))

    if not pdfs_na_pasta_atual:
        raise Exception(f"Erro: Sem PDF da revisão na pasta {pasta_rev_atual_CAMINHO_ANTIGO.name}. Verifique se essa pasta já não é a revisão em andamento.")

    status_callback(f"PDF encontrado: {pdfs_na_pasta_atual[0].name}", 30)

    #pdf_origem_nome = pdfs_na_pasta_atual[0].name
    docx_origem_nomes = [p.name for p in pasta_rev_atual_CAMINHO_ANTIGO.glob('*.docx')]
    pasta_rev_atual_nome = pasta_rev_atual_CAMINHO_ANTIGO.name

    tag = "[Em revisão] "
    if not diretorio_base.name.startswith(tag):
        status_callback(f"Adicionando tag ao diretório: {diretorio_base.name}...")
        try:
            novo_nome = tag + diretorio_base.name
            novo_caminho_completo = diretorio_base.parent / novo_nome
            diretorio_base.rename(novo_caminho_completo)
            diretorio_base = novo_caminho_completo 
            status_callback(f"Diretório renomeado para: {novo_nome}", 35)
        except Exception as e:
            raise Exception(f"Erro ao tentar renomear pasta raiz: {e}. (A pasta está em uso?)")
    else:
        status_callback("Tag [Em revisão] já existe. Prosseguindo...", 35)

    pasta_rev_atual = diretorio_base / pasta_rev_atual_nome 
    #pdf_origem = pasta_rev_atual / pdf_origem_nome
    docx_origem_list = [pasta_rev_atual / nome for nome in docx_origem_nomes]

    status_callback("\nIniciando modificações: Criando nova revisão...")
    
    rev_proxima_num = rev_atual_num + 1
    pasta_rev_proxima = diretorio_base / f"Rev.{rev_proxima_num}"

    if pasta_rev_proxima.exists():
        raise Exception(f"Erro: A pasta {pasta_rev_proxima.name} já existe! Nenhuma modificação foi feita.")

    # --- INÍCIO DAS MODIFICAÇÕES ---
    
    pasta_rev_proxima.mkdir()
    status_callback(f"Pasta criada: {pasta_rev_proxima.name}", 40)

    subpastas_iniciais = ["01-Auxiliares", "02-E-mails", "03-Fotos e Videos"]
    for subpasta in subpastas_iniciais:
        (pasta_rev_proxima / subpasta).mkdir()
        status_callback(f"Subpasta criada: {subpasta}")
    status_callback("Subpastas iniciais criadas.", 50)
    
    if aux_files_list:
        status_callback(f"Anexando {len(aux_files_list)} arquivos auxiliares...")
        aux_path = pasta_rev_proxima / "01-Auxiliares"
    for file_path_str in aux_files_list:
        if file_path_str:
            long_path_str = normalizar_caminho(file_path_str)

            if os.path.exists(long_path_str):
                try:
                    file_path = Path(long_path_str)
                    shutil.copy2(file_path, aux_path)
                    status_callback(f"Copiado: {file_path.name}")
                except Exception as e:
                    status_callback(f"Aviso: Falha ao copiar {os.path.basename(file_path_str)}. Erro: {e}")
            else:
                status_callback(f"Aviso: Caminho do arquivo auxiliar '{file_path_str}' não encontrado. Pulando.")


    pastas_para_copiar = ["04-JRA", "05-Desenhos"]
    for pasta_nome in pastas_para_copiar:
        src = pasta_rev_atual / pasta_nome
        dst = pasta_rev_proxima / pasta_nome
        if src.is_dir():
            shutil.copytree(src, dst)
            status_callback(f"Pasta copiada (com conteúdo): {pasta_nome}")
        else:
            status_callback(f"Aviso: Pasta '{pasta_nome}' não encontrada em {pasta_rev_atual.name}, não foi copiada.")
    status_callback("Pastas de projeto copiadas.", 70)

    pasta_old = pasta_rev_proxima / "06-OLD"
    pasta_old.mkdir()
    status_callback(f"Pasta criada: 06-OLD")
    
    #shutil.copy2(pdf_origem, pasta_old)
    #status_callback(f"PDF copiado para 06-OLD: {pdf_origem.name}")

    docx_origem_para_renomear = None
    if docx_origem_list:
        docx_origem_para_renomear = docx_origem_list[0] 
        for docx_path in docx_origem_list:
            shutil.copy2(docx_path, pasta_old)
            status_callback(f"DOCX copiado para 06-OLD: {docx_path.name}")
    else:
        status_callback(f"Aviso: Nenhum arquivo .docx encontrado em {pasta_rev_atual.name} para copiar para 06-OLD.")
    status_callback("Arquivos da revisão anterior movidos para 06-OLD.", 85)

    if docx_origem_para_renomear:
        padrao_docx_rev = re.compile(r'(\d+)(\.docx)$', re.IGNORECASE)
        nome_original = docx_origem_para_renomear.name
        novo_nome_docx = padrao_docx_rev.sub(f'{rev_proxima_num}\\2', nome_original)
        if novo_nome_docx == nome_original:
            status_callback(f"Aviso: Não foi possível renomear o DOCX '{nome_original}'. Padrão não encontrado.")
            shutil.copy2(docx_origem_para_renomear, pasta_rev_proxima / nome_original)
        else:
            shutil.copy2(docx_origem_para_renomear, pasta_rev_proxima / novo_nome_docx)
            status_callback(f"DOCX copiado e renomeado: {novo_nome_docx}")
    else:
        status_callback(f"Aviso: Nenhum arquivo .docx encontrado para copiar para a raiz da nova revisão.")

    status_callback("\nProcesso de Criação da Nova Revisão CONCLUÍDO!", 100)

def revisao_worker_thread(window, diretorio_base_str, aux_files_str):
    try:
        def update_gui(message, progress=None):
            window.write_event_value('-THREAD_UPDATE-', {'message': message, 'progress': progress})
        
        update_gui("Iniciando processo...", 0)
        aux_files_list = aux_files_str.split(';') if aux_files_str else []
        processar_revisao(diretorio_base_str, aux_files_list, update_gui)
        window.write_event_value('-THREAD_DONE-', None)
    except Exception as e:
        window.write_event_value('-THREAD_ERROR-', str(e))

def create_gui_revisao():
    sg.theme('GrayGrayGray')
    input_column = [
        [sg.Text("1. Selecione o Diretório Raiz do documento:\nEx: C:\\Users\\User\\Digicorner\\08-OPR\\07-PP\CDA\\01-Emissão\\001-001 - Passagem cabo W1")],
        [sg.Input(key="-DIR-", enable_events=True), sg.FolderBrowse("Procurar", target="-DIR-")],
        [sg.Text(" (A pasta que contém as 'Rev.1', 'Rev.2', etc.)", font=("Helvetica", 9))],
        [sg.Text("")],
        [sg.HSeparator()],
        [sg.Text("")],
        [sg.Checkbox("Anexar documentos auxiliares na pasta 01-Auxiliares? (VCPs, Referencias etc.)", key="-AUX_CHECK-", enable_events=True)],
        [
            sg.Text("Arquivos Auxiliares:", key="-AUX_TEXT-"), 
            sg.Input(key="-AUX_FILES-", readonly=True, disabled=True, enable_events=True), 
            sg.FilesBrowse("Procurar",
                          file_types=(("All Files", "*.*"),), 
                          target="-AUX_FILES-", 
                          key="-AUX_BROWSE-", 
                          disabled=True)
        ],
        #[sg.Text("Arquivos com nomes extensos não serão importados. Reduza o nome caso necessario.", font=("Helvetica", 9),text_color=("red"),)]
    ]
    status_column = [
        [sg.Text("Status do Processo")],
        [sg.Multiline(size=(50, 15), key="-STATUS-", autoscroll=True, disabled=True, reroute_cprint=True)],
        [sg.ProgressBar(100, orientation='h', size=(35, 20), key='-PROGRESS-')],
        [sg.Button("Executar", key="-RUN-"),
         sg.Button("Limpar", key="-CLEAR-"),
         sg.Button("Sair", key="-SAIR_REVISAO-", button_color=('white', 'firebrick'))]
    ]
    layout = [[sg.Column(input_column), sg.VSeperator(), sg.Column(status_column, element_justification='center')]]

    window = sg.Window("Criador de Novas Revisões v1.2", layout)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == "-SAIR_REVISAO-":
            break
        if event == "-AUX_CHECK-":
            is_checked = values["-AUX_CHECK-"]
            window["-AUX_FILES-"].update(disabled=not is_checked)
            window["-AUX_BROWSE-"].update(disabled=not is_checked)
            if not is_checked:
                window["-AUX_FILES-"].update("") 
        if event in ("-DIR-", "-AUX_CHECK-", "-AUX_FILES-"):
            dir_filled = bool(values["-DIR-"])
            aux_checked = values["-AUX_CHECK-"]
            aux_filled = bool(values["-AUX_FILES-"])
            run_enabled = False
            if dir_filled and not aux_checked: 
                run_enabled = True
            elif dir_filled and aux_checked and aux_filled: 
                run_enabled = True
            window["-RUN-"].update(disabled=not run_enabled)
        if event == "-CLEAR-":
            window["-DIR-"].update("")
            window["-STATUS-"].update("")
            window["-PROGRESS-"].update(0)
            window["-RUN-"].update(disabled=True)
            window["-AUX_CHECK-"].update(False)
            window["-AUX_FILES-"].update("", disabled=True)
            window["-AUX_BROWSE-"].update(disabled=True)
        if event == "-RUN-":
            dir_path = values["-DIR-"]
            if not os.path.isdir(dir_path):
                sg.popup_error(f"Erro: O diretório não foi encontrado:\n{dir_path}", title="Diretório Inválido")
                continue
            window["-RUN-"].update(disabled=True)
            window["-STATUS-"].update("") 
            window["-PROGRESS-"].update(0)
            threading.Thread(
                target=revisao_worker_thread,
                args=(window, values["-DIR-"], values["-AUX_FILES-"]),
                daemon=True
            ).start()
        if event == '-THREAD_UPDATE-':
            sg.cprint(values[event]['message']) 
            if values[event]['progress'] is not None:
                window['-PROGRESS-'].update(values[event]['progress'])
        if event == '-THREAD_DONE-':
            sg.popup_ok("Processo Concluído!", "A nova estrutura de revisão foi criada com sucesso.")
        if event == '-THREAD_ERROR-':
            error_message = values[event]
            sg.cprint(f"ERRO: {error_message}", colors='white on red')
            sg.popup_error(f"Ocorreu um erro durante o processo:\n\n{error_message}")
            window["-RUN-"].update(disabled=False) 
            
    window.close()


# --- CRIAÇÃO DE PDF FINAL ---

def convert_word_to_pdf(input_word_path, status_callback):
    word_app = None
    output_pdf_path = os.path.splitext(input_word_path)[0] + "_convertido.pdf"
    try:
        status_callback(f"Iniciando conversão de: {os.path.basename(input_word_path)}...")
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        doc = word_app.Documents.Open(os.path.abspath(input_word_path), ReadOnly=False)
        doc.Activate()
        doc.SaveAs(os.path.abspath(output_pdf_path), FileFormat=17) # 17 = PDF format
        doc.Close(False)
        status_callback(f"Conversão concluída: {os.path.basename(output_pdf_path)}")
        return os.path.abspath(output_pdf_path)
    except Exception as e:
        raise Exception(f"Erro durante a conversão do Word: {e}")
    finally:
        if word_app:
            word_app.Quit()

def manipulate_pdfs(pdf_modify, pdf_insert, start_page_replace, final_output_path, status_callback):
    try:
        reader_modify = PdfReader(pdf_modify)
        reader_insert = PdfReader(pdf_insert)
        num_pages_total_modify = len(reader_modify.pages)
        num_pages_to_insert = len(reader_insert.pages)
        status_callback("-" * 40)
        status_callback(f"PDF a modificar: '{os.path.basename(pdf_modify)}' ({num_pages_total_modify} páginas).")
        status_callback(f"PDF a inserir: '{os.path.basename(pdf_insert)}' ({num_pages_to_insert} páginas).")
        status_callback("-" * 40)
        if not (1 <= start_page_replace and (start_page_replace + num_pages_to_insert - 1) <= num_pages_total_modify):
            raise ValueError(
                f"Erro de lógica: A substituição de {num_pages_to_insert} páginas a partir da página {start_page_replace} "
                f"excede o total de {num_pages_total_modify} páginas do documento."
            )
        start_index = start_page_replace
        end_index = start_index + num_pages_to_insert
        status_callback(f"OK. Substituindo {num_pages_to_insert} páginas, começando na página {start_page_replace}.")
        writer = PdfWriter()
        for i in range(start_index):
            writer.add_page(reader_modify.pages[i])
        for page in reader_insert.pages:
            writer.add_page(page)
        for i in range(end_index, num_pages_total_modify):
            writer.add_page(reader_modify.pages[i])
        status_callback("Bloco de páginas substituído com sucesso.")
        with open(final_output_path, "wb") as f:
            writer.write(f)
        status_callback("-" * 40)
        status_callback(f"PROCESSO CONCLUÍDO!")
        status_callback(f"Arquivo final salvo como: {os.path.basename(final_output_path)}")
    except Exception as e:
        raise Exception(f"Erro during PDF manipulation: {e}")
    finally:
        try:
            os.remove(pdf_modify)
            status_callback(f"Arquivo intermediário '{os.path.basename(pdf_modify)}' removido.")
        except OSError as e:
            status_callback(f"Aviso: Não foi possível remover o arquivo intermediário: {e}")

def pdf_worker_thread(window, docx_file, pdf_file, start_page, output_file):
    try:
        def update_gui(message, progress=None):
            window.write_event_value('-THREAD_UPDATE-', {'message': message, 'progress': progress})
        
        update_gui("Iniciando processo...", 0)
        converted_pdf = convert_word_to_pdf(docx_file, lambda msg: update_gui(msg, 25))
        manipulate_pdfs(converted_pdf, pdf_file, start_page, output_file, lambda msg: update_gui(msg, 75))
        update_gui("Processo finalizado com sucesso!", 100)
        window.write_event_value('-THREAD_DONE-', None)
    except Exception as e:
        window.write_event_value('-THREAD_ERROR-', str(e))

def create_gui_pdf():
    sg.theme('GrayGrayGray')
    input_column = [
        [sg.Text("1. Arquivo Word Principal (.docx)")],
        [sg.Input(key="-DOCX-", readonly=True, enable_events=True), sg.FileBrowse("Procurar", file_types=(("Word Files", "*.docx *.doc"),))],
        [sg.Text("2. PDF com Desenhos para Inserir (.pdf)")],
        [sg.Input(key="-PDF-", readonly=True, enable_events=True), sg.FileBrowse("Procurar", file_types=(("PDF Files", "*.pdf"),))],
        [sg.Text("", size=(40,1), key="-PDF_INFO-", font=("Helvetica", 9))],
        [sg.Text("3. Iniciar Substituição na Página:")],
        [sg.Input("1", key="-START_PAGE-", size=(10, 1), enable_events=True)],
        [sg.Text("4. Salvar Arquivo Final Como...")],
        [sg.Input(key="-OUTPUT-", readonly=True, enable_events=True), sg.FileSaveAs("Salvar como...", file_types=(("PDF Files", "*.pdf"),))],
    ]
    status_column = [
        [sg.Text("Status do Processo")],
        [sg.Multiline(size=(50, 15), key="-STATUS-", autoscroll=True, disabled=True, reroute_cprint=True)],
        [sg.ProgressBar(100, orientation='h', size=(35, 20), key='-PROGRESS-')],
        [sg.Button("Executar", key="-RUN-"),
         sg.Button("Limpar", key="-CLEAR-"),
         sg.Button("Sair", key="-SAIR_PDF-", button_color=('white', 'firebrick'))]
    ]
    layout = [[sg.Column(input_column), sg.VSeperator(), sg.Column(status_column, element_justification='center')]]

    window = sg.Window("PDF Automático v2.0", layout)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == "-SAIR_PDF-":
            break
        if event in ("-DOCX-", "-PDF-", "-OUTPUT-", "-START_PAGE-"):
            all_filled = all(values[k] for k in ["-DOCX-", "-PDF-", "-OUTPUT-", "-START_PAGE-"])
            page_valid = values["-START_PAGE-"].isdigit() and int(values["-START_PAGE-"]) > 0
            window["-RUN-"].update(disabled=not (all_filled and page_valid))
        if event == "-PDF-" and values["-PDF-"]:
            try:
                page_count = len(PdfReader(values["-PDF-"]).pages)
                window["-PDF_INFO-"].update(f"-> PDF de {page_count} páginas selecionado.", text_color="cyan")
            except Exception:
                window["-PDF_INFO-"].update("-> Arquivo PDF inválido ou corrompido.", text_color="red")
        if event == "-CLEAR-":
            for key in ["-DOCX-", "-PDF-", "-OUTPUT-", "-START_PAGE-", "-PDF_INFO-", "-STATUS-"]:
                window[key].update("")
            window["-START_PAGE-"].update("1")
            window["-PROGRESS-"].update(0)
            window["-RUN-"].update(disabled=True)
        if event == "-RUN-":
            window["-RUN-"].update(disabled=True)
            window["-STATUS-"].update("")
            window["-PROGRESS-"].update(0)
            threading.Thread(
                target=pdf_worker_thread,
                args=(window, values["-DOCX-"], values["-PDF-"], int(values["-START_PAGE-"]), values["-OUTPUT-"]),
                daemon=True
            ).start()
        if event == '-THREAD_UPDATE-':
            sg.cprint(values[event]['message'])
            if values[event]['progress'] is not None:
                window['-PROGRESS-'].update(values[event]['progress'])
        if event == '-THREAD_DONE-':
            sg.popup_ok("Processo Concluído!", "O seu arquivo PDF foi gerado com sucesso.")
        if event == '-THREAD_ERROR-':
            sg.cprint(f"ERRO: {values[event]}", colors='white on red')
            sg.popup_error(f"Ocorreu um erro durante o processo:\n\n{values[event]}")
            window["-RUN-"].update(disabled=False)
            
    window.close()


# ---MENU PRINCIPAL---

def create_main_menu():
    sg.theme('GrayGrayGray')
    
    layout = [
        [sg.Text("RevMaker", font=("Helvetica", 16)),sg.Text(version, font=("Helvetica", 10),text_color=("grey"))],
        [sg.Text("Selecione a ferramenta desejada:")],
        [sg.Button("Criar nova pasta de revisão", key="-REVISAO-", mouseover_colors=("grey"),size=(40, 2),tooltip="Ferramenta para criar pastas de revisão.\nO que ela vai fazer:\n-Verificar se a pasta raiz (pasta do PP) está com a tag [Em revisão] e adicionar tag caso ausencia.\n-Criar pasta de revisão baseado na ultima pasta\n-Criar pastas internas\n-Salvar arquivos auxiliares na pasta 01-Auxiliares\n-Mover a pasta desenhos\n-Mover o arquivo .docx da rev antiga para a pasta 06-OLD\n-Criar novo .docx renomeado para a nova versão")],
        [sg.Button("Criação de PDF final (PP + Desenho)", key="-PDF-", size=(40, 2),mouseover_colors=("grey"),tooltip="Ferramenta para unificar .docx com os desenhos, gerando o PDF final.")],
        [sg.Button("GitHub", key = "-GIT-", size=(5, 1), button_color=('white', 'purple'))],
        [sg.Text("Fernando Carmo & Lucas Coelho\n       OperationsLTC@2025",font=("Helvetica", 7))]
        ]
    
    window = sg.Window("RevMaker Version 1.2 ", layout, element_justification='c')
    
    while True:
        event, values = window.read()
        
        if event == sg.WIN_CLOSED or event == "Sair":
            break
            
        if event == "-REVISAO-":
            window.hide()
            create_gui_revisao()
            window.un_hide()
            
        if event == "-PDF-":
            window.hide()
            create_gui_pdf()
            window.un_hide()
        
        if event == "-GIT-":
            webbrowser.open("https://github.com/LBCoelho/revMaker")
            
    window.close()


# ---ENTRADA---

if __name__ == "__main__":
    create_main_menu()