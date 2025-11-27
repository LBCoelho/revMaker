
import sys
import subprocess
import importlib.util
import os
import unicodedata
import webbrowser

print("Iniciando... Verificando dependências necessárias.")
version = "1.4.1"
required_packages = {
    'pywin32': 'win32com',
    'pypdf': 'pypdf',
    'FreeSimpleGUI': 'FreeSimpleGUI',
    'requests' : 'requests'
}

def verificar_ultima_versao(versao_atual):
  
    #Verifica se a versão atual é a última release do repositório GitHub.
  
    url = "https://api.github.com/repos/LBCoelho/revMaker/releases/latest"
    resposta = requests.get(url)
    
    if resposta.status_code == 200:
        dados = resposta.json()
        ultima_release = dados.get("tag_name")
        return versao_atual == ultima_release, ultima_release
    else:
        raise Exception(f"Erro ao acessar API do GitHub: {resposta.status_code}")

def check_and_install():
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
    if pacotes_instalados > 0:
        print("-" * 50)
        print("Instalação de dependências concluída. Reinicie o script.")
        print("-" * 50)
        sys.exit()
    else:
        print("Dependências já estão em dia.")

def normalizar_caminho(caminho):
    caminho_abs = os.path.abspath(caminho.replace('/', '\\'))
    caminho_normalizado = unicodedata.normalize('NFC', caminho_abs)
    if not caminho_normalizado.startswith('\\\\?\\'):
        caminho_normalizado = r'\\?\{}'.format(caminho_normalizado)
    return caminho_normalizado

check_and_install()
print("Todas as dependências estão prontas. Iniciando o aplicativo...")
print("-" * 50)

import threading
import re
import shutil
import requests
from pathlib import Path
import FreeSimpleGUI as sg
import win32com.client
from pypdf import PdfReader, PdfWriter

# --- PARTE 1: CRIAR NOVA PASTA DE REVISÃO (MODIFICADO) ---
def processar_revisao(diretorio_base_str, aux_files_list, copy_drawings_flag, status_callback, continuar_sem_pdf=False):
    diretorio_base = Path(normalizar_caminho(diretorio_base_str))
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
        raise Exception("Erro: Nenhuma pasta no formato 'Rev.X' foi encontrada. Verifique a formatação das pastas ou o diretório.")
    rev_atual_num, pasta_rev_atual_CAMINHO_ANTIGO = max(pastas_rev_encontradas, key=lambda item: item[0])
    status_callback(f"Revisão mais alta encontrada: {pasta_rev_atual_CAMINHO_ANTIGO.name}", 20)
    pdfs_na_pasta_atual = list(pasta_rev_atual_CAMINHO_ANTIGO.glob('*.pdf'))

    # --- Popup seguro via thread principal ---
    if not pdfs_na_pasta_atual and not continuar_sem_pdf:
        return {"acao": "confirmar_sem_pdf"}
    elif not pdfs_na_pasta_atual and continuar_sem_pdf:
        status_callback("Continuando sem PDF da revisão anterior.", 25)
    else:
        status_callback(f"PDF encontrado: {pdfs_na_pasta_atual[0].name}", 30)

    # Continua criação da nova revisão
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
    docx_origem_list = [pasta_rev_atual / nome for nome in docx_origem_nomes]
    status_callback("\nIniciando modificações: Criando nova revisão...")
    rev_proxima_num = rev_atual_num + 1
    pasta_rev_proxima = diretorio_base / f"Rev.{rev_proxima_num}"
    if pasta_rev_proxima.exists():
        raise Exception(f"Erro: A pasta {pasta_rev_proxima.name} já existe! Nenhuma modificação foi feita.")
    pasta_rev_proxima.mkdir()
    status_callback(f"Pasta criada: {pasta_rev_proxima.name}", 40)

    # Criação das subpastas
    subpastas_iniciais = ["01-Auxiliares", "02-E-mails", "03-Fotos e Videos"]
    for subpasta in subpastas_iniciais:
        (pasta_rev_proxima / subpasta).mkdir()
        status_callback(f"Subpasta criada: {subpasta}")
    status_callback("Subpastas iniciais criadas.", 50)

    # Copiar arquivos auxiliares
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

    # Copiar pastas adicionais
    pasta_jra = "04-JRA"
    src_jra = pasta_rev_atual / pasta_jra
    dst_jra = pasta_rev_proxima / pasta_jra
    if src_jra.is_dir():
        shutil.copytree(src_jra, dst_jra)
        status_callback(f"Pasta copiada (com conteúdo): {pasta_jra}")
    else:
        status_callback(f"Aviso: Pasta '{pasta_jra}' não encontrada em {pasta_rev_atual.name}, não foi copiada.")

    # Pasta desenhos
    pasta_desenhos = "05-Desenhos"
    src_desenhos = pasta_rev_atual / pasta_desenhos
    dst_desenhos = pasta_rev_proxima / pasta_desenhos
    dst_desenhos.mkdir()
    status_callback(f"Pasta criada: {pasta_desenhos}")
    if copy_drawings_flag:
        status_callback("Copiando desenhos da revisão anterior para a pasta OLD...")
        dst_desenhos_old = dst_desenhos / "OLD"
        dst_desenhos_old.mkdir()
        if src_desenhos.is_dir():
            files_copied_count = 0
            for item in src_desenhos.iterdir():
                if item.is_file():
                    try:
                        shutil.copy2(item, dst_desenhos_old)
                        files_copied_count += 1
                    except Exception as e:
                        status_callback(f"Aviso: Falha ao copiar desenho '{item.name}'. Erro: {e}")
            if files_copied_count > 0:
                status_callback(f" -> {files_copied_count} arquivos de desenho copiados para {dst_desenhos_old.name}")
            else:
                status_callback(f" -> Nenhum arquivo encontrado em {src_desenhos.name} para copiar.")
        else:
            status_callback(f"Aviso: Pasta de origem '{src_desenhos.name}' não encontrada, nada foi copiado.")
    else:
        status_callback("Pasta '05-Desenhos' criada (vazia, conforme solicitado).")

    # Pasta OLD
    pasta_old = pasta_rev_proxima / "06-OLD"
    pasta_old.mkdir()
    status_callback(f"Pasta criada: 06-OLD")
    docx_origem_para_renomear = None
    if docx_origem_list:
        docx_origem_para_renomear = docx_origem_list[0]
        for docx_path in docx_origem_list:
            shutil.copy2(docx_path, pasta_old)
            status_callback(f"DOCX copiado para 06-OLD: {docx_path.name}")
    else:
        status_callback(f"Aviso: Nenhum arquivo .docx encontrado em {pasta_rev_atual.name} para copiar para 06-OLD.")
    status_callback("Arquivos da revisão anterior movidos para 06-OLD.", 85)

    # Renomear DOCX
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
    os.startfile(pasta_rev_proxima)
    return str(pasta_rev_proxima).replace(r'\\?\\', '')

def revisao_worker_thread(window, diretorio_base_str, aux_files_str, copy_drawings_flag, continuar_sem_pdf=False):
    try:
        def update_gui(message, progress=None):
            window.write_event_value('-THREAD_UPDATE-', {'message': message, 'progress': progress})
        update_gui("Iniciando processo...", 0)
        aux_files_list = aux_files_str.split(';') if aux_files_str else []
        folder_path = processar_revisao(diretorio_base_str, aux_files_list, copy_drawings_flag, update_gui, continuar_sem_pdf)
        if isinstance(folder_path, dict) and folder_path.get("acao") == "confirmar_sem_pdf":
            window.write_event_value('-THREAD_CONFIRM-', None)
            return
        window.write_event_value('-THREAD_DONE-', folder_path)
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
        [sg.Checkbox("Copiar desenhos da revisão anterior? (Move para pasta OLD)", key="-COPY_DRAWINGS-", default=True)],
        [sg.Text("")],
        [sg.Checkbox("Anexar documentos auxiliares na pasta 01-Auxiliares? (VCPs, Referencias etc.)", key="-AUX_CHECK-", enable_events=True)],
        [sg.Input(key="-AUX_FILES-", readonly=True, disabled=True), sg.FilesBrowse("Procurar", key="-AUX_BROWSE-", disabled=True)]
    ]
    status_column = [
        [sg.Text("Status do Processo")],
        [sg.Multiline(size=(50, 15), key="-STATUS-", autoscroll=True, disabled=True, reroute_cprint=True)],
        [sg.ProgressBar(100, orientation='h', size=(35, 20), key='-PROGRESS-')],
        [sg.Button("Executar", key="-RUN-", disabled=True),
         sg.Button("Limpar", key="-CLEAR-"),
         sg.Button("Ajuda", key="-AJUDA_REVISAO-"),
         sg.Button("Sair", key="-SAIR_REVISAO-", button_color=('white', 'firebrick'))],
    ]
    layout = [[sg.Column(input_column), sg.VSeperator(), sg.Column(status_column)]]
    window = sg.Window(f"Criador de Novas Revisões v{version}", layout)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == "-SAIR_REVISAO-":
            break

        if event == "-AJUDA_REVISAO-":
            help_text = (
                "Passo a passo para Criar Nova Revisão:\n\n"
                "1. No campo '1. Selecione o Diretório Raiz...', clique em 'Procurar' e selecione a pasta principal do documento a ser revisado (a pasta que contém suas 'Rev.1', 'Rev.2', etc.).\n\n"
                "2. (Opcional) Marque a caixa 'Copiar desenhos...' para copiar os arquivos da '05-Desenhos' anterior para uma pasta 'OLD' na nova revisão. Se desmarcado, a pasta '05-Desenhos' será criada vazia.\n\n"
                "3. (Opcional) Marque a caixa 'Anexar documentos auxiliares...' para adicionar arquivos (VCPs, referências) à pasta '01-Auxiliares'.\n\n"
                "4. Se a caixa de 'Anexar' for marcada, clique em 'Procurar' ao lado de 'Arquivos Auxiliares' e selecione um ou mais arquivos.\n\n"
                "5. O botão 'Executar' será habilitado quando o diretório principal (Passo 1) for preenchido (e os arquivos auxiliares, se a caixa estiver marcada).\n\n"
                "6. Clique em 'Executar' e aguarde o processo terminar. A nova pasta será aberta automaticamente no final."
            )
            sg.popup_ok(help_text, title="Ajuda - Criador de Revisões")

        if event == "-AUX_CHECK-":
            is_checked = values["-AUX_CHECK-"]
            window["-AUX_FILES-"].update(disabled=not is_checked)
            window["-AUX_BROWSE-"].update(disabled=not is_checked)

        if event in ("-DIR-"):
            dir_ok = bool(values["-DIR-"])
            aux_ok = not values["-AUX_CHECK-"] or bool(values["-AUX_FILES-"])
            window["-RUN-"].update(disabled=not (dir_ok and aux_ok))

        if event == "-RUN-":
            threading.Thread(
                target=revisao_worker_thread,
                args=(window, values["-DIR-"], values["-AUX_FILES-"], values["-COPY_DRAWINGS-"]),
                daemon=True
            ).start()

        if event == "-CLEAR-":
            window["-DIR-"].update("")
            window["-STATUS-"].update("")
            window["-PROGRESS-"].update(0)
            window["-RUN-"].update(disabled=True)
            window["-AUX_CHECK-"].update(False)
            window["-AUX_FILES-"].update("", disabled=True)
            window["-AUX_BROWSE-"].update(disabled=True)
            window["-COPY_DRAWINGS-"].update(True)

        if event == '-THREAD_UPDATE-':
            sg.cprint(values[event]['message'])
            if values[event]['progress'] is not None:
                window['-PROGRESS-'].update(values[event]['progress'])

        if event == '-THREAD_CONFIRM-':
            resposta = sg.popup_yes_no("PDF da revisão anterior não foi encontrado, deseja continuar?", title="Aviso")
            if resposta == "Yes":
                threading.Thread(
                    target=revisao_worker_thread,
                    args=(window, values["-DIR-"], values["-AUX_FILES-"], values["-COPY_DRAWINGS-"], True),
                    daemon=True
                ).start()
            else:
                sg.popup_ok("Processo cancelado pelo usuário.")

        if event == '-THREAD_DONE-':
            sg.popup_ok("Processo Concluído!", "Nova revisão criada com sucesso.")
        if event == '-THREAD_ERROR-':
            sg.popup_error(f"Erro: {values[event]}")

    window.close()

# --- PARTE 2: CRIAÇÃO DE PDF FINAL (MODIFICADO) ---

def convert_word_to_pdf(input_word_path, status_callback):
    word_app = None
    doc = None
    wdDoNotSaveChanges = 0
    input_word_path_abs = os.path.abspath(input_word_path)
    output_pdf_path = os.path.splitext(input_word_path_abs)[0] + "_convertido.pdf"
    try:
        status_callback(f"Iniciando conversão de: {os.path.basename(input_word_path_abs)}...")
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = True
        word_app.DisplayAlerts = False
        status_callback("Abrindo documento (em modo de bypass)...")
        doc = word_app.Documents.Open(
            input_word_path_abs,
            ReadOnly=False,
            ConfirmConversions=False
        )
        word_app.Visible = False
        if doc is None:
            raise Exception("Falha ao abrir o documento. Verifique se o arquivo não está corrompido.")
        doc.Activate()
        status_callback("Salvando como PDF...")
        doc.SaveAs(output_pdf_path, FileFormat=17)
        doc.Close(wdDoNotSaveChanges)
        doc = None
        status_callback(f"Conversão concluída: {os.path.basename(output_pdf_path)}")
        word_app.Quit()
        word_app = None
        return output_pdf_path
    except Exception as e:
        if doc:
            doc.Close(wdDoNotSaveChanges)
        if word_app:
            word_app.Quit()
        raise Exception(f"Erro durante a conversão do Word: {e}")
    finally:
        doc = None
        word_app = None

def manipulate_pdfs(pdf_modify, pdf_insert, start_page_replace, final_output_path, status_callback):
    try:
        reader_modify = PdfReader(normalizar_caminho(pdf_modify))
        reader_insert = PdfReader(normalizar_caminho(pdf_insert))
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
        with open(normalizar_caminho(final_output_path), "wb") as f:
            writer.write(f)
        status_callback("-" * 40)
        status_callback(f"PROCESSO CONCLUÍDO!")
        status_callback(f"Arquivo final salvo como: {os.path.basename(final_output_path)}")
    except Exception as e:
        raise Exception(f"Erro during PDF manipulation: {e}")
    finally:
        try:
            os.remove(normalizar_caminho(pdf_modify))
            status_callback(f"Arquivo intermediário '{os.path.basename(pdf_modify)}' removido.")
        except OSError as e:
            status_callback(f"Aviso: Não foi possível remover o arquivo intermediário: {e}")

# MODIFICAÇÃO: agora recebe add_pdf (bool) e só manipula PDFs se for True
def pdf_worker_thread(window, docx_file, add_pdf, pdf_file, start_page, output_file):
    try:
        def update_gui(message, progress=None):
            window.write_event_value('-THREAD_UPDATE-', {'message': message, 'progress': progress})
        update_gui("Iniciando processo...", 0)
        converted_pdf = convert_word_to_pdf(docx_file, lambda msg: update_gui(msg, 25))
        if add_pdf and pdf_file:
            manipulate_pdfs(converted_pdf, pdf_file, start_page, output_file, lambda msg: update_gui(msg, 75))
        else:
            shutil.move(converted_pdf, output_file)
            update_gui("PDF gerado apenas a partir do Word (sem desenhos).")
        update_gui("Processo finalizado com sucesso!", 100)
        window.write_event_value('-THREAD_DONE-', None)
    except Exception as e:
        window.write_event_value('-THREAD_ERROR-', str(e))
    os.startfile(output_file)

def create_gui_pdf():
    sg.theme('GrayGrayGray')
    input_column = [
        [sg.Text("1. Arquivo Word Principal (.docx)")],
        [sg.Input(key="-DOCX-", readonly=True, enable_events=True), sg.FileBrowse("Procurar", file_types=(("Word Files", "*.docx *.doc"),))],
        [sg.Checkbox("Adicionar PDF de desenho?", key="-ADD_DRAWING_PDF-", default=True, enable_events=True)],
        [sg.Text("2. PDF com Desenhos para Inserir (.pdf)", key="-PDF_LABEL-")],
        [sg.Input(key="-PDF-", readonly=True, enable_events=True), sg.FileBrowse("Procurar", file_types=(("PDF Files", "*.pdf"),), key="-PDF_BROWSE-")],
        [sg.Text("", size=(40,1), key="-PDF_INFO-", font=("Helvetica", 9))],
        [sg.Text("3. Iniciar Substituição a partir de qual Página?")],
        [sg.Input("1", key="-START_PAGE-", size=(10, 1), enable_events=True)],
        [sg.HSeparator()],
        [sg.Checkbox("Salvar em local específico?", key="-CUSTOM_OUTPUT_CHECK-", default=False, enable_events=True)],
        [
            sg.Text("4. Salvar Arquivo Final Como:"),
            sg.Input(key="-OUTPUT-", readonly=True, enable_events=True, disabled=True),
            sg.FileSaveAs("Salvar como...", file_types=(("PDF Files", "*.pdf"),), disabled=True, key="-OUTPUT_BROWSE-")
        ]
    ]
    status_column = [
        [sg.Text("Status do Processo")],
        [sg.Multiline(size=(50, 15), key="-STATUS-", autoscroll=True, disabled=True, reroute_cprint=True)],
        [sg.ProgressBar(100, orientation='h', size=(35, 20), key='-PROGRESS-')],
        [sg.Button("Executar", key="-RUN-", disabled=True),
         sg.Button("Limpar", key="-CLEAR-"),
         sg.Button("Ajuda", key="-AJUDA_PDF-"),
         sg.Button("Sair", key="-SAIR_PDF-", button_color=('white', 'firebrick'))]
    ]
    layout = [[sg.Column(input_column), sg.VSeperator(), sg.Column(status_column, element_justification='center')]]
    window = sg.Window("PDF Automático v2.3", layout)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == "-SAIR_PDF-":
            break

        if event == "-AJUDA_PDF-":
            help_text = (
                "Passo a passo para Criar PDF Final:\n\n"
                "1. Em '1. Arquivo Word...', procure e selecione seu arquivo `.docx` principal.\n\n"
                "2. (Opcional) Marque ou desmarque 'Adicionar PDF de desenho?'.\n"
                "   - Se marcado, selecione o PDF de desenhos normalmente.\n"
                "   - Se desmarcado, o PDF será gerado apenas a partir do Word.\n\n"
                "3. Em '3. Iniciar Substituição...', digite o número da página do seu Word a partir de onde o PDF de desenhos deve começar. (Ex: '1' para começar na página 2).\n\n"
                "4. (Opcional) Marque a caixa 'Salvar em local específico?' se você quiser escolher um nome e local diferentes para o arquivo final.\n\n"
                "5. O botão 'Executar' será habilitado quando todos os campos obrigatórios estiverem preenchidos. Clique nele para iniciar."
            )
            sg.popup_ok(help_text, title="Ajuda - Criador de PDF")

        if event == "-ADD_DRAWING_PDF-":
            add_pdf = values["-ADD_DRAWING_PDF-"]
            window["-PDF-"].update(disabled=not add_pdf)
            window["-PDF_BROWSE-"].update(disabled=not add_pdf)
            window["-PDF_LABEL-"].update(text_color="black" if add_pdf else "grey")
            if not add_pdf:
                window["-PDF-"].update("")
                window["-PDF_INFO-"].update("")

        if event in ("-DOCX-", "-PDF-", "-OUTPUT-", "-START_PAGE-", "-CUSTOM_OUTPUT_CHECK-", "-ADD_DRAWING_PDF-"):
            docx_filled = bool(values["-DOCX-"])
            add_pdf = values["-ADD_DRAWING_PDF-"]
            pdf_filled = bool(values["-PDF-"]) if add_pdf else True
            custom_output = values["-CUSTOM_OUTPUT_CHECK-"]
            output_filled = bool(values["-OUTPUT-"]) if custom_output else True
            run_enabled = docx_filled and pdf_filled and output_filled
            window["-RUN-"].update(disabled=not run_enabled)

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
            window["-CUSTOM_OUTPUT_CHECK-"].update(False)
            window["-OUTPUT-"].update("", disabled=True)
            window["-OUTPUT_BROWSE-"].update(disabled=True)
            window["-ADD_DRAWING_PDF-"].update(True)
            window["-PDF-"].update(disabled=False)
            window["-PDF_BROWSE-"].update(disabled=False)
            window["-PDF_LABEL-"].update(text_color="black")

        if event == "-CUSTOM_OUTPUT_CHECK-":
            is_checked = values["-CUSTOM_OUTPUT_CHECK-"]
            window["-OUTPUT-"].update(disabled=not is_checked)
            window["-OUTPUT_BROWSE-"].update(disabled=not is_checked)
            if not is_checked:
                window["-OUTPUT-"].update("")

        if event == "-RUN-":
            window["-RUN-"].update(disabled=True)
            window["-STATUS-"].update("")
            window["-PROGRESS-"].update(0)
            try:
                start_page_num = int(values["-START_PAGE-"])
                if start_page_num <= 0:
                    raise ValueError("A página deve ser maior que 0")
            except ValueError:
                sg.popup_error("Número da Página Inválido. Deve ser um número inteiro maior que 0.")
                window["-RUN-"].update(disabled=False)
                continue
            final_output_path = ""
            if values["-CUSTOM_OUTPUT_CHECK-"]:
                final_output_path = values["-OUTPUT-"]
            else:
                docx_path = values["-DOCX-"]
                base_name = os.path.splitext(os.path.abspath(docx_path))[0]
                final_output_path = base_name + ".pdf"
            if not final_output_path:
                sg.popup_error("Erro: Caminho de saída não pôde ser determinado.")
                window["-RUN-"].update(disabled=False)
                continue
            threading.Thread(
                target=pdf_worker_thread,
                args=(
                    window,
                    values["-DOCX-"],
                    values["-ADD_DRAWING_PDF-"],
                    values["-PDF-"],
                    start_page_num,
                    final_output_path
                ),
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

# --- PARTE 3: MENU PRINCIPAL (sem alteração) ---

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
    if (version<ultima):
        window = sg.Window(f"RevMaker Version {version} ", layout, element_justification='c')
    else:
        window = sg.Window(f"RevMaker Version (Unreleased)", layout, element_justification='c')
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
            webbrowser.open(f"https://github.com/LBCoelho/revMaker")
    window.close()

# --- ENTRADA ---
if __name__ == "__main__":
    versao_atual = version  # Ajuste para a versão do seu programa
    atual, ultima = verificar_ultima_versao(versao_atual)

    if (version<ultima):
        if atual:
            print(f"Você está usando a última versão ({ultima})!")
        else:
            print(f"Existe uma versão mais recente disponível: {ultima}")
            print("Baixe aqui:", f"https://github.com/LBCoelho/revMaker/releases/download/{ultima}/revMaker.zip")
    else: 
        print("Versão de testes (Não publicada)")
    versao_atual = version  # Ajuste para a versão do seu programa
    atual, ultima = verificar_ultima_versao(versao_atual)
    if (version<ultima):
        if atual:
            print("Atual")
        else:
            resposta = sg.popup_yes_no(f"Existe uma versão mais recente disponível: {ultima}. Deseja baixar?", title="Aviso")
            if resposta == "Yes":
                webbrowser.open(f"https://digicorner.sharepoint.com/sites/SubseaEngineering/_layouts/15/download.aspx?SourceUrl=%2Fsites%2FSubseaEngineering%2FNuvem%2020%2F08%2DOPR%2F04%2DProjetos%2FRevMaker%2FrevMaker%2Eexe")

    create_main_menu()
    
