import os
import sys
import threading
import re
import shutil
from pathlib import Path
import FreeSimpleGUI as sg # Versão gratis do PySimpleGUI

# --- Parte 1: Lógica Principal (MODIFICADA) ---

def processar_revisao(diretorio_base_str, aux_files_list, status_callback):
    """
    Encontra a Rev. X, verifica o PDF, adiciona a tag [Em revisão] na pasta raiz,
    e então cria a estrutura da Rev. X+1.
    """
    
    # --- ETAPA 1: Análise com o caminho ANTIGO ---
    diretorio_base = Path(diretorio_base_str)
    status_callback(f"Analisando diretório: {diretorio_base}", 10) 
    
    padrao_rev = re.compile(r'^Rev\.\s(\d{1,2})$')
    pastas_rev_encontradas = []
    
    for item in diretorio_base.iterdir():
        if item.is_dir():
            match = padrao_rev.match(item.name)
            if match:
                numero_rev = int(match.group(1))
                pastas_rev_encontradas.append((numero_rev, item))

    if not pastas_rev_encontradas:
        raise Exception("Erro: Nenhuma pasta no formato 'Rev. X' (com espaço) foi encontrada.")

    # Armazena o caminho da Rev. X com base no caminho ANTIGO
    rev_atual_num, pasta_rev_atual_CAMINHO_ANTIGO = max(pastas_rev_encontradas, key=lambda item: item[0])
    
    status_callback(f"✅ Revisão mais alta encontrada: {pasta_rev_atual_CAMINHO_ANTIGO.name}", 20)

    # Verifica o PDF no caminho ANTIGO
    pdfs_na_pasta_atual = list(pasta_rev_atual_CAMINHO_ANTIGO.glob('*.pdf'))

    if not pdfs_na_pasta_atual:
        raise Exception(f"Erro: Sem PDF da revisão na pasta {pasta_rev_atual_CAMINHO_ANTIGO.name}. Verifique se essa pasta já não é a revisão em andamento.")

    status_callback(f"✅ PDF encontrado: {pdfs_na_pasta_atual[0].name}", 30)

    # --- ARMAZENA APENAS OS NOMES DOS ARQUIVOS ---
    # Isso é crucial. Não guardamos o caminho completo, pois ele vai mudar.
    pdf_origem_nome = pdfs_na_pasta_atual[0].name
    docx_origem_nomes = [p.name for p in pasta_rev_atual_CAMINHO_ANTIGO.glob('*.docx')]
    pasta_rev_atual_nome = pasta_rev_atual_CAMINHO_ANTIGO.name # Ex: "Rev. 2"


    # --- ETAPA 2: Renomeação da Pasta Raiz ---
    tag = "[Em revisão] "
    if not diretorio_base.name.startswith(tag):
        status_callback(f"Adicionando tag ao diretório: {diretorio_base.name}...")
        try:
            novo_nome = tag + diretorio_base.name
            novo_caminho_completo = diretorio_base.parent / novo_nome
            
            diretorio_base.rename(novo_caminho_completo)
            
            # 'diretorio_base' AGORA APONTA PARA O NOVO CAMINHO (ex: ...\[Em revisão] Projeto)
            diretorio_base = novo_caminho_completo 
            status_callback(f"   -> Diretório renomeado para: {novo_nome}", 35)
        except Exception as e:
            raise Exception(f"Erro ao tentar renomear pasta raiz: {e}. (A pasta está em uso?)")
    else:
        status_callback("Tag [Em revisão] já existe. Prosseguindo...", 35)

    # --- ETAPA 3: Definição dos NOVOS caminhos ---
    
    # ATUALIZA o caminho da pasta de origem (ex: "Rev. 2") para usar o NOVO 'diretorio_base'
    pasta_rev_atual = diretorio_base / pasta_rev_atual_nome 
    
    # Define o caminho completo do PDF de origem usando o NOVO caminho
    pdf_origem = pasta_rev_atual / pdf_origem_nome
    
    # Define a lista de caminhos DOCX de origem usando o NOVO caminho
    docx_origem_list = [pasta_rev_atual / nome for nome in docx_origem_nomes]

    status_callback("\nIniciando modificações: Criando nova revisão...")
    
    rev_proxima_num = rev_atual_num + 1
    # Cria a nova pasta de revisão dentro do NOVO 'diretorio_base'
    pasta_rev_proxima = diretorio_base / f"Rev. {rev_proxima_num}"

    if pasta_rev_proxima.exists():
        raise Exception(f"Erro: A pasta {pasta_rev_proxima.name} já existe! Nenhuma modificação foi feita.")

    # --- INÍCIO DAS MODIFICAÇÕES (Agora usando os caminhos corretos) ---
    
    pasta_rev_proxima.mkdir()
    status_callback(f"   -> Pasta criada: {pasta_rev_proxima.name}", 40)

    subpastas_iniciais = ["01-Auxiliares", "02-E-mails", "03-Fotos e Videos"]
    for subpasta in subpastas_iniciais:
        (pasta_rev_proxima / subpasta).mkdir()
        status_callback(f"      -> Subpasta criada: {subpasta}")
    
    status_callback("Subpastas iniciais criadas.", 50)
    
    # Copiar múltiplos arquivos auxiliares (Esta lógica já estava correta)
    if aux_files_list:
        status_callback(f"   -> Anexando {len(aux_files_list)} arquivos auxiliares...")
        aux_path = pasta_rev_proxima / "01-Auxiliares"
        for file_path_str in aux_files_list:
            if file_path_str and os.path.exists(file_path_str):
                try:
                    file_path = Path(file_path_str)
                    shutil.copy2(file_path, aux_path)
                    status_callback(f"      -> Copiado: {file_path.name}")
                except Exception as e:
                    status_callback(f"   -> ⚠️ Aviso: Falha ao copiar {os.path.basename(file_path_str)}. Erro: {e}")
            elif file_path_str:
                status_callback(f"   -> ⚠️ Aviso: Caminho do arquivo auxiliar '{file_path_str}' não encontrado. Pulando.")

    # Copiar 04-JRA e 05-Desenhos (AGORA CORRIGIDO)
    pastas_para_copiar = ["04-JRA", "05-Desenhos"]
    for pasta_nome in pastas_para_copiar:
        # Usa o 'pasta_rev_atual' ATUALIZADO
        src = pasta_rev_atual / pasta_nome
        dst = pasta_rev_proxima / pasta_nome
        
        if src.is_dir():
            shutil.copytree(src, dst)
            status_callback(f"   -> Pasta copiada (com conteúdo): {pasta_nome}")
        else:
            status_callback(f"   -> ⚠️ Aviso: Pasta '{pasta_nome}' não encontrada em {pasta_rev_atual.name}, não foi copiada.")

    status_callback("Pastas de projeto copiadas.", 70)

    # Criar 06-OLD
    pasta_old = pasta_rev_proxima / "06-OLD"
    pasta_old.mkdir()
    status_callback(f"   -> Pasta criada: 06-OLD")
    
    # Copiar PDF para 06-OLD (AGORA CORRIGIDO)
    # Usa o 'pdf_origem' ATUALIZADO
    shutil.copy2(pdf_origem, pasta_old)
    status_callback(f"      -> PDF copiado para 06-OLD: {pdf_origem.name}")

    # Copiar DOCX para 06-OLD (AGORA CORRIGIDO)
    docx_origem_para_renomear = None # Usaremos isso para a próxima etapa
    if docx_origem_list:
        # Pega o primeiro docx da lista para usar como base para renomear
        docx_origem_para_renomear = docx_origem_list[0] 
        
        # Itera sobre TODOS os docx encontrados e copia para a pasta OLD
        for docx_path in docx_origem_list:
            shutil.copy2(docx_path, pasta_old)
            status_callback(f"      -> DOCX copiado para 06-OLD: {docx_path.name}")
    else:
        status_callback(f"   -> ⚠️ Aviso: Nenhum arquivo .docx encontrado em {pasta_rev_atual.name} para copiar para 06-OLD.")

    status_callback("Arquivos da revisão anterior movidos para 06-OLD.", 85)

    # Copiar e renomear o .docx principal (AGORA CORRIGIDO)
    if docx_origem_para_renomear:
        padrao_docx_rev = re.compile(r'(\d+)(\.docx)$', re.IGNORECASE)
        # Usa o nome do primeiro docx encontrado
        nome_original = docx_origem_para_renomear.name
        novo_nome_docx = padrao_docx_rev.sub(f'{rev_proxima_num}\\2', nome_original)
        
        if novo_nome_docx == nome_original:
            status_callback(f"   -> ⚠️ Aviso: Não foi possível renomear o DOCX '{nome_original}'. Padrão não encontrado.")
            shutil.copy2(docx_origem_para_renomear, pasta_rev_proxima / nome_original)
        else:
            # Usa o 'docx_origem_para_renomear' ATUALIZADO como fonte
            shutil.copy2(docx_origem_para_renomear, pasta_rev_proxima / novo_nome_docx)
            status_callback(f"   -> DOCX copiado e renomeado: {novo_nome_docx}")
    else:
        status_callback(f"   -> ⚠️ Aviso: Nenhum arquivo .docx encontrado para copiar para a raiz da nova revisão.")

    status_callback("\n🎉 Processo de Criação da Nova Revisão CONCLUÍDO! 🎉", 100)

# --- Parte 2: Thread de Execução (MODIFICADA) ---

def revisao_worker_thread(window, diretorio_base_str, aux_files_str):
    """Roda a função 'processar_revisao' em uma thread separada para não travar a GUI."""
    try:
        def update_gui(message, progress=None):
            window.write_event_value('-THREAD_UPDATE-', {'message': message, 'progress': progress})
        
        update_gui("Iniciando processo...", 0)
        
        # --- Lógica da TAG movida para 'processar_revisao' ---
        
        # --- Processa múltiplos arquivos ---
        # A GUI retorna uma string separada por ';', se houver
        aux_files_list = aux_files_str.split(';') if aux_files_str else []

        # Executa a lógica principal, passando a lista de arquivos
        processar_revisao(diretorio_base_str, aux_files_list, update_gui)
        
        window.write_event_value('-THREAD_DONE-', None)

    except Exception as e:
        window.write_event_value('-THREAD_ERROR-', str(e))


# --- Parte 3: Criação da GUI e Loop de Eventos (MODIFICADO) ---

def create_gui_revisao():
    sg.theme('GrayGrayGray')

    # Layout MODIFICADO
    input_column = [
        [sg.Text("1. Selecione o Diretório Raiz do Projeto")],
        # Campo de diretório agora é editável (não é mais readonly)
        [sg.Input(key="-DIR-", enable_events=True), sg.FolderBrowse("Procurar", target="-DIR-")],
        [sg.Text(" (A pasta que contém as 'Rev. 1', 'Rev. 2', etc.)", font=("Helvetica", 9))],
        [sg.HSeparator()],
        # Textos e chaves atualizados
        [sg.Checkbox("Anexar documentos auxiliares?", key="-AUX_CHECK-", enable_events=True)],
        [
            sg.Text("Arquivos Auxiliares:", key="-AUX_TEXT-"), 
            # Input recebe a string de múltiplos arquivos
            sg.Input(key="-AUX_FILES-", readonly=True, disabled=True, enable_events=True), 
            # FileBrowse agora permite múltiplos arquivos de qualquer tipo
            sg.FilesBrowse("Procurar", 
                          file_types=(("All Files", "*.*"),), 
                          target="-AUX_FILES-", 
                          key="-AUX_BROWSE-", 
                          disabled=True)
        ]
    ]
    
    status_column = [
        [sg.Text("Status do Processo")],
        [sg.Multiline(size=(50, 15), key="-STATUS-", autoscroll=True, disabled=True, reroute_cprint=True)],
        [sg.ProgressBar(100, orientation='h', size=(35, 20), key='-PROGRESS-')],
        [sg.Button("Executar", key="-RUN-", font=("Helvetica", 12), size=(10,1), disabled=True, button_color=('white', 'green')),
         sg.Button("Limpar", key="-CLEAR-"),
         sg.Button("Sair", button_color=('white', 'firebrick'))]
    ]
    
    layout = [[sg.Column(input_column), sg.VSeperator(), sg.Column(status_column, element_justification='center')]]

    window = sg.Window("Criador de Novas Revisões", layout)

    # --- EVENT LOOP (MODIFICADO) ---
    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == "Sair":
            break

        # Lógica de Habilitar/Desabilitar (chaves atualizadas)
        if event == "-AUX_CHECK-":
            is_checked = values["-AUX_CHECK-"]
            window["-AUX_FILES-"].update(disabled=not is_checked)
            window["-AUX_BROWSE-"].update(disabled=not is_checked)
            if not is_checked:
                window["-AUX_FILES-"].update("") 

        # Lógica de Validação (chaves atualizadas)
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

        # Lógica dos Botões (chaves atualizadas)
        if event == "-CLEAR-":
            window["-DIR-"].update("")
            window["-STATUS-"].update("")
            window["-PROGRESS-"].update(0)
            window["-RUN-"].update(disabled=True)
            window["-AUX_CHECK-"].update(False)
            window["-AUX_FILES-"].update("", disabled=True)
            window["-AUX_BROWSE-"].update(disabled=True)

        if event == "-RUN-":
            # --- NOVO: Validação do diretório manual/selecionado ---
            dir_path = values["-DIR-"]
            if not os.path.isdir(dir_path):
                sg.popup_error(f"Erro: O diretório não foi encontrado:\n{dir_path}", title="Diretório Inválido")
                continue # Para a execução, mas mantém a GUI aberta
            # --- FIM DA VALIDAÇÃO ---

            window["-RUN-"].update(disabled=True)
            window["-STATUS-"].update("") 
            window["-PROGRESS-"].update(0)
            
            threading.Thread(
                target=revisao_worker_thread,
                args=(window, values["-DIR-"], values["-AUX_FILES-"]), # Passa a string de arquivos
                daemon=True
            ).start()
            
        # Lidar com eventos da Thread (Sem mudanças)
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

# --- Ponto de Entrada Principal ---
if __name__ == "__main__":
    create_gui_revisao()