import os
import sys
import threading
import win32com.client
from PyPDF2 import PdfReader, PdfWriter
import FreeSimpleGUI as sg # Versão gratis do PySimpleGUI

# --- Parte 1: Conversor de Word para PDF ---

def convert_word_to_pdf(input_word_path, status_callback):
    """Converte um arquivo Word para PDF e retorna o caminho do novo PDF."""
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

# --- Parte 2: Manipulação de PDFs ---

def manipulate_pdfs(pdf_modify, pdf_insert, start_page_replace, final_output_path, status_callback):
    """Substitui um bloco de páginas em um PDF por páginas de outro PDF."""
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
        
        start_index = start_page_replace - 1
        end_index = start_index + num_pages_to_insert
        
        status_callback(f"OK. Substituindo {num_pages_to_insert} páginas, começando na página {start_page_replace}.")

        writer = PdfWriter()

        # 1. Adiciona as paginas iniciais do PDF convertido
        for i in range(start_index):
            writer.add_page(reader_modify.pages[i])
        
        # 2. Adiciona todas as paginas do PDF
        for page in reader_insert.pages:
            writer.add_page(page)
        
        # 3. Adiciona as paginas finais do PDF convertido
        for i in range(end_index, num_pages_total_modify):
            writer.add_page(reader_modify.pages[i])
            
        status_callback("Bloco de páginas substituído com sucesso.")

        # 4. Salvamento
        with open(final_output_path, "wb") as f:
            writer.write(f)
        
        status_callback("-" * 40)
        status_callback(f"PROCESSO CONCLUÍDO!")
        status_callback(f"Arquivo final salvo como: {os.path.basename(final_output_path)}")
        
    except Exception as e:
        raise Exception(f"Erro durante a manipulação do PDF: {e}")
    finally:
        # 5. Remove o arquivo auxiliar
        try:
            os.remove(pdf_modify)
            status_callback(f"Arquivo intermediário '{os.path.basename(pdf_modify)}' removido.")
        except OSError as e:
            status_callback(f"Aviso: Não foi possível remover o arquivo intermediário: {e}")


def worker_thread(window, docx_file, pdf_file, start_page, output_file):
    """This function runs in a separate thread to prevent the GUI from freezing."""
    try:
        # Wrapper for status_callback to also update the progress bar
        def update_gui(message, progress=None):
            # Send an event to the GUI thread to update its components
            window.write_event_value('-THREAD_UPDATE-', {'message': message, 'progress': progress})
        
        update_gui("Iniciando processo...", 0)
        
        # Step 1: Convert Word to PDF
        converted_pdf = convert_word_to_pdf(docx_file, lambda msg: update_gui(msg, 25))
        
        # Step 2: Manipulate the PDFs
        manipulate_pdfs(converted_pdf, pdf_file, start_page, output_file, lambda msg: update_gui(msg, 75))
        
        update_gui("Processo finalizado com sucesso!", 100)
        window.write_event_value('-THREAD_DONE-', None)

    except Exception as e:
        # If an error occurs, send it back to the main thread
        window.write_event_value('-THREAD_ERROR-', str(e))


def create_gui():
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
        [sg.Button("Executar", key="-RUN-", font=("Helvetica", 12), size=(10,1), disabled=True, button_color=('white', 'green')),
         sg.Button("Limpar", key="-CLEAR-"),
         sg.Button("Sair", button_color=('white', 'firebrick'))]
    ]
    layout = [[sg.Column(input_column), sg.VSeperator(), sg.Column(status_column, element_justification='center')]]

    window = sg.Window("PDF Automático v2.0", layout)

    # --- EVENT LOOP ---
    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == "Sair":
            break

        # --- Validation Logic ---
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
            window["-STATUS-"].update("") # Clear status
            window["-PROGRESS-"].update(0) # Reset progress
            
            threading.Thread(
                target=worker_thread,
                args=(window, values["-DOCX-"], values["-PDF-"], int(values["-START_PAGE-"]), values["-OUTPUT-"]),
                daemon=True
            ).start()
            
        # --- Handle Events from Worker Thread ---
        if event == '-THREAD_UPDATE-':
            sg.cprint(values[event]['message']) # Use cprint to print to Multiline
            if values[event]['progress'] is not None:
                window['-PROGRESS-'].update(values[event]['progress'])
        
        if event == '-THREAD_DONE-':
            sg.popup_ok("Processo Concluído!", "O seu arquivo PDF foi gerado com sucesso.")
            # Keep Run button disabled until a change is made
        
        if event == '-THREAD_ERROR-':
            sg.cprint(f"ERRO: {values[event]}", colors='white on red')
            sg.popup_error(f"Ocorreu um erro durante o processo:\n\n{values[event]}")
            window["-RUN-"].update(disabled=False) # Re-enable run on error to allow retry
            
    window.close()

if __name__ == "__main__":
    # Add your logic functions (convert_word_to_pdf, manipulate_pdfs) here
    # ...
    create_gui()