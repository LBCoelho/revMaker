#import pdfAutomatico
#import pastaAutomatica
import subprocess
import sys

def menu():
        print("Escolha o Aplicativo desejado:\n")
        while True:
            print("1. Criação Automática de Pastas")
            print("2. Criação do PDF final (Abrir GUI)")
            print("3. Sair")

            escolha = input("Escolha uma opção: ")

            match escolha:
                case '1':
                    print("Abrindo criador de pastas \n\n")
                    try:
                        subprocess.run([sys.executable, "revFolderCreator.py"]) 
                    except:
                        print("ERRO: O arquivo 'revFolderCreator.py' não foi encontrado.")
                case '2':
                    print("Abrindo a aplicação de PDF...")
                    try:
                        subprocess.run([sys.executable, "pdfAutomatico.py"]) 
                    except FileNotFoundError:
                        print("ERRO: O arquivo 'pdfAutomatico.py' não foi encontrado.")

                case '3':
                    print("Saindo...")
                    break
menu()