import time
import openpyxl
import pyperclip
import pyautogui
from time import sleep
import tkinter as tk
import threading    


workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
pagina_produtos = workbook['Produtos']

print("""
           _ _                         _____             
     /\   | | |                       |  __ \            
    /  \  | | | ___  _ __  ___  ___   | |  | | _____   __
   / /\ \ | | |/ _ \| '_ \/ __|/ _ \  | |  | |/ _ \ \ / /
  / ____ \| | | (_) | | | \__ \ (_) | | |__| |  __/\ V / 
 /_/    \_\_|_|\___/|_| |_|___/\___/  |_____/ \___| \_/  
                                                         
                                                          """)
print('========Para finalizar a execução do programa, feche a janela do prompt de comando==========')
print('Para evitar erros, finalize o programa somente nos 5 segunos onde ele se encontra parado ao começar um novo produto!')

def programa():  
            for linha in pagina_produtos.iter_rows(min_row=2):
                nome_produto = linha[0].value
                pyperclip.copy(nome_produto)
                pyautogui.click(1414,184,duration=0.5)
                pyautogui.hotkey('tab')
                pyautogui.hotkey('ctrl','v')

                descricao = linha[1].value
                pyperclip.copy(descricao)
                pyautogui.hotkey('tab')
                pyautogui.hotkey('ctrl','v')

                categoria = linha[2].value
                pyperclip.copy(categoria)
                pyautogui.hotkey('tab')
                pyautogui.hotkey('ctrl','v')

                codigo_produto = linha[3].value
                pyperclip.copy(codigo_produto)
                pyautogui.hotkey('tab')
                pyautogui.hotkey('ctrl', 'v')

                peso = linha[4].value
                pyperclip.copy(peso)
                pyautogui.hotkey('tab')
                pyautogui.hotkey('ctrl', 'v')

                dimensoes = linha[5].value
                pyperclip.copy(dimensoes)
                pyautogui.hotkey('tab')
                pyautogui.hotkey('ctrl', 'v')

                pyautogui.hotkey('tab')
                pyautogui.hotkey('enter')
                sleep(2)

                preco = linha[6].value
                pyperclip.copy(preco)
                pyautogui.hotkey('tab')
                pyautogui.hotkey('ctrl', 'v')

                quantidade_em_estoque = linha[7].value
                pyperclip.copy(quantidade_em_estoque)
                pyautogui.hotkey('tab')
                pyautogui.hotkey('ctrl', 'v')

                data_de_validade = linha[8].value
                pyperclip.copy(data_de_validade)
                pyautogui.hotkey('tab')
                pyautogui.hotkey('ctrl', 'v')

                cor = linha[9].value
                pyperclip.copy(cor)
                pyautogui.hotkey('tab')
                pyautogui.hotkey('ctrl', 'v')

                tamanho = linha[10].value
                pyautogui.hotkey('tab')
                pyautogui.hotkey('enter')

                if tamanho == 'Pequeno':
                    pyautogui.hotkey('enter')

                elif tamanho == 'Médio':
                    pyautogui.hotkey('down')
                    pyautogui.hotkey('enter')

                else:
                    pyautogui.hotkey('down')
                    pyautogui.hotkey('down')
                    pyautogui.hotkey('enter')


                material = linha[11].value
                pyperclip.copy(material)
                pyautogui.hotkey('tab')
                pyautogui.hotkey('ctrl', 'v')

                pyautogui.hotkey('tab')
                pyautogui.hotkey('enter')
                sleep(2)


                fabricante = linha[12].value
                pyperclip.copy(fabricante)
                pyautogui.hotkey('tab')
                pyautogui.hotkey('ctrl', 'v')

                pais_origem = linha[13].value
                pyperclip.copy(pais_origem)
                pyautogui.hotkey('tab')
                pyautogui.hotkey('ctrl', 'v')

                observacoes = linha[14].value
                pyperclip.copy(observacoes)
                pyautogui.hotkey('tab')
                pyautogui.hotkey('ctrl', 'v')

                codigo_de_barras = linha[15].value
                pyperclip.copy(codigo_de_barras)
                pyautogui.hotkey('tab')
                pyautogui.hotkey('ctrl', 'v')

                localizacao_armazem = linha[16].value
                pyperclip.copy(localizacao_armazem)
                pyautogui.hotkey('tab')
                pyautogui.hotkey('ctrl', 'v')

                pyautogui.hotkey('tab')
                pyautogui.hotkey('enter')

                #Produto salvo no banco de dados
                pyautogui.hotkey('enter')
                sleep(2)

                #adicionar mais um
                pyautogui.hotkey('tab')
                pyautogui.hotkey('enter')

                sleep(5)        

minha_thread = threading.Thread(target=programa)
exitt = threading.Thread(target=exit)

janela = tk.Tk()
rotulo = tk.Label(janela, text="Aperte o botão para rodar o programa")
rotulo.pack()

botao_continuar = tk.Button(janela, text="Rodar", command=minha_thread.start)
botao_continuar.pack()
botao_sair = tk.Button(janela, text="Sair", command=exit)
botao_sair.pack()


janela.mainloop()
