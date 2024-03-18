import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
pagina_produtos = workbook['Produtos']

for linha in pagina_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1400,328,duration=0.5)
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1352,447,duration=0.5)
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1340,593,duration=0.5)
    pyautogui.hotkey('ctrl','v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1374,696,duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1402,802,duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1428,923,duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(1268,1000,duration=0.5)
    sleep(2)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1301,349,duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(1295,457,duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(1302,569,duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1314,656,duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    tamanho = linha[10].value
    pyautogui.click(1328,774,duration=0.5)

    if tamanho == 'Pequeno':
        pyautogui.click(1293,825,duration=0.5)
    elif tamanho == 'Médio':
        pyautogui.click(1344,850,duration=0.5)
    else:
        pyautogui.click(1317,833,duration=0.5)


    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1351,883,duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(1288,957,duration=0.5)
    sleep(2)


    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1312,367,duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1306,476,duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1290,613,duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1347,749,duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(1353,856,duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(1294,937,duration=0.5)

    #Produto salvo no banco de dados
    pyautogui.click(1774,204,duration=0.5)
    sleep(2)

    #adicionar mais um
    pyautogui.click(1568,664,duration=0.5)
    

