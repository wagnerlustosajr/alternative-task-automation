import openpyxl
import pyperclip
import pyautogui
from time import sleep

pyautogui.PAUSE = 0.7

# Entrar na planilha
workbook = openpyxl.load_workbook("produtos_ficticios.xlsx")
sheet_produtos = workbook["Produtos"]
# Copiar informações de um campo e colar no seu campo correspondente 
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(151,305,duration=1)
    pyautogui.hotkey("ctrl","v")

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(150,403 ,duration=1)
    pyautogui.hotkey("ctrl","v")
    
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(191,576,duration=1)
    pyautogui.hotkey("ctrl","v")

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(168,682,duration=1)
    pyautogui.hotkey("ctrl","v")

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(175,795,duration=1)
    pyautogui.hotkey("ctrl","v")

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(172,897 ,duration=1)
    pyautogui.hotkey("ctrl","v")

    pyautogui.click(189,970,duration=1)
    sleep(5)
    
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(177,330,duration=1)
    pyautogui.hotkey("ctrl","v")
    
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(180,438,duration=1)
    pyautogui.hotkey("ctrl","v")

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(171,550,duration=1)
    pyautogui.hotkey("ctrl","v")

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(174,652,duration=1)
    pyautogui.hotkey("ctrl","v")

    tamanho = linha[10].value
    pyautogui.click(182,765,duration=1)

    if tamanho == "Pequeno":
        pyautogui.click(205,800,duration=1)  
    elif tamanho == "Médio":
        pyautogui.click(203,831,duration=1)
    else: tamanho == "Grande"
    pyautogui.click(204,853,duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(177,868,duration=1)
    pyautogui.hotkey("ctrl","v")

    pyautogui.click(185,940,duration=1) #Botão próximo

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(168,352,duration=1)
    pyautogui.hotkey("ctrl","v")

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(175,462,duration=1)
    pyautogui.hotkey("ctrl","v")
    
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(192,560,duration=1)
    pyautogui.hotkey("ctrl","v")
    
    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(190,738,duration=1)
    pyautogui.hotkey("ctrl","v")
    
    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(204,839,duration=1)
    pyautogui.hotkey("ctrl","v")
   
    pyautogui.click(198,920,duration=1)  #Botão concluir
    pyautogui.click(1171,240,duration=1) #Confirmar produto "ok"
    pyautogui.click(976,642,duration=1)  #Adicionar mais 1