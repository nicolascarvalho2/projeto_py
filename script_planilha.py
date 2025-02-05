import openpyxl
import pyperclip
import pyautogui
from time import sleep
#Entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']
#Copiar informações de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    
    #Nome do Produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(192,372,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #Descricao
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(211,468,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #Categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(207,595,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #Codigo do Produto
    codigo_do_produto = linha[3].value
    pyperclip.copy(codigo_do_produto)
    pyautogui.click(262,681,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #Peso
    peso_kg = linha[4].value
    pyperclip.copy(peso_kg)
    pyautogui.click(241,766,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #Dimenssoes
    dimenssoes = linha[5].value
    pyperclip.copy(dimenssoes)
    pyautogui.click(282,844,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #Botao Prox
    pyautogui.click(175,908, duration=1)
    sleep(3)

    #Preco
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(161,416,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #Quantidade em Estoque
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(250,499,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #Data de validade
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(234,584,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #Cor
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(273,671,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #tamanho
    tamanho = linha[10].value
    pyautogui.click(256,756,duration=1)

    if tamanho == 'Pequeno':
        pyautogui.click(209,792,duration=1)
        
    elif tamanho == 'Médio':
        pyautogui.click(198,821,duration=1)
    
    else:
        pyautogui.click(223,854,duration=1)
           
    #Material
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(213,841,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #Botao Prox
    pyautogui.click(171,907, duration=1)
    sleep(3)

    #Fabricante
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(188,436,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #Pais de Origem
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(208,517,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #Observacoes
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(190,606,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #Codigo de Barras
    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(183,739,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #localizacao armazem
    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(164,828,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #Botao Concluir
    pyautogui.click(171,888,duration=1)
    #Botao confirmação inclusão
    pyautogui.click(654,190,duration=1)
    #Iniciar cadastro novamente
    pyautogui.click(481,645,duration=1)
    