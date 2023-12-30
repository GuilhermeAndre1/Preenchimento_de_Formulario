import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):

    #Nome do Produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(888,218, duration=1)
    pyautogui.hotkey('ctrl','V')

    #Descrição do Produto
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(957,315, duration=1)
    pyautogui.hotkey('ctrl','V')

    #Categoria do Produto
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(878,421, duration=1)
    pyautogui.hotkey('ctrl','V')

    #Código do Produto
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(899,494, duration=1)
    pyautogui.hotkey('ctrl','V')

    #Peso do produto
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(935,567, duration=1)
    pyautogui.hotkey('ctrl','V')

    #Dimensões do produto
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(911,644, duration=1)
    pyautogui.hotkey('ctrl','V')

    #Click para prosseguir de página
    pyautogui.click(892,698, duration=1)
    sleep(1)
    
    #Preço do produto
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(893,242, duration=1)
    pyautogui.hotkey('ctrl','V')

    #Quantidade em estoque do produto
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(884,322, duration=1)
    pyautogui.hotkey('ctrl','V')

    #Data de validade do produto
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(886,395, duration=1)
    pyautogui.hotkey('ctrl','V')

    #Cor do produto
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(886,473, duration=1)
    pyautogui.hotkey('ctrl','V')

    #Tamanho do produto
    tamanho = linha[10].value
    pyautogui.click(899,551, duration=1)

    if tamanho == 'Pequeno':
        pyautogui.click(913,584,duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(889,601,duration=1)
    else:
        pyautogui.click(876,615,duration=1)

    #Material do produto
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(876,626, duration=1)
    pyautogui.hotkey('ctrl','V')

    #Click para prosseguir de página 2
    pyautogui.click(892,689, duration=1)
    sleep(1)
    
    #Fabricante do Produto
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(890,260, duration=1)
    pyautogui.hotkey('ctrl','V')
    
    #País de origem do Produto
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(900,332, duration=1)
    pyautogui.hotkey('ctrl','V')
    
    #Observações do Produto
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(906,415, duration=1)
    pyautogui.hotkey('ctrl','V')

    #Código de barras do Produto
    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(893,533, duration=1)
    pyautogui.hotkey('ctrl','V')

    #Localização do Armazém onde se encontra o Produto
    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(900,612, duration=1)
    pyautogui.hotkey('ctrl','V')

    #Botão Concluir
    pyautogui.click(906,673, duration=1)
    #Botão OK
    pyautogui.click(1285,171, duration=1)
    #Botão Adicionar mais um Produto
    pyautogui.click(1066,462, duration=1)
