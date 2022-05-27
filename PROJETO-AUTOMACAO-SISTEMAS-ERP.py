import subprocess
import time
import pyautogui
import pandas as pd
import pyperclip

pyautogui.FAILSAFE = True


def encontrar_imagem(imagem):  # substituindo os while
    while not pyautogui.locateOnScreen(imagem, grayscale=True, confidence=0.90):
        time.sleep(1)
    encontrou = pyautogui.locateOnScreen(imagem, grayscale=True, confidence=0.90)  # posição x,y,largura,altura
    return encontrou


def escrever_texto(texto):
    pyperclip.copy(texto)
    pyautogui.hotkey('ctrl', 'v')

#  escrver a lógica:
#  abrir o ERP(fakturama)

subprocess.Popen([r'C:\Program Files\Fakturama2\Fakturama.exe'])
image = encontrar_imagem('logo-fak.png')

# como fazer para vários produtos
tabela_produtos = pd.read_excel('Produtos_fak.xlsx')
for linha in tabela_produtos.index:
    nome = tabela_produtos.loc[linha, 'Nome']
    id = tabela_produtos.loc[linha, 'ID']
    categoria = tabela_produtos.loc[linha, 'Categoria']
    gtin = tabela_produtos.loc[linha, 'GTIN']
    supplier = tabela_produtos.loc[linha, 'Supplier']
    descricao = tabela_produtos.loc[linha, 'Descrição']
    imagem = tabela_produtos.loc[linha, 'Imagem']
    preco = tabela_produtos.loc[linha, 'Preço']
    custo = tabela_produtos.loc[linha, 'Custo']
    estoque = tabela_produtos.loc[linha, 'Estoque']

    # *********CLICAR NO MENU NEW **************
    encontrou = encontrar_imagem('menu_new.png')
    pyautogui.click(pyautogui.center(encontrou))

    # ********** CLICAR EM NEW PRODUCT *************
    encontrou = encontrar_imagem('new_product.png')
    pyautogui.click(pyautogui.center(encontrou))

    #  ************* PREENCHER TODOS OS CAMPOS ***************
    encontrou = encontrar_imagem('item_number.png')
    pyautogui.click(pyautogui.center(encontrou))

    escrever_texto(str(id))
    pyautogui.press('tab')

    # name
    escrever_texto(str(nome))
    pyautogui.press('tab')
    # category
    escrever_texto(str(categoria))
    pyautogui.press('tab')
    # gtin
    escrever_texto(str(gtin))
    pyautogui.press('tab')
    # s.code
    escrever_texto(str(supplier))
    pyautogui.press('tab')
    # description
    escrever_texto(str(descricao))
    pyautogui.press('tab')

    # price gross
    preco_formatado = f'{preco:.2f}'.replace('.', ',')
    pyautogui.write(str(preco_formatado))
    pyautogui.press('tab')
    # cost price
    custo_formatado = f'{custo:.2f}'.replace('.', ',')
    pyautogui.write(str(custo_formatado))
    pyautogui.press('tab')

    pyautogui.press('tab')
    pyautogui.press('tab')
    # stock
    estoque_formatado = f'{estoque:.2f}'.replace('.', ',')
    pyautogui.write(str(estoque_formatado))
    pyautogui.write(str(estoque))

    # selecionar imagem
    encontrou = encontrar_imagem('select_photo.png')
    pyautogui.click(pyautogui.center(encontrou))

    encontrou = encontrar_imagem('nome_arq.png')
    pyautogui.click(pyautogui.center(encontrou))
    pyautogui.write(rf'"C:\Users\Jaque\Downloads\Imagens Produtos\{str(imagem)}"')
    pyautogui.press('enter')

    # ************* Clicar em salvar *****************
    encontrou = encontrar_imagem('save.png')
    pyautogui.click(pyautogui.center(encontrou))
