## Projeto de Automacao de Sistemas -ERP(Tovts, SAP, etc)
 
 Cadastrando, de forma automatizada, produtos de uma loja, no sistema ERP
- Exemplo de ERP : [`Fakturama`](https://www.fakturama.info/)
- Dados : Excel

## Bibliotecas

- Subprocess
- Pandas
- Time
- Pyautogui
- Pyperclip

## ERP de exemplo

![image](https://github.com/jackdepaulo/automacao-ERP/blob/6c2593ee31c695e77c3f8d1ed762985924f0a066/imagens/erp.png)

## Banco de Dados

Produtos a serem cadastrados:

![image](https://github.com/jackdepaulo/automacao-ERP/blob/5b645ed1ad4e44dc7774ae72742efe9e1d6c7e6b/imagens/dados.png)

---


**Função feita para localizar imagens, para que o "click" no item desejado, seja preciso. E assim, conseguir cadastrar produtos sem erros.**
```sh
def encontrar_imagem(imagem):
    while not pyautogui.locateOnScreen(imagem, grayscale=True, confidence=0.90):
        time.sleep(1)
    encontrou = pyautogui.locateOnScreen(imagem, grayscale=True, confidence=0.90)
    return encontrou
```

---

**Importando os Dados, e criando variáveis para cada coluna da Base de Dados, para cadastrar vários produtos ao mesmo tempo.**
```sh
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
```
---

## Produtos Cadastrados
![image](https://github.com/jackdepaulo/automacao-ERP/blob/6c2593ee31c695e77c3f8d1ed762985924f0a066/imagens/fakturama-resultado.png)



    
