import openpyxl
import pyautogui as pg
import pyperclip
from time import sleep
import webbrowser

# Abre o site do formulário de cadastro de produtos
webbrowser.open("https://cadastro-produtos-devaprender.netlify.app/index.html")
sleep(5)  # Aguarda o carregamento da página

# Carrega a planilha de produtos fictícios
workbook = openpyxl.load_workbook("data/produtos_ficticios.xlsx")
sheet = workbook["Produtos"]

# Itera sobre as linhas da planilha, começando pela segunda linha
for linha in sheet.iter_rows(min_row=2):
    # Copia o nome do produto e cola no campo correspondente no formulário
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pg.click(463,176, duration=1)
    pg.hotkey("ctrl", "v")

    # Copia a descrição do produto e cola no campo correspondente no formulário
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pg.click(471,281, duration=1)
    pg.hotkey("ctrl", "v")

    # Copia a categoria do produto e cola no campo correspondente no formulário
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pg.click(463,395, duration=1)
    pg.hotkey("ctrl", "v")

    # Copia o código do produto e cola no campo correspondente no formulário
    codigo_prod = linha[3].value
    pyperclip.copy(codigo_prod)
    pg.click(471,479, duration=1)
    pg.hotkey("ctrl", "v")

    # Copia o peso do produto e cola no campo correspondente no formulário
    peso = linha[4].value
    pyperclip.copy(peso)
    pg.click(464,567, duration=1)
    pg.hotkey("ctrl", "v")

    # Copia as dimensões do produto e cola no campo correspondente no formulário
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pg.click(475,657, duration=1)
    pg.hotkey("ctrl", "v")

    # Clica para pular para a próxima página no formulário
    pg.click(171,710, duration=1)
    sleep(3)  # Aguarda a próxima página carregar

    # Copia o preço do produto e cola no campo correspondente no formulário
    preco = linha[6].value
    pyperclip.copy(preco)
    pg.click(512,198, duration=1)
    pg.hotkey("ctrl", "v")

    # Copia a quantidade do produto e cola no campo correspondente no formulário
    quantidade = linha[7].value
    pyperclip.copy(quantidade)
    pg.click(511,286, duration=1)
    pg.hotkey("ctrl", "v")

    # Copia a validade do produto e cola no campo correspondente no formulário
    validade = linha[8].value
    pyperclip.copy(validade)
    pg.click(509,372, duration=1)
    pg.hotkey("ctrl", "v")

    # Copia a cor do produto e cola no campo correspondente no formulário
    cor = linha[9].value
    pyperclip.copy(cor)
    pg.click(516,460, duration=1)
    pg.hotkey("ctrl", "v")

    # Copia o tamanho do produto e seleciona o campo apropriado no formulário
    tamanho = linha[10].value
    pg.click(519,548, duration=1)  # Abre o campo de seleção
    if tamanho == "Pequeno":
        pg.click(508,578, duration=1)  # Seleciona o tamanho pequeno
    elif tamanho == "Médio":
        pg.click(508,605, duration=1)  # Seleciona o tamanho médio
    else:
        pg.click(508,367, duration=1)  # Seleciona o tamanho grande

    # Copia o material do produto e cola no campo correspondente no formulário
    material = linha[11].value
    pyperclip.copy(material)
    pg.click(508,630, duration=1)
    pg.hotkey("ctrl", "v")

    # Clica para pular para a próxima página no formulário
    pg.click(153,688, duration=1)
    sleep(3)  # Aguarda a próxima página carregar

    # Copia o fabricante do produto e cola no campo correspondente no formulário
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pg.click(460, 219, duration=1)
    pg.hotkey("ctrl", "v")

    # Copia o país de origem do produto e cola no campo correspondente no formulário
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pg.click(428, 307, duration=1)
    pg.hotkey("ctrl", "v")

    # Copia as observações do produto e cola no campo correspondente no formulário
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pg.click(420, 407, duration=1)
    pg.hotkey("ctrl", "v")

    # Copia o código de barras do produto e cola no campo correspondente no formulário
    codigo_barra = linha[15].value
    pyperclip.copy(codigo_barra)
    pg.click(429, 532, duration=1)
    pg.hotkey("ctrl", "v")

    # Copia a localização do armazém e cola no campo correspondente no formulário
    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pg.click(417, 614, duration=1)
    pg.hotkey("ctrl", "v")

    # Clica no botão para concluir o cadastro do produto
    pg.click(171, 671, duration=1)

    # Clica em "ok" para confirmar que o produto foi salvo no banco de dados
    pg.click(849, 180, duration=1)

    # Clica para adicionar mais um produto
    pg.click(689, 449, duration=1)
    sleep(3)  # Aguarda a nova página carregar
