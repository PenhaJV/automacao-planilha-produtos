import openpyxl
import pyautogui as pg
import pyperclip
from time import sleep

# entrar na planilha
workbook = openpyxl.load_workbook("data/produtos_ficticios.xlsx")
sheet = workbook["Produtos"]


# copiar informação de um campo e colar no campo correspondente
for linha in sheet.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pg.click(879, 207, duration=1)
    pg.hotkey("ctrl", "v")

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pg.click(877, 299, duration=1)
    pg.hotkey("ctrl", "v")

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pg.click(882, 411, duration=1)
    pg.hotkey("ctrl", "v")

    codigo_prod = linha[3].value
    pyperclip.copy(codigo_prod)
    pg.click(876, 485, duration=1)
    pg.hotkey("ctrl", "v")

    peso = linha[4].value
    pyperclip.copy(peso)
    pg.click(894, 567, duration=1)
    pg.hotkey("ctrl", "v")

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pg.click(886, 637, duration=1)
    pg.hotkey("ctrl", "v")

    # pular para a próxima página
    pg.click(878, 695, duration=1)
    sleep(3)

    preco = linha[6].value
    pyperclip.copy(preco)
    pg.click(894, 232, duration=1)
    pg.hotkey("ctrl", "v")

    quantidade = linha[7].value
    pyperclip.copy(quantidade)
    pg.click(898, 309, duration=1)
    pg.hotkey("ctrl", "v")

    validade = linha[8].value
    pyperclip.copy(validade)
    pg.click(887, 382, duration=1)
    pg.hotkey("ctrl", "v")

    cor = linha[9].value
    pyperclip.copy(cor)
    pg.click(886, 464, duration=1)
    pg.hotkey("ctrl", "v")

    tamanho = linha[10].value
    # abrir o campo de seleção
    pg.click(893, 545, duration=1)
    # escolher o campo equivalente
    if tamanho == "Pequeno":
        pg.click(903, 575, duration=1)
    elif tamanho == "Médio":
        pg.click(907, 601, duration=1)
    else:
        pg.click(907, 625, duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pg.click(907, 620, duration=1)
    pg.hotkey("ctrl", "v")

    # pular para a próxima página
    pg.click(887, 673, duration=1)
    sleep(3)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pg.click(938, 248, duration=1)
    pg.hotkey("ctrl", "v")

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pg.click(931, 325, duration=1)
    pg.hotkey("ctrl", "v")

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pg.click(926, 414, duration=1)
    pg.hotkey("ctrl", "v")

    codigo_barra = linha[15].value
    pyperclip.copy(codigo_barra)
    pg.click(925, 526, duration=1)
    pg.hotkey("ctrl", "v")

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pg.click(926, 602, duration=1)
    pg.hotkey("ctrl", "v")

    # concluir cadastro
    pg.click(887, 659, duration=1)

    # clicar em ok para confirmar que o produto foi salvo no banco de dados
    pg.click(1248, 186, duration=1)

    # clicar em add mais um produto
    pg.click(1081, 454, duration=1)
    sleep(3)
