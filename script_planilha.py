import openpyxl
#Entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']
#Copiar informações de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    descrição = linha[1].value
    categoria = linha[2].value
    codigo_do_produto = linha[3].value
    peso_kg = linha[4].value
    dimenssoes = linha[5].value
    preco = linha[6].value
    quantidade_em_estoque = linha[7].value
    data_de_validade = linha[8].value
    cor = linha[9].value
    tamanho = linha[10].value
    material = linha[11].value
    fabricante = linha[12].value
    pais_origem = linha[13].value
    observacoes = linha[14].value
    codigo_de_barras = linha[15].value
    localizacao_armazem = linha[16].value
    



