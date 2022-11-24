#FORMATACAO CONDICIONAL. CRIA UM ARQUIVO EXCEL COM VALORES E INSERE UMA COR CASO SEJA MAIOR OU IGUAL A 50 E OUTRA COR CASO SEJA MENOR DO QUE 50

import xlsxwriter as opcoesDoXlsxWriter
import os

#1 - indicando onde será criado o arquivo, seu nome e sua extensão. Importante a questão das barras duplas (testar).
nomeCaminhoArquivo = 'C:\\Users\\Windows\\Desktop\\Python Projetos\\xlsxwriter\FormatacaoCondicional.xlsx'
planilhaExcel = opcoesDoXlsxWriter.Workbook(nomeCaminhoArquivo)
sheetDados = planilhaExcel.add_worksheet("Dados") #Para renomear o nome da Sheet1 para Dados.

#variavel que armazena a formatação para números maiores ou iguais a 50
formatoMaior = planilhaExcel.add_format({ 'bg_color' : 'green',
                                          'font_color' : 'white'})

#variavel que armazena a formatação para números menores do que 50
formatoMenor = planilhaExcel.add_format({ 'bg_color' : 'red',
                                          'font_color' : 'white'})

#Aqui iremos inserir as colunas e os valores nas células
inserirDados = [
    ["Coluna 1", "Coluna 2", "Coluna 3", "Coluna 4"],
    [34, 50 ,12, 34],
    [23, 43, 76, 51],
    [43, 29, 34, 12],
    [29, 58 ,73, 19],
]

#
sheetDados.write('A1',"Célular com valores >= estão em verde e < 50 estão em vermelho")

#Criando um loop para verificar cada linha e seus valores
for linha, range in enumerate(inserirDados):
    sheetDados.write_row(linha + 2, 1, range)

#variavel e formatação caso seja maior ou igual a 50
sheetDados.conditional_format('B4:E7', {'type': 'cell',
                                        'criteria': '>=',
                                        'value': 50,
                                        'format' : formatoMaior})

#variavel e formatação caso seja menor do que 50
sheetDados.conditional_format('B4:E7', {'type': 'cell',
                                        'criteria': '<',
                                        'value': 50,
                                        'format' : formatoMenor})


#3 - Para fechar e salvar as informações
planilhaExcel.close()

#4 - Abrir o arquivo para verificar o resultado
os.startfile(nomeCaminhoArquivo)
