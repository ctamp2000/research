'''Ler arquivo txt e grava xlsx'''
import openpyxl
wb = openpyxl.Workbook()
sheet = wb.active
#arquivo = open('acm-temp.txt','r')  # nome do arquivo texto , modo de abertura;
arquivo = open('acm-completo.txt','r')  # nome do arquivo texto , modo de abertura;
c1 = sheet.cell(row=1, column=1)
c1.value = 'author'
c1 = sheet.cell(row=1, column=2)
c1.value = 'title'
c1 = sheet.cell(row=1, column=3)
c1.value = 'abstract'
planilha_linha = 2
for linha in arquivo:   # ler 1 linha de cada vez
    valor = linha
    posicao = linha.find('author')
    if posicao != -1: # encontrou 'author'
        c1 = sheet.cell(row=planilha_linha, column=1)  # author
        c1.value = linha[posicao+10:linha.__len__()]  # nome
 #       print(planilha_linha,'author',linha[posicao+9:linha.__len__()])
    else:
        posicao = linha.find('title')
        if posicao != -1: # encontrou 'title'
            posicao = linha.find('ktitle')
            if posicao == -1: # nao eh booktitle
                c1 = sheet.cell(row=planilha_linha, column=2)  # title
                c1.value = linha[posicao+10:linha.__len__()]  # nome
 #               print('title',linha[posicao+8:linha.__len__()])
        else:
            posicao = linha.find('abstract')
            if posicao != -1: # encontrou 'abstract'
                c1 = sheet.cell(row=planilha_linha, column=3)  # abstract
                c1.value = linha[posicao+12:linha.__len__()]  # nome
 #               print('abstract',linha[posicao+11:linha.__len__()])
                planilha_linha += 1
#    print('for', planilha_linha)
wb.save('C:\\Users\\ctamp\\Downloads\\acm.xlxs')

