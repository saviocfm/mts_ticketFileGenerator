import os
import xml.etree.ElementTree as ET
from datetime import datetime
from openpyxl import load_workbook

def getFatura(lote):
    faturas = [
[433706,39240361],
[449520,40793139],
[441888,40033635],
[449703,40795549],
[442017,40035226],
[449848,40797136],
[449617,40794035],
[449665,40794892],
[441879,40033478],
[449699,40795392],
[433908,39242914],
[441891,40033765],
[449708,40795679],
[441893,40033788],
[449709,40795702],
[449861,40797274],
[441936,40034416],
[449571,40796331],
[441935,40034414],
[449750,40796329],
[449853,40797220],
[426151,38543374],
[450029,40798817],
[434247,39246104],
[442217,40036917],
[450032,40798822],
[442219,40036918],
[450034,40798823],
[442218,40036919],
[450038,40798824],
[442226,40036920],
[450039,40798825],
[442228,40036921],
[450035,40798826],
[442227,40036922],
[450036,40798827],
[442229,40036923],
[450070,40799190]
]
    #print(lote)
    for i in faturas:
        if str(i[0]) == lote:
            return str(i[1])
            break
        else:
            continue


def buscarArquivos(pasta):
    nomesArquivos = []
    arquivosRetorno = []
    for diretorio, subpastas, arquivos in os.walk(pasta):
        for arquivo in arquivos:
            nomesArquivos.append(str(os.path.realpath(arquivo))[71:])
    for i in nomesArquivos:
        file = open(pasta+'/'+i, 'rb')
        arquivosRetorno.append({
            'nome': i,
            'dados': file.read()
        })
    return(arquivosRetorno)


def formataTransacao():
    arquivos = buscarArquivos('./xlsx')
    retorno = ''
    for arquivo in arquivos:
        wb = load_workbook(filename='./xlsx/'+arquivo['nome'])
        sheet = ''
        try:
            sheet = wb['Plan1']
        except:
            try:
                sheet = wb['Transações']
            except:
                sheet = wb['Planilha1']
        # A CODIGO TRANSACAO
        # D DATA TRANSACAO
        # E PLACA
        # S VALOR EMISSAO
        for i in range(2, len(sheet['A'])+1):
        #for i in range(2, 23):    
            #print(i)
            #print(sheet['A'+str(i)].value,type(sheet['D'+str(i)].value))
            #print(type(None))
            if(type(sheet['D'+str(i)].value) == type(None)):
                break
            else:
                linha = '00200'
                linha = linha + str(sheet['A'+str(i)].value)[:9]
                linha = linha + str(sheet['E'+str(i)].value)
                linha = linha + sheet['D'+str(i)].value.strftime("%d%m%Y")
                linha = linha + impostos(str(sheet['S'+str(i)].value), 10, 0)
                retorno = retorno+linha+'\n'
    return retorno


def formataFatura(xmls):
    retorno = []
    for xml in xmls:
        cabecalho = {
            'RazaoSocial': '',
            'Endereco': '',
            'Bairro': '',
            'Municipio': '',
            'CEP': '',
            'Cod_UF': '  ',
            'CNPJ_CPF': '              ',
            'Cod_Pais': '  ',
            'Inscricao_Estadu': '',
            'Num_NF': '',
            'Serie': '',
            'SubSerie': '000',
            'Dt_Lancamento': '        ',
            'Dt_Emissao': '        ',
            'CNPJ_Destino': '              ',
            'Chave': '00000000000000000000000000000000000000000000',
            #'Titulo': str(buscarArquivos('./fatura')[0]['nome'])
            'Titulo':getFatura(str(buscarArquivos('./xlsx')[0]['nome'])[0:6])
        }

        itemsDadosNf = []

        tree = ET.parse('./xml/'+xml['nome'])
        root = tree.getroot()
        detNumber = 0
        count = 10
        for elem in root.iter():
            if elem.tag[36:] == 'emit':
                for item in elem.iter():
                    if item.tag[36:] == 'xNome':
                        cabecalho['RazaoSocial'] = item.text

            if elem.tag[36:] == 'emit':
                for item1 in elem.iter():
                    if item1.tag[36:] == 'enderEmit':
                        for item2 in item1.iter():
                            if item2.tag[36:] == 'xLgr':
                                cabecalho['Endereco'] = item2.text

            if elem.tag[36:] == 'emit':
                for item1 in elem.iter():
                    if item1.tag[36:] == 'enderEmit':
                        for item2 in item1.iter():
                            if item2.tag[36:] == 'xBairro':
                                cabecalho['Bairro'] = item2.text

            if elem.tag[36:] == 'emit':
                for item1 in elem.iter():
                    if item1.tag[36:] == 'enderEmit':
                        for item2 in item1.iter():
                            if item2.tag[36:] == 'xMun':
                                cabecalho['Municipio'] = item2.text

            if elem.tag[36:] == 'emit':
                for item1 in elem.iter():
                    if item1.tag[36:] == 'enderEmit':
                        for item2 in item1.iter():
                            if item2.tag[36:] == 'CEP':
                                cabecalho['CEP'] = item2.text

            if elem.tag[36:] == 'emit':
                for item1 in elem.iter():
                    if item1.tag[36:] == 'enderEmit':
                        for item2 in item1.iter():
                            if item2.tag[36:] == 'UF':
                                cabecalho['Cod_UF'] = item2.text

            if elem.tag[36:] == 'emit':
                for item in elem.iter():
                    if item.tag[36:] == 'CNPJ':
                        cabecalho['CNPJ_CPF'] = item.text

            if elem.tag[36:] == 'emit':
                for item1 in elem.iter():
                    if item1.tag[36:] == 'enderEmit':
                        for item2 in item1.iter():
                            if item2.tag[36:] == 'cPais':
                                # cabecalho['Cod_Pais'] = item2.text
                                cabecalho['Cod_Pais'] = 'BR'

            if elem.tag[36:] == 'emit':
                for item in elem.iter():
                    if item.tag[36:] == 'IE':
                        cabecalho['Inscricao_Estadu'] = item.text


            if elem.tag[36:] == 'nNF':
                cabecalho['Num_NF'] = elem.text


            if elem.tag[36:] == 'serie':
                cabecalho['Serie'] = elem.text

            cabecalho['SubSerie'] = '   '


            if elem.tag[36:] == 'dhSaiEnt':
                cabecalho['Dt_Lancamento'] = str(
                    elem.text[8:10]+elem.text[5:7]+elem.text[0:4])


            if elem.tag[36:] == 'dhEmi':
                cabecalho['Dt_Emissao'] = str(
                    elem.text[8:10]+elem.text[5:7]+elem.text[0:4])


            if elem.tag[36:] == 'dest':
                for item1 in elem.iter():
                    if item1.tag[36:] == 'CNPJ':
                            cabecalho['CNPJ_Destino'] = item1.text


            if elem.tag[36:] == 'chNFe':
                cabecalho['Chave'] = elem.text

            #cabecalho['Titulo'] = '          '  # gerado na ticket

            cabecalho['RazaoSocial'] = acrescentarEspacos(
            cabecalho['RazaoSocial'], 70, 1)
            cabecalho['Endereco'] = acrescentarEspacos(
            cabecalho['Endereco'], 35, 1)
            cabecalho['Bairro'] = acrescentarEspacos(cabecalho['Bairro'], 35, 1)
            cabecalho['Municipio'] = acrescentarEspacos(
            cabecalho['Municipio'], 35, 1)
            cabecalho['CEP'] = acrescentarEspacos(cabecalho['CEP'], 10, 1)
            cabecalho['Cod_UF'] = acrescentarEspacos(cabecalho['Cod_UF'], 2, 1)
            cabecalho['CNPJ_CPF'] = acrescentarEspacos(
            cabecalho['CNPJ_CPF'], 14, 1)
            cabecalho['Cod_Pais'] = acrescentarEspacos(cabecalho['Cod_Pais'], 2, 1)
            cabecalho['Inscricao_Estadu'] = acrescentarZeros(
            cabecalho['Inscricao_Estadu'], 20, 0)
            cabecalho['Num_NF'] = acrescentarZeros(cabecalho['Num_NF'], 6, 0)
            cabecalho['Serie'] = acrescentarZeros(cabecalho['Serie'], 3, 0)
            cabecalho['SubSerie'] = acrescentarEspacos(cabecalho['SubSerie'], 3, 1)
            cabecalho['Dt_Lancamento'] = acrescentarEspacos(
            cabecalho['Dt_Lancamento'], 8, 1)
            cabecalho['Dt_Emissao'] = acrescentarEspacos(
            cabecalho['Dt_Emissao'], 8, 1)
            cabecalho['CNPJ_Destino'] = acrescentarEspacos(
            cabecalho['CNPJ_Destino'], 14, 1)
            cabecalho['Chave'] = acrescentarEspacos(cabecalho['Chave'], 44, 1)
            cabecalho['Titulo'] = acrescentarEspacos(cabecalho['Titulo'], 10, 1)

                ###############
                # itemDadosNf #
                ###############


            if elem.tag[36:] == 'det' and int(elem.attrib['nItem']) == detNumber + 1:
                detNumber = detNumber + 1
                itemDadosNf = {
                    'Id_Item': acrescentarEspacos(str(count),7,1),
                    'Cod_Produto': '00000000000000000000',
                    'Descricao_Nota': '                                        ',
                    'Quantidade': '0000000000000',
                    'Unidade de medida': '000',
                    'Vlr_Unitario': '000000000000000',
                    'CFOP': '0000000000',
                    'Base_Calculo_ICMS': '000000000000000',
                    'Aliquota_ICMS': '00000000',
                    'Vlr_Imposto_ICMS': '000000000000000',
                    'Base_Calculo_IPI': '000000000000000',
                    'Aliquota_IPI': '00000000',
                    'Vlr_Imposto_IPI': '000000000000000'
                }

                count = count + 10

                for elem2 in elem.iter():
                    if elem2.tag[36:] == 'prod':
                        produto = elem2
                    if elem2.tag[36:] == 'imposto':
                        imposto = elem2

                for item1 in produto.iter():
                    if item1.tag[36:] == 'cProd':
                        itemDadosNf['Cod_Produto'] = item1.text

                for item1 in produto.iter():
                    if item1.tag[36:] == 'xProd':
                        itemDadosNf['Descricao_Nota'] = item1.text

                for item1 in produto.iter():
                    if item1.tag[36:] == 'qCom':
                        itemDadosNf['Quantidade'] = item1.text

                for item1 in produto.iter():
                    if item1.tag[36:] == 'uCom':
                        itemDadosNf['Unidade de medida'] = item1.text

                for item1 in produto.iter():
                    if item1.tag[36:] == 'vProd':
                        itemDadosNf['Vlr_Unitario'] = '%.2f' % float(item1.text)
                        itemDadosNf['Vlr_Unitario'] = itemDadosNf['Vlr_Unitario'].replace('.',',')

                for item1 in produto.iter():
                    if item1.tag[36:] == 'CFOP':
                        itemDadosNf['CFOP'] = item1.text

                
                if elem.tag[36:] == 'ICMSTot':
                    for item1 in elem.iter():
                        if item1.tag[36:] == 'vBC':
                            itemDadosNf['Base_Calculo_ICMS'] = item1.text

            
                if elem.tag[36:] == 'ICMSTot':
                    for item1 in elem.iter():
                        if item1.tag[36:] == 'vICMS':
                            itemDadosNf['Vlr_Imposto_ICMS'] = item1.text
                # 'Aliquota_ICMS':'',
                # 'Base_Calculo_IPI':'',
                # 'Aliquota_IPI':'',
                # 'Vlr_Imposto_IPI':''
                itemDadosNf['Id_Item'] = acrescentarEspacos(
                    itemDadosNf['Id_Item'], 4, 1)
                itemDadosNf['Cod_Produto'] = acrescentarEspacos(
                    itemDadosNf['Cod_Produto'], 20, 1)
                itemDadosNf['Descricao_Nota'] = acrescentarEspacos(
                    itemDadosNf['Descricao_Nota'], 40, 1)
                itemDadosNf['Quantidade'] = acrescentarZeros(
                    itemDadosNf['Quantidade'], 13, 0)
                itemDadosNf['Unidade de medida'] = acrescentarEspacos(
                    itemDadosNf['Unidade de medida'], 3, 1)
                itemDadosNf['Vlr_Unitario'] = acrescentarZeros(
                    itemDadosNf['Vlr_Unitario'], 15, 0)
                itemDadosNf['CFOP'] = acrescentarEspacos(
                    itemDadosNf['CFOP'], 10, 1)
                itemDadosNf['Base_Calculo_ICMS'] = impostos(
                    itemDadosNf['Base_Calculo_ICMS'], 15, 0)
                itemDadosNf['Aliquota_ICMS'] = impostos(
                    itemDadosNf['Aliquota_ICMS'], 8, 0)
                itemDadosNf['Vlr_Imposto_ICMS'] = impostos(
                    itemDadosNf['Vlr_Imposto_ICMS'], 15, 0)
                itemDadosNf['Base_Calculo_IPI'] = impostos(
                    itemDadosNf['Base_Calculo_IPI'], 15, 0)
                itemDadosNf['Aliquota_IPI'] = impostos(
                    itemDadosNf['Aliquota_IPI'], 8, 0)
                itemDadosNf['Vlr_Imposto_IPI'] = impostos(
                    itemDadosNf['Vlr_Imposto_IPI'], 15, 0)

                itemsDadosNf.append(itemDadosNf)
        retorno.append(
            {
                'cabecalho': cabecalho,
                'itemsDadosNf': itemsDadosNf
            }
        )
    return retorno


def impostos(valor, tamanhoCampo, direcao):
    if len(valor) > tamanhoCampo:
        valor = valor[:(tamanhoCampo)]
    valor = '%.2f' % float(valor)
    valor = str(valor).replace('.', ',')
    while len(valor) < tamanhoCampo:
        if direcao == 1:  # direita
            valor = valor+'0'
        if direcao == 0:  # esquerda
            valor = '0'+valor
    return str(valor)


def acrescentarEspacos(valor, tamanhoCampo, direcao):
    if len(valor) > tamanhoCampo:
        valor = valor[:(tamanhoCampo)]
    while len(valor) < tamanhoCampo:
        if direcao == 1:  # direita
            valor = valor+' '
        if direcao == 0:  # esquerda
            valor = ' '+valor
    return str(valor)


def acrescentarZeros(valor, tamanhoCampo, direcao):
    if len(valor) > tamanhoCampo:
        valor = valor[:(tamanhoCampo)]
    while len(valor) < tamanhoCampo:
        if direcao == 1:  # direita
            valor = valor+'0'
        if direcao == 0:  # esquerda
            valor = '0'+valor
    return str(valor)


def formatarTxt(arquivos):
    retorno = ''
    faturas = formataFatura(arquivos)
    for i in faturas:
        cabecalho = i['cabecalho']
        itemsDadosNf = i['itemsDadosNf']
        retorno = retorno+cabecalho['RazaoSocial']+cabecalho['Endereco']+cabecalho['Bairro']+cabecalho['Municipio']+cabecalho['CEP']+cabecalho['Cod_UF']+cabecalho['CNPJ_CPF']+cabecalho['Cod_Pais'] + \
            cabecalho['Inscricao_Estadu']+cabecalho['Num_NF']+cabecalho['Serie']+cabecalho['SubSerie'] + \
            cabecalho['Dt_Lancamento']+cabecalho['Dt_Emissao'] + \
            cabecalho['CNPJ_Destino'] + \
            cabecalho['Chave']+cabecalho['Titulo']+'\n'
        for itemDadosNf in itemsDadosNf:
            retorno = retorno+itemDadosNf['Id_Item'] + itemDadosNf['Cod_Produto'] + itemDadosNf['Descricao_Nota'] + itemDadosNf['Quantidade'] + itemDadosNf['Unidade de medida']+itemDadosNf['Vlr_Unitario'] + \
                itemDadosNf['CFOP']+itemDadosNf['Base_Calculo_ICMS'] + itemDadosNf['Aliquota_ICMS'] + itemDadosNf['Vlr_Imposto_ICMS'] + \
                itemDadosNf['Base_Calculo_IPI'] + \
                itemDadosNf['Aliquota_IPI'] + itemDadosNf['Vlr_Imposto_IPI']+'\n'
    retorno = retorno + formataTransacao()
    return retorno


def salvarArquivoTxt(arquivo):
    nome = buscarArquivos('./xlsx')[0]['nome']
    f = open('./txt/'+nome[:15]+'.txt', 'w')
    f.write(arquivo)
    f.close


def main():
    arquivosXml = buscarArquivos('./xml')
    arquivoTxt = formatarTxt(arquivosXml)
    salvarArquivoTxt(arquivoTxt)
    print(arquivoTxt)
    print('\n\n\nfinalizou')


main()
