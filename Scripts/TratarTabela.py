'''Jordan Dias - 09/03/2022
Tratar Tabela: Tem o objetivo de transformar o texto das descrições para minúsculo e 
remover acentos das colunas de texto e remover tudo que for diferente de número das 
colunas de códigos exceto da coluna Código Seção.'''

def Tratar_Tabela(Input_Filename,Output_Filename):

#importando bibliotecas utilizadas
    import pandas as pd
    from unidecode import unidecode
    import openpyxl
    import xlsxwriter
    
#Leitura da planilha original
    tabela = pd.read_excel(Input_Filename)
    
           
#Deixando texto minusculo na descrição e deixando só numeros nos códigos
    for column in tabela.columns:
        tabela[column] = tabela[column].astype(str).str.lower()
        if "Código" in tabela[column].name:
            if tabela[column].name != "Código Seção":
                tabela[column] = (tabela[column].astype(str).str.findall('(\d+)').str.join(""))

#Substituindo caracteres especiais   
    for column in tabela.columns:
        for index, row in tabela.iterrows():
            row[column] = unidecode(row[column])
            
#Salvando planilha
    writer = pd.ExcelWriter(Output_Filename, engine='xlsxwriter')
    tabela.to_excel(writer, sheet_name="CNAE",index=False)
    worksheet = writer.sheets["CNAE"]
    
#Ajustando tamanho das colunas
    for idx, col in enumerate(tabela): 
        series = tabela[col]
        max_len = max((series.astype(str).map(len).max(),len(str(series.name))))+1
        worksheet.set_column(idx, idx, max_len)
    writer.save()
    