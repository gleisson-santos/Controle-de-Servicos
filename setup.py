import PySimpleGUI as sg
from openpyxl import load_workbook
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows

def criar_cadastro(): #função par criar tela
    sg.theme('Reddit')   
       
    layout = [
        [sg.Text('Informe os dados e Tipo do Serviço Executado')],
        [sg.Text('Tipo do Pavimento', size=(15, 1)), sg.InputText(key="tipo_pav")],
        [sg.Text('Equipe',            size=(15, 1)), sg.InputText(key="equipe")],
        [sg.Text('Data',              size=(15, 1)), sg.InputText(key="data")],
        [sg.Text('Metragem 1x',       size=(15, 1)), sg.InputText(key="metragem_1x")],
        [sg.Text('Metragem 2x',       size=(15, 1)), sg.InputText(key="metragem_2x")],
        [sg.Text('Priorizar',         size=(15, 1)), sg.InputText(key="priorizar")],
        [sg.Button('Adicionar Novo'), sg.Cancel()]
    ]

                   
    return sg.Window('Cadastro Serviço de Pavimento', layout=layout,finalize=True)

# Janela
janela = criar_cadastro()
lista_df = []
while True:
    event, values, = janela.read() #sai do aplicativo ao clicar no x
    if event == sg.WIN_CLOSED or event == 'Cancel':
        break

    if event == 'Adicionar Novo':  #renovar a seção para inserir novos dados
        janela.close()
        janela = criar_cadastro()
        
    lista_apoio = []
    tipo_pav        = values['tipo_pav']
    equipe          = values['equipe']
    data            = values['data']
    metragem_1x     = values['metragem_1x']  #cadeias de keys
    metragem_2x     = values['metragem_2x']
    priorizar       = values['priorizar']
    print(values)

    lista_apoio.append(tipo_pav)
    lista_apoio.append(equipe)
    lista_apoio.append(data)
    lista_apoio.append(metragem_1x) #lista provisoria
    lista_apoio.append(metragem_2x)
    lista_apoio.append(priorizar)
    
    lista_df.append(lista_apoio)    #lista final

df = pd.DataFrame(lista_df, columns=["Tipo do Pavimento", "Equipe", "Data", "Metragem_1x", "Metragem_2x", "Priorizar"])

rows = df.values.tolist()
workbook = load_workbook("Cadastro.xlsx")
sheet = workbook['Sheet1']
for row in rows:
    sheet.append(row)
workbook.save('Cadastro.xlsx')

janela.close()