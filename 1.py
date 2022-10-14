import PySimpleGUI as sg
from openpyxl import load_workbook
import pandas as pd
from datetime import datetime, date
import datetime as dt

#sg.popup('Atenção!', 'Lembre de preencher todos os campos corretamente.Ex: Data = dd/mm/aaaa!')
def criar_cadastro():  # função par criar tela
    sg.theme('Reddit')

    layout = [

        [sg.Text('Informe os dados e Tipo do Serviço Executado')],

        [sg.Text('Tipo Pavimento',  size=(15, 1)), sg.Combo(['Asfaltico', 'Cimentado', 'Blocos'], key="tipo_pav")],
        [sg.Text('Equipe',          size=(15, 1)), sg.InputText(key="equipe")],
        [sg.Text('Data Lançamento', size=(15, 1)), sg.InputText(key="data_lancamento")],
        [sg.Text('Metragem 1x',     size=(15, 1)), sg.InputText(key="metragem_1x")],
        [sg.Text('Metragem 2x',     size=(15, 1)), sg.InputText(key="metragem_2x")],
        [sg.Text('Priorizar',       size=(15, 1)), sg.InputText(key="priorizar")],
        [sg.Button('Adicionar Novo'), sg.Exit('Sair')]
    ]

    return sg.Window('Cadastro Serviço de Pavimento', layout=layout, finalize=True)


def append_df_to_excel(df, excel_path):
    df_excel = pd.read_excel(excel_path)  # vai ler o arquivo criado
    result = pd.concat([df_excel, df], ignore_index=True)
    result.to_excel(excel_path, index=False)

# Janela
janela = criar_cadastro()
lista_df = []
while True:
    event, values, = janela.read()  # sai do aplicativo ao clicar no x
    # if (event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT or event == 'Sair') and sg.popup_yes_no('Deseja realmente sair?') == 'Yes':
    if event == sg.WIN_CLOSED or event == 'Sair':
        break
    if event == 'Adicionar Novo':  # renovar a seção para inserir novos dados
        janela.close()
        janela = criar_cadastro()

    lista_apoio = []

    tipo_pav =          values['tipo_pav']
    equipe =            values['equipe']
    data_lancamento =   values['data_lancamento']
    metragem_1x =       values['metragem_1x']  # cadeias de keys
    metragem_2x =       values['metragem_2x']
    priorizar =         values['priorizar']
    print(values)

    lista_apoio.append(tipo_pav)
    lista_apoio.append(equipe)
    lista_apoio.append(data_lancamento)
    lista_apoio.append(metragem_1x)  # lista provisoria
    lista_apoio.append(metragem_2x)
    lista_apoio.append(priorizar)

    lista_df.append(lista_apoio)  # lista final

# Tratativa da planilha com Pandas

df = pd.DataFrame(lista_df,
                  columns=["Tipo do Pavimento", "Equipe", "Data Lançamento", "Metragem_1x", "Metragem_2x", "Priorizar"])


df['Metragem_1x'] = df['Metragem_1x'].str.replace(',','.').astype(float)
df['Metragem_2x'] = df['Metragem_2x'].str.replace(',','.').astype(float)


df['Data Atual'] = datetime.today()  # criando coluna com a data atual
df['Data Atual'] = pd.to_datetime(df['Data Atual'].dt.strftime('%m/%d/%Y'))
df['Data Atual'] = pd.to_datetime(df['Data Atual'], format='%d-%m-%Y')

df['Data Lançamento'] = pd.to_datetime(df['Data Lançamento'], format='%d/%m/%Y')
df['Atraso'] = df['Data Atual'] - df['Data Lançamento']


append_df_to_excel(df, r"Cadastro.xlsx")

print(df.dtypes)
janela.close()