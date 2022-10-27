# Controle de Servicos
 Pequeno sistema de cadastro e controle de demandas e serviços.. Ja Ja exportnado os itens para Exel. 

   
   
   
   
   
   ```python
    layout = [      
        
        [sg.Text('Informe os dados e Tipo do Serviço Executado')],
        
        [sg.Text('Tipo Pavimento',    size=(15, 1)), sg.Combo(['Asfaltico', 'Cimentado', 'Blocos'], key="tipo_pav")],
        [sg.Text('Equipe',            size=(15, 1)), sg.InputText(key="equipe")],
        [sg.Text('Data Lançamento',   size=(15, 1)), sg.InputText(key="data_lancamento")],
        [sg.Text('Metragem 1x',       size=(15, 1)), sg.InputText(key="metragem_1x")],
        [sg.Text('Metragem 2x',       size=(15, 1)), sg.InputText(key="metragem_2x")],
        [sg.Text('Priorizar',         size=(15, 1)), sg.InputText(key="priorizar")],
        [sg.Button('Adicionar Novo'), sg.Exit('Sair')]
    ]
 ```
 
 
 ![image](https://user-images.githubusercontent.com/33934341/195911171-40809fe1-cb6f-4c2c-949e-08752782269e.png)


```python
#Tratativa da planilha com Pandas

df = pd.DataFrame(lista_df, columns=["Tipo do Pavimento", "Equipe", "Data Lançamento", "Metragem_1x", "Metragem_2x", "Priorizar"])

df['Data Atual'] = datetime.now().strftime('%d/%m/%Y')                            #criando coluna com a data atual
df['Data Atual'] = pd.to_datetime(df['Data Atual'], format='%d/%m/%Y')              #convertendo coluna para data
df['Data Lançamento'] = pd.to_datetime(df['Data Lançamento'], format='%d/%m/%Y')    #convertendo coluna para data
df['Atraso'] = df['Data Atual'] - df['Data Lançamento']                             #criando coluna com dif de data

append_df_to_excel(df, r"Cadastro.xlsx")

janela.close()
```
