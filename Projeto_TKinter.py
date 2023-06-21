import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
import requests
from tkinter.filedialog import askopenfilename
import pandas as pd
from datetime import datetime
import numpy as np

window = tk.Tk()

window.title("Ferramenta de Cotação de Moedas")

requisicao = requests.get('https://economia.awesomeapi.com.br/json/last/USD-BRL,USD-BRLT,CAD-BRL,EUR-BRL,GBP-BRL,ARS-BRL,BTC-BRL,LTC-BRL,JPY-BRL,CHF-BRL,AUD-BRL,CNY-BRL,ILS-BRL,ETH-BRL,XRP-BRL,EUR-USD,CAD-USD,GBP-USD,ARS-USD,JPY-USD,CHF-USD,AUD-USD,CNY-USD,ILS-USD,BTC-USD,LTC-USD,ETH-USD,XRP-USD,BRL-USD,BRL-EUR,USD-EUR,CAD-EUR,GBP-EUR,ARS-EUR,JPY-EUR,CHF-EUR,AUD-EUR,CNY-EUR,ILS-EUR,BTC-EUR,LTC-EUR,ETH-EUR,XRP-EUR,DOGE-BRL,DOGE-EUR,DOGE-USD,USD-JPY,USD-CHF,USD-CAD,NZD-USD,USD-ZAR,USD-TRY,USD-MXN,USD-PLN,USD-SEK,USD-SGD,USD-DKK,USD-NOK,USD-ILS,USD-HUF,USD-CZK,USD-THB,USD-AED,USD-JOD,USD-KWD,USD-HKD,USD-SAR,USD-INR,USD-KRW,FJD-USD,GHS-USD,KYD-USD,SGD-USD,USD-ALL,USD-AMD,USD-ANG,USD-ARS,USD-AUD,USD-BBD,USD-BDT,USD-BGN,USD-BHD,USD-BIF,USD-BND,USD-BOB,USD-BSD,USD-BWP,USD-BZD,USD-CLP,USD-CNY,USD-COP,USD-CRC,USD-CUP,USD-DJF,USD-DOP,USD-DZD,USD-EGP,USD-ETB,USD-FJD,USD-GBP,USD-GEL,USD-GHS,USD-GMD,USD-GNF,USD-GTQ,USD-HNL,USD-HRK,USD-HTG,USD-IDR,USD-IQD,USD-IRR,USD-ISK,USD-JMD,USD-KES,USD-KHR,USD-KMF,USD-KZT,USD-LAK,USD-LBP,USD-LKR,USD-LSL,USD-LYD,USD-MAD,USD-MDL,USD-MGA,USD-MKD,USD-MMK,USD-MOP,USD-MRO,USD-MUR,USD-MVR,USD-MWK,USD-MYR,USD-NAD,USD-NGN,USD-NIO,USD-NPR,USD-NZD,USD-OMR,USD-PAB,USD-PEN,USD-PGK,USD-PHP,USD-PKR,USD-PYG,USD-QAR,USD-RON,USD-RSD,USD-RWF,USD-SCR,USD-SDG,USD-SOS,USD-STD,USD-SVC,USD-SYP,USD-SZL,USD-TND,USD-TTD,USD-TWD,USD-TZS,USD-UAH,USD-UGX,USD-UYU,USD-UZS,USD-VEF,USD-VND,USD-VUV,USD-XAF,USD-XCD,USD-XOF,USD-XPF,USD-YER,USD-ZMK,AED-USD,DKK-USD,HKD-USD,MXN-USD,NOK-USD,PLN-USD,RUB-USD,SAR-USD,SEK-USD,TRY-USD,TWD-USD,VEF-USD,ZAR-USD,UYU-USD,PYG-USD,CLP-USD,COP-USD,PEN-USD,NIO-USD,BOB-USD,KRW-USD,EGP-USD,USD-BYN,USD-MZN,INR-USD,JOD-USD,KWD-USD,USD-AZN,USD-CNH,USD-KGS,USD-TJS,USD-RUB,MYR-USD,UAH-USD,HUF-USD,IDR-USD,USD-AOA,VND-USD,BYN-USD,XBR-USD,THB-USD,PHP-USD,USD-TMT,XAGG-USD,USD-MNT,USD-AFN,AFN-USD,SYP-USD,IRR-USD,IQD-USD,USD-NGNI,USD-ZWL,BRL-ARS,BRL-AUD,BRL-CAD,BRL-CHF,BRL-CLP,BRL-DKK,BRL-HKD,BRL-JPY,BRL-MXN,BRL-SGD,SGD-BRL,AED-BRL,BRL-AED,BRL-BBD,BRL-BHD,BRL-CNY,BRL-CZK,BRL-EGP,BRL-GBP,BRL-HUF,BRL-IDR,BRL-ILS,BRL-INR,BRL-ISK,BRL-JMD,BRL-JOD,BRL-KES,BRL-KRW,BRL-LBP,BRL-LKR,BRL-MAD,BRL-MYR,BRL-NAD,BRL-NOK,BRL-NPR,BRL-NZD,BRL-OMR,BRL-PAB,BRL-PHP,BRL-PKR,BRL-PLN,BRL-QAR,BRL-RON,BRL-RUB,BRL-SAR,BRL-SEK,BRL-THB,BRL-TRY,BRL-VEF,BRL-XAF,BRL-XCD,BRL-XOF,BRL-ZAR,BRL-TWD,DKK-BRL,HKD-BRL,MXN-BRL,NOK-BRL,NZD-BRL,PLN-BRL,SAR-BRL,SEK-BRL,THB-BRL,TRY-BRL,TWD-BRL,VEF-BRL,ZAR-BRL,BRL-PYG,BRL-UYU,BRL-COP,BRL-PEN,BRL-BOB,CLP-BRL,PYG-BRL,UYU-BRL,COP-BRL,PEN-BRL,BOB-BRL,RUB-BRL,INR-BRL,EUR-GBP,EUR-JPY,EUR-CHF,EUR-AUD,EUR-CAD,EUR-NOK,EUR-DKK,EUR-PLN,EUR-NZD,EUR-SEK,EUR-ILS,EUR-TRY,EUR-THB,EUR-ZAR,EUR-MXN,EUR-SGD,EUR-HUF,EUR-HKD,EUR-CZK,EUR-KRW,BHD-EUR,EUR-AED,EUR-AFN,EUR-ALL,EUR-ANG,EUR-ARS,EUR-BAM,EUR-BBD,EUR-BDT,EUR-BGN,EUR-BHD,EUR-BIF,EUR-BND,EUR-BOB,EUR-BSD,EUR-BWP,EUR-BYN,EUR-BZD,EUR-CLP,EUR-CNY,EUR-COP,EUR-CRC,EUR-CUP,EUR-CVE,EUR-DJF,EUR-DOP,EUR-DZD,EUR-EGP,EUR-ETB,EUR-FJD,EUR-GHS,EUR-GMD,EUR-GNF,EUR-GTQ,EUR-HNL,EUR-HRK,EUR-HTG,EUR-IDR,EUR-INR,EUR-IQD,EUR-IRR,EUR-ISK,EUR-JMD,EUR-JOD,EUR-KES,EUR-KHR,EUR-KWD,EUR-KYD,EUR-KZT,EUR-LAK,EUR-LBP,EUR-LKR,EUR-LSL,EUR-LYD,EUR-MAD,EUR-MDL,EUR-MGA,EUR-MKD,EUR-MMK,EUR-MOP,EUR-MRO,EUR-MUR,EUR-MWK,EUR-MYR,EUR-NAD,EUR-NGN,EUR-NIO,EUR-NPR,EUR-OMR,EUR-PAB,EUR-PEN,EUR-PGK,EUR-PHP,EUR-PKR,EUR-PYG,EUR-QAR,EUR-RON,EUR-RSD,EUR-RWF,EUR-SAR,EUR-SCR,EUR-SDG,EUR-SDR,EUR-SOS,EUR-STD,EUR-SVC,EUR-SYP,EUR-SZL,EUR-TND,EUR-TTD,EUR-TWD,EUR-TZS,EUR-UAH,EUR-UGX,EUR-UYU,EUR-UZS,EUR-VEF,EUR-VND,EUR-XAF,EUR-XOF,EUR-XPF,EUR-ZMK,GHS-EUR,NZD-EUR,SGD-EUR,AED-EUR,DKK-EUR,EUR-XCD,HKD-EUR,MXN-EUR,NOK-EUR,PLN-EUR,RUB-EUR,SAR-EUR,SEK-EUR,TRY-EUR,TWD-EUR,VEF-EUR,ZAR-EUR,MAD-EUR,KRW-EUR,EGP-EUR,EUR-MZN,INR-EUR,JOD-EUR,KWD-EUR,EUR-AZN,EUR-AMD,EUR-TJS,EUR-RUB,HUF-EUR,GEL-EUR,EUR-GEL,IDR-EUR,EUR-AOA,BYN-EUR,XAGG-EUR,PEN-EUR')
dicionario_moedas = requisicao.json()
lista_moedas = list(dicionario_moedas.keys())


# print(dicionario_moedas.keys())

def pegar_cotacao():
    moeda = combobox_selecionarmoeda.get()
    dicionario_moedas2 = requisicao.json()[f'{moeda}']
    code = dicionario_moedas2['code']
    codein = dicionario_moedas2['codein']
    data_cotacao = calendario_moeda.get()
    ano = data_cotacao[-4:]
    mes = data_cotacao[3:5]
    dia = data_cotacao[:2]
    link = f"https://economia.awesomeapi.com.br/json/daily/{code}-{codein}/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}"
    requisicao_moedas = requests.get(link)
    cotacao = requisicao_moedas.json()[0]['bid']
    label_textocotacao['text'] = f"A cotação da {moeda} no dia {data_cotacao} foi de: R${cotacao}"


def selecionar_arquivo():
    caminho_arquivo = askopenfilename(title="Selecione o Arquivo de Moeda")
    var_caminhoarquivo.set(caminho_arquivo)
    if caminho_arquivo:
        label_arquivoselecionado['text'] = f"Arquivo Selecionado: {caminho_arquivo}"


def atualizar_cotacoes():
    try:
        df = pd.read_excel(var_caminhoarquivo.get())
        moedas = df.iloc[:, 0]
        data_inicial = calendario_datainicial.get()
        ano_inicial = data_inicial[-4:]
        mes_inicial = data_inicial[3:5]
        dia_inicial = data_inicial[:2]
        data_final = calendario_datafinal.get()
        ano_final = data_final[-4:]
        mes_final = data_final[3:5]
        dia_final = data_final[:2]

        for moeda in moedas:
            link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}/?" \
                   f"start_date={ano_inicial}{mes_inicial}{dia_inicial}&" \
                   f"end_date={ano_final}{mes_final}{dia_final}"
            print(link)
            requisicao_moeda = requests.get(link)
            cotacoes = requisicao_moeda.json()
            for cotacao in cotacoes:
                timestamp = int(cotacao["timestamp"])
                bid = float(cotacao["bid"])
                data = datetime.fromtimestamp(timestamp)
                data = data.strftime('%d/%m/%Y')
                if data not in df:
                    # print(data)
                    df[data] = np.nan

                df.loc[df.iloc[:, 0] == moeda, data] = bid
                print(bid)
        df.to_excel("Teste.xlsx")
        label_atualizarcotacoes['text'] = "Aquivo Atualizado com Sucesso"
    except:
        label_atualizarcotacoes['text'] = "Selecione um arquivo Excel no Formato Correto"


label_cotacaomoeda = tk.Label(text="Cotação de 1 moeda específica", borderwidth=2, relief='solid')
label_cotacaomoeda.grid(row=0, column=0, padx=10, pady=10, sticky='nsew', columnspan=3)

label_selecionarmoeda = tk.Label(text="Selecionar Moeda", anchor='w')
label_selecionarmoeda.grid(row=1, column=0, padx=10, pady=10, sticky='nsew', columnspan=2)

combobox_selecionarmoeda = ttk.Combobox(values=lista_moedas)
combobox_selecionarmoeda.grid(row=1, column=2, pady=10, padx=10, sticky='nsew')

label_selecionardia = tk.Label(text="Selecionar o dia que deseja pegar a cotação", anchor='w')
label_selecionardia.grid(row=2, column=0, padx=10, pady=10, sticky='nsew', columnspan=2)

calendario_moeda = DateEntry(year=2023, locale='pt_br')
calendario_moeda.grid(row=2, column=2, padx=10, pady=10, sticky='nsew')

label_textocotacao = tk.Label(text="")
label_textocotacao.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')

botao_pegarcotacao = tk.Button(text="Pegar Cotação", command=pegar_cotacao)
botao_pegarcotacao.grid(row=3, column=2, padx=10, pady=10, sticky='nsew')

label_cotacaovariasmoeda = tk.Label(text="Cotação de Múltiplas Moedas", borderwidth=2, relief='solid')
label_cotacaovariasmoeda.grid(row=4, column=0, padx=10, pady=10, sticky='nsew', columnspan=3)

label_selecionararquivo = tk.Label(text="Selecione um arquivo em Excel com as Moedas na Coluna A")
label_selecionararquivo.grid(row=5, column=0, padx=10, pady=10, sticky='nsew', columnspan=2)

var_caminhoarquivo = tk.StringVar()

botao_selecionararquivo = tk.Button(text="Clique aqui para selecionar", command=selecionar_arquivo)
botao_selecionararquivo.grid(row=5, column=2, padx=10, pady=10, sticky='nsew', columnspan=2)

label_arquivoselecionado = tk.Label(text="Nenhum Arquivo Selecionado", anchor='e')
label_arquivoselecionado.grid(row=6, column=0, padx=10, pady=10, sticky='nsew', columnspan=3)

label_datainicial = tk.Label(text="Data Inicial", anchor='w')
label_datainicial.grid(row=7, column=0, padx=10, pady=10, sticky='nsew')

label_datafinal = tk.Label(text="Data Final", anchor='w')
label_datafinal.grid(row=8, column=0, padx=10, pady=10, sticky='nsew')

calendario_datainicial = DateEntry(year=2023, locale='pt_br')
calendario_datainicial.grid(row=7, column=1, padx=10, pady=10, sticky='nsew')

calendario_datafinal = DateEntry(year=2023, locale='pt_br')
calendario_datafinal.grid(row=8, column=1, padx=10, pady=10, sticky='nsew')

botao_atualizarcotacoes = tk.Button(text="Atualizar Cotações", command=atualizar_cotacoes)
botao_atualizarcotacoes.grid(row=9, column=0, padx=10, pady=10, sticky='nsew')

label_atualizarcotacoes = tk.Label(text="")
label_atualizarcotacoes.grid(row=9, column=1, columnspan=2, padx=10, pady=10, sticky='nsew')

botao_fechar = tk.Button(text='Fechar', command=window.quit)
botao_fechar.grid(row=10, column=2, padx=10, pady=10, sticky='nsew')

window.mainloop()
