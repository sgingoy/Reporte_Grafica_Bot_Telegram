import requests
import pandas as pd
import matplotlib.pyplot as plt
import json
import glob, os
import time



hojas = list()
minutosT = 0
token_bot = '1942879271:AAHUiumoNSs5WaeVc9TzkV6bPllD05tfTVU'
ruta_fotos = r'C:\Users\ASUS TUF\Desktop\BotGraficas\Fotos'



class TelegranBot():
    def __init__(self):
        self.token = token_bot
        self.group = -524501869
        self.channel = None

    def get_me(self):
        url = f"https://api.telegram.org/bot{self.token}/getMe"
        response = requests.get(url)
        if response.status_code == 200:
            salida = json.loads(response.text)
            return salida
        return None

    def get_updates(self):
        url = f"https://api.telegram.org/bot{self.token}/getUpdates"
        response = requests.get(url)
        if response.status_code == 200:
            salida = json.loads(response.text)
            return salida
        return None

    def get_last_update(self):
        url = f"https://api.telegram.org/bot{self.token}/getUpdates?offset=-1"
        response = requests.get(url)
        if response.status_code == 200:
            salida = json.loads(response.text)
            for result in salida ["result"]:
                ultimo_mensaje = result["message"]["text"]
            return ultimo_mensaje
        return None

    def send_message_to_group(self,message):
        url = f"https://api.telegram.org/bot{self.token}/sendMessage"
        data = {"chat_id": self.group, "text":message}
        response = requests.post(url,data=data)
        if response.status_code == 200:
            salida = json.loads(response.text)
            return salida
        return None

    def send_photo_group(self, filename,caption):
        url = f"https://api.telegram.org/bot{self.token}/sendPhoto"
        data = {"chat_id": self.group, "caption": caption}
        os.chdir(ruta_fotos)
        files = {"photo":(filename,open(filename,'rb'))}
        response = requests.post(url, data=data, files = files)
        if response.status_code == 200:
            salida = json.loads(response.text)
            return salida
        error = json.loads(response.text)
        error_code = error['error_code']
        description = error['description']
        msg = f'Error:{error_code}.Description:{description}'
        raise Exception(msg)

class Hoja():

    def seleccionar_hoja(self, numero_hoja):
        df = pd.read_excel(r'C:\Users\ASUS TUF\Desktop\BotGraficas\9.GEPACK SEPTIEMBRE.xlsm', sheet_name=numero_hoja, skiprows=33, nrows=56, usecols='C:X')
        return df

    def llenar_datos(self,df):
        for i, row in df.iterrows():
            Aviso = row["Aviso"]
            Check_Datos = row["Check Datos"]
            Check_Maquina = row["Check Maquina"]
            hora = row["Hora"]
            Orden = row["Orden"]
            Prod_Horaria = row["Prod. Horaria"]
            Turno = row["Turno"]
            Causa = row["Causa"]
            Codigo = row["Codigo"]
            Equipo = row["Equipo"]
            Subconjunto = row["Subconjunto"]
            Componente = row["Componente"]
            Modo_de_Fallo = row["Modo de Fallo"]
            Minutos = row["Minutos"]
            Descripción_Breve = row["Descripción Breve"]
            TEXTO_AMPLIADO = row["TEXTO AMPLIADO"]
            T_no_uso = row["T. No Uso"]
            T_programadas = row["T. Programadas"]
            T_externas = row["T. Externas"]
            T_internas = row["T. Internas"]
            Hora1 = row["Hora"]
            Turno1 = row["Turno"]
            hojas.append(
                Informe(Aviso, Check_Datos, Check_Maquina, hora, Orden, Prod_Horaria, Turno, Causa, Codigo, Equipo,
                        Subconjunto, Componente, Modo_de_Fallo, Minutos, Descripción_Breve, TEXTO_AMPLIADO,
                        T_no_uso, T_programadas, T_externas, T_internas, Hora1, Turno1))
        return  hojas

    def reporte(self,hojas,hx):
        problemas = " "
        hx.sort_values('Minutos Totales', inplace=True)
        maximo_buscar = hx["Equipo"].value_counts().index.tolist()[-1]
        print(maximo_buscar)
        for i in range(len(hojas)):
            if (maximo_buscar == hojas[i].Equipo):
                problemas += hojas[i].Equipo+" : "+ hojas[i].Descripción_Breve+" : "+ str(hojas[i].Minuto)+" Minutos"+"\n"
        return  problemas

    def reporte2(self,hojas,hx):
        problemas2 = " "
        hx.sort_values('Minutos Totales', inplace=True)
        maximo_buscar2 = hx["Equipo"].value_counts().index.tolist()[-2]
        print(maximo_buscar2)
        for i in range(len(hojas)):
            if (maximo_buscar2 == hojas[i].Equipo):
                problemas2 += hojas[i].Equipo+" : "+ hojas[i].Descripción_Breve+" : "+ str(hojas[i].Minuto)+" Minutos"+"\n"
        return  problemas2

   # def reporte(self, hojas, listaMayor):
    #    problemas = " "
    #    lista_nombre = df.Equipo.value_counts().index.tolist()
    #    for o in range(len(lista_nombre)):
    #        for i in range(len(hojas)):
    #            if (listaMayor == hojas[i].Equipo):
    #                problemas += hojas[i].Equipo + " : " + hojas[i].Descripción_Breve + " : " + str(
    #                    hojas[i].Minuto) + "\n"
    #    return problemas

    def crea_dataframe(self,df, hojas):
        minutosT = 0
        lista_nombre = df.Equipo.value_counts().index.tolist()

        data0 = {'Equipo': [],
                 'Minutos Totales': []}
        hx = pd.DataFrame(data0, columns=['Equipo', 'Minutos Totales'])
        for o in range(len(lista_nombre)):
            for i in range(len(hojas)):
                if (lista_nombre[o] == hojas[i].Equipo):
                    minutosT += hojas[i].Minuto
            data = {'Equipo': [lista_nombre[o]],
                    'Minutos Totales': [minutosT]}
            hf = pd.DataFrame(data, columns=['Equipo', 'Minutos Totales'])
            hx = hx.append(hf, ignore_index=True)
            minutosT = 0
        return hx

    def Crear_grafico(self,hx,tipog_grafico):
        hx.groupby('Equipo')["Minutos Totales"].sum().plot(kind=tipog_grafico,legend = 'Reverse',figsize=(15,5))
        os.chdir(ruta_fotos)
        plt.savefig("prueba.png", dpi=500)
        return "prueba.png"

class Informe:
    def __init__(self, Aviso, Check_Datos, Check_Maquina, Hora, Orden, Prod_Horaria, Turno, Causa, Codigo, Equipo,
             Subconjunto, Componente, Modo_Fallo
             , Minuto, Descripción_Breve, TEXTO_AMPLIADO, T_No_Uso, T_Programadas, T_Externas, T_Internas, Hora1,
             Turno1):
        self.Aviso = Aviso
        self.Check_Datos = Check_Datos
        self.Check_Maquina = Check_Maquina
        self.Hora = Hora
        self.Orden = Orden
        self.Prod_Horaria = Prod_Horaria
        self.Turno = Turno
        self.Causa = Causa
        self.Codigo = Codigo
        self.Equipo = Equipo
        self.Subconjunto = Subconjunto
        self.Componente = Componente
        self.Modo_Fallo = Modo_Fallo
        self.Minuto = Minuto
        self.Descripción_Breve = Descripción_Breve
        self.TEXTO_AMPLIADO = TEXTO_AMPLIADO
        self.T_No_Uso = T_No_Uso
        self.T_Programadas = T_Programadas
        self.T_Externas = T_Externas
        self.T_Internas = T_Internas
        self.Hora1 = Hora1
        self.Turno1 = Turno1

bots = TelegranBot()
df = pd.read_excel(r'9.GEPACK SEPTIEMBRE.xlsm', sheet_name=1, skiprows=33, nrows=56, usecols='C:X')
excel_hoja = Hoja()

while(True):
    if(bots.get_last_update() == "hola bot"):
        bots.send_message_to_group("Dime un dia del mes")
        time.sleep(7)
        opcion_hoja = str(bots.get_last_update())
        if 0 < int(opcion_hoja) < 32:
            bots.send_message_to_group("Seleccionando....")
            df = excel_hoja.seleccionar_hoja(opcion_hoja)
         #   serie = df.Equipo.value_counts()
         #   serie = serie.sort_values(ascending=False)
         #   listaMayor = df.Equipo.value_counts().index.tolist()[0]
            hojas = excel_hoja.llenar_datos(df)
            hx = excel_hoja.crea_dataframe(df, hojas)
            reporte = excel_hoja.reporte(hojas,hx)
            reporte2 = excel_hoja.reporte2(hojas,hx)
            bots.send_message_to_group("Generando Grafico....")
            #time.sleep(7)
            opcion_grafico = bots.get_last_update()
            #if(opcion_grafico == "barras"):
            #bots.send_message_to_group("Un momento por favor..")
            file_name = excel_hoja.Crear_grafico(hx,'barh')
            bots.send_photo_group(file_name,"Grafica del dia "+ opcion_hoja + " maximo "+ str(hx["Minutos Totales"].sort_values().tolist()[-1]))
            bots.send_message_to_group(reporte)
            bots.send_message_to_group(reporte2)
            #else:
            #    bots.send_message_to_group("No tengo la opcion de crear ese grafico..")
        else:
            bots.send_message_to_group("El valor tiene que ser: (1-31)")