'''
Se necesita tener instalado <<openpyxl>>, si no se tiene usar pip para instalarlo
pip install openpyxl
'''

#*------------- Librerias necesarias -------------*#

import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import *
import tkinter.messagebox as tmsg
from tkinter import filedialog, ttk
import time

#!------ Función Para Generar el EBCDIC Header -----!#

createdFiles = [] #* Lista para reportar el número de archivos generados.

'''
Esta función depende de 3 variables:
1.  Archivo Excel con toda la información de cada una de las líneas
2.  Archivo de texto con la descripción de almacenamiento de los headers
    en los diferentes bytes y la secuencia de procesamiento
3.  Tipo de proceso (Gathers, In-In, etc)

Una vez el archivo Excel se carga, éste se lee con Pandas y se genera un dataframe
con el cual se definen diferentes variables según su posición dentro del dataframe,
las cuales serviran para llenar cada uno de los parámetros de adquisición y geofísicos
requeridos para generar las primeras 16 líneas del EBCDIC header (C1 - C16), información
que es obtenida de cada una de las columnas del dataframe.

Una vez se tienen las 16 líneas del EBCDIC header, este se concatena con el archivo de
texto que tiene la información de posición de bytes y secuencia de procesamiento, y que va
de la línea C17 a la C40.
'''

def textHeader():
    try:
        #! Declaración de archivos de entrada y ruta de salida
        excelFile = pathXlsx.get()
        textFile = pathTxt.get()
        path = pathToSave.get()
        userInput = userSelect.get()

        if excelFile != '' and textFile != '' and path != '':

            #! Lectura del archivo Excel y creación del dataframe
            df = pd.read_excel(excelFile)

            #! Ciclo for para generar cada uno de los archivos de texto EBCDIC header
            for i in range(len(df)):

                #! Definición de variables necesarias para generar el EBCDIC header
                lineName = df.iloc[i,0]
                survey = df.iloc[i,1]
                contract = df.iloc[i,2]
                basin = df.iloc[i,3]
                country = df.iloc[i,4]
                area = df.iloc[i,5]
                client = df.iloc[i,6]
                recordDate = df.iloc[i,7]
                recordedBy = df.iloc[i,8]
                processedBy = df.iloc[i,9]
                processingDate = df.iloc[i,10]
                fstSp = df.iloc[i,11]
                lstSp = df.iloc[i,12]
                fstRcvr = df.iloc[i,13]
                lstRcvr = df.iloc[i,14]
                totSps = df.iloc[i,15]
                fstCdp = df.iloc[i,16]
                lstCdp = df.iloc[i,17]
                fold = df.iloc[i,18]
                sampInt = df.iloc[i,19]
                recLength = df.iloc[i,20]
                spInt = df.iloc[i,21]
                grpInt = df.iloc[i,22]
                totChan = df.iloc[i,23]
                srcType = df.iloc[i,24]

                #! Definición del sufijo que describe el nombre del EBCDIC header
                nameSufix = ''
                if userInput == 'G':
                    nameSufix = 'Gathers'
                elif userInput == 'I':
                    nameSufix = 'In-In'
                elif userInput == 'O':
                    nameSufix = 'Out-Out'
                elif userInput == 'V':
                    nameSufix = 'Vels'
                elif userInput == 'E':
                    nameSufix = 'Eta'

                #! Creación del archivo de texto EBCDIC header
                with open(path + "/" + lineName + "_" + nameSufix + "_hdr.txt", "w") as file:
                    file.write("C 1 LINEA: " + lineName + "    PROGRAMA: " + survey + "\n")
                    if userInput == 'G':
                        file.write("C 2 PROCESO: CDP GATHERS MIGRADOS PRE APILADO EN TIEMPO\n")
                    elif userInput == 'I':
                        file.write("C 2 PROCESO: MIGRACION PRE APILADO EN TIEMPO IN-IN (CON FILTROS Y ESCALARES)\n")
                    elif userInput == 'O':
                        file.write(
                            "C 2 PROCESO: MIGRACION PRE APILADO EN TIEMPO OUT-OUT (SIN FILTROS NI ESCALARES)\n")
                    elif userInput == 'V':
                        file.write(
                            "C 2 PROCESO: VELOCIDADES DE MIGRACION PRE APILADO EN TIEMPO\n")
                    elif userInput == 'E':
                        file.write(
                            "C 2 PROCESO: CAMPO DE ANISOTROPIA ETA\n")
                    file.write("C 3 CONTRATO: " + contract +"\n")
                    file.write("C 4 CUENCA: " + basin + "\n")
                    file.write("C 5 PAIS Y AREA: " + country + ", " + area + "\n")
                    file.write("C 6 CLIENTE: " + client + "\n")
                    file.write("C 7 FECHA DE REGISTRO: " + str(recordDate) + "\n")
                    file.write("C 8 REGISTRADO POR: " + recordedBy + "\n")
                    file.write("C 9 PROCESADO POR: " + processedBy +
                                "    PROCESADO PARA: " + client + "\n")
                    file.write("C10 FECHA DE PROCESAMIENTO: " + processingDate + "\n")
                    file.write("C11 RANGO DE FUENTES: " + str(fstSp) +
                                "-" + str(lstSp) + "    RANGO DE RECEPTORES: " + str(fstRcvr) + "-" + str(lstRcvr) + "\n")
                    file.write("C12 NUMERO DE FUENTES: " + str(totSps) + "\n")
                    file.write("C13 RANGO DE CDP: " + str(fstCdp) + "-" +
                                str(lstCdp) + "    CUBRIMIENTO: " + str(fold) + "\n")
                    file.write("C14 RAZON DE MUESTREO: " + str(sampInt) +
                                " ms     LONGITUD DE REGISTRO: " + str(recLength) + " s\n")
                    file.write("C15 INT. DE FUENTES: " + str(spInt) +
                                " m    INT. DE RECEPTORES: " + str(grpInt) + " m    CANALES: " + str(totChan) + "\n")
                    file.write("C16 TIPO DE FUENTE: " + srcType + "\n")

                    #! Lee el archivo con la secuencia de procesamiento y lo concatena al EBCDIC header
                    with open(textFile) as f:
                        for line in f:
                            file.write(line)

            #! Lista con el nombre de cada uno de los archivos EBCDIC Header creados
            os.chdir(path)
            createdFiles = [
                f for f in os.listdir() if os.path.isfile(f) and f.endswith(nameSufix + '_hdr.txt')
                ]
            #! Imprime por pantalla la cantidad de archivos generados y el nombre de cada uno
            if len(createdFiles) != 0:
                tmsg.showinfo(
                    message ='Se generaron satisfactoriamente ' + str(len(createdFiles)
                    ) + ' archivos EBCDIC Header pertenecientes a datos PSTM ' + nameSufix,
                    title='Reporte de Ejecución'
                    )
            else:
                tmsg.showinfo(
                    message ='No se generó ningún archivo EBCDIC Header',
                    title='Reporte de Ejecución'
                    )
        else:
            tmsg.showerror(
                message='Revisar que los archivo XLSX, TXT y la ruta de salida existan',
                title='Error'
                )

    except OSError:
        tmsg.showerror(
            message = "Algo salió mal.\nPor Favor Revisar los Parámetros de Entrada",
            title = 'Error')


#*------------- Lógica Gráfica (Función Principal) -------------*#

if __name__ == "__main__":

    #! Window Definition and Configuration
    root = Tk()
    root.geometry('1200x640')
    root.resizable(0,0)
    root.iconbitmap('./img/binaryCode.ico')
    root.title('EBCDIC Generator')

    #! Importar Imágenes
    imgXlsx = PhotoImage(file='./img/imgExcel.png')
    imgTxt = PhotoImage(file='./img/imgTxt.png')
    imgPathToSave = PhotoImage(file='./img/folder1.png')
    imgMenuExit = PhotoImage(file='./img/imgShutdown.png')

    #! Definición de Funciones Gráficas Auxiliares
    def selectFileXlsx():
        #* Seleccionar Archivo XLSX
        filetypesXlsx = (
            ('xlsx files', '*.xlsx'),
            ('xls files', '*.xls')
        )
        fileNameXlsx = filedialog.askopenfilename(
            title='Open file',
            initialdir='/',
            filetypes=filetypesXlsx
            )
        pathXlsx.set(fileNameXlsx)

    def selectFileTxt():
        #* Seleccionar Archivo TXT
        filetypesTxt = (
            (('text files', '*.txt'),)
        )
        fileNameTxt = filedialog.askopenfilename(
            title='Open file',
            initialdir='/',
            filetypes=filetypesTxt
            )
        pathTxt.set(fileNameTxt)

    def selectFolder():
        #* Seleccion folder
        filePath = filedialog.askdirectory()
        pathToSave.set(filePath)

    def closeWindow():
        #* Confirmación Cierre de Ventana
        closeWin = tmsg.askyesno(
            message = "¿Está seguro que desea cerrar la apliacación?",
            title = "Confirmar cerrar"
        )
        if closeWin:
            root.quit()

    def showInfo():
        showInf = tmsg.showinfo(
            message = "Esta aplicación sigue el formato del Manual de Entrega de Información Técnica del B.I.P. y depende de 3 variables:\n\n1.  Archivo Excel con la información de parámetros de adquisición y geofísicos de cada una de las líneas sísmicas, para generar las líneas C1 a C16 del EBCDIC Header del Archivo SEG-Y\n2.  Archivo de texto con la descripción de la posición de grabación de los headers en los diferentes bytes y la secuencia de procesamiento para generar las líneas C17 a C40 del EBCDIC Header del Archivo SEG-Y\n3.  Tipo de proceso (Gathers, In-In, Out-Out, etc)\n\nVer anexo 1 de  Normatividad para la entrega de información técnica al BIP:\n\nhttps://www2.sgc.gov.co/ProgramasDeInvestigacion/BancoInformacionPetrolera/Paginas/normatividad-entrega-informacion-tecnica-BIP.aspx\n\nAutor: Oscar Lancheros (olancheros@gmail.com)",
            title = "Acerca de"
        )

    #! Código Barra Menú
    menuBar = tk.Menu()
    menuFile = tk.Menu(menuBar, tearoff=False)

    menuFile.add_command(
        label='Seleccionar XLSX',
        font='Roboto 10',
        image=imgXlsx,
        command=selectFileXlsx,
        compound=tk.LEFT
        )

    menuFile.add_command(
        label='Seleccionar TXT',
        font='Roboto 10',
        image=imgTxt,
        command=selectFileTxt,
        compound=tk.LEFT
        )

    menuFile.add_separator()

    menuFile.add_command(
        label='Seleccionar Ruta Salida',
        font='Roboto 10',
        image=imgPathToSave,
        command=selectFolder,
        compound=tk.LEFT
        )

    menuFile.add_separator()

    menuFile.add_command(
        label='Salir',
        font='Roboto 10',
        image=imgMenuExit,
        command=closeWindow,
        compound=tk.LEFT
        )

    menuHelp = tk.Menu(menuBar, tearoff=False)

    menuHelp.add_command(
        label='Acerca de',
        font='Roboto 10',
        command=showInfo,
        compound=tk.LEFT
        )

    menuBar.add_cascade(menu=menuFile, label='Archivo')
    menuBar.add_cascade(menu=menuHelp, label='Ayuda')
    root.config(menu=menuBar)

    #! Código Frame0: Header Frame
    hdrFrm = Frame(root, bd=4, width=1200, height=50)
    hdrFrm.pack(pady=4, fill=X)

    hdr = Label(
        hdrFrm,
        text='Aplicación para Generar el Encabezado Textual (EBCDIC Header)\nde Datos Sísmicos 2D Procesados y en formato SEG-Y\nSegún el Manual de Entrega de Información Técnica del\nBanco de Información Petrolera del Servicio Geológico Colombiano.',
        font='Roboto 15 bold',
        pady=10,
        padx=10
        )

    hdr.pack(pady=1,fill=X)

    #! Código Frame1: Selección Datos de Entrada y Folder de Salida
    userInput = StringVar()
    pathXlsx = StringVar()
    pathTxt = StringVar()
    pathToSave = StringVar()

    mainFrame = Frame(root, width=1200, height=100)
    mainFrame = LabelFrame(root, text=' Seleccionar Datasets ', font='Roboto 14 bold')
    mainFrame.pack(fill='both', expand='True', padx=20, pady=50)

    #! Código Selección Archivo XLSX
    labelXlsx = Label(
        mainFrame,
        text='Seleccionar Archivo XLSX:',
        font='Roboto 13'
        )
    labelXlsx.grid(row=0, column=0, sticky="e", padx=20, pady=10)

    buttonXlsx = Button(
        mainFrame,
        image=imgXlsx,
        cursor="hand2",
        command=selectFileXlsx,
        bg='#e1e1e1'
        )
    buttonXlsx.grid(row=0, column=1, sticky='w', pady=10)

    entryXlsx = Entry(
        mainFrame,
        textvariable=pathXlsx,
        width=80,
        font='Roboto 13',
        state='readonly')
    entryXlsx.grid(row=0, column=2, sticky='e', padx=25, pady=10)

    #! Código Selección Archivo TXT
    labelTxt = Label(
        mainFrame,
        text='Seleccionar Archivo TXT:',
        font='Roboto 13'
        )
    labelTxt.grid(row=1, column=0, sticky="w", padx=20)

    buttonTxt = Button(
        mainFrame,
        image=imgTxt,
        cursor="hand2",
        command=selectFileTxt,
        bg='#e1e1e1'
        )
    buttonTxt.grid(row=1, column=1, sticky='w')

    entryTxt = Entry(
        mainFrame,
        textvariable=pathTxt,
        width=80,
        font='Roboto 13',
        state='readonly'
        )
    entryTxt.grid(row=1, column=2, sticky='w', padx=25)

    #! Código Selección Ruta Salida
    labelPathToSave = Label(
        mainFrame,
        text='Seleccionar Ruta Salida:',
        font='Roboto 13'
        )
    labelPathToSave.grid(row=2, column=0, sticky="w", padx=20)

    buttonPathToSave = Button(
        mainFrame,
        image=imgPathToSave,
        cursor="hand2",
        command=selectFolder,
        bg='#e1e1e1'
        )
    buttonPathToSave.grid(row=2, column=1, sticky='w')

    entryPathToSave = Entry(
        mainFrame,
        textvariable=pathToSave,
        width=80,
        font='Roboto 13',
        state='readonly'
        )
    entryPathToSave.grid(row=2, column=2, sticky='w', padx=25, pady=20)

    #! Código Frame2: Selección Tipo de Dato Sísmico por RadioButton
    mainFrame2 = Frame(root, width=1200, height=10)
    mainFrame2.pack(fill='both', expand='True', padx=20)

    labelData = Label(
        mainFrame2,
        text='Seleccionar Tipo de Dato:',
        font='Roboto 14 bold'
        )
    labelData.grid(row=0, column=0, sticky='W', padx=20, pady=5)

    userSelect = StringVar()

    G = Radiobutton(
        mainFrame2,
        text='PSTM Gathers',
        value='G',
        variable=userSelect,
        font='Roboto 13'
        ).grid(row=0, column=1, sticky='W', padx=27, pady=5)

    I = Radiobutton(
        mainFrame2,
        text='PSTM In-In',
        value='I',
        variable=userSelect,
        font='Roboto 13'
        ).grid(row=0, column=2, sticky='W', padx=27, pady=5)

    O = Radiobutton(
        mainFrame2,
        text='PSTM Out-Out',
        value='O',
        variable=userSelect,
        font='Roboto 13'
        ).grid(row=0, column=3, sticky='W', padx=27, pady=5)

    V = Radiobutton(
        mainFrame2,
        text='PSTM Vels',
        value='V',
        variable=userSelect,
        font='Roboto 13'
        ).grid(row=0, column=4, sticky='W', padx=27, pady=5)

    E = Radiobutton(
        mainFrame2,
        text='PSTM Aniso',
        value='E',
        variable=userSelect,
        font='Roboto 13'
        ).grid(row=0, column=5, sticky='W', padx=27, pady=5)

    userSelect.set('G')

    #! Código Frame3: Botones Ejecutar & Salir
    mainFrame3 = Frame(root, width=1200, height=80)
    mainFrame3.pack(fill='both', expand='True', padx=10)

    buttonExe = Button(
        mainFrame3,
        text='Ejecutar',
        font='Roboto 14',
        cursor="hand2",
        command=textHeader,
        bg='#e1e1e1'
        )
    buttonExe.place(relx=0.38, rely=0.1, width=90)

    buttonQuit = Button(
        mainFrame3,
        text='Salir',
        font='Roboto 14',
        cursor="hand2",
        command=closeWindow,
        bg='#e1e1e1'
        )
    buttonQuit.place(relx=0.55, rely=0.1, width=90)

    root.protocol("WM_DELETE_WINDOW", closeWindow)
    root.mainloop()

#*----------------- Fin Programa -----------------*#
