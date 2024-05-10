import openpyxl
from openpyxl.drawing.image import Image  # Importa Image para añadir imágenes a los archivos Excel
from win32com import client
import math
import os
import pandas as pd
import tkinter as tk
from tkinter import ttk

CONTRATO=1782
ORDEN_CONTRACTUAL="CMM-2023-000130"
COMERCIAL='LUIS CAÑÓN'
CLIENTE='Universidad Nacional Abierta y a Distancia (UNAD)'
RUTA_DE_ALMACENAMIENTO=r'\\172.16.0.9\Depto tecnico\2 0 2 3\2. CONTRATOS\1782 UNAD DOTACIÓN\2. LISTAS DE CHEQUEO'


def crear_lista_de_chequeo(rutaPlantilla, rutaScript, cliente, comercial, contrato, orden_c, ruta_de_guardado, pdseries):
    
    wb_LC = openpyxl.load_workbook(rutaPlantilla)  # Loads the Excel workbook.
    sheet= wb_LC.worksheets[0]

    ###########
    #ENCABEZADO
    ###########
    celda_cliente=sheet["E4:K4"]
    celda_cliente[0][0].value= cliente

    celda_comercial=sheet["E5:K5"]
    celda_comercial[0][0].value= comercial

    celda_ciudad=sheet["N4:Q4"]
    celda_ciudad[0][0].value= pdseries['CIUDAD']

    celda_fel=sheet["N5:Q5"]
    celda_fel[0][0].value= pdseries.iloc[7]

    celda_rev=sheet["N6:Q6"]
    celda_rev[0][0].value= pdseries.iloc[8]

    celda_contrato=sheet["E6:K6"]
    celda_contrato[0][0].value= contrato

    celda_orden=sheet["E7:K7"]
    celda_orden[0][0].value= orden_c

    celda_gestor=sheet["N7:Q7"]
    celda_gestor[0][0].value= pdseries.iloc[6]

    celda_ped=sheet["E8:K8"]
    celda_ped[0][0].value= "N/A"

    celda_caj=sheet["N8:Q8"]
    celda_caj[0][0].value= "N/A"

    celda_nombre=sheet["E10:K10"]
    celda_nombre[0][0].value= pdseries.iloc[2]

    celda_ref=sheet["N10:Q10"]
    celda_ref[0][0].value= pdseries.iloc[4]

    celda_marca=sheet["E11:K11"]
    celda_marca[0][0].value= pdseries.iloc[3]

    celda_serial=sheet["N11:Q11"]
    celda_serial[0][0].value= pdseries.iloc[5]


    ###################
    #CONEXION ELECTRICA
    ###################
    if pdseries.iloc[10]=='110V':
        sheet["H15"]='X'
    elif pdseries.iloc[10]=='220V':
        sheet["K15"]='X'

    sheet["M15"]=pdseries.iloc[11]

    celda_fase=sheet["P15:Q15"]
    celda_fase[0][0].value= pdseries.iloc[12]

    ###################
    #PARAMETROS
    ###################
    celdasDeX=['19','20','21','22','23','24','25','27','28','29','30','31','33']
    columnasEnBaseDeDatos=[13,15,17,19,21,23,25,27,29,31,33,35,37]
    for i in range(0,len(celdasDeX)):
        filaActual=celdasDeX[i]
        if pdseries.iloc[columnasEnBaseDeDatos[i]]=='SI':
            sheet['G'+filaActual]='X'

        elif pdseries.iloc[columnasEnBaseDeDatos[i]]=='NO':
            sheet['I'+filaActual]='X'

        else:
            sheet['K'+filaActual]='X'

        observaciones=sheet["M"+filaActual+":Q"+filaActual]
        observaciones[0][0].value= pdseries.iloc[columnasEnBaseDeDatos[i]+1]

    ###################
    #VARIABLES REV.
    ###################
    celda_variable1=sheet["C37:D37"]
    celda_variable1[0][0].value= pdseries.iloc[39]

    celda_variable1o=sheet["E37:K37"]
    celda_variable1o[0][0].value= pdseries.iloc[40]

    celda_variable2=sheet["C38:D38"]
    celda_variable2[0][0].value= pdseries.iloc[41]

    celda_variable2o=sheet["E38:K38"]
    celda_variable2o[0][0].value= pdseries.iloc[42]

    celda_variable3=sheet["M37:N37"]
    celda_variable3[0][0].value= pdseries.iloc[43]
    
    celda_variable3o=sheet["O37:Q37"]
    celda_variable3o[0][0].value= pdseries.iloc[44]

    celda_variable4=sheet["M38:N38"]
    celda_variable4[0][0].value= pdseries.iloc[45]
    
    celda_variable4o=sheet["O38:Q38"]
    celda_variable4o[0][0].value= pdseries.iloc[46]

    ######################
    #OBSERVACIONES Y FINAL
    ######################
    celda_observaciones=sheet["C41:Q44"]
    celda_observaciones[0][0].value= pdseries.iloc[47]

    celda_realizado=sheet["E52:M52"]
    celda_realizado[0][0].value= pdseries.iloc[48]

    celda_revisado=sheet["E53:M53"]
    celda_revisado[0][0].value= pdseries.iloc[49]

    consecutivo=str(pdseries.iloc[1])

    carpeta_de_almacenamiento= os.path.join(ruta_de_guardado,pdseries['CIUDAD'],consecutivo+" "+pdseries.iloc[2])
    sheet["J46"]=carpeta_de_almacenamiento
    sheet["J48"]=os.path.join(carpeta_de_almacenamiento,'REGISTRO AUDIOVISUAL')

    # Ajustar la altura de cada fila basándose en el contenido más largo que excede el ancho fijo
    # Ancho fijo de columna, por ejemplo, 20 caracteres
    fixed_column_width = 20
    for row in sheet.iter_rows(min_row=19):
        max_lines = 1  # Mínimo una línea por fila
        for cell in row:
            # Calcular cuántas líneas se necesitan para el texto en esta celda
            lines_needed = math.ceil(len(str(cell.value)) / fixed_column_width)
            if lines_needed > max_lines:
                max_lines = lines_needed
        # Ajustar la altura de la fila; asumir que cada línea necesita aproximadamente 15 puntos de altura
                if lines_needed==2:
                    sheet.row_dimensions[cell.row].height = max_lines * 15
                else:
                    sheet.row_dimensions[cell.row].height = max_lines * 9

    img = Image(os.path.join(rutaScript, 'encabezado.png'))
    sheet.add_image(img, 'C3')

    

    carpeta=os.path.join(rutaScript,'LC',consecutivo+" "+pdseries.iloc[2])
    os.makedirs(carpeta)
    os.makedirs(os.path.join(carpeta,'REGISTRO AUDIOVISUAL'))

    nombrearchivo=consecutivo+" "+pdseries.iloc[2]+'.xlsx'
    sheet.title=consecutivo
    archivo=os.path.join(carpeta,nombrearchivo)
    wb_LC.save(archivo)

    xsl2pdf(archivo)

    
def main(comercial, cliente, contrato, orden_c,ruta):
    


    # Nombre del archivo de la plantilla de Excel para la lista de chequeo.
    plantilla='106.xlsx'
    baseDeDatosParaRevisionDeListas='baseDeDatosParaRevisionDeListas.xlsx'
    script_actual = os.path.realpath(__file__)  # Obtiene la ruta absoluta del script en ejecución
    script_directory = os.path.dirname(script_actual)  # Obtiene el directorio donde se encuentra el script

    ruta_plantilla=os.path.join(script_directory,plantilla)
    ruta_basedatos=os.path.join(script_directory,baseDeDatosParaRevisionDeListas)


    # Lee los datos desde un archivo Excel específico, usándolos como base para rellenar la plantilla.
    # El primer argumento es la ruta del archivo y el segundo argumento establece la primera columna como índice o etiqueta.
    df = pd.read_excel(ruta_basedatos)

    for indice, fila in df.iterrows():
        crear_lista_de_chequeo(ruta_plantilla,script_directory,cliente,comercial,contrato,orden_c,ruta,fila)

    print(len(os.listdir(os.path.join(script_directory,'LC'))))

def crear_label_entry(master, texto, row):
    """
    Crea un label y un entry con su variable asociada en una fila específica del grid.

    Args:
    master : tk.Widget - Widget padre en el que se colocarán el label y el entry.
    texto : str - Texto para el label.
    row : int - Fila del grid donde colocar el label y el entry.

    Returns:
    tk.StringVar - Variable asociada al entry.
    """
    # Crear y posicionar el label
    label = ttk.Label(master, text=texto)
    label.grid(row=row, column=0, sticky=tk.W, pady=5,padx=(10,0))

    # Crear la variable para el entry
    var = tk.StringVar()

    # Crear y posicionar el entry
    entry = ttk.Entry(master, textvariable=var,width=100)
    entry.grid(row=row, column=1, pady=5,padx=(5,10))

    return var

def xsl2pdf(file_location):
    app = client.DispatchEx("Excel.Application")
    app.Interactive = False
    app.Visible = False

    workbook=app.Workbooks.open(file_location)
    output = os.path.splitext(file_location)[0]

    workbook.ActiveSheet.ExportAsFixedFormat(0,output)
    workbook.Close()

def funcion_de_boton():
        com=entry_vars["Comercial:"].get()
        cont=entry_vars["Contrato:"].get()
        ordc=entry_vars["Orden Contractual:"].get()
        clie=entry_vars["Cliente:"].get()
        ruta=entry_vars["Ruta de Almacenamiento:"].get()
        main(com,clie,cont,ordc,ruta)

if __name__ == '__main__':  
    root = tk.Tk()
    root.title("Crear Listas de Chequeo")
    root.iconbitmap()
    root.grid_rowconfigure(1, weight=1)
    root.grid_columnconfigure(1, weight=1)

    labels = ["Comercial:", "Contrato:", "Orden Contractual:", "Cliente:","Ruta de Almacenamiento:"]

    # Diccionario para almacenar las variables asociadas a cada entry
    entry_vars = {}

    # Crear los labels y entries en el grid
    for i, label_text in enumerate(labels):
        entry_vars[label_text] = crear_label_entry(root, label_text, i)

    entry_vars["Comercial:"].set(COMERCIAL)
    entry_vars["Contrato:"].set(CONTRATO)
    entry_vars["Orden Contractual:"].set(ORDEN_CONTRACTUAL)
    entry_vars["Cliente:"].set(CLIENTE)
    entry_vars["Ruta de Almacenamiento:"].set(RUTA_DE_ALMACENAMIENTO)

                                

    button = ttk.Button(root, text="CREAR LISTAS DE CHEQUEO",command=funcion_de_boton)
    button.grid(row=5, column=0,columnspan=2, padx=100, pady=(10, 10))

    






    # Start the application
    root.mainloop()









