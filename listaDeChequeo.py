import openpyxl
import pandas as pd
import shutil

# Nombre del archivo de la plantilla de Excel para la lista de chequeo.
plantilla='106.xlsx'

# Define la ruta del directorio donde se encuentran los archivos relevantes.
ruta= '/home/kaelectro/Documentos/LC/'
#ruta= '/home/ka_97/Projects/Electroequipos/LCX/Documentos/'
#ruta= '/home/kaelectro/Documentos/PC Y VIDEOBEAM/'

# Lee los datos desde un archivo Excel específico, usándolos como base para rellenar la plantilla.
# El primer argumento es la ruta del archivo y el segundo argumento establece la primera columna como índice o etiqueta.
data = pd.read_excel(ruta+'InfoParaPlantilla.xlsx',index_col=0)

# Extrae información específica del DataFrame para su uso posterior.
# Esto incluye número de contrato, orden, ciudad, marca, referencia y nombres de equipos.
contrato=1780
orden="CMM-2023-000136"
fechaLlegada='28/12/2023'
fechaRevision='29/12/2023'
comercial='Luis Cañon'
cliente='Universidad Nacional Abierta y a Distancia (UNAD)'
ciudad=data['Ciudad']
marca=data['Marca']
serial=data['Serial']
nombres=data['Nombre del equipo']
gestor=data['Nombre del Gestor']

def crearLibro(ruta, plantilla, consecutivo):
    # Define el nombre del archivo y copia la plantilla a un nuevo archivo Excel.
    nombre_archivo = consecutivo + '.xlsx'
    shutil.copy(ruta + plantilla, ruta+nombre_archivo)
    # Carga el libro de Excel copiado y cambia el título de la hoja de trabajo.
    wb = openpyxl.load_workbook(ruta + nombre_archivo)
    wb['OPF05V4'].title=consecutivo
    return wb

def actualizar_celdas(wb, cliente, comercial, ciudad, fechaLlegada, fechaRevision, contrato, orden, 
                      gestor, nombre, ref, marca, serial):
    sheet = wb[consecutivo]

    celda_cliente=sheet["E4:K4"]
    celda_cliente[0][0].value= cliente

    celda_comercial=sheet["E5:K5"]
    celda_comercial[0][0].value= comercial

    celda_ciudad=sheet["N4:Q4"]
    celda_ciudad[0][0].value= ciudad

    celda_fel=sheet["N5:Q5"]
    celda_fel[0][0].value= fechaLlegada

    celda_rev=sheet["N6:Q6"]
    celda_rev[0][0].value= fechaRevision

    celda_contrato=sheet["E6:K6"]
    celda_contrato[0][0].value= contrato

    celda_orden=sheet["E7:K7"]
    celda_orden[0][0].value= orden

    celda_gestor=sheet["N7:Q7"]
    celda_gestor[0][0].value= gestor

    celda_ped=sheet["E8:K8"]
    celda_ped[0][0].value= "N/A"

    celda_caj=sheet["N8:Q8"]
    celda_caj[0][0].value= "N/A"

    celda_nombre=sheet["E10:K10"]
    celda_nombre[0][0].value= nombre

    celda_ref=sheet["N10:Q10"]
    celda_ref[0][0].value= ref

    celda_marca=sheet["E11:K11"]
    celda_marca[0][0].value= marca

    celda_serial=sheet["N11:Q11"]
    celda_serial[0][0].value= serial

    return wb

i=0
for item,nombre in nombres.items():
    i+=1
    # Crea un identificador único para cada informe.
    consecutivo=f'LC-{i}-{nombre}'

    wb=crearLibro(ruta,plantilla,consecutivo)
    wb=actualizar_celdas(wb,cliente,comercial,ciudad[item],fechaLlegada,fechaRevision,
                         contrato,orden,gestor[item],nombre,item,marca[item],serial[item])

    wb.save(ruta+consecutivo+".xlsx")
    wb.close()
    