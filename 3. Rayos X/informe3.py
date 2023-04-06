
import openpyxl
#import pandas as pd
import matplotlib.pyplot as plt
from scipy import stats
import numpy as np

# Informe practica 3: Rayos X

def main():


    None

    # ACTIVIDAD 1
    #actividad1()

    # ACTIVIDAD 2       
    #actividad2()

    # ACTIVIDAD 3
    #actividad3()

    # ACTIVIDAD 4

    actividad4()


def actividad1():

    # desde 10 hasta 532
    # columna A es angulo
    # columna B es intensidad
    
    lista_angulo = getListFromCells("A10", "A532", "Actividad1")
    lista_intensidad = getListFromCells("B10", "B532", "Actividad1")

    # encontramos k_beta en 20.3
    # encontramos k_alpha en 22.6

    lam = 1.54
    theta = 22.6
    d = lam / (2*np.sin(np.radians(theta)))

    #print(f'Valor de d = {d}')

    bragg = lambda theta : 2*d*np.sin(np.radians(theta))    
    lista_longitudes = [bragg(theta) for theta in lista_angulo]
    
    #print(f'Valor de lambda para k_beta = {lista_longitudes[175]}')
    #print(f'Valor de lambda para k_alpha = {lista_longitudes[198]}')

    graph(lista_longitudes, lista_intensidad, "Intensidad vs Longitud de onda", r"Longitud de onda ($\AA$)", "Intensidad (Impulsos/s)", "Actividad 1 - Intensidad vs Longitud de onda", yerror=400)

def actividad2():

    # desde 10 hasta 532
    # columna A es angulo
    # columna B es intensidad
    
    #lista_angulo = getListFromCells("A10", "A532", "Actividad2")
    #lista_intensidad = getListFromCells("B10", "B532", "Actividad2")

    # se hizo de 10 a 16 porque Jose nos lo sugirió, debido al pico de la actividad1 debido a bajos angulos

    # lista de angulos de 10 a 16
    lista_angulos = [ i for i in range(10, 17) ]


    # ALUMINIO

    lista_grosor = [ 0.1, 0.08, 0.06, 0.04, 0.02 ]
    
    # lista de intensidades para cada grosor
    lista_columnas = [ "C", "E", "G", "I", "K" ]
    lista_filas = [ str(i) for i in range(23, 30) ]


    lista_lista_intensidades = []

    for j in lista_filas:
        lista_intensidades = []
        for k in lista_columnas:
            lista_intensidades.append( getCellValue(k+j, "Actividad2") / 1423 ) # se divide entre el valor mayor para normalizar
        
        #print(lista_intensidades)
        lista_lista_intensidades.append( lista_intensidades )

    # con semilog
    lista_pendientes = graphMultiple(lista_angulos, lista_grosor, lista_lista_intensidades, "Intensidad vs Grosor de Aluminio", "Grosor (mm)", "Intensidad (Impulsos / s)", "Actividad 2 - Intensidad vs Grosor de Aluminio")
    #print(lista_pendientes)
    


    # Zinc

    lista_grosor = [ 0.1, 0.075, 0.05, 0.025]
    
    # lista de intensidades para cada grosor
    lista_columnas = [ "C", "E", "G", "I"]
    lista_filas = [ str(i) for i in range(37, 44) ]

    lista_lista_intensidades = []

    for j in lista_filas:
        lista_intensidades = []
        for k in lista_columnas:
            lista_intensidades.append( getCellValue(k+j, "Actividad2") / 640) # se divide entre el valor mayor para normalizar
        
        lista_lista_intensidades.append( lista_intensidades )

    lista_pendientes = graphMultiple(lista_angulos, lista_grosor, lista_lista_intensidades, "Intensidad vs Grosor de Zinc", "Grosor (mm)", "Intensidad (Impulsos / s)", "Actividad 2 - Intensidad vs Grosor de Zinc")
    #print(lista_pendientes)

def actividad3():
    
    lista_angulo = getListFromCells("B7", "B57", "Actividad3")
    #print(lista_angulo)
    bragg = lambda theta : 2*2.014*np.sin(np.radians(theta))  
    lista_longitudes = [bragg(theta) for theta in lista_angulo]



    # Corriente constante 1mA

    lista_voltajes = [ 3*i + 11 for i in range(9) ]

    # lista de intensidades para cada grosor
    lista_columnas = ["C", "D", "E", "F", "G", "H", "I", "J", "K"] 
    lista_filas = [ str(i) for i in range(7, 58) ]


    lista_lista_intensidades = []

    for j in lista_columnas:
        lista_intensidades = []
        for k in lista_filas:
            lista_intensidades.append( getCellValue(j+k, "Actividad3") ) # se divide entre el valor mayor para normalizar
        
        #print(lista_intensidades)
        lista_lista_intensidades.append( lista_intensidades )


    graphMultiple(lista_voltajes, lista_longitudes, lista_lista_intensidades, "Intensidad vs Longitud de onda variando voltaje", r"Longitud de onda ($\AA$)", "Intensidad (Impulsos / s)", "Actividad 3 - Intensidad vs Longitud de onda variando voltajes", semilog=False, variable="Voltaje", unidades="kV" ) 
                  



    # Voltaje constante 35kV

    lista_corrientes = [ round(i*0.1 + 0.1, 1) for i in range(10) ]

    # lista de intensidades para cada grosor
    lista_columnas = ["N", "P", "R", "T",  "V",  "X", "Z", "AB", "AD", "AF"]
    lista_filas = [ str(i) for i in range(7, 58) ]


    lista_lista_intensidades = []

    for j in lista_columnas:
        lista_intensidades = []
        for k in lista_filas:
            lista_intensidades.append( getCellValue(j+k, "Actividad3") ) # se divide entre el valor mayor para normalizar
        
        #print(lista_intensidades)
        lista_lista_intensidades.append( lista_intensidades )


    graphMultiple(lista_corrientes, lista_longitudes, lista_lista_intensidades, "Intensidad vs Longitud de onda variando corriente", r"Longitud de onda ($\AA$)", "Intensidad (Impulsos / s)", "Actividad 3 - Intensidad vs Longitud de onda variando corriente", semilog=False, variable = "Corriente", unidades = "mA") 

def actividad4():
    None
    lista_angulos = getListFromCells("B21", "B81", "Actividad4")

    lista_voltajes = [i for i in reversed(range(13,37,2))]


    # Primeros valores 

    lista_columnas = ["C", "E", "G", "I", "K", "M"]
    lista_filas = [ str(i) for i in range(21, 82) ]


    lista_lista_intensidades = []

    for j in lista_columnas:
        lista_intensidades = []
        for k in lista_filas:
            lista_intensidades.append( getCellValue(j+k, "Actividad4") )
        
        #print(lista_intensidades)
        lista_lista_intensidades.append( lista_intensidades )

    

    # Segundos valores

    lista_columnas = ["O", "Q", "S", "U", "W", "Y"] 
    lista_filas = [ str(i) for i in range(11, 72) ]


    for j in lista_columnas:
        lista_intensidades = []
        for k in lista_filas:
            lista_intensidades.append( getCellValue(j+k, "Actividad4") ) # se divide entre el valor mayor para normalizar
        
        print(lista_intensidades)
        lista_lista_intensidades.append( lista_intensidades )


    graphMultiple(lista_voltajes, lista_angulos, lista_lista_intensidades, "Intensidad vs Angulo variando voltaje", "Angulo (°)", "Intensidad (Impulsos / s)", "Actividad 4 - Intensidad vs Angulo variando voltaje", semilog=False, variable="Voltaje", unidades="kV" )          
    #graphMultiple(lista_voltajes, lista_angulos, lista_lista_intensidades, "Intensidad vs Longitud de onda variando voltaje", r"Longitud de onda ($\AA$)", "Intensidad (Impulsos / s)", "Actividad 3 - Intensidad vs Longitud de onda variando voltajes", semilog=False, variable="Voltaje", unidades="kV" ) 
               


def graphMultiple( lista_angulos, lista_grosor, lista_lista_intensidades, title, xlabel, ylabel, filename, xerror=None, yerror=None, semilog = True, variable = "Angulo", unidades = "°"):

    plt.style.use(['science','no-latex'])

    fig = plt.figure()
    ax = fig.add_subplot()
    ax.set_title(title, fontsize=8)
    ax.set_xlabel(xlabel,fontsize=8)
    ax.set_ylabel(ylabel,fontsize=8)

    color_list = ['r', 'b', 'g', 'k', 'm',  'c', 'y']

    lista_pendientes = []

    j = 0
    for i in lista_angulos:
        #plt.scatter(lista_grosor, lista_lista_intensidades[j], label=f'Angulo = {i}', marker='o', s=2)
        
        slope, intercept, r_value, p_value, std_err = stats.linregress(lista_grosor, lista_lista_intensidades[j])
        lista_pendientes.append(slope)

        if semilog:
            plt.semilogy(lista_grosor, lista_lista_intensidades[j], label=f'{variable} = {i}{unidades}', marker='o', c=lighten_color(color_list[j], 0.7))
            plt.semilogy(lista_grosor, slope*np.array(lista_grosor) + intercept, marker='2', c=lighten_color(color_list[j], 0.4)) 
        else:    
            plt.plot(lista_grosor, lista_lista_intensidades[j], label=f'{variable} = {i}{unidades}', marker='o',  linewidth=1, markersize=1)

        j += 1

    plt.legend( loc='upper right', numpoints = 1, fontsize=5)
    plt.grid()
    plt.savefig(filename + ".png", dpi=300, bbox_inches='tight')
    #plt.show()

    return lista_pendientes

def lighten_color(color, amount=0.5):
    """
    Lightens the given color by multiplying (1-luminosity) by the given amount.
    Input can be matplotlib color string, hex string, or RGB tuple.

    Examples:
    >> lighten_color('g', 0.3)
    >> lighten_color('#F034A3', 0.6)
    >> lighten_color((.3,.55,.1), 0.5)
    """
    import matplotlib.colors as mc
    import colorsys
    try:
        c = mc.cnames[color]
    except:
        c = color
    c = colorsys.rgb_to_hls(*mc.to_rgb(c))
    return colorsys.hls_to_rgb(c[0], 1 - amount * (1 - c[1]), c[2])

def graph(x, y, title, xlabel, ylabel, filename, xerror=None, yerror=None):

    plt.style.use(['science','no-latex'])

    fig = plt.figure()
    ax = fig.add_subplot()
    ax.set_title(title, fontsize=8)
    ax.set_xlabel(xlabel,fontsize=8)
    ax.set_ylabel(ylabel,fontsize=8)

    plt.scatter(x, y, label='Datos experimentales', color='blue', marker='o', s=2)

    if xerror is not None:
        plt.errorbar(x, y,xerr=xerror, fmt='none', ecolor='black', capsize=1, elinewidth=0.5)
    if yerror is not None:
        plt.errorbar(x, y,yerr=yerror, fmt='none', ecolor='black', capsize=0.8, elinewidth=0.2)

    plt.legend( loc='upper right', numpoints = 1, fontsize=6)
    plt.grid()
    plt.savefig(filename + ".png", dpi=300, bbox_inches='tight')
    #plt.show()


def getCellValue(celda, hoja_trabajo=None): 

    # Selección de la hoja de trabajo
    if hoja_trabajo is None:
        hoja_trabajo = libro_excel.active
    else:
        hoja_trabajo = libro_excel[hoja_trabajo]

    #print(hoja_trabajo[celda])

    return hoja_trabajo[celda].value

def getListFromCells(celda_inicio, celda_fin, hoja_trabajo=None):
    # OJO: No funciona para celdas que sean calculadas con fórmulas

    # Selección de la hoja de trabajo
    if hoja_trabajo is None:
        hoja_trabajo = libro_excel.active
    else:
        hoja_trabajo = libro_excel[hoja_trabajo]

    # Creación de la lista vacía
    lista_valores = []
        
    # Recorrido de las celdas entre las celdas de inicio y fin
    for celda in hoja_trabajo[celda_inicio:celda_fin]:
        
        # print(celda)
        # print(celda[0].value)
        # print(type(celda[0].value))
        
        if type(celda[0].value) == str:
            #print("Celda es STRING")
            lista_valores.append(float(celda[0].value.replace(',', '.')))

        elif type(celda[0].value) == float:
            lista_valores.append(celda[0].value)

        elif type(celda[0].value) == int:
            lista_valores.append(celda[0].value)

        else:
            print("ERROR: No se reconoce el tipo de dato")
            

    return lista_valores



if __name__ == "__main__":

    # Ruta del archivo de Excel
    archivo_excel = r"C:\Users\eduar\OneDrive - Universidad de los Andes\OTROS\Sofi-Edu\2023-10\Laboratorio Intermedio\Práctica 3 - Rayos X\RayosX_Grupo6.xlsx"

    # Carga del libro de Excel
    libro_excel = openpyxl.load_workbook(archivo_excel)

    main()

    