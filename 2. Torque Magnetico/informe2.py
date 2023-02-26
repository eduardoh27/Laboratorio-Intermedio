
import openpyxl
#import pandas as pd
import matplotlib.pyplot as plt
from scipy import stats
import numpy as np

def main():
    


    # ACTIVIDAD 1       
    #actividad1()



    # ACTIVIDAD 2       

    lista_T2 = [2.732409, 2.4964, 2.32410025, 2.02920025, 1.89475225, 1.830609]
    
    lista_B1 = [1543.686323, 1397.428731, 1289.98968, 1105.70544, 1021.763564, 981.3542689]

    #graph(lista_B1, lista_T2, "T^2 vs 1/B", "1/B (1/T)", "T^2 (s^2)", "Actividad 2 - T2 vs 1B")


    # ACTIVIDAD 3
    #actividad3()


def graph(x, y, title, xlabel, ylabel, filename):

    plt.style.use(['science','no-latex'])

    # linear regression
    slope, intercept, r_value, p_value, std_err = stats.linregress(x, y)

    fig = plt.figure()
    ax = fig.add_subplot()
    ax.set_title(title, fontsize=8)
    ax.set_xlabel(xlabel,fontsize=8)
    ax.set_ylabel(ylabel,fontsize=8)

    # change the fontsize of the xtick and ytick labels
    #plt.rc('xtick', labelsize=7)
    #plt.rc('ytick', labelsize=7)

    plt.scatter(x, y, label='Datos experimentales', color='red', marker='o', s=10)
    plt.plot(x, slope*np.array(x) + intercept, 'b', label=f'Ajuste lineal: m={slope:.5f}, b={intercept:.5f}')
    plt.legend( loc='lower right', numpoints = 1, fontsize=6)
    plt.grid()
    plt.savefig(filename + ".png", dpi=300, bbox_inches='tight')
    #plt.show()





def getListFromCells(celda_inicio, celda_fin):
    # OJO: No funciona para celdas que sean calculadas con fórmulas

    # Ruta del archivo de Excel
    archivo_excel = r"C:\Users\eduar\OneDrive - Universidad de los Andes\OTROS\Sofi-Edu\2023-10\Laboratorio Intermedio\Práctica 2 - Torque Magnético\Labfolder Table.xlsx"

    # Carga del libro de Excel
    libro_excel = openpyxl.load_workbook(archivo_excel)

    # Selección de la hoja de trabajo
    hoja_trabajo = libro_excel.active

    # Creación de la lista vacía
    lista_valores = []
        
    # Recorrido de las celdas entre las celdas de inicio y fin
    for celda in hoja_trabajo[celda_inicio:celda_fin]:
        #print(celda)
        if type(celda[0].value) != float:

            lista_valores.append(float(celda[0].value.replace(',', '.')))
        else:
            lista_valores.append(celda[0].value)

    return lista_valores


def actividad1():
    
    lista_B = getListFromCells("F11", "F17")

    lista_r = getListFromCells("E11", "E17")

    graph(lista_r, lista_B, "B vs r", "r (m)", "B (T)", "Actividad 1 - B vs r")

def actividad3():
    
    lista_Omega = getListFromCells("D33", "D39")
    lista_B = getListFromCells("E33", "E39")

    graph(lista_B, lista_Omega, "Omega vs B", "B (T)", "Omega (rad/s)", "Actividad 3 - Omega vs B")

if __name__ == "__main__":
    main()

    