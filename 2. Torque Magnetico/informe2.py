
import openpyxl
#import pandas as pd
import matplotlib.pyplot as plt

def main():
    
    # Ruta del archivo de Excel
    archivo_excel = r"C:\Users\eduar\OneDrive - Universidad de los Andes\OTROS\Sofi-Edu\2023-10\Laboratorio Intermedio\Práctica 2 - Torque Magnético\Labfolder Table.xlsx"

    # Carga del libro de Excel
    libro_excel = openpyxl.load_workbook(archivo_excel)

    # Selección de la hoja de trabajo
    hoja_trabajo = libro_excel.active



    # ACTIVIDAD 1       

    lista_B = getListFromCells("C11", "C17", hoja_trabajo)
    
    lista_r = getListFromCells("A11", "A17", hoja_trabajo)

    graph(lista_B, lista_r, "B vs r", "B (T)", "r (m)", "B vs r")
    




def graph(x, y, title, xlabel, ylabel, filename):

    plt.style.use(['science','no-latex'])

    fig = plt.figure()
    ax = fig.add_subplot()
    ax.set_title(title)
    ax.set_xlabel(xlabel,fontsize=10)
    ax.set_ylabel(ylabel,fontsize=10)
    plt.scatter(x, y)
    plt.savefig(filename + ".png", dpi=300, bbox_inches='tight')
    plt.show()



def getListFromCells(celda_inicio, celda_fin, hoja_trabajo):

    # Creación de la lista vacía
    lista_valores = []
        
    # Recorrido de las celdas entre las celdas de inicio y fin
    for celda in hoja_trabajo[celda_inicio:celda_fin]:
        lista_valores.append(float(celda[0].value.replace(',', '.')))

    # Impresión de la lista generada
    #print(lista_valores)

    return lista_valores


if __name__ == "__main__":
    main()

    