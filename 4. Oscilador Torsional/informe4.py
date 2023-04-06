
import openpyxl
#import pandas as pd
import matplotlib.pyplot as plt
from scipy import stats
import numpy as np

# Informe practica 3: Rayos X

def main():

    # ACTIVIDAD 1
    #actividad1()

    # ACTIVIDAD 2       
    #actividad2()

    # ACTIVIDAD 3
    actividad3()

    # ACTIVIDAD 4
    #actividad4()
    
    # ACTIVIDAD 5
    #actividad5()

    None


def actividad1():

    # torque = -k * theta

    hoja_trabajo = "Actividad1"    

    # PESA POR LA IZQUIERDA

    masas_data = getListFromCells('A3', 'A10', hoja_trabajo) # (g)
    posiciones_data = getListFromCells('B3', 'B10', hoja_trabajo) # (rad)
    radio_data = 1.4725 # (cm)
    radio_data = 1.2

    pesos = [gravedad * (masa / 1000) for masa in masas_data] # (N)
    deltas = [posicion - 3 for posicion in posiciones_data] # (rad)
    radio = radio_data / 100 # (m)
    torques = [peso * radio for peso in pesos] # (N*m)
    
    error_r = 0.05/1000 # (m)
    errores_torque = [error_r * peso for peso in pesos] # (N*m)
    errores_angulo = [0.1 for i in range(len(deltas))] # (rad)

    print(f"errores_torque: {errores_torque}")

    graph(deltas, torques, "Torque vs. Delta", "Delta (rad)", "Torque (N*m)", "Actividad1_izquierda", direccion="L", yerror=errores_torque, xerror=errores_angulo)

    # PESA POR LA DERECHA

    masas_data = getListFromCells('D3', 'D10', hoja_trabajo) # (g)
    posiciones_data = getListFromCells('E3', 'E10', hoja_trabajo) # (rad)

    pesos = [gravedad * (masa / 1000) for masa in masas_data] # (N)
    deltas = [posicion - 3 for posicion in posiciones_data] # (rad)
    radio = radio_data / 100 # (m)
    torques = [peso * radio for peso in pesos] # (N*m)

    errores_torque = [error_r * peso for peso in pesos] # (N*m)

    graph(deltas, torques, "Torque vs. Delta", "Delta (rad)", "Torque (N*m)", "Actividad1_derecha", yerror=errores_torque, xerror=errores_angulo)

def actividad2():
    
    hoja_trabajo = "Actividad2"

    N_latones = getListFromCells('B2', 'B10', hoja_trabajo)
    periodos = getListFromCells('C2', 'C10', hoja_trabajo)

    print(f"N_latones: {N_latones}")
    print(f"periodos: {periodos}")

    errores_periodo = [0.01 for i in range(len(periodos))]

    pendiente = graph(N_latones, periodos, "Periodo vs. N Latones", "N Latones", "Periodo (s)", "Actividad2", 
                      residuals=True, regresion=True, yerror=errores_periodo, direccion="L", pendiente=True)
    
    # calcular I
    I = pendiente / (2 * np.pi) # (kg*m^2)
    #I = (pendiente / ((2 * np.pi)**2))**2 # (kg*m^2)
    k = I * (2 * np.pi)**2 # (N*m^2)

    print(f"pendiente: {pendiente}")
    print(f"I: {I}")
    print(f"k: {k}")

def actividad3():
    
    hoja_trabajo = "Actividad3"

    posiciones = getListFromCells('B2', 'B10', hoja_trabajo) # (rad)
    corrientes = getListFromCells('C2', 'C10', hoja_trabajo) # (A)

    # ignorar primeros 2 datos
    #posiciones = posiciones[2:]
    #corrientes = corrientes[2:]

    print(f"posiciones: {posiciones}")
    print(f"corrientes: {corrientes}")

    deltas = [posicion - 3 for posicion in posiciones] # (rad)

    pendiente = graph(deltas, corrientes, "Corriente vs. $\Delta\\theta$", "$\Delta\\theta$ (rad)", "Corriente (A)", "Actividad3", 
                      residuals=True, regresion=True, pendiente=True, direccion="L")

    momentos_magneticos = [pendiente * corriente for corriente in corrientes] # (A*m^2)

    print(f"pendiente: {pendiente}")
    print(f"momentos_magneticos: {momentos_magneticos}")

def actividad4():
    None

def actividad5():
    
    hoja_trabajo = "Actividad5"

    # Conjunto A

    frecuencias_A = getListFromCells('A16', 'A25', hoja_trabajo) # (Hz)
    amplitudes_A  = getListFromCells('B16', 'B25', hoja_trabajo) # (V)

    # Conjunto B

    frecuencias = getListFromCells('I17', 'I26', hoja_trabajo) # (Hz)
    amplitudes  = getListFromCells('J17', 'J26', hoja_trabajo) # (V)

    y_error = [0.01 for i in range(len(amplitudes))]
    x_error = [0.1 for i in range(len(frecuencias))]

    graph([frecuencias_A, frecuencias], [amplitudes_A, amplitudes], "Amplitud vs. Frecuencia", 
          "Frecuencia (Hz)", "Amplitud (V)", "Actividad5", residuals=False, regresion=False, direccion="L", 
          yerror=y_error, xerror=x_error, multiple=True)


def graph(x, y, title, xlabel, ylabel, filename, xerror=None, yerror=None, 
          residuals = True, regresion = True, pendiente = False, direccion = "R", multiple=False):

    fig = plt.figure()
    ax = fig.add_subplot()
    ax.set_title(title, fontsize=8)
    ax.set_xlabel(xlabel,fontsize=8)
    ax.set_ylabel(ylabel,fontsize=8)
    
    if multiple == False:

        if regresion:
            slope, intercept, r_value, p_value, std_err = stats.linregress(x, y)
            ax.plot(x, slope*np.array(x) + intercept, 'b', label=f'Ajuste lineal: m={slope:.5f}, b={intercept:.5f}')
            
        ax.scatter(x, y, label='Datos experimentales', color='blue', marker='o', s=2)
        
        if xerror is not None:
            plt.errorbar(x, y, xerr=xerror, fmt='none', ecolor='black', capsize=1, elinewidth=0.5)
        if yerror is not None:
            plt.errorbar(x, y,yerr=yerror, fmt='none', ecolor='black', capsize=0.8, elinewidth=0.2)
    
    else:
        for k in range(len(x)):

            i = x[k]
            j = y[k]

            if regresion:
                slope, intercept, r_value, p_value, std_err = stats.linregress(i, j)
                ax.plot(i, slope*np.array(i) + intercept, 'b', label=f'Ajuste lineal: m={slope:.5f}, b={intercept:.5f}')
                
            ax.scatter(i, j, label='Datos experimentales', marker='o', s=8)
 
            if xerror is not None:
                i_error = xerror
                #i_error = xerror[k]
                plt.errorbar(i, j, xerr=i_error, fmt='none', ecolor='black', capsize=1, elinewidth=0.2)
            if yerror is not None:
                j_error = yerror
                #j_error = yerror[k]
                plt.errorbar(i, j, yerr=j_error, fmt='none', ecolor='black', capsize=0.8, elinewidth=0.2)

            

    if direccion == "R":
        legend = ax.legend( loc='upper right', numpoints = 1, fontsize=6)
    else:
        legend = ax.legend( loc='upper left', numpoints = 1, fontsize=6)
    plt.grid()
    plt.savefig(ruta_imagenes + "\\"+ filename + ".png", dpi=300, bbox_inches='tight')
    #plt.show()

    if residuals:
        graph_residuals(x, y, slope, intercept, title, xlabel, ylabel, filename + "_residuals") 
    
    if pendiente:
        return slope


def graph_residuals(x, y, slope, intercept, title, 
                   xlabel, ylabel, filename):
    
    residuals = [y[i] - (slope * x[i] + intercept) for i in range(len(x))]

    plt.style.use(['science','no-latex'])
    
    fig = plt.figure()
    ax = fig.add_subplot()

    ax.scatter(x, residuals, label='Residuos', color='blue', marker='o', s=2)
    ax.axhline(y=0, color='r', linestyle='-')

    plt.title(title + " - Residuos", fontsize=8)
    plt.xlabel(xlabel,fontsize=8)
    plt.ylabel("Residuos",fontsize=8)
    plt.grid()
    plt.savefig(ruta_imagenes + "\\"+filename+".png", dpi=300, bbox_inches='tight')
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
    archivo_excel = r"C:\Users\eduar\OneDrive - Universidad de los Andes\OTROS\Sofi-Edu\2023-10\Laboratorio Intermedio\Practica 4 - Oscilador Torsional\OsciladorTorsional.xlsx"
    
    # Carga del libro de Excel
    libro_excel = openpyxl.load_workbook(archivo_excel)

    ruta_imagenes = r"C:\Users\eduar\OneDrive - Universidad de los Andes\OTROS\Sofi-Edu\2023-10\Laboratorio Intermedio\Laboratorio-Intermedio-202310\4. Oscilador Torsional"
    
    plt.style.use(['science','no-latex'])

    gravedad = 9.81 # (m/s^2)

    main()

    