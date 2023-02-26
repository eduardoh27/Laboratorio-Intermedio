import pandas as pd
import matplotlib.pyplot as plt


def main():

    voltajesA = turnColumnIntoList('A', "Voltage U1", 1726)
    corrientesB = turnColumnIntoList('B', "Corriente IA", 1726)

    print(voltajesA)
    print(corrientesB)


    fig = plt.figure()
    ax = fig.add_subplot()
    ax.set_title('$V_o$ vs $V_i$')
    ax.set_xlabel('$V_i$ (V)',fontsize=10)
    ax.set_ylabel('$V_o$ (V)',fontsize=10)
    #ax.set_xlim(-5,12)
    #ax.set_ylim(5,15)
    plt.plot(voltajesA, corrientesB)
    #plt.savefig('tarea4edu.png', dpi=300, bbox_inches='tight')
    plt.show()


def turnColumnIntoList(colLetter: str, colName: str, lastExcelRow: int):

    df = pd.read_excel(r"C:\Users\eduar\Downloads\datos frank hertz para python.xlsx", usecols=colLetter) 
    np_array = df[colName].values[:lastExcelRow-1]
    
    return np_array


if __name__ == "__main__":
    main()

    