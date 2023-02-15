import numpy as np
import matplotlib.pyplot as plt

def main():
    
    temperaturas = np.linspace(300, 500, 1000)
    caminos = l_bar(temperaturas)



    fig = plt.figure()
    ax = fig.add_subplot()
    ax.set_title('$l$ vs $T$')
    ax.set_xlabel('$T\,[K]$',fontsize=10)
    ax.set_ylabel('$l\,[\mu m]$   ',fontsize=10)
    #ax.set_xlim(-5,12)
    #ax.set_ylim(5,15)

    plt.plot(temperaturas, caminos, c='b')
    
    plt.savefig('preinforme1.png', dpi=300, bbox_inches='tight')
    plt.show()
    

def l_bar (temperaturas):
    r = 155 * (10**(-12))
    sigma = np.pi * (r**2)
    return 1 / (np.sqrt(2) * sigma * n(temperaturas)) 

def n(T):
    k_B = 1.38065 * (10**(-23))
    #k_B = 8.617333 * (10**(-5))
    return P(T) / (k_B*T)

def P(T):
    return 8.7 * (10**(9-(3110/T)))


if __name__ == '__main__':
    k_B = 1.38065 * (10**(-23))
    r = 155 * (10**(-12))
    sigma = np.pi * (r**2)
    print((k_B*300)/(P(300) *sigma))
    main()
