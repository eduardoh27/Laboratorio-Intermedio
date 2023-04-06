import numpy as np
import matplotlib.pyplot as plt
from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import PolynomialFeatures

# Generar datos de ejemplo
np.random.seed(0)
n = 50
months = np.arange(n)
sales = 5000 + 100 * months + np.random.normal(0, 500, n)

# Ajustar el Modelo A (regresión lineal simple con el tiempo)
X_A = months.reshape(-1, 1)
model_A = LinearRegression().fit(X_A, sales  wsedafqzwerfwsdd

# Ajustar el Modelo B (regresión polinomial de grado 4 con el tiempo)
poly = PolynomialFeatures(degree=4)
X_B_poly = poly.fit_transform(X_A)
model_B = LinearRegression().fit(X_B_poly, sales)
predictions_B = model_B.predict(X_B_poly)

# Graficar las predicciones de los dos modelos junto con los datos reales
plt.figure()
plt.scatter(months, sales, label="Datos reales", color="black")
plt.plot(months, predictions_A, label="Modelo A", color="blue")
plt.xlabel("Meses")
plt.ylabel("Ventas")
plt.legend()
plt.savefig( "./ockham1.png", dpi=300, bbox_inches='tight')

plt.figure()
plt.scatter(months, sales, label="Datos reales", color="black")
plt.plot(months, predictions_B, label="Modelo B", color="green")
plt.xlabel("Meses")
plt.ylabel("Ventas")
plt.legend()
plt.savefig( "./ockham2.png", dpi=300, bbox_inches='tight')