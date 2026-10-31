import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# -------------------------
# Données
# -------------------------

df = pd.DataFrame(
    [
        [68, 15, 5, 20],
        [119, 54, 29, 84],
        [26, 14, 14, 17],
        [7, 10, 16, 94]
    ],
    index=["Noir", "Brun", "Roux", "Blond"],
    columns=["Marron", "Noisette", "Vert", "Bleu"]
)

print("Tableau de contingence :\n")
print(df)

# -------------------------
# AFC
# -------------------------

data = df.values

# total
n = np.sum(data)

# matrice des fréquences
F = data / n

print("\nMatrice F :\n", F)

# masses lignes et colonnes
fi = np.sum(F, axis=1).reshape(-1, 1)
fj = np.sum(F, axis=0).reshape(1, -1)

print("\nMasses lignes fi :\n", fi)
print("\nMasses colonnes fj :\n", fj)

# matrices diagonales
Di = np.diag(fi.flatten())
Dj = np.diag(fj.flatten())

# matrices inverses racine
Di_inv_sqrt = np.diag(1 / np.sqrt(fi.flatten()))
Dj_inv_sqrt = np.diag(1 / np.sqrt(fj.flatten()))

print("\nDi^-1/2 :\n", Di_inv_sqrt)
print("\nDj^-1/2 :\n", Dj_inv_sqrt)

# matrice X
X = Di_inv_sqrt @ (F - fi @ fj) @ Dj_inv_sqrt

print("\nMatrice X :\n", X)

# matrice S
S = X.T @ X

print("\nMatrice S = X^T X :\n", S)

# valeurs propres et vecteurs propres
eigenvalues, eigenvectors = np.linalg.eig(S)

# tri décroissant
idx = np.argsort(eigenvalues)[::-1]
eigenvalues = eigenvalues[idx]
eigenvectors = eigenvectors[:, idx]

print("\nValeurs propres :\n", eigenvalues)
print("\nVecteurs propres :\n", eigenvectors)

# coordonnées colonnes
Ckj = Dj_inv_sqrt @ eigenvectors

# coordonnées lignes
Cki = X @ eigenvectors

print("\nCoordonnées lignes Cki :\n", Cki)
print("\nCoordonnées colonnes Ckj :\n", Ckj)

# -------------------------
# Graphique AFC
# -------------------------

plt.figure(figsize=(7,7))

x_row = Cki[:,0]
y_row = Cki[:,1]

x_col = Ckj[:,0]
y_col = Ckj[:,1]

plt.scatter(x_row, y_row, label="Cheveux")

for i, txt in enumerate(df.index):
    plt.text(x_row[i], y_row[i], txt)

plt.scatter(x_col, y_col, label="Yeux")

for i, txt in enumerate(df.columns):
    plt.text(x_col[i], y_col[i], txt)

plt.axhline(0)
plt.axvline(0)

plt.title("Analyse Factorielle des Correspondances")
plt.legend()

plt.show()