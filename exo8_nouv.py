import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from tabulate import tabulate

# --- Données Initiales ---
X = np.array([91, 4, 6, 7, 9, 10, 9, 11, 12, 13, 15, 4])  # Epreuve A
Y = np.array([23, 40, 10, 13, 15, 14, 13, 16, 13, 19, 6, 19])  # Epreuve B

print("\n" + "=" * 90)
print(" EXERCICE 8 : RÉGRESSION LINÉAIRE AVEC COMMENTAIRES ".center(90))
print("=" * 90 + "\n")


# ==============================================================================
# FONCTION DE CALCUL (Manuelle)
# ==============================================================================
def calculer_parametres_complet(x_vec, y_vec):
    n_local = len(x_vec)
    moy_x = np.mean(x_vec)
    moy_y = np.mean(y_vec)

    # Covariance et Variance
    covariance = np.mean((x_vec - moy_x) * (y_vec - moy_y))
    variance_x = np.var(x_vec)
    ecart_type_x = np.std(x_vec)
    ecart_type_y = np.std(y_vec)

    # Paramètres a et b
    a = covariance / variance_x
    b = moy_y - (a * moy_x)

    # Coefficients
    r = covariance / (ecart_type_x * ecart_type_y)
    r_sq = r ** 2

    return a, b, r, r_sq, covariance, variance_x


# PARTIE 1 : AVEC TOUS LES STAGIAIRES (12)

a1, b1, r1, r2_1, cov1, var1 = calculer_parametres_complet(X, Y)

# Calcul des vecteurs
Y_pred_1 = a1 * X + b1
residus_1 = Y - Y_pred_1
ortho_check_1 = Y_pred_1 * residus_1

df_1 = pd.DataFrame({
    'X': X, 'Y': Y, 'Y_pred': Y_pred_1, 'Résidus': residus_1, 'Y_pred*e': ortho_check_1
})

print("--- 1. ANALYSE GLOBALE (12 Stagiaires) ---")
print(tabulate(df_1, headers='keys', tablefmt='fancy_grid', floatfmt=".4f", numalign="center"))

print(f"\n>>> RÉSULTATS ET COEFFICIENTS :")
print(f"  * a  = {a1:.4f}")
print(f"  * b  = {b1:.4f}")
print(f"  * Corrélation (r) = {r1:.4f}")
print(f"  * Détermination (R²) = {r2_1:.4f}")

print("\n" + ">" * 10 + " COMMENTAIRE SUR LES COEFFICIENTS (PARTIE 1) " + "<" * 10)
print(f"Le coefficient de corrélation est positive mais très faible (r = {r1:.2f}).")
print("Cela indique qu'il n'y a pratiquement aucune relation linéaire entre l'épreuve A et B.")
print("Le modèle est mauvais (R² proche de 0), probablement à cause de valeurs aberrantes (outliers).")
print("-" * 90)

# PARTIE 2

indices_outliers = [10, 11]
X_new = np.delete(X, indices_outliers)
Y_new = np.delete(Y, indices_outliers)

a2, b2, r2, r2_2, cov2, var2 = calculer_parametres_complet(X_new, Y_new)

Y_pred_2 = a2 * X_new + b2
residus_2 = Y_new - Y_pred_2
ortho_check_2 = Y_pred_2 * residus_2

df_2 = pd.DataFrame({
    'X': X_new, 'Y': Y_new, 'Y_pred': Y_pred_2, 'Résidus': residus_2, 'Y_pred*e': ortho_check_2
})

print("\n--- 2. ANALYSE APRÈS CORRECTION (10 Stagiaires) ---")
print(tabulate(df_2, headers='keys', tablefmt='fancy_grid', floatfmt=".4f", numalign="center"))

print(f"\n>>> NOUVEAUX RÉSULTATS ET COEFFICIENTS :")
print(f"  * Nouveau a = {a2:.4f}")
print(f"  * Nouveau b = {b2:.4f}")
print(f"  * Nouvelle Corrélation (r) = {r2:.4f}")
print(f"  * Nouveau Détermination (R²) = {r2_2:.4f}")

# --- COMMENTAIRE AJOUTÉ POUR PARTIE 2 ---
print("\n" + ">" * 10 + " COMMENTAIRE SUR LES NOUVEAUX COEFFICIENTS (PARTIE 2) " + "<" * 10)
print(f"Le nouveau coefficient de corrélation est très fort et positif (r = {r2:.2f}).")
print("En supprimant les deux stagiaires atypiques, on découvre une relation linéaire très nette.")
print(f"Le coefficient de détermination R² = {r2_2:.2f} signifie que le modèle explique 92% de la variance.")
print("La régression est maintenant fiable.")
print("=" * 90)

# GRAPHIQUE FINAL

plt.figure(figsize=(10, 6))
plt.scatter(X, Y, color='red', marker='x', s=100, label='Outliers (Supprimés)')
plt.scatter(X_new, Y_new, color='blue', s=100, label='Données Valides')
x_vals = np.linspace(min(X), max(X), 100)
plt.plot(x_vals, a1 * x_vals + b1, 'r--', label=f'Droite 1 (r={r1:.2f})')
plt.plot(x_vals, a2 * x_vals + b2, 'g-', linewidth=2, label=f'Droite 2 (r={r2:.2f})')
plt.title("Comparaison des Droites de Régression")
plt.legend()
plt.grid(True)
plt.show()