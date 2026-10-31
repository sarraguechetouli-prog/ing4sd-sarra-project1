import numpy as np
import pandas as pd

# ---------------------------
# 1) Tableau de données
# ---------------------------
df = pd.DataFrame({
    'V1': [100, 90, np.nan, 95],
    'V2': [30, 45, 56, np.nan],
    'V3': [np.nan, 40, 80, 98]
}, index=['I0', 'I1', 'I2', 'I3'])

print("=== Données originales ===")
print(df, "\n")


# ---------------------------
# 2) Distance EUCLIDIENNE corrigée (méthode du PROF)
# ---------------------------
def euclidean_distance_corrected(x, y):
    """

    dist = sqrt( Weight * somme((xi - yi)^2) )

    où Weight = p / k
    p : nombre total de variables
    k : variables présentes chez les deux lignes
    """

    mask = ~np.isnan(x) & ~np.isnan(y)

    k = np.sum(mask)  # nb valeurs disponibles
    p = len(x)  # nb total de colonnes

    if k == 0:
        return np.inf

    diff2 = np.sum((x[mask] - y[mask]) ** 2)

    Weight = p - 1 / k  # formule exacte du cours

    return np.sqrt(Weight * diff2)


# ---------------------------
# 3) Imputation d'une cellule avec affichage
# ---------------------------
def impute_cell_verbose(df, row_idx, col_idx, k=2):
    print(f"\n=== Imputation de {col_idx} pour {row_idx} ===")

    target_row = df.loc[row_idx].values
    distances = []

    for other_idx in df.index:
        if other_idx == row_idx:
            continue

        other_row = df.loc[other_idx].values
        d = euclidean_distance_corrected(target_row, other_row)
        print(f"Distance({row_idx}, {other_idx}) = {d:.4f}")
        distances.append((other_idx, d))

    distances.sort(key=lambda x: x[1])

    print("\n--> Distances triées (du plus proche au plus loin) :")
    for idx, d in distances:
        print(f"{idx} : {d:.4f}")

    neighbors = [idx for idx, d in distances[:k]]

    print(f"\n--> Les {k} plus proches voisins : {neighbors}")

    valeurs_voisins = df.loc[neighbors, col_idx]
    mean_value = valeurs_voisins.mean()

    print(f"--> Valeurs des voisins dans {col_idx} : {list(valeurs_voisins)}")
    print(f"--> Valeur imputée pour {col_idx}({row_idx}) = {mean_value}\n")

    return mean_value


# ---------------------------
# 4) KNN colonne par colonne
# ---------------------------
df_imputed = df.copy()

for col in df_imputed.columns:  # ordre V1 → V2 → V3
    for row in df_imputed.index:
        if pd.isna(df_imputed.loc[row, col]):
            df_imputed.loc[row, col] = impute_cell_verbose(df_imputed, row, col, k=2)

print("\n=== Tableau final après KNN (colonne par colonne) ===")
print(df_imputed)