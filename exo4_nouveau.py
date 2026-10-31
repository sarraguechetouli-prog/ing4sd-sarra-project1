import pandas as pd
import numpy as np
import math
from tabulate import tabulate

data = {
    'V1': [80, 44, np.nan, 60, np.nan],
    'V2': [30, np.nan, 85, 70, 54],
    'V3': [7, 10, 25, 74, 49],
    'V4': [14, 0, 5, 9, 20]
}

df = pd.DataFrame(data, index=['I0', 'I1', 'I2', 'I3', 'I4'])
nb_colonnes = df.shape[1]
k = 2

print("\n" + "=" * 80)
print(" EXERCICE 5 ".center(80))
print("=" * 80 + "\n")

print("Tableau Initial :")
print(tabulate(df, headers='keys', tablefmt='fancy_grid', numalign="center", stralign="center"))
print("\n")

df_resultat = df.copy()

for nom_colonne in df.columns:
    lignes_manquantes = df_resultat.index[df_resultat[nom_colonne].isna()]

    if len(lignes_manquantes) == 0:
        continue

    for cible_ligne in lignes_manquantes:
        print(f"\n>>> Traitement : {cible_ligne} (Variable {nom_colonne}) <<<")

        vecteur_cible = df_resultat.loc[cible_ligne].values
        tableau_calculs = []
        candidats_details = []

        for candidat_ligne in df_resultat.index:
            if cible_ligne == candidat_ligne:
                continue

            vecteur_candidat = df_resultat.loc[candidat_ligne].values

            diff_text_parts = []
            somme_diff_carre = 0
            nb_coord_presentes = 0

            for i in range(nb_colonnes):
                val_cible = vecteur_cible[i]
                val_candidat = vecteur_candidat[i]

                if not np.isnan(val_cible) and not np.isnan(val_candidat):
                    diff = val_cible - val_candidat
                    somme_diff_carre += diff ** 2
                    nb_coord_presentes += 1
                    diff_text_parts.append(f"({int(val_cible)}-{int(val_candidat)})²")

            if nb_coord_presentes > 0:
                poids = (nb_colonnes - 1) / nb_coord_presentes
                distance = math.sqrt(poids * somme_diff_carre)

                partie_somme = " + ".join(diff_text_parts)
                formule_str = f"√[ {poids:.2f} * ({partie_somme}) ]"

                tableau_calculs.append([
                    f"{cible_ligne} - {candidat_ligne}",
                    formule_str,
                    f"{distance:.4f}"
                ])

                val_cand = df_resultat.loc[candidat_ligne, nom_colonne]
                candidats_details.append((distance, val_cand, candidat_ligne))

        headers = ["Paire (Ii - Ij)", "Formule de Distance", "Distance Résultat"]
        print(tabulate(tableau_calculs, headers=headers, tablefmt='fancy_grid', stralign="left"))

        candidats_details.sort(key=lambda x: x[0])
        voisins_top_k = candidats_details[:k]

        print(f"\n-> Les {k} plus proches voisins sélectionnés (Top-k) :")
        valeurs_pour_calcul = []

        for dist, val, nom in voisins_top_k:
            statut = "NaN (Ignoré dans la moyenne)" if np.isnan(val) else f"{val} (Utilisé)"
            print(f"   * Voisin : {nom} | Distance : {dist:.4f} | Valeur : {statut}")

            if not np.isnan(val):
                valeurs_pour_calcul.append(val)

        if len(valeurs_pour_calcul) > 0:
            moyenne = sum(valeurs_pour_calcul) / len(valeurs_pour_calcul)
            calcul_str = f"({' + '.join([str(v) for v in valeurs_pour_calcul])}) / {len(valeurs_pour_calcul)}"
            print(f"\n-> Calcul Imputation : {calcul_str} = {moyenne:.2f}")

            df_resultat.loc[cible_ligne, nom_colonne] = moyenne
        else:
            print("\n-> Impossible d'imputer : Les k voisins les plus proches sont tous NaN !")

        print("-" * 80)

print("\n" + "=" * 80)
print(" RÉSULTAT FINAL ".center(80))
print("=" * 80)
print(tabulate(df_resultat, headers='keys', tablefmt='fancy_grid', numalign="center", floatfmt=".2f"))