import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
import tkinter as tk
from tkinter import ttk, scrolledtext
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import io
import sys


class AFCAnalyzer:
    def __init__(self):
        self.data = None
        self.row_names = None
        self.col_names = None
        self.total = None
        self.profile_row = None
        self.profile_col = None
        self.mass_row = None
        self.mass_col = None
        self.chi2 = None
        self.phi2 = None
        self.inertia = None
        self.svd_results = None

    def load_data(self, data_matrix, row_names, col_names):
        """Charge les données et calcule les matrices nécessaires"""
        self.data = np.array(data_matrix, dtype=float)
        self.row_names = row_names
        self.col_names = col_names

        # Tableau des contingences
        self.contingency = self.data

        # Effectif total
        self.total = np.sum(self.data)

        # Profils lignes (fréquences conditionnelles)
        self.profile_row = self.data / np.sum(self.data, axis=1)[:, np.newaxis]

        # Profils colonnes
        self.profile_col = self.data / np.sum(self.data, axis=0)[np.newaxis, :]

        # Masses des lignes (poids)
        self.mass_row = np.sum(self.data, axis=1) / self.total

        # Masses des colonnes
        self.mass_col = np.sum(self.data, axis=0) / self.total

        # Profil moyen (barycentre)
        self.avg_profile_row = self.mass_col
        self.avg_profile_col = self.mass_row

        print("✓ Données chargées avec succès")

    def compute_chi2(self):
        """Calcule le chi2 et l'inertie"""
        # Effectifs théoriques sous indépendance
        expected = np.outer(np.sum(self.data, axis=1), np.sum(self.data, axis=0)) / self.total

        # Contribution au chi2
        chi2_contrib = (self.data - expected) ** 2 / expected

        # Chi2 total
        self.chi2 = np.sum(chi2_contrib)

        # Phi2
        self.phi2 = self.chi2 / self.total

        # Inertie totale
        self.inertia = self.chi2 / self.total

        print(f"✓ Chi2 calculé: {self.chi2:.4f}")
        print(f"✓ Inertie totale: {self.inertia:.4f}")

        return chi2_contrib

    def compute_svd(self):
        """Effectue la décomposition en valeurs singulières pour l'AFC"""
        # Construction de la matrice à décomposer pour les lignes
        n_rows, n_cols = self.data.shape

        # Matrice des écarts à l'indépendance standardisée
        D_r_inv_sqrt = np.diag(1.0 / np.sqrt(self.mass_row))
        D_c_inv_sqrt = np.diag(1.0 / np.sqrt(self.mass_col))

        # Matrice des fréquences relatives
        F = self.data / self.total

        # Matrice S = D_r^(-1/2) * (F - r*c) * D_c^(-1/2)
        r = self.mass_row.reshape(-1, 1)
        c = self.mass_col.reshape(1, -1)
        F_indep = r @ c

        S = D_r_inv_sqrt @ (F - F_indep) @ D_c_inv_sqrt

        # SVD
        U, s, Vt = np.linalg.svd(S, full_matrices=False)

        # Valeurs singulières
        self.eigenvalues = s ** 2

        # Coordonnées des lignes
        self.row_coords = np.zeros((n_rows, len(s)))
        for k in range(len(s)):
            self.row_coords[:, k] = U[:, k] * s[k]

        # Coordonnées des colonnes
        self.col_coords = np.zeros((n_cols, len(s)))
        for k in range(len(s)):
            self.col_coords[:, k] = (1.0 / np.sqrt(self.mass_col)) * Vt[k, :].T * s[k]

        # Contributions et cos2
        self.row_contrib = np.zeros((n_rows, len(s)))
        self.col_contrib = np.zeros((n_cols, len(s)))
        self.row_cos2 = np.zeros((n_rows, len(s)))
        self.col_cos2 = np.zeros((n_cols, len(s)))

        for k in range(len(s)):
            # Contributions
            self.row_contrib[:, k] = (self.mass_row * self.row_coords[:, k] ** 2) / self.eigenvalues[k]
            self.col_contrib[:, k] = (self.mass_col * self.col_coords[:, k] ** 2) / self.eigenvalues[k]

            # Cos2 (qualité de représentation)
            row_dist2 = np.sum(self.row_coords ** 2, axis=1)
            col_dist2 = np.sum(self.col_coords ** 2, axis=1)

            self.row_cos2[:, k] = self.row_coords[:, k] ** 2 / row_dist2
            self.col_cos2[:, k] = self.col_coords[:, k] ** 2 / col_dist2

        self.svd_results = {
            'U': U, 's': s, 'Vt': Vt,
            'eigenvalues': self.eigenvalues,
            'row_coords': self.row_coords,
            'col_coords': self.col_coords
        }

        print("✓ Décomposition SVD effectuée")

        return self.svd_results

    def get_inertia_explained(self):
        """Retourne le pourcentage d'inertie expliquée"""
        total_inertia = np.sum(self.eigenvalues)
        explained = self.eigenvalues / total_inertia * 100
        cumulative = np.cumsum(explained)

        return explained, cumulative

    def get_interpretation_summary(self):
        """Génère un résumé interprétatif des résultats"""
        explained, cumulative = self.get_inertia_explained()

        summary = []
        summary.append("=" * 60)
        summary.append("RÉSULTATS DE L'ANALYSE FACTORIELLE DES CORRESPONDANCES")
        summary.append("=" * 60)
        summary.append("")

        # Statistiques générales
        summary.append("📊 STATISTIQUES GÉNÉRALES:")
        summary.append(f" • Effectif total: {int(self.total)} individus")
        summary.append(f" • Chi2 d'indépendance: {self.chi2:.4f}")
        summary.append(f" • Inertie totale: {self.inertia:.4f}")
        summary.append("")

        # Valeurs propres
        summary.append("📈 VALEURS PROPRES ET INERTIE EXPLIQUÉE:")
        for i, (ev, expl, cum) in enumerate(zip(self.eigenvalues, explained, cumulative)):
            summary.append(f" • Axe {i + 1}: {ev:.4f} ({expl:.2f}% - cumulé: {cum:.2f}%)")
        summary.append("")

        # Interprétation des axes
        summary.append("🔍 INTERPRÉTATION DES AXES PRINCIPAUX:")

        # Axe 1
        summary.append(f"\n AXE 1 ({explained[0]:.2f}% de l'inertie):")

        # Contributions des lignes à l'axe 1
        row_contrib1 = [(self.row_names[i], self.row_contrib[i, 0]) for i in range(len(self.row_names))]
        row_contrib1.sort(key=lambda x: x[1], reverse=True)

        summary.append(" ➤ Contributions des modalités lignes (sujets):")
        for name, contrib in row_contrib1[:3]:
            summary.append(f" - {name}: {contrib:.3f}")

        # Contributions des colonnes à l'axe 1
        col_contrib1 = [(self.col_names[i], self.col_contrib[i, 0]) for i in range(len(self.col_names))]
        col_contrib1.sort(key=lambda x: x[1], reverse=True)

        summary.append(" ➤ Contributions des modalités colonnes (sources):")
        for name, contrib in col_contrib1:
            summary.append(f" - {name}: {contrib:.3f}")

        # Qualité de représentation
        summary.append("\n 📍 QUALITÉ DE REPRÉSENTATION (cos²):")
        summary.append(" ➤ Lignes:")
        for i, name in enumerate(self.row_names):
            cos2_axis1 = self.row_cos2[i, 0]
            summary.append(f" - {name}: {cos2_axis1:.3f}")

        summary.append(" ➤ Colonnes:")
        for i, name in enumerate(self.col_names):
            cos2_axis1 = self.col_cos2[i, 0]
            summary.append(f" - {name}: {cos2_axis1:.3f}")

        # Synthèse des correspondances
        summary.append("\n💡 SYNTHÈSE DES CORRESPONDANCES:")

        # Regrouper les points proches
        row_pos_axis1 = [(self.row_names[i], self.row_coords[i, 0]) for i in range(len(self.row_names))]
        col_pos_axis1 = [(self.col_names[i], self.col_coords[i, 0]) for i in range(len(self.col_names))]

        all_points = row_pos_axis1 + col_pos_axis1
        all_points.sort(key=lambda x: x[1])

        summary.append(" Ordre des modalités sur l'axe 1 (du négatif au positif):")
        for name, pos in all_points:
            summary.append(f" {name}: {pos:.3f}")

        summary.append("\n INTERPRÉTATION: Les modalités proches sur l'axe sont associées,")
        summary.append(" celles éloignées sont dissociées dans les réponses des individus.")

        return "\n".join(summary)


class AFCAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Analyse Factorielle des Correspondances - Exercice n°6")
        self.root.geometry("1200x800")

        # Style
        self.style = ttk.Style()
        self.style.theme_use('clam')

        # Données de l'exercice
        self.data_matrix = [
            [10, 20, 30],  # PI
            [10, 30, 20],  # PR
            [20, 10, 30],  # FD
            [20, 30, 10]  # RS
        ]
        self.row_names = ['PI', 'PR', 'FD', 'RS']
        self.col_names = ['Journal', 'Radio', 'Télévision']

        # Initialisation de l'analyseur
        self.analyzer = AFCAnalyzer()

        # Création de l'interface
        self.create_widgets()

    def create_widgets(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Titre
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 10))

        title_label = ttk.Label(title_frame, text="Analyse Factorielle des Correspondances (AFC)",
                                font=('Arial', 16, 'bold'))
        title_label.pack()

        subtitle_label = ttk.Label(title_frame,
                                   text="Exercice n°6 : Sources d'information et sujets d'actualité",
                                   font=('Arial', 12))
        subtitle_label.pack()

        # Frame pour les données
        data_frame = ttk.LabelFrame(main_frame, text="📋 Tableau de contingence", padding="10")
        data_frame.pack(fill=tk.X, pady=(0, 10))

        # Afficher le tableau
        data_display = ttk.Frame(data_frame)
        data_display.pack()

        # En-têtes
        headers = ['Sujet / Source'] + self.col_names
        for j, header in enumerate(headers):
            label = ttk.Label(data_display, text=header, font=('Arial', 10, 'bold'),
                              relief=tk.RIDGE, width=12)
            label.grid(row=0, column=j, padx=1, pady=1)

        # Données
        for i, row_name in enumerate(self.row_names):
            # Nom de la ligne
            label = ttk.Label(data_display, text=row_name, font=('Arial', 10, 'bold'),
                              relief=tk.RIDGE, width=12)
            label.grid(row=i + 1, column=0, padx=1, pady=1)

            # Valeurs
            for j, val in enumerate(self.data_matrix[i]):
                label = ttk.Label(data_display, text=str(val), relief=tk.RIDGE, width=12)
                label.grid(row=i + 1, column=j + 1, padx=1, pady=1)

        # Totaux
        col_totals = np.sum(self.data_matrix, axis=0)
        row_totals = np.sum(self.data_matrix, axis=1)
        total = np.sum(self.data_matrix)

        # Ligne des totaux
        label = ttk.Label(data_display, text="Total", font=('Arial', 10, 'bold'),
                          relief=tk.RIDGE, width=12)
        label.grid(row=5, column=0, padx=1, pady=1)

        for j, total_val in enumerate(col_totals):
            label = ttk.Label(data_display, text=str(total_val), font=('Arial', 10, 'bold'),
                              relief=tk.RIDGE, width=12)
            label.grid(row=5, column=j + 1, padx=1, pady=1)

        label = ttk.Label(data_display, text=f"Total: {total}", font=('Arial', 10, 'bold'),
                          relief=tk.RIDGE, width=12)
        label.grid(row=5, column=4, padx=1, pady=1)

        # Frame pour les boutons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)

        ttk.Button(button_frame, text="Exécuter l'AFC",
                   command=self.run_analysis).pack(side=tk.LEFT, padx=5)

        ttk.Button(button_frame, text="Afficher le nuage de points",
                   command=self.show_cloud).pack(side=tk.LEFT, padx=5)

        ttk.Button(button_frame, text="Afficher les profils",
                   command=self.show_profiles).pack(side=tk.LEFT, padx=5)

        # Frame pour les résultats
        result_frame = ttk.LabelFrame(main_frame, text="📊 Résultats de l'analyse", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True)

        # Notebook pour les onglets
        self.notebook = ttk.Notebook(result_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Onglet des résultats textuels
        self.text_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.text_frame, text="Résultats textuels")

        self.text_result = scrolledtext.ScrolledText(self.text_frame, wrap=tk.WORD,
                                                     width=80, height=20)
        self.text_result.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Onglet des graphiques
        self.graph_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.graph_frame, text="Graphiques")

        # Frame pour les contrôles du graphique
        graph_control = ttk.Frame(self.graph_frame)
        graph_control.pack(fill=tk.X, pady=5)

        ttk.Label(graph_control, text="Axe X:").pack(side=tk.LEFT, padx=5)
        self.axis_x = ttk.Combobox(graph_control, values=[1, 2], width=5)
        self.axis_x.set(1)
        self.axis_x.pack(side=tk.LEFT, padx=5)

        ttk.Label(graph_control, text="Axe Y:").pack(side=tk.LEFT, padx=5)
        self.axis_y = ttk.Combobox(graph_control, values=[1, 2], width=5)
        self.axis_y.set(2)
        self.axis_y.pack(side=tk.LEFT, padx=5)

        ttk.Button(graph_control, text="Mettre à jour",
                   command=self.update_graph).pack(side=tk.LEFT, padx=20)

        # Canvas pour le graphique
        self.fig, self.ax = plt.subplots(figsize=(8, 6))
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.graph_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Onglet des interprétations
        self.interp_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.interp_frame, text="Interprétation")

        self.interp_text = scrolledtext.ScrolledText(self.interp_frame, wrap=tk.WORD,
                                                     width=80, height=20)
        self.interp_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def run_analysis(self):
        """Exécute l'analyse AFC"""
        # Charger les données
        self.analyzer.load_data(self.data_matrix, self.row_names, self.col_names)

        # Calculs
        self.analyzer.compute_chi2()
        self.analyzer.compute_svd()

        # Afficher les résultats
        self.display_results()

        # Afficher l'interprétation
        self.display_interpretation()

        # Afficher le graphique
        self.plot_cloud()

    def display_results(self):
        """Affiche les résultats dans l'onglet texte"""
        self.text_result.delete(1.0, tk.END)

        output = []
        output.append("=" * 60)
        output.append("RÉSULTATS DÉTAILLÉS DE L'ANALYSE FACTORIELLE DES CORRESPONDANCES")
        output.append("=" * 60)
        output.append("")

        # 1. Nuages de points
        output.append("1️⃣ NUAGES DE POINTS N(I) et N(J)")
        output.append("-" * 40)

        output.append("\nNuage N(I) - Profils lignes (sujets d'actualité):")
        for i, name in enumerate(self.row_names):
            output.append(f" {name}: [{', '.join([f'{x:.3f}' for x in self.analyzer.profile_row[i]])}]")

        output.append("\nNuage N(J) - Profils colonnes (sources d'information):")
        for j, name in enumerate(self.col_names):
            output.append(f" {name}: [{', '.join([f'{x:.3f}' for x in self.analyzer.profile_col[:, j]])}]")

        # 2. Masses
        output.append("\n\n2️⃣ MASSES DES POINTS")
        output.append("-" * 40)

        output.append("\nMasses des lignes (poids des sujets):")
        for i, name in enumerate(self.row_names):
            output.append(f" {name}: {self.analyzer.mass_row[i]:.3f}")

        output.append("\nMasses des colonnes (poids des sources):")
        for j, name in enumerate(self.col_names):
            output.append(f" {name}: {self.analyzer.mass_col[j]:.3f}")

        # 3. Distances du chi2
        output.append("\n\n3️⃣ DISTANCES DU CHI2")
        output.append("-" * 40)

        # Distance entre lignes
        output.append("\nDistances entre profils lignes (sujets):")
        for i in range(len(self.row_names)):
            for j in range(i + 1, len(self.row_names)):
                # Calcul de la distance du chi2
                dist = 0
                for k in range(len(self.col_names)):
                    diff = (self.analyzer.profile_row[i, k] - self.analyzer.profile_row[j, k])
                    dist += diff ** 2 / self.analyzer.avg_profile_row[k]
                dist = np.sqrt(dist)
                output.append(f" d({self.row_names[i]}, {self.row_names[j]}) = {dist:.4f}")

        # 4. Inertie
        output.append("\n\n4️⃣ INERTIE TOTALE")
        output.append("-" * 40)
        output.append(f" Inertie totale: {self.analyzer.inertia:.4f}")
        output.append(f" Phi2: {self.analyzer.phi2:.4f}")
        output.append(f" Chi2: {self.analyzer.chi2:.4f}")

        # 5. Axes factoriels
        output.append("\n\n5️⃣ AXES FACTORIELS")
        output.append("-" * 40)

        explained, cumulative = self.analyzer.get_inertia_explained()

        for i in range(len(self.analyzer.eigenvalues)):
            output.append(f"\nAxe {i + 1}:")
            output.append(f" Valeur propre λ{i + 1} = {self.analyzer.eigenvalues[i]:.4f}")
            output.append(f" Inertie expliquée: {explained[i]:.2f}%")
            output.append(f" Inertie cumulée: {cumulative[i]:.2f}%")

        # 6. Coordonnées
        output.append("\n\n6️⃣ COORDONNÉES FACTORIELLES")
        output.append("-" * 40)

        output.append("\nCoordonnées des lignes (sujets):")
        for i, name in enumerate(self.row_names):
            coords = [f"{self.analyzer.row_coords[i, k]:.4f}" for k in range(min(3, len(self.analyzer.eigenvalues)))]
            output.append(f" {name}: F1={coords[0]}, F2={coords[1] if len(coords) > 1 else '-'}")

        output.append("\nCoordonnées des colonnes (sources):")
        for j, name in enumerate(self.col_names):
            coords = [f"{self.analyzer.col_coords[j, k]:.4f}" for k in range(min(3, len(self.analyzer.eigenvalues)))]
            output.append(f" {name}: F1={coords[0]}, F2={coords[1] if len(coords) > 1 else '-'}")

        # 7. Contributions
        output.append("\n\n7️⃣ CONTRIBUTIONS À L'INERTIE DES AXES")
        output.append("-" * 40)

        output.append("\nContributions des lignes à l'axe 1:")
        for i, name in enumerate(self.row_names):
            output.append(f" {name}: {self.analyzer.row_contrib[i, 0]:.4f}")

        output.append("\nContributions des colonnes à l'axe 1:")
        for j, name in enumerate(self.col_names):
            output.append(f" {name}: {self.analyzer.col_contrib[j, 0]:.4f}")

        # 8. Qualité de représentation
        output.append("\n\n8️⃣ QUALITÉ DE REPRÉSENTATION (COS²)")
        output.append("-" * 40)

        output.append("\nCos² des lignes sur l'axe 1:")
        for i, name in enumerate(self.row_names):
            output.append(f" {name}: {self.analyzer.row_cos2[i, 0]:.4f}")

        output.append("\nCos² des colonnes sur l'axe 1:")
        for j, name in enumerate(self.col_names):
            output.append(f" {name}: {self.analyzer.col_cos2[j, 0]:.4f}")

        self.text_result.insert(1.0, "\n".join(output))

    def display_interpretation(self):
        """Affiche l'interprétation dans l'onglet dédié"""
        self.interp_text.delete(1.0, tk.END)

        summary = self.analyzer.get_interpretation_summary()

        # Ajouter des explications pédagogiques
        pedagogical = """

📚 RAPPEL DES RÈGLES D'INTERPRÉTATION DE L'AFC:
─────────────────────────────────────────────────

1️⃣ PROXIMITÉ ENTRE POINTS:
   • Deux points LIGNES sont proches s'ils ont des profils similaires
   • Deux points COLONNES sont proches s'ils ont des profils similaires
   • Un point LIGNE et un point COLONNE sont proches si la modalité ligne est 
     sur-représentée dans la modalité colonne (et vice-versa)

2️⃣ INTERPRÉTATION DES AXES:
   • Un axe oppose généralement des modalités aux profils différents
   • La contribution mesure l'importance d'un point dans la construction de l'axe
   • Le cos² mesure la qualité de représentation du point sur l'axe

3️⃣ POUR NOTRE EXERCICE:
   • Les sujets d'actualité (PI, PR, FD, RS) sont les modalités LIGNES
   • Les sources d'information (Journal, Radio, Télévision) sont les modalités COLONNES
   • L'effectif total est de 60 individus
"""

        self.interp_text.insert(1.0, summary + pedagogical)

    def plot_cloud(self, ax_x=1, ax_y=2):
        """Trace le nuage de points factoriel"""
        self.ax.clear()

        ax_x_idx = ax_x - 1
        ax_y_idx = ax_y - 1

        # Récupérer les coordonnées
        row_x = self.analyzer.row_coords[:, ax_x_idx]
        row_y = self.analyzer.row_coords[:, ax_y_idx]
        col_x = self.analyzer.col_coords[:, ax_x_idx]
        col_y = self.analyzer.col_coords[:, ax_y_idx]

        # Tracer les points
        self.ax.scatter(row_x, row_y, c='blue', s=100, marker='o', label='Sujets (N(I))')
        self.ax.scatter(col_x, col_y, c='red', s=100, marker='s', label='Sources (N(J))')

        # Ajouter les étiquettes
        for i, name in enumerate(self.row_names):
            self.ax.annotate(name, (row_x[i], row_y[i]),
                             xytext=(5, 5), textcoords='offset points',
                             fontsize=10, fontweight='bold')

        for j, name in enumerate(self.col_names):
            self.ax.annotate(name, (col_x[j], col_y[j]),
                             xytext=(5, 5), textcoords='offset points',
                             fontsize=10, fontweight='bold', color='red')

        # Axes
        self.ax.axhline(y=0, color='gray', linestyle='-', linewidth=0.5)
        self.ax.axvline(x=0, color='gray', linestyle='-', linewidth=0.5)

        # Titres et labels
        explained, _ = self.analyzer.get_inertia_explained()

        self.ax.set_xlabel(f'Axe {ax_x} ({explained[ax_x_idx]:.1f}% d\'inertie)')
        self.ax.set_ylabel(f'Axe {ax_y} ({explained[ax_y_idx]:.1f}% d\'inertie)')
        self.ax.set_title('Plan factoriel de l\'AFC')

        # Légende
        self.ax.legend(loc='best')

        # Grille
        self.ax.grid(True, alpha=0.3)

        # Égaliser les axes
        max_abs = max(np.abs(row_x).max(), np.abs(row_y).max(),
                      np.abs(col_x).max(), np.abs(col_y).max())
        self.ax.set_xlim(-max_abs * 1.1, max_abs * 1.1)
        self.ax.set_ylim(-max_abs * 1.1, max_abs * 1.1)

        self.canvas.draw()

    def update_graph(self):
        """Met à jour le graphique avec les axes sélectionnés"""
        ax_x = int(self.axis_x.get())
        ax_y = int(self.axis_y.get())
        self.plot_cloud(ax_x, ax_y)

    def show_cloud(self):
        """Affiche l'onglet du graphique et met à jour le graphique"""
        self.notebook.select(self.graph_frame)
        self.update_graph()

    def show_profiles(self):
        """Affiche les profils dans l'onglet texte"""
        self.notebook.select(self.text_frame)
        # Les profils sont déjà affichés dans les résultats


def main():
    root = tk.Tk()
    app = AFCAnalyzerGUI(root)

    # Exécuter automatiquement l'analyse au démarrage
    app.run_analysis()

    root.mainloop()


if __name__ == "__main__":
    main()
