import sys
import numpy as np
import pandas as pd
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.patches as mpatches
import os


# ─────────────────────────────────────────────
#  DESIGN SYSTEM
# ─────────────────────────────────────────────
PALETTE = {
    "bg":          "#0F1117",
    "surface":     "#1A1D27",
    "surface2":    "#222536",
    "border":      "#2E3349",
    "accent":      "#6C63FF",
    "accent2":     "#00D4AA",
    "accent3":     "#FF6B6B",
    "text":        "#E8EAF6",
    "text_muted":  "#8892B0",
    "text_dim":    "#4A5568",
    "success":     "#00D4AA",
    "warning":     "#FFB347",
    "danger":      "#FF6B6B",
}

STYLESHEET = f"""
    QMainWindow, QWidget {{
        background-color: {PALETTE['bg']};
        color: {PALETTE['text']};
        font-family: 'Segoe UI', 'SF Pro Display', sans-serif;
        font-size: 13px;
    }}

    /* ── Sidebar ── */
    #sidebar {{
        background-color: {PALETTE['surface']};
        border-right: 1px solid {PALETTE['border']};
    }}

    /* ── Section headers ── */
    #section_label {{
        color: {PALETTE['text_muted']};
        font-size: 10px;
        font-weight: 700;
        letter-spacing: 2px;
        text-transform: uppercase;
        padding: 0px 0px 6px 0px;
    }}

    /* ── Cards ── */
    #card {{
        background-color: {PALETTE['surface2']};
        border: 1px solid {PALETTE['border']};
        border-radius: 10px;
        padding: 16px;
    }}

    /* ── QGroupBox ── */
    QGroupBox {{
        background-color: {PALETTE['surface2']};
        border: 1px solid {PALETTE['border']};
        border-radius: 10px;
        margin-top: 18px;
        padding: 14px 12px 12px 12px;
        font-weight: 600;
        font-size: 12px;
        color: {PALETTE['text_muted']};
        letter-spacing: 1px;
        text-transform: uppercase;
    }}
    QGroupBox::title {{
        subcontrol-origin: margin;
        subcontrol-position: top left;
        padding: 2px 10px;
        left: 12px;
        background-color: {PALETTE['surface2']};
        color: {PALETTE['text_muted']};
        border-radius: 4px;
    }}

    /* ── Inputs ── */
    QTextEdit, QPlainTextEdit {{
        background-color: {PALETTE['bg']};
        border: 1px solid {PALETTE['border']};
        border-radius: 8px;
        color: {PALETTE['text']};
        padding: 10px;
        selection-background-color: {PALETTE['accent']};
        font-family: 'Consolas', 'Fira Code', monospace;
        font-size: 12px;
        line-height: 1.6;
    }}
    QTextEdit:focus, QPlainTextEdit:focus {{
        border: 1px solid {PALETTE['accent']};
    }}

    /* ── ComboBox ── */
    QComboBox {{
        background-color: {PALETTE['surface']};
        border: 1px solid {PALETTE['border']};
        border-radius: 7px;
        color: {PALETTE['text']};
        padding: 7px 12px;
        min-height: 32px;
        selection-background-color: {PALETTE['accent']};
    }}
    QComboBox:hover {{
        border: 1px solid {PALETTE['accent']};
    }}
    QComboBox::drop-down {{
        border: none;
        width: 28px;
    }}
    QComboBox::down-arrow {{
        image: none;
        border-left: 5px solid transparent;
        border-right: 5px solid transparent;
        border-top: 6px solid {PALETTE['text_muted']};
        margin-right: 8px;
    }}
    QComboBox QAbstractItemView {{
        background-color: {PALETTE['surface2']};
        border: 1px solid {PALETTE['border']};
        border-radius: 7px;
        color: {PALETTE['text']};
        selection-background-color: {PALETTE['accent']};
        padding: 4px;
    }}

    /* ── SpinBox ── */
    QDoubleSpinBox, QSpinBox {{
        background-color: {PALETTE['surface']};
        border: 1px solid {PALETTE['border']};
        border-radius: 7px;
        color: {PALETTE['text']};
        padding: 7px 10px;
        min-height: 32px;
    }}
    QDoubleSpinBox:hover, QSpinBox:hover {{
        border: 1px solid {PALETTE['accent']};
    }}
    QDoubleSpinBox::up-button, QDoubleSpinBox::down-button,
    QSpinBox::up-button, QSpinBox::down-button {{
        background-color: {PALETTE['surface2']};
        border: none;
        border-radius: 4px;
        width: 20px;
    }}

    /* ── Buttons ── */
    QPushButton {{
        background-color: {PALETTE['surface2']};
        color: {PALETTE['text']};
        border: 1px solid {PALETTE['border']};
        border-radius: 8px;
        padding: 9px 16px;
        font-weight: 600;
        font-size: 12px;
        min-height: 36px;
    }}
    QPushButton:hover {{
        background-color: {PALETTE['surface']};
        border: 1px solid {PALETTE['accent']};
        color: {PALETTE['accent']};
    }}
    QPushButton:pressed {{
        background-color: {PALETTE['accent']};
        color: white;
        border: 1px solid {PALETTE['accent']};
    }}

    /* Primary button */
    #btn_run {{
        background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
            stop:0 {PALETTE['accent']}, stop:1 #8B5CF6);
        color: white;
        border: none;
        font-size: 13px;
        font-weight: 700;
        letter-spacing: 0.5px;
        min-height: 42px;
        border-radius: 10px;
    }}
    #btn_run:hover {{
        background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
            stop:0 #7C74FF, stop:1 #9D6BFF);
    }}
    #btn_run:pressed {{
        background-color: #5A52D5;
    }}

    /* Excel button */
    #btn_excel {{
        background-color: #1A4D30;
        color: #00D4AA;
        border: 1px solid #00D4AA44;
        font-weight: 600;
    }}
    #btn_excel:hover {{
        background-color: #1F5A38;
        border: 1px solid {PALETTE['success']};
    }}

    /* Export button */
    #btn_export {{
        background-color: #3D2B0F;
        color: {PALETTE['warning']};
        border: 1px solid #FFB34744;
        font-weight: 600;
    }}
    #btn_export:hover {{
        background-color: #4A3415;
        border: 1px solid {PALETTE['warning']};
    }}

    /* ── TabWidget ── */
    QTabWidget::pane {{
        background-color: {PALETTE['surface']};
        border: 1px solid {PALETTE['border']};
        border-radius: 10px;
        padding: 4px;
    }}
    QTabBar::tab {{
        background-color: transparent;
        color: {PALETTE['text_muted']};
        border: none;
        padding: 10px 18px;
        border-radius: 8px;
        font-weight: 600;
        font-size: 12px;
        margin-right: 2px;
        min-width: 110px;
    }}
    QTabBar::tab:selected {{
        background-color: {PALETTE['surface2']};
        color: {PALETTE['text']};
        border-bottom: 2px solid {PALETTE['accent']};
    }}
    QTabBar::tab:hover:!selected {{
        background-color: {PALETTE['surface2']};
        color: {PALETTE['text']};
    }}

    /* ── Scrollbars ── */
    QScrollBar:vertical {{
        background: {PALETTE['surface']};
        width: 8px;
        border-radius: 4px;
        margin: 0;
    }}
    QScrollBar::handle:vertical {{
        background: {PALETTE['border']};
        border-radius: 4px;
        min-height: 30px;
    }}
    QScrollBar::handle:vertical:hover {{
        background: {PALETTE['text_dim']};
    }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
    QScrollBar:horizontal {{
        background: {PALETTE['surface']};
        height: 8px;
        border-radius: 4px;
        margin: 0;
    }}
    QScrollBar::handle:horizontal {{
        background: {PALETTE['border']};
        border-radius: 4px;
        min-width: 30px;
    }}
    QScrollBar::handle:horizontal:hover {{
        background: {PALETTE['text_dim']};
    }}
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0; }}

    /* ── TableWidget ── */
    QTableWidget {{
        background-color: {PALETTE['bg']};
        border: 1px solid {PALETTE['border']};
        border-radius: 8px;
        gridline-color: {PALETTE['border']};
        color: {PALETTE['text']};
        font-size: 12px;
        selection-background-color: {PALETTE['accent']}44;
    }}
    QTableWidget::item {{
        padding: 8px 12px;
        border-bottom: 1px solid {PALETTE['surface2']};
    }}
    QTableWidget::item:selected {{
        background-color: {PALETTE['accent']}33;
        color: {PALETTE['text']};
    }}
    QHeaderView::section {{
        background-color: {PALETTE['surface2']};
        color: {PALETTE['text_muted']};
        border: none;
        border-bottom: 1px solid {PALETTE['border']};
        padding: 10px 12px;
        font-weight: 700;
        font-size: 11px;
        letter-spacing: 1px;
        text-transform: uppercase;
    }}

    /* ── Labels ── */
    QLabel {{
        color: {PALETTE['text']};
        background: transparent;
    }}
    #label_muted {{
        color: {PALETTE['text_muted']};
        font-size: 12px;
    }}
    #status_label {{
        color: {PALETTE['success']};
        font-size: 11px;
        padding: 6px 10px;
        background-color: {PALETTE['success']}18;
        border: 1px solid {PALETTE['success']}44;
        border-radius: 6px;
    }}
    #path_label {{
        color: {PALETTE['text_muted']};
        font-size: 11px;
        padding: 6px 10px;
        background-color: {PALETTE['surface']};
        border: 1px solid {PALETTE['border']};
        border-radius: 6px;
    }}

    /* ── Splitter ── */
    QSplitter::handle {{
        background-color: {PALETTE['border']};
        width: 1px;
    }}

    /* ── MessageBox ── */
    QMessageBox {{
        background-color: {PALETTE['surface']};
    }}
    QMessageBox QLabel {{
        color: {PALETTE['text']};
    }}
"""


def mpl_dark_style():
    """Apply dark style to matplotlib figures."""
    return {
        "figure.facecolor":  PALETTE["surface"],
        "axes.facecolor":    PALETTE["bg"],
        "axes.edgecolor":    PALETTE["border"],
        "axes.labelcolor":   PALETTE["text_muted"],
        "axes.titlecolor":   PALETTE["text"],
        "xtick.color":       PALETTE["text_muted"],
        "ytick.color":       PALETTE["text_muted"],
        "grid.color":        PALETTE["border"],
        "text.color":        PALETTE["text"],
        "figure.edgecolor":  PALETTE["surface"],
    }


# ─────────────────────────────────────────────
#  METRIC CARD WIDGET
# ─────────────────────────────────────────────
class MetricCard(QFrame):
    def __init__(self, title, value="—", accent=None, parent=None):
        super().__init__(parent)
        self._accent = accent or PALETTE["accent"]
        self.setObjectName("card")
        self.setFixedHeight(80)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 10, 14, 10)
        lay.setSpacing(4)

        self._title_lbl = QLabel(title.upper())
        self._title_lbl.setStyleSheet(f"color:{PALETTE['text_dim']};font-size:10px;font-weight:700;letter-spacing:1.5px;")
        lay.addWidget(self._title_lbl)

        self._value_lbl = QLabel(str(value))
        self._value_lbl.setStyleSheet(f"color:{self._accent};font-size:20px;font-weight:700;")
        lay.addWidget(self._value_lbl)

        # Accent bar at top
        self.setStyleSheet(f"""
            MetricCard {{
                background-color:{PALETTE['surface2']};
                border:1px solid {PALETTE['border']};
                border-top: 3px solid {self._accent};
                border-radius:10px;
            }}
        """)

    def set_value(self, v):
        self._value_lbl.setText(str(v))


# ─────────────────────────────────────────────
#  BADGE LABEL
# ─────────────────────────────────────────────
def badge(text, color=None):
    color = color or PALETTE["accent"]
    lbl = QLabel(text)
    lbl.setStyleSheet(f"""
        background-color:{color}22;
        color:{color};
        border:1px solid {color}66;
        border-radius:5px;
        padding:2px 8px;
        font-size:11px;
        font-weight:700;
    """)
    lbl.setFixedHeight(22)
    return lbl


# ─────────────────────────────────────────────
#  SEPARATOR
# ─────────────────────────────────────────────
def hline():
    f = QFrame()
    f.setFrameShape(QFrame.HLine)
    f.setStyleSheet(f"color:{PALETTE['border']};")
    return f


# ─────────────────────────────────────────────
#  MAIN APP
# ─────────────────────────────────────────────
class ACPApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.data = None
        self.results = {}
        self.initUI()

    def initUI(self):
        self.setWindowTitle("ACP Studio — Analyse en Composantes Principales")
        self.setGeometry(80, 60, 1560, 900)
        self.setMinimumSize(1200, 700)

        # Central widget
        central = QWidget()
        self.setCentralWidget(central)
        root = QHBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── SIDEBAR ──────────────────────────────────────────
        sidebar = QWidget()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(320)
        sidebar_lay = QVBoxLayout(sidebar)
        sidebar_lay.setContentsMargins(20, 24, 20, 20)
        sidebar_lay.setSpacing(0)

        # App title
        title_row = QHBoxLayout()
        dot = QLabel("●")
        dot.setStyleSheet(f"color:{PALETTE['accent']};font-size:20px;")
        title_row.addWidget(dot)
        app_title = QLabel("ACP Studio")
        app_title.setStyleSheet(f"color:{PALETTE['text']};font-size:18px;font-weight:700;letter-spacing:1px;")
        title_row.addWidget(app_title)
        title_row.addStretch()
        sidebar_lay.addLayout(title_row)

        sub = QLabel("Analyse en Composantes Principales")
        sub.setStyleSheet(f"color:{PALETTE['text_muted']};font-size:11px;margin-left:26px;margin-bottom:20px;")
        sidebar_lay.addWidget(sub)
        sidebar_lay.addSpacing(8)
        sidebar_lay.addWidget(hline())
        sidebar_lay.addSpacing(16)

        # ── Data section ──
        sec1 = QLabel("DONNÉES")
        sec1.setObjectName("section_label")
        sidebar_lay.addWidget(sec1)
        sidebar_lay.addSpacing(8)

        self.data_text = QTextEdit()
        self.data_text.setPlaceholderText(
            "Collez vos données ici\n(virgules ou espaces)\n\nEx:\n1,2,3\n4,5,6")
        self.data_text.setFixedHeight(130)
        sidebar_lay.addWidget(self.data_text)
        sidebar_lay.addSpacing(8)

        btn_row = QHBoxLayout()
        btn_example = QPushButton("Exemple")
        btn_example.clicked.connect(self.load_example_data)
        btn_row.addWidget(btn_example)

        btn_excel = QPushButton("📂  Importer Excel")
        btn_excel.setObjectName("btn_excel")
        btn_excel.clicked.connect(self.load_excel_data)
        btn_row.addWidget(btn_excel)
        sidebar_lay.addLayout(btn_row)

        self.excel_info_label = QLabel("")
        self.excel_info_label.setObjectName("status_label")
        self.excel_info_label.hide()
        sidebar_lay.addWidget(self.excel_info_label)
        sidebar_lay.addSpacing(20)
        sidebar_lay.addWidget(hline())
        sidebar_lay.addSpacing(16)

        # ── ACP Parameters ──
        sec2 = QLabel("PARAMÈTRES")
        sec2.setObjectName("section_label")
        sidebar_lay.addWidget(sec2)
        sidebar_lay.addSpacing(8)

        lbl_type = QLabel("Type d'ACP")
        lbl_type.setObjectName("label_muted")
        sidebar_lay.addWidget(lbl_type)
        self.acp_type = QComboBox()
        self.acp_type.addItems(["ACP Normée", "ACP Non normée"])
        self.acp_type.currentTextChanged.connect(self.on_acp_type_changed)
        sidebar_lay.addWidget(self.acp_type)
        sidebar_lay.addSpacing(10)

        lbl_dtype = QLabel("Type de données")
        lbl_dtype.setObjectName("label_muted")
        sidebar_lay.addWidget(lbl_dtype)
        self.data_type = QComboBox()
        self.data_type.addItems(["Variables homogènes", "Variables hétérogènes"])
        self.data_type.setEnabled(False)
        sidebar_lay.addWidget(self.data_type)
        sidebar_lay.addSpacing(10)

        params_row = QHBoxLayout()
        params_row.setSpacing(10)

        left_p = QVBoxLayout()
        lbl_k = QLabel("Critère Kaiser (λ)")
        lbl_k.setObjectName("label_muted")
        left_p.addWidget(lbl_k)
        self.kaiser_threshold = QDoubleSpinBox()
        self.kaiser_threshold.setRange(0.1, 10.0)
        self.kaiser_threshold.setValue(1.0)
        self.kaiser_threshold.setSingleStep(0.1)
        left_p.addWidget(self.kaiser_threshold)
        params_row.addLayout(left_p)

        right_p = QVBoxLayout()
        lbl_i = QLabel("Inertie cumulée")
        lbl_i.setObjectName("label_muted")
        right_p.addWidget(lbl_i)
        self.inertia_threshold = QDoubleSpinBox()
        self.inertia_threshold.setRange(0, 100)
        self.inertia_threshold.setValue(80)
        self.inertia_threshold.setSuffix("%")
        right_p.addWidget(self.inertia_threshold)
        params_row.addLayout(right_p)

        sidebar_lay.addLayout(params_row)
        sidebar_lay.addSpacing(20)

        # ── Run button ──
        self.btn_run = QPushButton("▶  Exécuter l'ACP")
        self.btn_run.setObjectName("btn_run")
        self.btn_run.clicked.connect(self.run_acp)
        sidebar_lay.addWidget(self.btn_run)
        sidebar_lay.addSpacing(16)
        sidebar_lay.addWidget(hline())
        sidebar_lay.addSpacing(16)

        # ── Export section ──
        sec3 = QLabel("EXPORT")
        sec3.setObjectName("section_label")
        sidebar_lay.addWidget(sec3)
        sidebar_lay.addSpacing(8)

        btn_export = QPushButton("⬇  Exporter pour Tableau")
        btn_export.setObjectName("btn_export")
        btn_export.clicked.connect(self.export_for_tableau)
        sidebar_lay.addWidget(btn_export)
        sidebar_lay.addSpacing(8)

        self.tableau_path_label = QLabel("Aucun fichier exporté")
        self.tableau_path_label.setObjectName("path_label")
        self.tableau_path_label.setWordWrap(True)
        sidebar_lay.addWidget(self.tableau_path_label)

        sidebar_lay.addStretch()

        # ── Metric cards at bottom of sidebar ──
        metrics_grid = QGridLayout()
        metrics_grid.setSpacing(8)
        self.card_n     = MetricCard("Individus", "—", PALETTE["accent"])
        self.card_p     = MetricCard("Variables", "—", PALETTE["accent2"])
        self.card_axes  = MetricCard("Axes retenus", "—", PALETTE["warning"])
        self.card_iner  = MetricCard("Inertie expl.", "—", PALETTE["danger"])
        metrics_grid.addWidget(self.card_n,    0, 0)
        metrics_grid.addWidget(self.card_p,    0, 1)
        metrics_grid.addWidget(self.card_axes, 1, 0)
        metrics_grid.addWidget(self.card_iner, 1, 1)
        sidebar_lay.addLayout(metrics_grid)

        root.addWidget(sidebar)

        # ── MAIN CONTENT ─────────────────────────────────────
        content = QWidget()
        content_lay = QVBoxLayout(content)
        content_lay.setContentsMargins(0, 0, 0, 0)
        content_lay.setSpacing(0)

        # Top bar
        topbar = QWidget()
        topbar.setFixedHeight(52)
        topbar.setStyleSheet(f"background-color:{PALETTE['surface']};border-bottom:1px solid {PALETTE['border']};")
        topbar_lay = QHBoxLayout(topbar)
        topbar_lay.setContentsMargins(24, 0, 24, 0)

        self.status_badge = badge("En attente", PALETTE["text_dim"])
        topbar_lay.addWidget(self.status_badge)
        topbar_lay.addStretch()

        self.type_badge = badge("ACP Normée", PALETTE["accent"])
        topbar_lay.addWidget(self.type_badge)

        content_lay.addWidget(topbar)

        # Tabs
        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        content_lay.addWidget(self.tabs)

        # ── Tab 1: Résultats numériques ──
        tab_results = QWidget()
        tr_lay = QVBoxLayout(tab_results)
        tr_lay.setContentsMargins(20, 16, 20, 20)
        tr_lay.setSpacing(12)

        self.results_text = QTextEdit()
        self.results_text.setReadOnly(True)
        self.results_text.setFont(QFont("Consolas", 11))
        self.results_text.setStyleSheet(f"""
            QTextEdit {{
                background-color:{PALETTE['bg']};
                border:1px solid {PALETTE['border']};
                border-radius:10px;
                color:{PALETTE['text']};
                padding:16px;
                line-height:1.7;
            }}
        """)
        tr_lay.addWidget(self.results_text)
        self.tabs.addTab(tab_results, "📊  Résultats")

        # ── Tab 2: Vecteurs propres ──
        tab_eigen = QWidget()
        te_lay = QVBoxLayout(tab_eigen)
        te_lay.setContentsMargins(20, 16, 20, 20)
        self.eigenvectors_text = QTextEdit()
        self.eigenvectors_text.setReadOnly(True)
        self.eigenvectors_text.setFont(QFont("Consolas", 11))
        self.eigenvectors_text.setStyleSheet(f"""
            QTextEdit {{
                background-color:{PALETTE['bg']};
                border:1px solid {PALETTE['border']};
                border-radius:10px;
                color:{PALETTE['text']};
                padding:16px;
            }}
        """)
        te_lay.addWidget(self.eigenvectors_text)
        self.tabs.addTab(tab_eigen, "🧮  Vecteurs propres")

        # ── Tab 3: Graphiques ──
        tab_plots = QWidget()
        tp_lay = QVBoxLayout(tab_plots)
        tp_lay.setContentsMargins(20, 16, 20, 20)
        tp_lay.setSpacing(16)

        plot_titles_row = QHBoxLayout()
        lbl_ind = QLabel("Projection des individus")
        lbl_ind.setStyleSheet(f"color:{PALETTE['text_muted']};font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;")
        lbl_var = QLabel("Cercle des corrélations")
        lbl_var.setStyleSheet(f"color:{PALETTE['text_muted']};font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;")
        plot_titles_row.addWidget(lbl_ind)
        plot_titles_row.addWidget(lbl_var)
        tp_lay.addLayout(plot_titles_row)

        plots_row = QHBoxLayout()
        plots_row.setSpacing(16)

        with plt.rc_context(mpl_dark_style()):
            self.figure_individus = Figure(figsize=(6, 5), tight_layout=True)
            self.canvas_individus = FigureCanvas(self.figure_individus)
            self.canvas_individus.setStyleSheet(f"border-radius:10px;border:1px solid {PALETTE['border']};")

            self.figure_variables = Figure(figsize=(6, 5), tight_layout=True)
            self.canvas_variables = FigureCanvas(self.figure_variables)
            self.canvas_variables.setStyleSheet(f"border-radius:10px;border:1px solid {PALETTE['border']};")

        plots_row.addWidget(self.canvas_individus)
        plots_row.addWidget(self.canvas_variables)
        tp_lay.addLayout(plots_row)

        # Scree plot
        scree_title = QLabel("Éboulis des valeurs propres (Scree Plot)")
        scree_title.setStyleSheet(f"color:{PALETTE['text_muted']};font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;")
        tp_lay.addWidget(scree_title)

        with plt.rc_context(mpl_dark_style()):
            self.figure_scree = Figure(figsize=(12, 2.5), tight_layout=True)
            self.canvas_scree = FigureCanvas(self.figure_scree)
            self.canvas_scree.setFixedHeight(180)
            self.canvas_scree.setStyleSheet(f"border-radius:10px;border:1px solid {PALETTE['border']};")

        tp_lay.addWidget(self.canvas_scree)
        self.tabs.addTab(tab_plots, "📈  Graphiques")

        # ── Tab 4: Contributions ──
        tab_contrib = QWidget()
        tc_lay = QVBoxLayout(tab_contrib)
        tc_lay.setContentsMargins(20, 16, 20, 20)

        self.contrib_text = QTextEdit()
        self.contrib_text.setReadOnly(True)
        self.contrib_text.setFont(QFont("Consolas", 11))
        self.contrib_text.setStyleSheet(f"""
            QTextEdit {{
                background-color:{PALETTE['bg']};
                border:1px solid {PALETTE['border']};
                border-radius:10px;
                color:{PALETTE['text']};
                padding:16px;
            }}
        """)
        tc_lay.addWidget(self.contrib_text)
        self.tabs.addTab(tab_contrib, "📋  Contributions")

        # ── Tab 5: Règles appliquées ──
        tab_rules = QWidget()
        tru_lay = QVBoxLayout(tab_rules)
        tru_lay.setContentsMargins(20, 16, 20, 20)

        self.rules_text = QTextEdit()
        self.rules_text.setReadOnly(True)
        self.rules_text.setFont(QFont("Consolas", 11))
        self.rules_text.setStyleSheet(f"""
            QTextEdit {{
                background-color:{PALETTE['bg']};
                border:1px solid {PALETTE['border']};
                border-radius:10px;
                color:{PALETTE['text']};
                padding:16px;
            }}
        """)
        tru_lay.addWidget(self.rules_text)
        self.tabs.addTab(tab_rules, "📖  Détail calculs")

        # ── Tab 6: Export Tableau ──
        tab_tableau = QWidget()
        tt_lay = QVBoxLayout(tab_tableau)
        tt_lay.setContentsMargins(20, 16, 20, 20)
        tt_lay.setSpacing(12)

        info_box = QTextEdit()
        info_box.setReadOnly(True)
        info_box.setFixedHeight(200)
        info_box.setStyleSheet(f"""
            QTextEdit {{
                background-color:{PALETTE['bg']};
                border:1px solid {PALETTE['border']};
                border-radius:10px;
                color:{PALETTE['text']};
                padding:14px;
                font-size:12px;
            }}
        """)
        info_box.setHtml(f"""
        <div style="color:{PALETTE['text']};font-family:'Segoe UI';line-height:1.7;">
        <b style="color:{PALETTE['accent']};">Comment utiliser dans Tableau :</b><br>
        <ol>
          <li>Cliquer sur <b>Exporter pour Tableau</b> dans la barre latérale</li>
          <li>Ouvrir <b>Tableau Desktop</b></li>
          <li>Se connecter → Fichier texte → sélectionner un CSV exporté</li>
          <li>Créer un <b>scatter plot</b> (Axe1 vs Axe2), un graphique en barres des valeurs propres, ou un cercle de corrélation</li>
        </ol>
        <b style="color:{PALETTE['accent2']};">Fichiers exportés :</b>
        tableau_individus.csv · tableau_variables.csv · tableau_valeurs_propres.csv · tableau_contributions.csv
        </div>
        """)
        tt_lay.addWidget(info_box)

        preview_lbl = QLabel("APERÇU DES DONNÉES")
        preview_lbl.setObjectName("section_label")
        tt_lay.addWidget(preview_lbl)

        self.tableau_preview = QTableWidget()
        self.tableau_preview.setAlternatingRowColors(True)
        self.tableau_preview.setStyleSheet(f"""
            QTableWidget {{ alternate-background-color: {PALETTE['surface2']}; }}
        """)
        tt_lay.addWidget(self.tableau_preview)
        self.tabs.addTab(tab_tableau, "🗃  Export Tableau")

        root.addWidget(content, 1)

    # ─────────────────────────────────────────
    #  SLOTS / HELPERS
    # ─────────────────────────────────────────
    def on_acp_type_changed(self, text):
        self.data_type.setEnabled(text == "ACP Non normée")
        self.type_badge.setText(text)

    def _set_status(self, text, color=None):
        color = color or PALETTE["accent2"]
        self.status_badge.setText(text)
        self.status_badge.setStyleSheet(f"""
            background-color:{color}22;
            color:{color};
            border:1px solid {color}66;
            border-radius:5px;
            padding:2px 8px;
            font-size:11px;
            font-weight:700;
        """)

    def load_example_data(self):
        self.data_text.setText("2,3,5\n3,4,6\n4,5,7\n5,6,8\n6,7,9")

    def load_excel_data(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Ouvrir un fichier Excel", "",
            "Fichiers Excel (*.xlsx *.xls);;Tous les fichiers (*)")
        if not file_path:
            return
        try:
            xl = pd.ExcelFile(file_path)
            sheet_names = xl.sheet_names
            if len(sheet_names) > 1:
                sheet, ok = QInputDialog.getItem(
                    self, "Choisir une feuille",
                    "Sélectionnez la feuille :", sheet_names, 0, False)
                if not ok:
                    return
            else:
                sheet = sheet_names[0]

            df = pd.read_excel(file_path, sheet_name=sheet, header=None)
            first_row = df.iloc[0]
            has_header = any(isinstance(val, str) for val in first_row)
            if has_header:
                df = pd.read_excel(file_path, sheet_name=sheet, header=0)
                df_numeric = df.select_dtypes(include=[np.number])
            else:
                df_numeric = df.apply(pd.to_numeric, errors='coerce').dropna(axis=1, how='all')

            df_numeric = df_numeric.fillna(df_numeric.mean())
            if df_numeric.empty:
                QMessageBox.warning(self, "Attention", "Aucune donnée numérique trouvée.")
                return

            lines = []
            for _, row in df_numeric.iterrows():
                lines.append(','.join(str(round(v, 6)) for v in row))
            self.data_text.setText('\n'.join(lines))

            n_rows, n_cols = df_numeric.shape
            fname = os.path.basename(file_path)
            self.excel_info_label.setText(f"✔  {fname}  ·  {n_rows} × {n_cols}")
            self.excel_info_label.show()

        except Exception as e:
            QMessageBox.critical(self, "Erreur d'import", str(e))

    def parse_data(self):
        text = self.data_text.toPlainText().strip()
        if not text:
            QMessageBox.warning(self, "Erreur", "Veuillez entrer des données")
            return None
        try:
            lines = text.split('\n')
            data = []
            for line in lines:
                if line.strip():
                    values = line.replace(',', ' ').split()
                    data.append([float(x) for x in values])
            return np.array(data)
        except Exception as e:
            QMessageBox.warning(self, "Erreur", f"Erreur de parsing: {str(e)}")
            return None

    # ─────────────────────────────────────────
    #  ACP NORMÉE
    # ─────────────────────────────────────────
    def run_acp_normee(self, X):
        rules = []
        N, p = X.shape

        rules.append("1) Centrage-réduction: Z = (X - mean)/std")
        means = np.mean(X, axis=0)
        stds = np.std(X, axis=0, ddof=0)
        stds = np.where(stds < 1e-10, 1.0, stds)
        Z = (X - means) / stds
        rules.append(f"   Z (centrée-réduite):\n{np.round(Z, 4)}")

        rules.append("\n2) Matrice de corrélation R = (1/N) Z^T Z")
        R = (1 / N) * Z.T @ Z
        rules.append(f"   R =\n{np.round(R, 4)}")

        rules.append("\n3) Valeurs propres λ et vecteurs propres U de R")
        eigenvalues, eigenvectors = np.linalg.eig(R)
        eigenvalues = np.maximum(eigenvalues, 0)
        idx = eigenvalues.argsort()[::-1]
        eigenvalues = eigenvalues[idx]
        eigenvectors = eigenvectors[:, idx]
        rules.append(f"   λ: {np.round(eigenvalues, 4)}")
        rules.append(f"   Trace(R) = {np.trace(R):.4f}  ≈  Σλ = {np.sum(eigenvalues):.4f}")
        for j in range(p):
            rules.append(f"   U{j+1} = {np.round(eigenvectors[:, j], 4)}")

        threshold = self.kaiser_threshold.value()
        rules.append(f"\n4) Critère de Kaiser: λ > {threshold}")
        kaiser_indices = np.where(eigenvalues > threshold)[0]
        rules.append(f"   Axes retenus: {list(kaiser_indices + 1)}")

        U = eigenvectors
        rules.append(f"\n5) Axes principaux U:\n{np.round(U, 4)}")

        rules.append("\n6) Composantes principales C = Z U")
        C = Z @ U
        rules.append(f"   C:\n{np.round(C, 4)}")

        rules.append("\n7) Coordonnées variables: Coord(j,k) = √λ_k · U_k")
        sqrt_eigenvalues = np.sqrt(np.maximum(eigenvalues, 0))
        var_coords = sqrt_eigenvalues * U
        rules.append(f"   Coordonnées variables:\n{np.round(var_coords, 4)}")

        rules.append("\n8) Cosinus carrés des variables:")
        cos2_vars = np.zeros((p, p))
        for k in range(p):
            if eigenvalues[k] > 1e-10:
                cos2_vars[:, k] = var_coords[:, k] ** 2 / np.sum(var_coords ** 2, axis=1)
        rules.append(f"   Cos²:\n{np.round(cos2_vars, 4)}")

        self.results = {
            'type': 'normee',
            'Z': Z, 'R': R,
            'eigenvalues': eigenvalues,
            'eigenvectors': U,
            'components': C,
            'var_coords': var_coords,
            'means': means, 'stds': stds,
            'kaiser_indices': kaiser_indices,
            'cos2_vars': cos2_vars,
        }
        return rules

    # ─────────────────────────────────────────
    #  ACP NON NORMÉE
    # ─────────────────────────────────────────
    def run_acp_non_normee(self, X):
        rules = []
        N, p = X.shape

        rules.append("1) Centrage: X' = X - g")
        g = np.mean(X, axis=0)
        X_centered = X - g
        rules.append(f"   g = {np.round(g, 4)}\n   X':\n{np.round(X_centered, 4)}")

        rules.append("\n2) Matrice variance-covariance V = (1/N) X'^T X'")
        V = (1 / N) * X_centered.T @ X_centered
        rules.append(f"   V =\n{np.round(V, 4)}")

        rules.append("\n3) Métrique M")
        data_type = self.data_type.currentText()
        if data_type == "Variables homogènes":
            M = np.eye(p)
            rules.append("   Variables homogènes: M = I")
        else:
            variances = np.var(X, axis=0, ddof=0)
            variances = np.where(variances < 1e-10, 1.0, variances)
            M = np.diag(1 / variances)
            rules.append(f"   Variables hétérogènes: M = D(1/σ²)\n   M =\n{np.round(M, 4)}")

        rules.append("\n4) Axes principaux de VM")
        VM = V @ M
        rules.append(f"   VM =\n{np.round(VM, 4)}")
        eigenvalues, eigenvectors = np.linalg.eig(VM)
        eigenvalues = np.maximum(eigenvalues, 0)
        idx = eigenvalues.argsort()[::-1]
        eigenvalues = eigenvalues[idx]
        eigenvectors = eigenvectors[:, idx]
        rules.append(f"   λ: {np.round(eigenvalues, 4)}")
        for j in range(p):
            rules.append(f"   U{j+1} = {np.round(eigenvectors[:, j], 4)}")

        total_inertia = np.sum(eigenvalues)
        if total_inertia > 0:
            cum_inertia = np.cumsum(eigenvalues) / total_inertia * 100
        else:
            cum_inertia = np.zeros_like(eigenvalues)

        threshold = self.inertia_threshold.value()
        n_ax = np.where(cum_inertia >= threshold)[0]
        n_axes = (n_ax[0] + 1) if len(n_ax) > 0 else 1

        rules.append(f"   Inertie cumulée: {np.round(cum_inertia, 2)}%")
        rules.append(f"   Axes pour {threshold}%: {n_axes}")

        U = eigenvectors
        rules.append("\n5) Composantes principales: C = X' M U")
        C = X_centered @ M @ U
        rules.append(f"   C:\n{np.round(C, 4)}")

        rules.append("\n6) Cos²(θ_ik) = (C_i^k)² / ||X_i||²_M")
        norms_X = np.sum((X_centered @ M) * X_centered, axis=1)
        norms_X = np.where(norms_X < 1e-10, 1.0, norms_X)
        cos2 = np.zeros((N, p))
        for k in range(p):
            cos2[:, k] = C[:, k] ** 2 / norms_X
        rules.append(f"   Cos²:\n{np.round(cos2, 4)}")

        rules.append("\n7) Corrélation variables-composantes")
        cor_vars = np.zeros((p, p))
        for j in range(p):
            for k in range(p):
                cov = np.cov(X_centered[:, j], C[:, k], ddof=0)[0, 1]
                std_x = np.std(X_centered[:, j], ddof=0)
                std_c = np.std(C[:, k], ddof=0)
                if std_x > 1e-10 and std_c > 1e-10:
                    cor_vars[j, k] = cov / (std_x * std_c)
        rules.append(f"   Corrélations:\n{np.round(cor_vars, 4)}")

        self.results = {
            'type': 'non_normee',
            'X_centered': X_centered, 'V': V, 'M': M,
            'eigenvalues': eigenvalues,
            'eigenvectors': U,
            'components': C,
            'cos2': cos2,
            'cor_vars': cor_vars,
            'means': g,
            'n_axes_retained': n_axes,
        }
        return rules

    # ─────────────────────────────────────────
    #  RUN ACP
    # ─────────────────────────────────────────
    def run_acp(self):
        X = self.parse_data()
        if X is None:
            return

        self._set_status("Calcul en cours…", PALETTE["warning"])
        QApplication.processEvents()

        self.rules_text.clear()
        self.results_text.clear()
        self.eigenvectors_text.clear()

        try:
            if self.acp_type.currentText() == "ACP Normée":
                rules = self.run_acp_normee(X)
                self.display_results_normee()
            else:
                rules = self.run_acp_non_normee(X)
                self.display_results_non_normee()

            self.rules_text.setText('\n'.join(rules))
            self.display_eigenvectors()
            self.plot_results()
            self.display_contributions()

            # Update metric cards
            N, p = X.shape
            self.card_n.set_value(str(N))
            self.card_p.set_value(str(p))
            res = self.results
            if res['type'] == 'normee':
                retained = len(res['kaiser_indices'])
                total = np.sum(res['eigenvalues'])
                if total > 0 and retained > 0:
                    iner = np.sum(res['eigenvalues'][:retained]) / total * 100
                else:
                    iner = 0
            else:
                retained = res['n_axes_retained']
                total = np.sum(res['eigenvalues'])
                iner = np.sum(res['eigenvalues'][:retained]) / total * 100 if total > 0 else 0

            self.card_axes.set_value(str(retained))
            self.card_iner.set_value(f"{iner:.1f}%")

            # Tableau preview
            df_i, df_v, df_e, _ = self.prepare_data_for_tableau()
            if df_i is not None:
                self.preview_tableau_data(df_i, df_v, df_e)

            self._set_status("✓  ACP terminée", PALETTE["success"])
            self.tabs.setCurrentIndex(2)  # switch to plots

        except Exception as e:
            self._set_status("✗  Erreur", PALETTE["danger"])
            QMessageBox.critical(self, "Erreur", f"Une erreur s'est produite:\n{str(e)}")

    # ─────────────────────────────────────────
    #  DISPLAY RESULTS
    # ─────────────────────────────────────────
    def display_results_normee(self):
        res = self.results
        sep = "─" * 52
        text = f"{'ACP NORMÉE':^52}\n{sep}\n\n"

        text += "MATRICE DE CORRÉLATION  R\n"
        text += str(np.round(res['R'], 4)) + "\n\n" + sep + "\n\n"

        text += "VALEURS PROPRES\n"
        total = np.sum(res['eigenvalues'])
        cum = 0
        for i, val in enumerate(res['eigenvalues']):
            pct = val / total * 100 if total > 0 else 0
            cum += pct
            bar = "█" * int(pct / 2)
            text += f"  λ{i+1} = {val:8.4f}   {pct:6.2f}%  cumul {cum:6.2f}%  {bar}\n"
        text += "\n" + sep + "\n\n"

        text += f"AXES RETENUS  (Kaiser λ > {self.kaiser_threshold.value()})\n"
        text += f"  {list(res['kaiser_indices'] + 1)}\n\n" + sep + "\n\n"

        text += "COORDONNÉES DES INDIVIDUS (3 premiers axes)\n"
        n = min(3, res['components'].shape[1])
        text += str(np.round(res['components'][:, :n], 4)) + "\n\n" + sep + "\n\n"

        text += "COORDONNÉES DES VARIABLES (3 premiers axes)\n"
        n = min(3, res['var_coords'].shape[1])
        text += str(np.round(res['var_coords'][:, :n], 4))

        self.results_text.setText(text)

    def display_results_non_normee(self):
        res = self.results
        sep = "─" * 52
        text = f"{'ACP NON NORMÉE':^52}\n{sep}\n\n"

        text += "MATRICE VARIANCE-COVARIANCE  V\n"
        text += str(np.round(res['V'], 4)) + "\n\n"
        text += "MÉTRIQUE  M\n"
        text += str(np.round(res['M'], 4)) + "\n\n" + sep + "\n\n"

        text += "VALEURS PROPRES\n"
        total = np.sum(res['eigenvalues'])
        cum = 0
        for i, val in enumerate(res['eigenvalues']):
            pct = val / total * 100 if total > 0 else 0
            cum += pct
            bar = "█" * int(pct / 2)
            text += f"  λ{i+1} = {val:8.4f}   {pct:6.2f}%  cumul {cum:6.2f}%  {bar}\n"
        text += "\n" + sep + "\n\n"

        text += f"AXES RETENUS  (seuil {self.inertia_threshold.value()}%)\n"
        text += f"  {list(range(1, res['n_axes_retained']+1))}\n\n" + sep + "\n\n"

        text += "COMPOSANTES PRINCIPALES (3 premiers axes)\n"
        n = min(3, res['components'].shape[1])
        text += str(np.round(res['components'][:, :n], 4))

        self.results_text.setText(text)

    def display_eigenvectors(self):
        res = self.results
        sep = "─" * 50
        text = f"{'VECTEURS PROPRES':^50}\n{sep}\n\n"
        eigenvalues = res['eigenvalues']
        eigenvectors = res['eigenvectors']
        p = len(eigenvalues)

        for j in range(p):
            text += f"AXE {j+1}  ·  λ = {eigenvalues[j]:.6f}\n"
            text += sep + "\n"
            for i, val in enumerate(eigenvectors[:, j]):
                text += f"  U{j+1}[{i+1}] = {val:.6f}\n"
            text += f"\n  Norme : {np.linalg.norm(eigenvectors[:, j]):.6f}\n"
            text += f"\n  Contributions des variables :\n"
            contribs = eigenvectors[:, j] ** 2 / np.sum(eigenvectors[:, j] ** 2) * 100
            for i, c in enumerate(contribs):
                bar = "▓" * int(c / 5)
                text += f"  Var{i+1}: {c:6.2f}%  {bar}\n"
            text += "\n" + "═" * 50 + "\n\n"

        self.eigenvectors_text.setText(text)

    # ─────────────────────────────────────────
    #  PLOTS
    # ─────────────────────────────────────────
    def plot_results(self):
        res = self.results
        style = mpl_dark_style()

        # ── Individus ──
        self.figure_individus.clear()
        with plt.rc_context(style):
            ax1 = self.figure_individus.add_subplot(111)
            components = res['components']
            if components.shape[1] >= 2:
                scatter = ax1.scatter(
                    components[:, 0], components[:, 1],
                    c=range(len(components)),
                    cmap='plasma', s=80, alpha=0.9,
                    edgecolors=PALETTE['surface2'], linewidths=1.5, zorder=3)
                for i, (x, y) in enumerate(components[:, :2]):
                    ax1.annotate(
                        f'I{i+1}', (x, y),
                        fontsize=9, color=PALETTE['text'],
                        xytext=(7, 5), textcoords='offset points',
                        bbox=dict(boxstyle='round,pad=0.25',
                                  fc=PALETTE['surface2'], ec=PALETTE['border'], alpha=0.85))

                total = np.sum(res['eigenvalues'])
                p1 = res['eigenvalues'][0] / total * 100 if total > 0 else 0
                p2 = res['eigenvalues'][1] / total * 100 if total > 0 else 0
                ax1.set_xlabel(f'CP1  ({p1:.1f}%)', fontsize=10, labelpad=8)
                ax1.set_ylabel(f'CP2  ({p2:.1f}%)', fontsize=10, labelpad=8)

            ax1.axhline(0, color=PALETTE['border'], lw=0.8)
            ax1.axvline(0, color=PALETTE['border'], lw=0.8)
            ax1.grid(True, alpha=0.2, linestyle='--', color=PALETTE['border'])
            ax1.set_title('Individus', fontsize=12, fontweight='bold', pad=10)
            ax1.tick_params(colors=PALETTE['text_muted'])
            self.figure_individus.patch.set_facecolor(PALETTE['surface'])
            ax1.set_facecolor(PALETTE['bg'])
        self.canvas_individus.draw()

        # ── Variables ──
        self.figure_variables.clear()
        with plt.rc_context(style):
            ax2 = self.figure_variables.add_subplot(111)
            if res['type'] == 'normee':
                var_coords = res['var_coords']
            else:
                var_coords = res['cor_vars']

            if var_coords.shape[1] >= 2:
                circle = plt.Circle((0, 0), 1, fill=False,
                                    color=PALETTE['border'], linestyle='--', lw=1.5, alpha=0.6)
                ax2.add_patch(circle)
                ax2.axhline(0, color=PALETTE['border'], lw=0.7)
                ax2.axvline(0, color=PALETTE['border'], lw=0.7)

                colors_var = ['#6C63FF', '#00D4AA', '#FF6B6B', '#FFB347', '#A78BFA', '#34D399']
                for j in range(var_coords.shape[0]):
                    x, y = var_coords[j, 0], var_coords[j, 1]
                    norm = np.sqrt(x**2 + y**2)
                    if norm > 1:
                        x, y = x / norm, y / norm

                    c = colors_var[j % len(colors_var)]
                    ax2.annotate("",
                                 xy=(x, y), xytext=(0, 0),
                                 arrowprops=dict(arrowstyle="->",
                                                 color=c, lw=2.0))
                    ax2.annotate(f'V{j+1}', (x, y),
                                 fontsize=9, color=c,
                                 xytext=(8, 4), textcoords='offset points',
                                 fontweight='bold')

                ax2.set_xlim(-1.3, 1.3)
                ax2.set_ylim(-1.3, 1.3)
                ax2.set_aspect('equal')

            ax2.grid(True, alpha=0.15, linestyle=':', color=PALETTE['border'])
            ax2.set_title('Cercle des corrélations', fontsize=12, fontweight='bold', pad=10)
            ax2.tick_params(colors=PALETTE['text_muted'])
            self.figure_variables.patch.set_facecolor(PALETTE['surface'])
            ax2.set_facecolor(PALETTE['bg'])
        self.canvas_variables.draw()

        # ── Scree plot ──
        self.figure_scree.clear()
        with plt.rc_context(style):
            ax3 = self.figure_scree.add_subplot(111)
            eigenvalues = res['eigenvalues']
            total = np.sum(eigenvalues)
            pcts = eigenvalues / total * 100 if total > 0 else np.zeros_like(eigenvalues)
            cum_pcts = np.cumsum(pcts)
            axes_labels = [f'Axe {i+1}' for i in range(len(eigenvalues))]

            bars = ax3.bar(axes_labels, pcts, color=PALETTE['accent'],
                           alpha=0.8, width=0.5, zorder=3,
                           edgecolor=PALETTE['surface2'], linewidth=0.5)

            ax3_r = ax3.twinx()
            ax3_r.plot(axes_labels, cum_pcts, 'o-',
                       color=PALETTE['accent2'], lw=2, ms=5, zorder=4)
            ax3_r.set_ylabel("Inertie cumulée %", color=PALETTE['accent2'], fontsize=9)
            ax3_r.tick_params(colors=PALETTE['accent2'])
            ax3_r.set_ylim(0, 110)
            ax3_r.axhline(80, color=PALETTE['warning'], lw=1.2,
                          linestyle='--', alpha=0.7)

            for bar, pct in zip(bars, pcts):
                ax3.text(bar.get_x() + bar.get_width() / 2,
                         bar.get_height() + 0.4,
                         f'{pct:.1f}%', ha='center', va='bottom',
                         fontsize=9, color=PALETTE['text'])

            ax3.set_ylabel("Inertie %", color=PALETTE['text_muted'], fontsize=9)
            ax3.tick_params(colors=PALETTE['text_muted'])
            ax3.set_facecolor(PALETTE['bg'])
            ax3.grid(True, alpha=0.15, axis='y', linestyle='--')
            self.figure_scree.patch.set_facecolor(PALETTE['surface'])
        self.canvas_scree.draw()

    # ─────────────────────────────────────────
    #  CONTRIBUTIONS
    # ─────────────────────────────────────────
    def display_contributions(self):
        res = self.results
        sep = "─" * 50
        text = f"{'CONTRIBUTIONS':^50}\n{sep}\n\n"

        if res['type'] == 'normee':
            components = res['components']
            eigenvalues = res['eigenvalues']
            var_coords = res['var_coords']
            N = components.shape[0]

            text += "CONTRIBUTIONS DES INDIVIDUS AUX AXES\n" + sep + "\n"
            for k in range(min(3, len(eigenvalues))):
                pct_ax = eigenvalues[k] / np.sum(eigenvalues) * 100 if np.sum(eigenvalues) > 0 else 0
                text += f"\nAxe {k+1}  (λ={eigenvalues[k]:.4f}, {pct_ax:.1f}%)\n"
                if eigenvalues[k] > 1e-10:
                    contrib = components[:, k] ** 2 / (N * eigenvalues[k])
                    for i, c in sorted(enumerate(contrib), key=lambda x: -x[1]):
                        if c > 0.05:
                            bar = "▓" * int(c * 100 / 5)
                            text += f"  Ind{i+1}: {c:.4f}  ({c*100:.1f}%)  {bar}\n"

            text += f"\n\n{'CONTRIBUTIONS DES VARIABLES AUX AXES':^50}\n{sep}\n"
            for k in range(min(3, len(eigenvalues))):
                text += f"\nAxe {k+1}\n"
                if eigenvalues[k] > 1e-10:
                    cv = var_coords[:, k] ** 2 / (var_coords.shape[0] * eigenvalues[k])
                    for j, c in enumerate(cv):
                        bar = "▓" * int(c * 100 / 5)
                        text += f"  Var{j+1}: {c:.4f}  ({c*100:.1f}%)  {bar}\n"
        else:
            components = res['components']
            cos2 = res['cos2']
            text += "QUALITÉ DE REPRÉSENTATION (cos²)\n" + sep + "\n"
            for k in range(min(3, components.shape[1])):
                text += f"\nAxe {k+1}\n"
                for i in np.argsort(cos2[:, k])[::-1][:5]:
                    if cos2[i, k] > 0.1:
                        bar = "▓" * int(cos2[i, k] * 20)
                        text += f"  Ind{i+1}: {cos2[i, k]:.4f}  {bar}\n"

        self.contrib_text.setText(text)

    # ─────────────────────────────────────────
    #  TABLEAU EXPORT
    # ─────────────────────────────────────────
    def prepare_data_for_tableau(self):
        res = self.results
        if not res:
            return None, None, None, None

        components = res['components']
        n_indiv, n_axes = components.shape

        individus_data = []
        for i in range(n_indiv):
            for j in range(n_axes):
                individus_data.append({
                    'Individu': f'Ind{i+1}', 'Axe': f'Axe{j+1}',
                    'Coordonnee': components[i, j],
                    'Inertie_axe': res['eigenvalues'][j] if j < len(res['eigenvalues']) else 0,
                    'Type': 'Individu'})
        df_individus = pd.DataFrame(individus_data)

        var_data = res['var_coords'] if res['type'] == 'normee' else res['cor_vars']
        variables_data = []
        for i in range(var_data.shape[0]):
            for j in range(min(n_axes, var_data.shape[1])):
                variables_data.append({
                    'Variable': f'Var{i+1}', 'Axe': f'Axe{j+1}',
                    'Coordonnee': var_data[i, j],
                    'Cos2': res.get('cos2_vars', np.zeros((var_data.shape[0], n_axes)))[i, j] if res['type'] == 'normee' else 0,
                    'Type': 'Variable'})
        df_variables = pd.DataFrame(variables_data)

        total = np.sum(res['eigenvalues'])
        eigen_data = []
        for j, val in enumerate(res['eigenvalues']):
            pct = val / total * 100 if total > 0 else 0
            eigen_data.append({
                'Axe': f'Axe{j+1}', 'Valeur_propre': val,
                'Inertie_pourcent': pct,
                'Inertie_cumulee': np.sum(res['eigenvalues'][:j+1]) / total * 100 if total > 0 else 0})
        df_eigen = pd.DataFrame(eigen_data)

        contrib_data = []
        if res['type'] == 'normee':
            for j in range(min(3, n_axes)):
                if res['eigenvalues'][j] > 1e-10:
                    ci = components[:, j] ** 2 / (n_indiv * res['eigenvalues'][j])
                    for i, c in enumerate(ci):
                        if c > 0.05:
                            contrib_data.append({
                                'Element': f'Ind{i+1}', 'Axe': f'Axe{j+1}',
                                'Contribution': c, 'Type': 'Individu'})
        df_contrib = pd.DataFrame(contrib_data)
        return df_individus, df_variables, df_eigen, df_contrib

    def export_for_tableau(self):
        if not self.results:
            QMessageBox.warning(self, "Attention", "Veuillez d'abord exécuter l'ACP")
            return
        try:
            df_i, df_v, df_e, df_c = self.prepare_data_for_tableau()
            if df_i is None:
                return
            folder = QFileDialog.getExistingDirectory(self, "Choisir le dossier d'export")
            if not folder:
                return

            df_i.to_csv(os.path.join(folder, "tableau_individus.csv"),    index=False, encoding='utf-8-sig')
            df_v.to_csv(os.path.join(folder, "tableau_variables.csv"),    index=False, encoding='utf-8-sig')
            df_e.to_csv(os.path.join(folder, "tableau_valeurs_propres.csv"), index=False, encoding='utf-8-sig')
            if not df_c.empty:
                df_c.to_csv(os.path.join(folder, "tableau_contributions.csv"), index=False, encoding='utf-8-sig')

            self.tableau_path_label.setText(f"📁  {folder}")
            self.preview_tableau_data(df_i, df_v, df_e)
            QMessageBox.information(self, "Succès",
                f"Données exportées avec succès\n→ {folder}")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", str(e))

    def preview_tableau_data(self, df_i, df_v, df_e):
        preview = []
        for _, row in df_i.head(10).iterrows():
            preview.append([row['Individu'], row['Axe'], f"{row['Coordonnee']:.3f}", "Individu"])
        for _, row in df_v.head(10).iterrows():
            preview.append([row['Variable'], row['Axe'], f"{row['Coordonnee']:.3f}", "Variable"])

        self.tableau_preview.setRowCount(len(preview))
        self.tableau_preview.setColumnCount(4)
        self.tableau_preview.setHorizontalHeaderLabels(['Élément', 'Axe', 'Coordonnée', 'Type'])

        accent_colors = {
            "Individu": QColor(108, 99, 255, 30),
            "Variable": QColor(0, 212, 170, 30),
        }
        for i, row in enumerate(preview):
            for j, val in enumerate(row):
                item = QTableWidgetItem(str(val))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                if j == 3:
                    item.setForeground(QColor(PALETTE["accent"] if val == "Individu" else PALETTE["accent2"]))
                self.tableau_preview.setItem(i, j, item)

        self.tableau_preview.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)


# ─────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────
def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(STYLESHEET)

    # High DPI
    app.setAttribute(Qt.AA_UseHighDpiPixmaps)

    window = ACPApp()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()