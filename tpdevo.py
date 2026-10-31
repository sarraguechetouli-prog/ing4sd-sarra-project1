import os
import io
import re
import json
import unicodedata
from datetime import datetime

import pandas as pd
from flask import (Flask, render_template_string, request, redirect,
                   url_for, send_file, flash)

app = Flask(__name__)
app.secret_key = "change-moi"

FICHIER = "regles_expert.xlsx"          # règles publiées
BROUILLON = "brouillon.json"            # règles en attente de validation
IMPORTS = ["import_en_cours.xlsx", "import_en_cours.xlsm", "import_en_cours.csv"]
COLONNES = ["Categorie", "Condition", "Type", "Operateur", "Valeur",
            "Unite", "Obligatoire", "Question", "Details"]
OPS_NOMBRE = [">=", "<=", ">", "<", "==", "entre"]
SYMBOLES = {">=": "≥", "<=": "≤", ">": ">", "<": "<", "==": "="}
NUM = r"\d+(?:[.,]\d+)?"
NON_OBLIG = {"non", "n", "0", "false", "facultatif", "recommande", "recommandee",
             "conseille", "souhaitable", "optionnel", "secondaire", "faible"}
MOTS_VIDES = {"et", "a", "ou", "au", "de", "du", "des", "le", "la", "les",
              "par", "pour", "si", "sur", "en", "dans"}

ROLES = [("enonce", "Énoncé de la condition"), ("categorie", "Catégorie"),
         ("valeur", "Valeur / seuil"), ("min", "Valeur minimale"),
         ("max", "Valeur maximale"), ("operateur", "Opérateur (≥, ≤…)"),
         ("unite", "Unité"), ("obligatoire", "Obligatoire ?"),
         ("question", "Question personnalisée"), ("details", "Info à afficher"),
         ("ignorer", "— Ignorer —")]

SYN = {
    "Min": ["min", "mini", "minimum", "minimal", "valeur min", "valeur minimale"],
    "Max": ["max", "maxi", "maximum", "maximal", "valeur max", "valeur maximale"],
    "Unite": ["unite", "unites", "unit"],
    "Categorie": ["categorie", "groupe", "theme", "section", "domaine", "chapitre",
                  "famille", "rubrique", "type de regle"],
    "Obligatoire": ["obligatoire", "caractere", "priorite", "importance", "obligation"],
    "Operateur": ["operateur", "comparaison", "relation", "signe"],
    "Question": ["question", "question posee"],
    "Type": ["type", "type de donnee", "nature"],
    "Valeur": ["valeur", "seuil", "limite", "requis", "exige", "valeur limite",
               "valeur requise", "norme", "admissible", "valeur admissible"],
    "Condition": ["condition", "conditions", "regle", "regles", "critere", "criteres",
                  "exigence", "exigences", "designation", "description", "parametre",
                  "element", "intitule", "libelle", "caracteristique", "verification",
                  "disposition", "prescription", "enonce"],
}
ORDRE_ROLES = ["Min", "Max", "Unite", "Categorie", "Obligatoire", "Operateur",
               "Question", "Type", "Valeur", "Condition"]
CORRESP = {"Condition": "enonce", "Categorie": "categorie", "Valeur": "valeur",
           "Min": "min", "Max": "max", "Unite": "unite", "Obligatoire": "obligatoire",
           "Question": "question", "Operateur": "operateur", "Type": "ignorer"}
MAX_COLS = 20


# ---------------------------------------------------------------
#  Outils
# ---------------------------------------------------------------
def normaliser(txt):
    txt = unicodedata.normalize("NFD", str(txt))
    txt = "".join(c for c in txt if unicodedata.category(c) != "Mn")
    return txt.strip().lower()


def cellule(x):
    if x is None:
        return ""
    s = str(x).strip()
    if s.lower() in ("nan", "none", "nat"):
        return ""
    if re.fullmatch(r"-?\d+\.0+", s):
        s = s.split(".")[0]
    return s


def nb(s):
    return float(str(s).strip().replace(",", "."))


def fmt(x):
    return f"{x:g}"


def bornes(valeur):
    parties = re.split(r"\s*[;-]\s*", str(valeur).strip())
    if len(parties) != 2:
        raise ValueError("format attendu : min-max (ex : 10-20)")
    a, b = nb(parties[0]), nb(parties[1])
    return min(a, b), max(a, b)


def extraire_nombre(s):
    m = re.search(NUM, s or "")
    if not m:
        return None, ""
    reste = s[m.end():].strip()
    u = re.match(r"([A-Za-zµ°%²³]+(?:/[A-Za-z²³]+)?)", reste)
    unite = u.group(1) if u and normaliser(u.group(1)) not in MOTS_VIDES else ""
    return m.group(0).replace(",", "."), unite


def detect_op(t):
    """Opérateur explicite dans un texte normalisé, sinon None."""
    if not t:
        return None
    if re.search(r">=|≥|=>|superieur\w*\s+ou\s+egal|au moins|pas moins|minimum|minimal|\bmini?\b", t):
        return ">="
    if re.search(r"<=|≤|=<|inferieur\w*\s+ou\s+egal|au plus|pas plus|ne (?:doit )?pas depasser|"
                 r"ne depasse pas|maximum|maximal|\bmaxi?\b|limite a|plafonne", t):
        return "<="
    if re.search(r">|superieur|plus de", t):
        return ">"
    if re.search(r"<|inferieur|moins de", t):
        return "<"
    if re.search(r"==|\begal|identique|exactement|(?<![<>=!])=(?!=)", t):
        return "=="
    return None


def normaliser_op(s):
    alias = {">=": ">=", "≥": ">=", "=>": ">=", "<=": "<=", "≤": "<=", "=<": "<=",
             ">": ">", "<": "<", "==": "==", "=": "==", "egal": "==", "entre": "entre"}
    return alias.get(normaliser(s), "")


# ---------------------------------------------------------------
#  Lecture du fichier importé
# ---------------------------------------------------------------
def chemin_import():
    for p in IMPORTS:
        if os.path.exists(p):
            return p
    return None


def supprimer_imports():
    for p in IMPORTS:
        if os.path.exists(p):
            os.remove(p)


def lire_feuilles(chemin):
    if chemin.endswith(".csv"):
        for enc in ("utf-8-sig", "latin-1"):
            try:
                df = pd.read_csv(chemin, sep=None, engine="python", header=None,
                                 dtype=str, encoding=enc)
                return {"Règles": df}
            except Exception:
                continue
        raise ValueError("Fichier CSV illisible.")
    return pd.read_excel(chemin, sheet_name=None, header=None, dtype=str)


def lignes_feuille(df):
    lignes = [[cellule(c) for c in row] for row in df.fillna("").values.tolist()]
    return [l for l in lignes if any(l)]


def role_entete(h):
    h = normaliser(h)
    if not h:
        return None
    for role in ORDRE_ROLES:
        if h in SYN[role]:
            return role
    for role in ORDRE_ROLES:
        if role == "Type":
            continue
        for s in SYN[role]:
            if len(s) >= 3 and re.search(r"\b" + re.escape(s) + r"\b", h):
                return role
    return None


def detecter(lignes, entete=None):
    """Retourne (ligne d'en-tête [0 = aucune], rôles proposés par colonne)."""
    nb_cols = min(max(len(l) for l in lignes), MAX_COLS)

    def roles_ligne(row):
        trouves = {}
        for c, h in enumerate(row[:nb_cols]):
            r = role_entete(h)
            if r and r not in trouves:
                trouves[r] = c
        return trouves

    if entete is None:
        meilleur, idx = 0, None
        for k, row in enumerate(lignes[:15]):
            n = len(roles_ligne(row))
            if n > meilleur:
                meilleur, idx = n, k
        if idx is not None:
            entete = idx + 1
        else:
            p = lignes[0]
            entete = 1 if (sum(1 for c in p if c) >= 2 and all(len(c) <= 40 for c in p)) else 0

    roles = roles_ligne(lignes[entete - 1]) if 0 < entete <= len(lignes) else {}
    pre = ["ignorer"] * nb_cols
    for r, c in roles.items():
        pre[c] = CORRESP[r]
    reconnues = set(roles.values())

    if "enonce" not in pre:
        data = lignes[entete:]
        meilleur, col = -1, 0
        for c in range(nb_cols):
            if c in reconnues:
                continue
            score = sum(len(l[c]) for l in data
                        if c < len(l) and l[c] and not re.fullmatch(r"[\d.,\s]+", l[c]))
            if score > meilleur:
                meilleur, col = score, c
        pre[col] = "enonce"

    if 0 < entete <= len(lignes):
        for c in range(nb_cols):
            if pre[c] == "ignorer" and c not in reconnues and c < len(lignes[entete - 1]) \
                    and lignes[entete - 1][c]:
                pre[c] = "details"
    return entete, pre


# ---------------------------------------------------------------
#  Construction des règles à partir du mapping choisi par l'expert
# ---------------------------------------------------------------
def comprendre_seuil(enonce, val, vmin, vmax, unite, op_col, entete_val):
    """Retourne (op, valeur, unité, sûr) uniquement si un seuil EXPLICITE existe."""
    if vmin or vmax:
        a, ua = extraire_nombre(vmin)
        b, ub = extraire_nombre(vmax)
        u = unite or ua or ub
        if a and b:
            return "entre", f"{a}-{b}", u, True
        if a:
            return ">=", a, u, True
        if b:
            return "<=", b, u, True
    op_col = normaliser_op(op_col)
    for source, est_col in ((val, True), (enonce, False)):
        if not source:
            continue
        tn = normaliser(source)
        nombres = re.findall(NUM, tn)
        if not nombres:
            continue
        m = re.search(rf"({NUM})\s*(?:et|a|-|–|;)\s*({NUM})", tn)
        if m and ("entre" in tn or re.fullmatch(rf"\s*{NUM}\s*(?:-|–|;|a)\s*{NUM}\s*\S*\s*", tn)):
            a, b = m.group(1).replace(",", "."), m.group(2).replace(",", ".")
            u = unite
            if not u:
                mu = re.match(r"\s*([A-Za-zµ°%²³]+)", source[m.end():])
                if mu and normaliser(mu.group(1)) not in MOTS_VIDES:
                    u = mu.group(1)
            return "entre", f"{a}-{b}", u, est_col
        if len(nombres) > 1:
            continue                      # trop ambigu : question de conformité
        op = op_col or detect_op(tn)
        sur = est_col and bool(op)
        if not op and est_col:
            op = detect_op(normaliser(enonce)) or detect_op(entete_val)
        if not op:
            continue
        n, u = extraire_nombre(source)
        return op, n, unite or u, sur
    return None


def exigence_txt(r):
    if r["Type"] != "nombre":
        return ""
    u = f" {r['Unite']}" if r["Unite"] else ""
    if r["Operateur"] == "entre":
        a, b = bornes(r["Valeur"])
        return f"entre {fmt(a)} et {fmt(b)}{u}"
    return f"{SYMBOLES[r['Operateur']]} {fmt(nb(r['Valeur']))}{u}"


def valider_regle(r):
    r = {c: str(r.get(c, "") or "").strip() for c in COLONNES}
    if not r["Condition"]:
        return r, "énoncé vide"
    if r["Type"] not in ("nombre", "oui/non"):
        return r, "type invalide"
    if r["Type"] == "nombre":
        if r["Operateur"] not in OPS_NOMBRE:
            return r, "opérateur invalide"
        try:
            if r["Operateur"] == "entre":
                a, b = bornes(r["Valeur"])
                r["Valeur"] = f"{fmt(a)}-{fmt(b)}"
            else:
                r["Valeur"] = fmt(nb(r["Valeur"]))
        except ValueError as e:
            return r, f"seuil invalide ({e})"
    else:
        r["Operateur"], r["Valeur"], r["Unite"] = "==", "oui", ""
    r["Obligatoire"] = "non" if normaliser(r["Obligatoire"]) in NON_OBLIG else "oui"
    if not r["Categorie"]:
        r["Categorie"] = "Général"
    return r, None


def construire_regles(feuilles, nom, entete, mapping, toutes):
    brutes = {str(n): lignes_feuille(d) for n, d in feuilles.items()}
    brutes = {n: l for n, l in brutes.items() if l}
    cibles = list(brutes.items()) if toutes else [(nom, brutes.get(nom, []))]
    plusieurs = len(cibles) > 1

    cols = {}
    for c, role in mapping.items():
        cols.setdefault(role, []).append(c)

    regles, alertes = [], []
    for nom_f, lignes in cibles:
        ent = lignes[entete - 1] if 0 < entete <= len(lignes) else []

        def val(row, role):
            c = cols.get(role)
            return row[c[0]] if c and c[0] < len(row) else ""

        entete_val = ""
        if "valeur" in cols and cols["valeur"][0] < len(ent):
            entete_val = normaliser(ent[cols["valeur"][0]])
        cat_cour = nom_f if plusieurs else ""

        for row in lignes[entete:]:
            enonce = val(row, "enonce")
            if not enonce:
                continue
            if val(row, "categorie"):
                cat_cour = val(row, "categorie")
            al = []
            details = []
            for c in cols.get("details", []):
                if c < len(row) and row[c]:
                    titre = ent[c] if c < len(ent) and ent[c] else "Info"
                    details.append(f"{titre} : {row[c]}")

            unite = val(row, "unite")
            seuil = comprendre_seuil(enonce, val(row, "valeur"), val(row, "min"),
                                     val(row, "max"), unite, val(row, "operateur"), entete_val)
            r = {"Categorie": cat_cour or "Général", "Condition": enonce,
                 "Type": "oui/non", "Operateur": "==", "Valeur": "oui", "Unite": "",
                 "Obligatoire": val(row, "obligatoire"), "Question": val(row, "question"),
                 "Details": ""}
            if seuil:
                op, v, u, sur = seuil
                r.update(Type="nombre", Operateur=op, Valeur=v, Unite=u)
                if not sur:
                    al.append(f"mesure déduite du texte : {exigence_txt(r)}")
            else:
                ref = " ".join(x for x in (val(row, "valeur"), val(row, "min"),
                                           val(row, "max"), unite) if x)
                if ref:
                    details.append(f"Valeur de référence : {ref}")
                    al.append("valeur non convertie en seuil → question de conformité")
            r["Details"] = " · ".join(details)

            rv, err = valider_regle(r)
            if err:
                al.append("à corriger : " + err)
            regles.append(r if err else rv)
            alertes.append(al)
    return regles, alertes


# ---------------------------------------------------------------
#  Stockage des règles publiées
# ---------------------------------------------------------------
def charger_regles():
    if not os.path.exists(FICHIER):
        return None
    df = pd.read_excel(FICHIER, dtype=str).fillna("")
    for c in COLONNES:
        if c not in df.columns:
            df[c] = ""
    return df[COLONNES].to_dict("records") or None


def sauver_regles(regles):
    pd.DataFrame(regles, columns=COLONNES).to_excel(FICHIER, index=False)


def question_txt(r):
    if r["Question"]:
        return r["Question"]
    if r["Type"] == "nombre":
        return "Quelle est la valeur de ce critère pour votre bâtiment ?"
    return "Votre bâtiment respecte-t-il cette condition ?"


def preparer(regles):
    out = []
    for i, r in enumerate(regles):
        r = dict(r)
        r["i"] = i
        r["exigence"] = exigence_txt(r)
        r["question"] = question_txt(r)
        out.append(r)
    return out


def grouper(regles):
    groupes = {}
    for r in regles:
        groupes.setdefault(r["Categorie"], []).append(r)
    return groupes


def evaluer(r, saisie):
    try:
        if r["Type"] == "nombre":
            x = nb(saisie)
            if r["Operateur"] == "entre":
                a, b = bornes(r["Valeur"])
                return a <= x <= b
            s = nb(r["Valeur"])
            return {">=": x >= s, "<=": x <= s, ">": x > s,
                    "<": x < s, "==": x == s}[r["Operateur"]]
        return normaliser(saisie) == "oui"
    except Exception:
        return False


# ---------------------------------------------------------------
#  Gabarits HTML
# ---------------------------------------------------------------
LAYOUT = """
<!doctype html>
<html lang="fr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SismoCheck – Vérification antisismique</title>
<style>
:root{--primary:#2563eb;--primary-d:#1e3a8a;--bg:#eef2f7;--card:#fff;--txt:#0f172a;
 --muted:#64748b;--line:#e2e8f0;--ok:#16a34a;--ok-bg:#dcfce7;--ko:#dc2626;--ko-bg:#fee2e2;
 --warn:#d97706;--warn-bg:#fef3c7;--shadow:0 4px 14px rgba(15,23,42,.07)}
@media (prefers-color-scheme:dark){:root{--bg:#0b1220;--card:#111a2e;--txt:#e5eaf5;
 --muted:#93a1bd;--line:#223052;--ok-bg:#12351f;--ko-bg:#3b1414;--warn-bg:#3b2a0b;
 --shadow:0 4px 14px rgba(0,0,0,.4)}}
*{box-sizing:border-box}
[hidden]{display:none!important}
body{margin:0;font-family:system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;
 background:var(--bg);color:var(--txt);line-height:1.5}
header{background:linear-gradient(120deg,#0f2557,var(--primary));color:#fff;
 box-shadow:0 4px 18px rgba(0,0,0,.25)}
.bar{max-width:1150px;margin:auto;padding:14px 22px;display:flex;align-items:center;
 justify-content:space-between;flex-wrap:wrap;gap:10px}
.logo{font-size:21px;font-weight:800;text-decoration:none;color:#fff;display:flex;
 align-items:center;gap:10px}
.logo i{font-style:normal;background:rgba(255,255,255,.18);border-radius:10px;padding:4px 9px}
.logo span{opacity:.7;font-weight:500;font-size:15px}
nav a{color:#dbeafe;text-decoration:none;padding:9px 16px;border-radius:99px;font-weight:600;
 font-size:14px;margin-left:4px;transition:.15s}
nav a:hover{background:rgba(255,255,255,.16)}
nav a.actif{background:#fff;color:var(--primary-d)}
main{max-width:1150px;margin:26px auto;padding:0 22px 50px;animation:in .3s ease}
@keyframes in{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}
.card{background:var(--card);border-radius:16px;padding:26px;margin-bottom:20px;
 box-shadow:var(--shadow);border:1px solid var(--line)}
h1{margin:0 0 6px;font-size:28px}h2{margin:0 0 14px;font-size:19px}
.sub{color:var(--muted);margin:0 0 18px}
.btn{display:inline-block;background:var(--primary);color:#fff;border:0;padding:11px 20px;
 border-radius:12px;font-weight:700;font-size:15px;cursor:pointer;text-decoration:none;
 transition:.15s;box-shadow:0 2px 6px rgba(37,99,235,.3)}
.btn:hover{background:var(--primary-d);transform:translateY(-1px)}
.btn.sec{background:var(--line);color:var(--txt);box-shadow:none}
.btn.sec:hover{filter:brightness(.93)}
.btn.danger{background:var(--ko-bg);color:var(--ko);box-shadow:none}
.btn.big{padding:14px 28px;font-size:17px}
.btn.sm{padding:6px 12px;font-size:13px}
.flash{padding:13px 18px;border-radius:12px;margin-bottom:16px;font-weight:600;white-space:pre-line}
.flash.ok{background:var(--ok-bg);color:var(--ok)}.flash.err{background:var(--ko-bg);color:var(--ko)}
.grid{display:grid;gap:16px;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));margin-bottom:20px}
.stat{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:18px;
 text-align:center;box-shadow:var(--shadow)}
.stat b{display:block;font-size:32px;color:var(--primary)}
.stat span{color:var(--muted);font-size:13px}
.hero{text-align:center;padding:56px 24px;background:linear-gradient(160deg,var(--card),var(--bg))}
.hero h1{font-size:36px}
.hero .actions{margin-top:24px;display:flex;gap:12px;justify-content:center;flex-wrap:wrap}
.step{display:flex;gap:14px;align-items:flex-start}
.step i{flex:none;width:38px;height:38px;border-radius:50%;background:var(--primary);color:#fff;
 display:flex;align-items:center;justify-content:center;font-style:normal;font-weight:800}
.steps{display:flex;gap:10px;margin-bottom:20px;flex-wrap:wrap}
.steps div{flex:1;min-width:150px;padding:12px 16px;border-radius:12px;background:var(--card);
 border:1px solid var(--line);color:var(--muted);font-weight:700;font-size:14px}
.steps div.on{background:var(--primary);color:#fff;border-color:var(--primary)}
.drop{border:2px dashed var(--primary);border-radius:16px;padding:32px;text-align:center;
 background:var(--bg)}
.drop input{margin:14px 0}
table{border-collapse:collapse;width:100%;font-size:14px}
th{background:var(--bg);text-align:left;padding:10px;font-size:12px;text-transform:uppercase;
 letter-spacing:.5px;color:var(--muted);white-space:nowrap;vertical-align:bottom}
td{padding:8px 10px;border-top:1px solid var(--line);vertical-align:top}
.tw{overflow-x:auto}
.badge{display:inline-block;padding:3px 10px;border-radius:99px;font-size:12px;font-weight:700;margin:1px}
.b-ok{background:var(--ok-bg);color:var(--ok)}.b-ko{background:var(--ko-bg);color:var(--ko)}
.b-oblig{background:#dbeafe;color:#1e40af}.b-fac{background:var(--line);color:var(--muted)}
.b-warn{background:var(--warn-bg);color:var(--warn)}
code{background:var(--bg);padding:2px 6px;border-radius:5px;font-size:13px}
input[type=number],input[type=text],select,textarea{width:100%;padding:10px 12px;
 border:1.5px solid var(--line);border-radius:12px;font-size:15px;background:var(--card);
 color:var(--txt);font-family:inherit}
input[type=text]{max-width:380px}
input:focus,select:focus,textarea:focus{outline:0;border-color:var(--primary);
 box-shadow:0 0 0 3px rgba(37,99,235,.25)}
.edit input,.edit select,.edit textarea{padding:7px 9px;font-size:13px;border-radius:8px;min-width:90px}
.edit td:nth-child(2) textarea{min-width:300px}
.edit td:nth-child(8) textarea,.edit td:nth-child(9) textarea{min-width:180px}
.prev th select{min-width:150px;font-size:13px;padding:7px;font-weight:700}
.prev td{max-width:260px;font-size:13px}
.prev tr.ent td{background:var(--warn-bg);font-weight:700}
.opts{display:flex;gap:16px;flex-wrap:wrap;align-items:center;margin-bottom:16px}
.opts label{font-weight:600;display:flex;gap:8px;align-items:center}
.opts select{width:auto}
.progress{position:sticky;top:0;z-index:5;background:var(--card);padding:12px 20px;
 border-radius:0 0 16px 16px;box-shadow:var(--shadow);margin-bottom:16px;
 border:1px solid var(--line);border-top:0}
.track{height:10px;background:var(--line);border-radius:99px;overflow:hidden;margin-top:6px}
.fill{height:100%;width:0;background:linear-gradient(90deg,#3b82f6,var(--ok));transition:.3s}
.chips{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:16px}
.chip{border:1.5px solid var(--line);background:var(--card);color:var(--muted);border-radius:99px;
 padding:7px 15px;font-weight:700;font-size:13px;cursor:pointer;transition:.15s}
.chip.on{background:var(--primary);border-color:var(--primary);color:#fff}
.q{display:flex;gap:14px;padding:20px 0;border-top:1px solid var(--line)}
.q.first{border-top:0}
.q .num{flex:none;width:32px;height:32px;border-radius:50%;background:var(--primary);
 color:#fff;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:14px}
.q .corps{flex:1;min-width:0}
.enonce{background:var(--bg);border-left:4px solid var(--primary);padding:12px 16px;
 border-radius:10px;font-weight:600;margin-bottom:8px}
.details{font-size:13px;color:var(--muted);margin-bottom:8px}
.qtxt{font-weight:700;margin:10px 0 8px}
.hint{font-size:13px;color:var(--muted);margin-bottom:8px}.hint b{color:var(--primary)}
.unit{display:flex;align-items:center;gap:10px;max-width:340px}
.unit span{color:var(--muted);font-weight:700}
.pills{display:flex;gap:10px;flex-wrap:wrap}
.pills input{display:none}
.pills label{padding:11px 26px;border:1.5px solid var(--line);border-radius:12px;cursor:pointer;
 font-weight:700;background:var(--card);transition:.15s}
.pills label:hover{border-color:var(--primary)}
.pills input:checked+label{background:var(--primary);border-color:var(--primary);color:#fff}
.nav{display:flex;justify-content:space-between;gap:10px;margin-top:6px}
.verdict{display:flex;align-items:center;gap:30px;flex-wrap:wrap;border-radius:18px;padding:28px}
.verdict.ok{background:var(--ok-bg)}.verdict.ko{background:var(--ko-bg)}
.ring{--p:0;--c:var(--ok);width:140px;height:140px;border-radius:50%;flex:none;
 background:conic-gradient(var(--c) calc(var(--p)*1%),rgba(148,163,184,.35) 0);
 display:flex;align-items:center;justify-content:center}
.ring div{width:104px;height:104px;border-radius:50%;background:var(--card);display:flex;
 align-items:center;justify-content:center;font-size:28px;font-weight:800}
.verdict h1{margin:0;font-size:27px}
.todo{border-left:5px solid var(--ko);background:var(--ko-bg);padding:12px 16px;
 border-radius:10px;margin-bottom:10px}
.todo small{color:var(--muted)}
.cbar{display:flex;align-items:center;gap:12px;margin:10px 0}
.cbar span.n{width:180px;font-weight:700}
.cbar .track{flex:1;margin:0}.cbar .fill{width:var(--w)}
.vide{text-align:center;padding:44px 10px}.vide .ico{font-size:54px}
footer{text-align:center;color:var(--muted);font-size:13px;padding-bottom:30px}
@media print{header,.noprint,footer{display:none}body{background:#fff;color:#000}.card{box-shadow:none}}
</style></head><body>
<header><div class="bar">
 <a class="logo" href="{{ url_for('accueil') }}"><i>🏢</i> SismoCheck <span>| antisismique</span></a>
 <nav>
  <a href="{{ url_for('accueil') }}" class="{{ 'actif' if actif=='accueil' }}">Accueil</a>
  <a href="{{ url_for('expert') }}" class="{{ 'actif' if actif=='expert' }}">Espace expert</a>
  <a href="{{ url_for('verifier') }}" class="{{ 'actif' if actif=='verifier' }}">Vérifier un bâtiment</a>
 </nav></div></header>
<main>
{% for cat, m in get_flashed_messages(with_categories=true) %}
 <div class="flash {{ cat }}">{{ m }}</div>{% endfor %}
{{ body|safe }}
</main>
<footer>Outil de pré-vérification – ne remplace pas une étude réglementaire complète.</footer>
</body></html>
"""

ACCUEIL = """
<div class="card hero">
 <div style="font-size:60px">🏗️</div>
 <h1>Ce bâtiment est-il antisismique ?</h1>
 <p class="sub">L'expert importe son tableau Excel de conditions.<br>
 Le questionnaire reprend ses énoncés <b>mot pour mot</b>.</p>
 <div class="actions">
  <a class="btn big" href="{{ url_for('verifier') }}">Vérifier un bâtiment</a>
  <a class="btn big sec" href="{{ url_for('expert') }}">Espace expert</a>
 </div>
</div>
<div class="grid">
 <div class="card step"><i>1</i><div><h2>Import du tableau</h2>
  <p class="sub">L'expert charge son fichier Excel.</p></div></div>
 <div class="card step"><i>2</i><div><h2>Choix des colonnes</h2>
  <p class="sub">Il indique où se trouvent les énoncés des conditions.</p></div></div>
 <div class="card step"><i>3</i><div><h2>Questionnaire & verdict</h2>
  <p class="sub">L'utilisateur répond, le bâtiment est évalué.</p></div></div>
</div>
"""

EXPERT = """
<div class="steps"><div class="on">① Importer</div><div>② Indiquer les colonnes</div>
 <div class="{{ 'on' if existe }}">③ Vérifier et publier</div></div>
{% if attente %}
<div class="flash" style="background:var(--warn-bg);color:var(--warn)">
 Des règles sont en attente de validation.
 <a class="btn sm" href="{{ url_for('editer', brouillon=1) }}">Continuer →</a></div>
{% endif %}
<div class="card">
 <h1>Espace expert</h1>
 <p class="sub">Importez votre tableau de conditions antisismiques (Excel ou CSV).
 Chaque ligne devient une question de conformité posée à l'utilisateur.</p>
 <form method="post" action="{{ url_for('importer') }}" enctype="multipart/form-data">
  <div class="drop">
   <div style="font-size:42px">📄</div>
   <b>{{ 'Remplacer le tableau de règles' if existe else 'Choisir le fichier Excel de l\\'expert' }}</b><br>
   <input type="file" name="fichier" accept=".xlsx,.xlsm,.csv" required><br>
   <button class="btn">Importer</button>
   <a class="btn sec" href="{{ url_for('editer') }}">✏ Saisir à la main</a>
   <a class="btn sec" href="{{ url_for('modele') }}">⬇ Modèle Excel</a>
  </div>
 </form>
</div>
{% if existe %}
<div class="grid">
 <div class="stat"><b>{{ regles|length }}</b><span>conditions</span></div>
 <div class="stat"><b>{{ nb_mesure }}</b><span>avec mesure chiffrée</span></div>
 <div class="stat"><b>{{ nb_cat }}</b><span>catégories</span></div>
 <div class="stat"><b style="font-size:17px;line-height:48px">{{ maj }}</b><span>dernière mise à jour</span></div>
</div>
<div class="card">
 <h2>Conditions en vigueur</h2>
 <div class="tw"><table>
 <tr><th>#</th><th>Catégorie</th><th>Énoncé</th><th>Mode</th><th>Exigence</th><th>Niveau</th></tr>
 {% for r in regles %}
 <tr><td>{{ r.i+1 }}</td><td>{{ r.Categorie }}</td><td>{{ r.Condition }}</td>
 <td>{{ 'Mesure' if r.Type=='nombre' else 'Oui/Non' }}</td><td>{{ r.exigence or '—' }}</td>
 <td><span class="badge {{ 'b-oblig' if r.Obligatoire=='oui' else 'b-fac' }}">
 {{ 'Obligatoire' if r.Obligatoire=='oui' else 'Facultatif' }}</span></td></tr>
 {% endfor %}
 </table></div>
 <div style="margin-top:16px;display:flex;gap:10px;flex-wrap:wrap">
  <a class="btn" href="{{ url_for('editer') }}">✏ Modifier</a>
  <a class="btn sec" href="{{ url_for('exporter') }}">⬇ Exporter</a>
  <form method="post" action="{{ url_for('reinitialiser') }}"
   onsubmit="return confirm('Supprimer toutes les règles ?')">
   <button class="btn danger">🗑 Supprimer</button></form>
 </div>
</div>
{% else %}
<div class="card vide"><div class="ico">📭</div><h2>Aucune règle publiée</h2>
 <p class="sub">Le formulaire utilisateur sera disponible après publication.</p></div>
{% endif %}
"""

COLONNES_PAGE = """
<div class="steps"><div>① Importer</div><div class="on">② Indiquer les colonnes</div>
 <div>③ Vérifier et publier</div></div>
<div class="card">
 <h1>Que contient chaque colonne ?</h1>
 <p class="sub">Choisissez au-dessus de chaque colonne son rôle. Exactement une colonne doit être
 <b>« Énoncé de la condition »</b> : son texte sera repris tel quel dans le questionnaire.
 {{ apercu|length }} lignes affichées sur {{ total }}.</p>
 <form method="post" action="{{ url_for('construire') }}">
  <div class="opts">
   {% if feuilles|length > 1 %}
   <label>Feuille
    <select name="feuille">{% for f in feuilles %}<option {{ 'selected' if f==feuille }}>{{ f }}</option>{% endfor %}</select></label>
   <label><input type="checkbox" name="toutes" value="1"> Importer toutes les feuilles (même structure)</label>
   {% else %}<input type="hidden" name="feuille" value="{{ feuille }}">{% endif %}
   <label>Ligne d'en-tête <input type="number" name="entete" min="0" max="15" value="{{ entete }}" style="width:80px"></label>
   <small style="color:var(--muted)">0 = le tableau n'a pas de titres de colonnes</small>
   <button class="btn sec sm" formaction="{{ url_for('colonnes') }}" formmethod="get">↻ Actualiser l'aperçu</button>
  </div>
  <div class="tw"><table class="prev">
   <thead><tr><th>Ligne</th>
   {% for c in range(nb_cols) %}
    <th><select name="map{{ c }}">
     {% for code, label in roles %}<option value="{{ code }}" {{ 'selected' if pre[c]==code }}>{{ label }}</option>{% endfor %}
    </select></th>
   {% endfor %}</tr></thead>
   <tbody>
   {% for row in apercu %}
    <tr class="{{ 'ent' if loop.index==entete }}"><td>{{ loop.index }}</td>
    {% for c in range(nb_cols) %}<td>{{ row[c] if c < row|length else '' }}</td>{% endfor %}</tr>
   {% endfor %}
   </tbody>
  </table></div>
  <div style="margin-top:18px;display:flex;gap:10px;flex-wrap:wrap">
   <button class="btn big">Générer les questions →</button>
   <a class="btn sec big" href="{{ url_for('expert') }}">Annuler</a>
  </div>
 </form>
</div>
"""

EDITEUR = """
{% macro ligne(r, al) %}
<tr>
 <td><input type="text" name="cat" value="{{ r.Categorie }}" placeholder="Catégorie"></td>
 <td><textarea name="cond" rows="3" placeholder="Énoncé de la condition">{{ r.Condition }}</textarea></td>
 <td><select name="typ">
  <option value="oui/non" {{ 'selected' if r.Type!='nombre' }}>Conformité (Oui/Non)</option>
  <option value="nombre" {{ 'selected' if r.Type=='nombre' }}>Mesure (nombre)</option></select></td>
 <td class="m"><select name="op">
  {% for o in ['>=','<=','>','<','==','entre'] %}<option value="{{ o }}" {{ 'selected' if r.Operateur==o }}>{{ o }}</option>{% endfor %}
 </select></td>
 <td class="m"><input type="text" name="val" value="{{ r.Valeur if r.Type=='nombre' else '' }}" placeholder="25 ou 10-15"></td>
 <td class="m"><input type="text" name="unite" value="{{ r.Unite }}" style="min-width:60px"></td>
 <td><select name="oblig">
  <option value="oui" {{ 'selected' if r.Obligatoire!='non' }}>oui</option>
  <option value="non" {{ 'selected' if r.Obligatoire=='non' }}>non</option></select></td>
 <td><textarea name="question" rows="2" placeholder="(automatique)">{{ r.Question }}</textarea></td>
 <td><textarea name="details" rows="2" placeholder="Info affichée sous l'énoncé">{{ r.Details }}</textarea></td>
 <td>{% if al %}{% for a in al %}<span class="badge b-warn">⚠ {{ a }}</span>{% endfor %}
  {% else %}<span class="badge b-ok">✓</span>{% endif %}</td>
 <td><button type="button" class="btn danger sm" onclick="this.closest('tr').remove()">✕</button></td>
</tr>
{% endmacro %}

<div class="steps"><div>① Importer</div><div>② Indiquer les colonnes</div>
 <div class="on">③ Vérifier et publier</div></div>
<div class="card">
 <h1>{{ 'Vérifiez les questions générées' if brouillon else 'Éditeur de règles' }}</h1>
 <p class="sub">Chaque ligne = une question. <b>Conformité</b> : l'utilisateur répond Oui / Non / Sans objet
 à l'énoncé. <b>Mesure</b> : il saisit une valeur comparée au seuil (ex. <code>25</code> avec <code>≥</code>,
 ou <code>10-15</code> avec <code>entre</code>). Supprimez avec ✕ les lignes qui ne sont pas des conditions
 (titres de sections par exemple).</p>
 <form method="post" action="{{ url_for('enregistrer') }}">
  <div class="tw"><table class="edit">
   <thead><tr><th>Catégorie</th><th>Énoncé de la condition</th><th>Mode</th><th>Opérateur</th>
   <th>Seuil</th><th>Unité</th><th>Oblig.</th><th>Question perso.</th><th>Info</th>
   <th>Contrôle</th><th></th></tr></thead>
   <tbody id="corps">
   {% for r in regles %}{{ ligne(r, alertes[loop.index0]) }}{% endfor %}
   </tbody>
  </table></div>
  <template id="tpl">{{ ligne({}, []) }}</template>
  <div style="margin-top:16px;display:flex;gap:10px;flex-wrap:wrap">
   <button type="button" class="btn sec" onclick="ajouter()">＋ Ajouter une condition</button>
   <button class="btn">✔ Valider et publier</button>
   <a class="btn sec" href="{{ url_for('expert') }}">Annuler</a>
  </div>
 </form>
</div>
<script>
function adapter(tr) {
  const mesure = tr.querySelector('[name=typ]').value === 'nombre';
  tr.querySelectorAll('.m').forEach(function (td) {
    td.style.visibility = mesure ? 'visible' : 'hidden';
  });
}
function ajouter() {
  const corps = document.getElementById('corps');
  corps.appendChild(document.getElementById('tpl').content.cloneNode(true));
  adapter(corps.lastElementChild);
}
document.querySelectorAll('#corps tr').forEach(adapter);
document.addEventListener('change', function (e) {
  if (e.target.name === 'typ') adapter(e.target.closest('tr'));
});
</script>
"""

VERIFIER = """
{% if not regles %}
<div class="card vide"><div class="ico">🔒</div>
 <h1>Vérification indisponible</h1>
 <p class="sub">L'expert n'a pas encore publié ses règles antisismiques.</p>
 <a class="btn" href="{{ url_for('expert') }}">Aller à l'espace expert</a></div>
{% else %}
<form id="f" method="post" action="{{ url_for('resultat') }}" novalidate>
<div class="progress">
 <b id="plabel">0 / {{ regles|length }} réponses</b>
 <div class="track"><div class="fill" id="fill"></div></div>
</div>
<div class="chips">
 <button type="button" class="chip on">🏢 Bâtiment</button>
 {% for cat in groupes %}<button type="button" class="chip">{{ loop.index }}. {{ cat }}</button>{% endfor %}
</div>

<div class="card etape">
 <h2>🏢 Le bâtiment</h2>
 <p class="sub">{{ regles|length }} conditions à vérifier, en {{ groupes|length }} étape(s).</p>
 <label><b>Nom du bâtiment</b> <small>(facultatif)</small></label><br>
 <input type="text" name="nom_batiment" placeholder="Ex : Résidence Les Oliviers" style="margin-top:6px">
</div>

{% for cat, liste in groupes.items() %}
<div class="card etape" hidden>
 <h2>{{ loop.index }}. {{ cat }} <span class="badge b-oblig">{{ liste|length }} condition{{ 's' if liste|length>1 }}</span></h2>
 {% for r in liste %}
 <div class="q {{ 'first' if loop.first }}">
  <div class="num">{{ r.i+1 }}</div>
  <div class="corps">
   <div class="enonce">{{ r.Condition }}</div>
   {% if r.Details %}<div class="details">{{ r.Details }}</div>{% endif %}
   <div class="qtxt">{{ r.question }}
    {% if r.Obligatoire != 'oui' %}<span class="badge b-fac">facultatif</span>{% endif %}</div>
   {% if r.Type == 'nombre' %}
    <div class="hint">Exigence : <b>{{ r.exigence }}</b></div>
    <div class="unit"><input type="number" step="any" name="r{{ r.i }}" required>
    {% if r.Unite %}<span>{{ r.Unite }}</span>{% endif %}</div>
   {% else %}
    <div class="pills">
     <input type="radio" id="r{{ r.i }}o" name="r{{ r.i }}" value="oui" required><label for="r{{ r.i }}o">✔ Oui</label>
     <input type="radio" id="r{{ r.i }}n" name="r{{ r.i }}" value="non" required><label for="r{{ r.i }}n">✘ Non</label>
     <input type="radio" id="r{{ r.i }}a" name="r{{ r.i }}" value="na" required><label for="r{{ r.i }}a">Sans objet</label>
    </div>
   {% endif %}
  </div>
 </div>
 {% endfor %}
</div>
{% endfor %}

<div class="nav">
 <button type="button" class="btn sec" id="prev">← Précédent</button>
 <button type="button" class="btn" id="next">Suivant →</button>
 <button class="btn big" id="send" hidden>Lancer la vérification ✔</button>
</div>
</form>
<script>
const form = document.getElementById('f');
const etapes = [...document.querySelectorAll('.etape')];
const chips = [...document.querySelectorAll('.chip')];
const total = {{ regles|length }};
let cur = 0;

function aller(i) {
  cur = i;
  etapes.forEach(function (e, k) { e.hidden = (k !== i); });
  chips.forEach(function (c, k) { c.classList.toggle('on', k === i); });
  document.getElementById('prev').style.visibility = i === 0 ? 'hidden' : 'visible';
  document.getElementById('next').hidden = (i === etapes.length - 1);
  document.getElementById('send').hidden = (i !== etapes.length - 1);
  window.scrollTo({top: 0, behavior: 'smooth'});
}
document.getElementById('prev').onclick = function () { aller(cur - 1); };
document.getElementById('next').onclick = function () {
  const bad = etapes[cur].querySelector(':invalid');
  if (bad) { bad.reportValidity(); return; }
  aller(cur + 1);
};
chips.forEach(function (c, k) { c.onclick = function () { aller(k); }; });
form.addEventListener('submit', function (ev) {
  for (let k = 0; k < etapes.length; k++) {
    const bad = etapes[k].querySelector(':invalid');
    if (bad) { ev.preventDefault(); aller(k); bad.reportValidity(); return; }
  }
});

function maj() {
  let ok = 0;
  for (let i = 0; i < total; i++) {
    let rempli = false;
    form.querySelectorAll('[name=r' + i + ']').forEach(function (e) {
      if (e.type === 'radio') { if (e.checked) rempli = true; }
      else if (e.value.trim() !== '') rempli = true;
    });
    if (rempli) ok++;
  }
  document.getElementById('fill').style.width = (100 * ok / total) + '%';
  document.getElementById('plabel').textContent = ok + ' / ' + total + ' réponses';
}
form.addEventListener('input', maj);
form.addEventListener('change', maj);
maj(); aller(0);
</script>
{% endif %}
"""

RESULTAT = """
<div class="verdict {{ 'ok' if conforme else 'ko' }}">
 <div class="ring" style="--p:{{ score }};--c:{{ 'var(--ok)' if conforme else 'var(--ko)' }}">
  <div>{{ score }}%</div></div>
 <div>
  <h1>{{ '✅ Bâtiment conforme' if conforme else '❌ Bâtiment non conforme' }}</h1>
  <p style="margin:6px 0 0">
  {% if nom %}<b>{{ nom }}</b> · {% endif %}
  {{ nb_ok }} condition(s) respectée(s) sur {{ total }}{% if nb_na %} ({{ nb_na }} sans objet){% endif %}.
  {% if conforme %}Toutes les conditions obligatoires sont satisfaites.
  {% else %}{{ echecs|length }} condition(s) obligatoire(s) non respectée(s).{% endif %}</p>
 </div>
</div>

<div class="card" style="margin-top:20px">
 <h2>📊 Conformité par catégorie</h2>
 {% for c in cats %}
 <div class="cbar"><span class="n">{{ c.nom }}</span>
  <div class="track"><div class="fill" style="--w:{{ c.pct }}%"></div></div>
  <b>{{ c.ok }}/{{ c.total }}</b></div>
 {% endfor %}
</div>

{% if echecs %}
<div class="card">
 <h2>⚠ Points à corriger</h2>
 {% for l in echecs %}
 <div class="todo"><b>{{ l.condition }}</b><br>
  {% if l.exigence %}Exigé : {{ l.exigence }} · {% endif %}Réponse : <b>{{ l.saisie }}</b><br>
  <small>{{ l.categorie }}</small></div>
 {% endfor %}
</div>
{% endif %}

<div class="card">
 <h2>Détail des conditions</h2>
 {% for cat, liste in groupes.items() %}
 <h3 style="margin:18px 0 6px">{{ cat }}</h3>
 <div class="tw"><table>
 <tr><th>Condition</th><th>Réponse</th><th>Niveau</th><th>Statut</th></tr>
 {% for l in liste %}
 <tr><td>{{ l.condition }}
  {% if l.exigence %}<br><small style="color:var(--muted)">Exigé : {{ l.exigence }}</small>{% endif %}</td>
 <td>{{ l.saisie }}</td>
 <td><span class="badge {{ 'b-oblig' if l.oblig else 'b-fac' }}">{{ 'Obligatoire' if l.oblig else 'Facultatif' }}</span></td>
 <td>{% if l.ok is none %}<span class="badge b-fac">Sans objet</span>
  {% elif l.ok %}<span class="badge b-ok">Respectée</span>
  {% else %}<span class="badge {{ 'b-ko' if l.oblig else 'b-warn' }}">Non respectée</span>{% endif %}</td></tr>
 {% endfor %}
 </table></div>
 {% endfor %}
</div>
<div class="noprint" style="display:flex;gap:10px;flex-wrap:wrap">
 <a class="btn" href="{{ url_for('verifier') }}">Nouvelle vérification</a>
 <button class="btn sec" onclick="window.print()">🖨 Imprimer le rapport</button>
</div>
"""


def page(gabarit, actif="", **ctx):
    corps = render_template_string(gabarit, **ctx)
    return render_template_string(LAYOUT, body=corps, actif=actif)


# ---------------------------------------------------------------
#  Routes : accueil et espace expert
# ---------------------------------------------------------------
@app.route("/")
def accueil():
    return page(ACCUEIL, actif="accueil")


@app.route("/expert")
def expert():
    brut = charger_regles()
    regles = preparer(brut) if brut else []
    maj = ""
    if brut:
        maj = datetime.fromtimestamp(os.path.getmtime(FICHIER)).strftime("%d/%m/%Y %H:%M")
    return page(
        EXPERT, actif="expert", existe=bool(brut), regles=regles, maj=maj,
        attente=os.path.exists(BROUILLON),
        nb_mesure=sum(1 for r in regles if r["Type"] == "nombre"),
        nb_cat=len({r["Categorie"] for r in regles}),
    )


@app.route("/expert/importer", methods=["POST"])
def importer():
    f = request.files.get("fichier")
    nom = (f.filename or "").lower() if f else ""
    if not nom.endswith((".xlsx", ".xlsm", ".csv")):
        flash("Format non supporté : utilisez .xlsx, .xlsm ou .csv.", "err")
        return redirect(url_for("expert"))
    supprimer_imports()
    f.save("import_en_cours" + os.path.splitext(nom)[1])
    return redirect(url_for("colonnes"))


@app.route("/expert/colonnes")
def colonnes():
    chemin = chemin_import()
    if not chemin:
        return redirect(url_for("expert"))
    try:
        feuilles = lire_feuilles(chemin)
    except Exception as e:
        flash(f"Fichier illisible : {e}", "err")
        return redirect(url_for("expert"))
    valides = {str(n): lignes_feuille(d) for n, d in feuilles.items()}
    valides = {n: l for n, l in valides.items() if l}
    if not valides:
        flash("Le fichier est vide.", "err")
        return redirect(url_for("expert"))

    feuille = request.args.get("feuille")
    if feuille not in valides:
        feuille = next(iter(valides))
    lignes = valides[feuille]
    entete_arg = request.args.get("entete", type=int)
    entete, pre = detecter(lignes, entete_arg)
    nb_cols = len(pre)
    return page(
        COLONNES_PAGE, actif="expert", feuilles=list(valides), feuille=feuille,
        entete=entete, pre=pre, nb_cols=nb_cols, roles=ROLES,
        apercu=lignes[:10], total=len(lignes),
    )


@app.route("/expert/construire", methods=["POST"])
def construire():
    chemin = chemin_import()
    if not chemin:
        return redirect(url_for("expert"))
    feuille = request.form.get("feuille", "")
    entete = request.form.get("entete", type=int, default=0)
    mapping = {}
    for c in range(MAX_COLS):
        role = request.form.get(f"map{c}")
        if role and role != "ignorer":
            mapping[c] = role
    if list(mapping.values()).count("enonce") != 1:
        flash("Choisissez exactement UNE colonne « Énoncé de la condition ».", "err")
        return redirect(url_for("colonnes", feuille=feuille, entete=entete))
    try:
        regles, alertes = construire_regles(lire_feuilles(chemin), feuille, entete,
                                            mapping, bool(request.form.get("toutes")))
    except Exception as e:
        flash(f"Import impossible : {e}", "err")
        return redirect(url_for("colonnes", feuille=feuille, entete=entete))
    if not regles:
        flash("Aucune condition trouvée avec ce choix de colonnes.", "err")
        return redirect(url_for("colonnes", feuille=feuille, entete=entete))
    with open(BROUILLON, "w", encoding="utf-8") as fh:
        json.dump({"regles": regles, "alertes": alertes}, fh, ensure_ascii=False)
    flash(f"{len(regles)} condition(s) importée(s). Vérifiez les questions puis publiez.", "ok")
    return redirect(url_for("editer", brouillon=1))


@app.route("/expert/editer")
def editer():
    brouillon = request.args.get("brouillon") == "1" and os.path.exists(BROUILLON)
    if brouillon:
        with open(BROUILLON, encoding="utf-8") as fh:
            d = json.load(fh)
        regles, alertes = d["regles"], d["alertes"]
    else:
        regles = charger_regles() or []
        alertes = [[] for _ in regles]
    if not regles:
        regles, alertes = [{}], [[]]
    return page(EDITEUR, actif="expert", regles=regles, alertes=alertes, brouillon=brouillon)


@app.route("/expert/enregistrer", methods=["POST"])
def enregistrer():
    champs = {"cat": "Categorie", "cond": "Condition", "typ": "Type", "op": "Operateur",
              "val": "Valeur", "unite": "Unite", "oblig": "Obligatoire",
              "question": "Question", "details": "Details"}
    listes = {k: request.form.getlist(k) for k in champs}
    brutes = []
    for i in range(len(listes["cond"])):
        d = {champs[k]: listes[k][i] if i < len(listes[k]) else "" for k in champs}
        if d["Condition"].strip():
            brutes.append(d)

    valides, erreurs, alertes = [], [], []
    for n, d in enumerate(brutes, start=1):
        r, err = valider_regle(d)
        valides.append(r)
        if err:
            erreurs.append(f"Condition {n} : {err}")
            alertes.append(["à corriger : " + err])
        else:
            alertes.append([])

    if not valides:
        flash("Ajoutez au moins une condition.", "err")
        return redirect(url_for("editer"))
    if erreurs:
        with open(BROUILLON, "w", encoding="utf-8") as fh:
            json.dump({"regles": brutes, "alertes": alertes}, fh, ensure_ascii=False)
        flash("Publication refusée :\n" + "\n".join(erreurs[:6]), "err")
        return redirect(url_for("editer", brouillon=1))

    sauver_regles(valides)
    if os.path.exists(BROUILLON):
        os.remove(BROUILLON)
    supprimer_imports()
    flash(f"{len(valides)} condition(s) publiée(s). Le formulaire est à jour.", "ok")
    return redirect(url_for("expert"))


@app.route("/expert/exporter")
def exporter():
    if not os.path.exists(FICHIER):
        return redirect(url_for("expert"))
    return send_file(FICHIER, as_attachment=True, download_name="regles_antisismiques.xlsx")


@app.route("/expert/modele")
def modele():
    ex = pd.DataFrame([
        ["Sol", "Le bâtiment est fondé sur un sol ferme ou rocheux.", "", "", "oui"],
        ["Structure", "Le nombre d'étages ne doit pas dépasser 5.", "", "", "oui"],
        ["Structure", "Présence de voiles de contreventement dans les deux directions.", "", "", "oui"],
        ["Matériaux", "Résistance minimale du béton", "25", "MPa", "oui"],
        ["Matériaux", "Espacement des cadres en zone critique", "≤ 10", "cm", "non"],
    ], columns=["Catégorie", "Condition", "Valeur", "Unité", "Obligatoire"])
    flux = io.BytesIO()
    ex.to_excel(flux, index=False)
    flux.seek(0)
    return send_file(flux, as_attachment=True, download_name="modele_conditions.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/expert/reinitialiser", methods=["POST"])
def reinitialiser():
    for f in (FICHIER, BROUILLON):
        if os.path.exists(f):
            os.remove(f)
    supprimer_imports()
    flash("Les règles ont été supprimées.", "ok")
    return redirect(url_for("expert"))


# ---------------------------------------------------------------
#  Routes : utilisateur
# ---------------------------------------------------------------
@app.route("/verifier")
def verifier():
    brut = charger_regles()
    regles = preparer(brut) if brut else []
    return page(VERIFIER, actif="verifier", regles=regles, groupes=grouper(regles))


@app.route("/resultat", methods=["POST"])
def resultat():
    brut = charger_regles()
    if not brut:
        return redirect(url_for("verifier"))
    lignes = []
    for r in preparer(brut):
        saisie = request.form.get(f"r{r['i']}", "").strip()
        if r["Type"] == "nombre":
            ok = evaluer(r, saisie)
            affiche = f"{saisie} {r['Unite']}".strip()
        else:
            ok = None if saisie == "na" else evaluer(r, saisie)
            affiche = {"oui": "Oui", "non": "Non", "na": "Sans objet"}.get(saisie, saisie)
        lignes.append({
            "categorie": r["Categorie"], "condition": r["Condition"],
            "exigence": r["exigence"], "saisie": affiche,
            "oblig": r["Obligatoire"] == "oui", "ok": ok,
        })

    echecs = [l for l in lignes if l["ok"] is False and l["oblig"]]
    evaluees = [l for l in lignes if l["ok"] is not None]
    nb_ok = sum(1 for l in evaluees if l["ok"])
    groupes = {}
    for l in lignes:
        groupes.setdefault(l["categorie"], []).append(l)
    cats = []
    for nom, liste in groupes.items():
        ev = [l for l in liste if l["ok"] is not None]
        ok = sum(1 for l in ev if l["ok"])
        cats.append({"nom": nom, "ok": ok, "total": len(ev),
                     "pct": round(100 * ok / len(ev)) if ev else 100})
    return page(
        RESULTAT, actif="verifier", groupes=groupes, echecs=echecs, cats=cats,
        conforme=not echecs, nb_ok=nb_ok, total=len(evaluees),
        nb_na=len(lignes) - len(evaluees),
        score=round(100 * nb_ok / len(evaluees)) if evaluees else 100,
        nom=request.form.get("nom_batiment", "").strip(),
    )


if __name__ == "__main__":
    app.run(debug=True)gh auth login