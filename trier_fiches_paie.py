#!/usr/bin/env python3
"""
Tri de fiches de paie (PDF numériques) + récapitulatif annuel + fiche de report
pour la déclaration de revenus.

Tout reste en local. Les originaux ne sont jamais modifiés : le script COPIE
les fichiers dans un dossier de sortie, rangés par année.

Utilisation en ligne de commande :
    pip install -r requirements.txt
    python trier_fiches_paie.py "C:\\chemin\\vers\\les\\pdf" -o resultat
    python trier_fiches_paie.py "C:\\chemin\\vers\\les\\pdf" -o resultat --declarant 2
    python trier_fiches_paie.py "C:\\chemin\\vers\\les\\pdf" -o resultat --debug

--declarant : 1 = déclarant 1 (case 1AJ), 2 = déclarant 2 / conjoint (1BJ),
              3 = personne à charge 1 (1CJ), 4 = personne à charge 2 (1DJ).

Pour une version avec fenêtre : python app.py

Gérer une mise en page inconnue :
    1. lancer avec --debug : le texte brut de chaque PDF est écrit dans
       resultat/debug/*.txt
    2. repérer comment le net imposable / la période / l'employeur y apparaissent
    3. ajouter une expression dans un fichier patterns.json (à côté du
       programme), sans toucher au code. Exemple :
       {
         "net_imposable": ["montant fiscal"],
         "employeur": ["agence\\s*:\\s*(.+)"]
       }
"""
import argparse
import copy
import hashlib
import html
import json
import re
import shutil
import sys
from collections import defaultdict
from pathlib import Path

try:
    import pdfplumber
    from openpyxl import Workbook
    from openpyxl.styles import Font
except ImportError:  # pragma: no cover
    sys.exit("Il manque des modules. Installe-les avec :\n  pip install -r requirements.txt")

# --------------------------------------------------------------------------
# Motifs de lecture (plusieurs variantes par champ pour couvrir plusieurs
# mises en page). Les motifs sont testés dans l'ordre, du plus fiable au moins
# fiable.
# --------------------------------------------------------------------------
PATTERNS_DEFAUT = {
    # libellés (regex, insensible à la casse) suivis d'un montant
    "net_imposable": [
        r"net\s+imposable",
        r"net\s+fiscal",
        r"revenu\s+net\s+imposable",
        r"salaire\s+net\s+imposable",
        r"net\s+soumis\s+[àa]\s+l.?imp[oô]t",
    ],
    "cumul_net_imposable": [
        r"cumul\w*\s+.{0,25}net\s+imposable",
        r"net\s+imposable\s+(?:cumul\w*|annuel|depuis)",
        r"net\s+imposable\s+cumul",
        r"cumul\s+annuel\s+net\s+imposable",
    ],
    "brut": [
        r"salaire\s+brut",
        r"total\s+brut",
        r"brut\s+total",
        r"^\s*brut\b",
        r"r[ée]mun[ée]ration\s+brute",
    ],
    "net_a_payer": [
        r"net\s+[àa]\s+payer\s+avant\s+imp[oô]t",
        r"net\s+[àa]\s+payer",
        r"net\s+pay[ée]",
        r"net\s+vers[ée]",
    ],
    # employeur : motifs avec un groupe capturant le nom
    "employeur": [
        # « : » ou tiret suivi d'une espace (pas un signe moins collé à un montant)
        r"(?:employeur|raison\s+sociale|soci[ée]t[ée]|entreprise|agence)\s*(?::|-\s)\s*(.+)",
    ],
}
PATTERNS = copy.deepcopy(PATTERNS_DEFAUT)

# Agences / enseignes courantes d'intérim (reconnues n'importe où dans le texte)
EMPLOYEURS_CONNUS = [
    "adecco", "randstad", "manpower", "synergie", "crit", "partnaire", "supplay",
    "proman", "daco", "start people", "temporis", "interaction", "samsic",
    "actual", "gi group", "walters", "terre d'int", "triangle", "domino",
    "inter-contrats", "groupe daco", "ctt", "adéquat", "adequat", "job & box",
]

MOIS = {
    "janvier": 1, "fevrier": 2, "février": 2, "mars": 3, "avril": 4, "mai": 5,
    "juin": 6, "juillet": 7, "aout": 8, "août": 8, "septembre": 9,
    "octobre": 10, "novembre": 11, "decembre": 12, "décembre": 12,
}
NOMS_MOIS = ["", "janv", "fév", "mars", "avr", "mai", "juin", "juil", "août",
             "sept", "oct", "nov", "déc"]

# Cases « traitements et salaires » de la déclaration de revenus (formulaire 2042)
# selon la personne concernée. Vérifier le numéro sur le formulaire de chaque année.
CASES_DECLARANT = {
    "1": ("Déclarant 1", "1AJ"),
    "2": ("Déclarant 2 (conjoint)", "1BJ"),
    "3": ("Personne à charge 1", "1CJ"),
    "4": ("Personne à charge 2", "1DJ"),
}

SEP_MILLIERS = r"[ \u00a0\u202f.]"
MONTANT = re.compile(
    r"(?<![\d/])(-?\d{1,3}(?:" + SEP_MILLIERS + r"\d{3})*[,.]\d{2}|-?\d+[,.]\d{2})(?![\d/])"
)

_perso_charge = False


def dossier_programme():
    """Dossier où se trouve le programme (ou le .exe une fois compilé)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def charger_patterns_perso(log=print):
    """Ajoute les motifs de patterns.json (s'il existe) en tête de liste."""
    global _perso_charge
    if _perso_charge:
        return
    p = dossier_programme() / "patterns.json"
    if p.exists():
        try:
            extra = json.loads(p.read_text(encoding="utf-8"))
            for cle, liste in extra.items():
                PATTERNS.setdefault(cle, [])
                PATTERNS[cle] = list(liste) + PATTERNS[cle]
            log(f"Motifs personnalisés chargés depuis {p.name}")
        except Exception as e:
            log(f"Attention : patterns.json illisible ({e})")
    _perso_charge = True


def en_nombre(txt):
    s = re.sub(r"[ \u00a0\u202f]", "", txt)
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    return float(s)


def arrondi_euro(x):
    """Arrondi à l'euro le plus proche (0,5 arrondi vers le haut)."""
    return int(x + 0.5) if x >= 0 else -int(-x + 0.5)


def format_euro(x, decimales=0):
    """12345.6 -> '12 346' (espace insécable fine pour les milliers)."""
    return f"{x:,.{decimales}f}".replace(",", "\u202f").replace(".", ",")


def lire_texte(pdf_path):
    morceaux = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            morceaux.append(page.extract_text() or "")
    return "\n".join(morceaux)


def montants_ligne(ligne):
    return [en_nombre(m) for m in MONTANT.findall(ligne)]


def chercher_montants(lignes, motifs, exclure=None, apres_label_seulement=True):
    """Retourne la liste des montants de la 1re ligne qui correspond à un motif
    (ou de la ligne suivante si la ligne du libellé n'en contient pas)."""
    for motif in motifs:
        rx = re.compile(motif, re.I)
        for i, ligne in enumerate(lignes):
            if exclure and re.search(exclure, ligne, re.I):
                continue
            m = rx.search(ligne)
            if not m:
                continue
            reste = ligne[m.end():] if apres_label_seulement else ligne
            vals = montants_ligne(reste)
            if not vals and i + 1 < len(lignes):
                vals = montants_ligne(lignes[i + 1])
            if vals:
                return vals
    return None


def detecter_periode(texte):
    """Retourne (annee, mois) ou None. Plusieurs façons de lire la période."""
    t = texte.lower()
    # 1) "du 01/03/2026 au 31/03/2026" -> on prend la date de fin
    m = re.search(
        r"du\s+(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})\s+au\s+(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})", t)
    if m:
        mois, an = int(m.group(5)), int(m.group(6))
        an += 2000 if an < 100 else 0
        if 1 <= mois <= 12:
            return an, mois
    # 2) "période : 03/2026" / "bulletin de paie 03/2026" / "mois : 03-2026"
    m = re.search(r"(?:p[ée]riode|mois|paie|salaire|bulletin)[^\n\d]{0,30}(\d{1,2})[/.\-](\d{4})", t)
    if m and 1 <= int(m.group(1)) <= 12:
        return int(m.group(2)), int(m.group(1))
    # 3) "mars 2026" (nom du mois), 1re occurrence
    noms = "|".join(sorted(MOIS, key=len, reverse=True))
    m = re.search(r"\b(" + noms + r")\s+(?:de\s+)?(\d{4})\b", t)
    if m:
        return int(m.group(2)), MOIS[m.group(1)]
    # 4) mm/aaaa isolé
    m = re.search(r"\b(0?[1-9]|1[0-2])[/.\-](20\d{2})\b", t)
    if m:
        return int(m.group(2)), int(m.group(1))
    return None


def nettoyer_nom(nom):
    nom = re.sub(r"\s{2,}.*", "", nom.strip())          # coupe à la 1re grosse coupure
    nom = re.sub(r"\b(siret|siren|ape|naf|adresse|tel|urssaf).*", "", nom, flags=re.I)
    nom = re.sub(r"[^\w\s&'\-]", " ", nom)
    nom = re.sub(r"\s+", " ", nom).strip(" -_")
    return nom[:40]


def detecter_employeur(texte, lignes):
    bas = texte.lower()
    # 1) libellé explicite "Employeur : XXX"
    for motif in PATTERNS["employeur"]:
        for m in re.finditer(motif, texte, re.I):
            nom = nettoyer_nom(m.group(1))
            # un vrai nom contient des lettres (pas seulement un montant)
            if len(nom) >= 2 and re.search(r"[A-Za-zÀ-ÿ]{2}", nom):
                return nom
    # 2) enseignes connues
    for nom in EMPLOYEURS_CONNUS:
        if nom in bas:
            return nom.title()
    # 3) 1re ligne "utile" du haut du document (souvent le nom de l'entreprise)
    for ligne in lignes[:8]:
        l = ligne.strip()
        if len(l) < 3 or re.search(r"bulletin|paie|salaire|p[ée]riode|n°|\d{5}", l, re.I):
            continue
        if re.search(r"[A-Za-zÀ-ÿ]{3,}", l):
            nom = nettoyer_nom(l)
            if len(nom) >= 3:
                return nom
    return None


def hash_fichier(p):
    h = hashlib.sha1()
    with open(p, "rb") as f:
        for bloc in iter(lambda: f.read(65536), b""):
            h.update(bloc)
    return h.hexdigest()


def nom_sur(nom):
    return re.sub(r"[^\w\-]+", "_", nom).strip("_") or "Inconnu"


def _lignes_de_mots(mots, tolerance=2.5):
    """Regroupe les mots d'une page en lignes (même hauteur), de haut en bas."""
    lignes = []
    for w in sorted(mots, key=lambda m: (m["top"], m["x0"])):
        if lignes and abs(w["top"] - lignes[-1][0]["top"]) <= tolerance:
            lignes[-1].append(w)
        else:
            lignes.append([w])
    return [sorted(l, key=lambda m: m["x0"]) for l in lignes]


def _regrouper_montants(mots):
    """Recolle les montants que l'extraction a coupés en deux ('1' + '234,56')."""
    sortie = []
    for w in mots:
        if (sortie
                and re.fullmatch(r"\d{3}(?:[,.]\d{2})?", w["text"])
                and re.fullmatch(r"-?\d{1,3}(?: \d{3})*", sortie[-1]["text"])
                and w["x0"] - sortie[-1]["x1"] < 5):
            prec = sortie[-1]
            sortie[-1] = {"text": prec["text"] + " " + w["text"], "x0": prec["x0"],
                          "x1": w["x1"], "top": prec["top"]}
        else:
            sortie.append(dict(w))
    return sortie


def extraire_tableau_cumuls(pdf_path, tolerance=6):
    """Lit le net imposable dans un tableau « Cumuls » : le titre de colonne est sur une
    ligne et les montants sur les lignes « Mensuel » / « Annuel » en dessous. Les montants
    sont rattachés au titre par leur position horizontale sur la page.

    Retourne {"mois": montant du mois, "cumul": montant cumulé}, clés absentes si introuvable.
    """
    res = {}
    roles = {"mensuel": "mois", "mois": "mois", "courant": "mois", "mensuelle": "mois",
             "annuel": "cumul", "annuelle": "cumul", "cumul": "cumul", "cumulé": "cumul",
             "cumule": "cumul", "cumuls": "cumul"}
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                lignes = _lignes_de_mots(page.extract_words())
                for i, ligne in enumerate(lignes):
                    textes = [w["text"].lower() for w in ligne]
                    pos = next((j for j in range(len(ligne) - 1)
                                if textes[j] == "net" and textes[j + 1].startswith("imposable")), None)
                    if pos is None:
                        continue
                    x0, x1 = ligne[pos]["x0"], ligne[pos + 1]["x1"]
                    for suivante in lignes[i + 1:i + 7]:
                        role = roles.get(suivante[0]["text"].lower().strip(":"))
                        if role is None or role in res:
                            continue
                        meilleur, meilleur_ov = None, 0
                        for w in _regrouper_montants(suivante[1:]):
                            if not MONTANT.fullmatch(w["text"]):
                                continue
                            ov = min(w["x1"], x1 + tolerance) - max(w["x0"], x0 - tolerance)
                            if ov > meilleur_ov:
                                meilleur, meilleur_ov = w, ov
                        if meilleur is not None:
                            res[role] = en_nombre(meilleur["text"])
                    if res:
                        return res
    except Exception:  # PDF inhabituel : on laisse les autres méthodes de lecture
        return res
    return res


def analyser(pdf_path, dossier_debug=None):
    texte = lire_texte(pdf_path)
    if dossier_debug is not None:
        (dossier_debug / (pdf_path.stem + ".txt")).write_text(texte, encoding="utf-8")
    res = {"fichier": pdf_path, "avert": []}
    if len(texte.strip()) < 40:
        res["avert"].append("pas de texte (scan ? il faudrait de l'OCR)")
        return res
    lignes = [l for l in texte.splitlines() if l.strip()]

    per = detecter_periode(texte)
    res["annee"], res["mois"] = per if per else (None, None)
    if not per:
        res["avert"].append("période non trouvée")

    res["employeur"] = detecter_employeur(texte, lignes)
    if not res["employeur"]:
        res["avert"].append("employeur non trouvé")

    # net imposable du mois : on ignore les lignes de cumul
    v = chercher_montants(lignes, PATTERNS["net_imposable"], exclure=r"cumul|annuel")
    res["net_imposable"] = v[0] if v else None
    vc = chercher_montants(lignes, PATTERNS["cumul_net_imposable"])
    res["cumul"] = vc[-1] if vc else None

    # 2e méthode : tableau « Cumuls » (titre de colonne + lignes Mensuel / Annuel)
    if res["net_imposable"] is None or res["cumul"] is None:
        tableau = extraire_tableau_cumuls(pdf_path)
        if res["net_imposable"] is None:
            res["net_imposable"] = tableau.get("mois")
        if res["cumul"] is None:
            res["cumul"] = tableau.get("cumul")

    if res["net_imposable"] is None:
        res["avert"].append("net imposable non trouvé")
    elif v and len(v) > 1:
        res["avert"].append(f"plusieurs montants sur la ligne ({', '.join(f'{x:.2f}' for x in v)}) : 1er retenu")

    vb = chercher_montants(lignes, PATTERNS["brut"], exclure=r"cumul|annuel")
    res["brut"] = vb[0] if vb else None
    vn = chercher_montants(lignes, PATTERNS["net_a_payer"])
    res["net_a_payer"] = vn[0] if vn else None
    return res


def calculer_mois_manquants(par_cle):
    """Pour chaque employeur : les mois sans fiche entre sa 1re et sa dernière fiche.
    Retourne {(employeur, annee): [noms de mois]}."""
    par_emp = defaultdict(set)
    for (an, emp), fiches in par_cle.items():
        for f in fiches:
            par_emp[emp].add((f["annee"], f["mois"]))
    manques = defaultdict(list)
    for emp, vus in par_emp.items():
        vus_tries = sorted(vus)
        a, m = vus_tries[0]
        fin = vus_tries[-1]
        while (a, m) <= fin:
            if (a, m) not in vus:
                manques[(emp, a)].append(NOMS_MOIS[m])
            m += 1
            if m == 13:
                a, m = a + 1, 1
    return manques


def rappels(case):
    """Rappels affichés avec la fiche de report. `case` ex. '1AJ' -> lettre 'A'."""
    p = case[1]
    return [
        "Si la déclaration est préremplie, compare avec ces montants. Si le montant prérempli est "
        "inexact, il faut le barrer et inscrire le montant exact dans la case.",
        "Les indemnités de fin de mission d'intérim sont imposables ; elles sont normalement déjà "
        "comprises dans le net imposable des fiches.",
        "En principe, un revenu se déclare l'année où il a été versé : une paie de décembre versée en "
        "janvier peut relever de l'année suivante.",
        f"Non calculés ici : frais réels (case 1{p}K), allocations chômage (case 1{p}P), "
        f"pensions et retraites (case 1{p}S).",
        "Apprentis et étudiants : une partie du salaire peut être exonérée (voir la notice des impôts).",
        "Les numéros de case peuvent changer d'une année à l'autre : vérifie sur le formulaire de "
        "l'année concernée.",
        "Cette fiche n'est pas un conseil fiscal. Vérifie les montants avant de les recopier.",
    ]


def ecrire_html(chemin, libelle, case, annees, avert_globaux, lignes_rappels):
    h = html.escape
    parts = [
        "<!doctype html><html lang='fr'><head><meta charset='utf-8'>",
        "<title>Fiche de report - déclaration de revenus</title>",
        "<style>body{font-family:Segoe UI,Arial,sans-serif;max-width:820px;margin:2rem auto;"
        "padding:0 1rem;color:#222}h1{font-size:1.4rem}table{border-collapse:collapse;width:100%;"
        "margin:.5rem 0 1.5rem}th,td{border:1px solid #bbb;padding:.45rem .6rem;text-align:left}"
        "th{background:#f0f0f0}td.n{text-align:right;white-space:nowrap}.ok{color:#176b2c}"
        ".ko{color:#a33;font-weight:600}.case{font-size:1.3rem;font-weight:700}"
        ".alerte{background:#fff4e5;border:1px solid #e0a040;padding:.6rem .8rem;margin:.5rem 0}"
        "@media print{body{margin:0}}</style></head><body>",
        "<h1>Fiche de report pour la déclaration de revenus</h1>",
        f"<p>Personne concernée : <b>{h(libelle)}</b> &mdash; case « traitements et salaires » : "
        f"<span class='case'>{h(case)}</span></p>",
    ]
    for a in avert_globaux:
        parts.append(f"<div class='alerte'>{h(a)}</div>")
    parts.append("<table><tr><th>Année des revenus</th><th>Case</th><th>Montant à reporter (€)</th>"
                 "<th>Montant exact (€)</th><th>Fiches</th><th>Fiabilité</th></tr>")
    for an in sorted(annees):
        d = annees[an]
        ok = not d["raisons"]
        etat = "<span class='ok'>OK</span>" if ok else "<span class='ko'>À vérifier</span>"
        parts.append(
            f"<tr><td>{an}</td><td>{h(case)}</td><td class='n'><b>{format_euro(arrondi_euro(d['total']))}</b></td>"
            f"<td class='n'>{format_euro(d['total'], 2)}</td><td>{d['nb']}</td><td>{etat}</td></tr>")
        if not ok:
            parts.append("<tr><td colspan='6'><ul>" +
                         "".join(f"<li>{h(r)}</li>" for r in d["raisons"]) + "</ul></td></tr>")
    parts.append("</table><h2>Détail par employeur</h2><table><tr><th>Année</th><th>Employeur</th>"
                 "<th>Fiches</th><th>Net imposable (€)</th></tr>")
    for an in sorted(annees):
        for emp, somme, nb in annees[an]["employeurs"]:
            parts.append(f"<tr><td>{an}</td><td>{h(emp)}</td><td>{nb}</td>"
                         f"<td class='n'>{format_euro(somme, 2)}</td></tr>")
    parts.append("</table><h2>À savoir</h2><ul>" +
                 "".join(f"<li>{h(r)}</li>" for r in lignes_rappels) + "</ul></body></html>")
    Path(chemin).write_text("".join(parts), encoding="utf-8")


def traiter(dossier, sortie="tri_fiches", debug=False, log=print, declarant="1"):
    """Trie les PDF de `dossier`, écrit le résultat dans `sortie`.

    `declarant` : "1" à "4" (voir CASES_DECLARANT) = personne dont ce sont les fiches.

    Retourne un dict : xlsx, html, sortie (Path), nb_pdf, doublons, lues,
    a_verifier, totaux (année -> montant), case, aide (année -> dict).
    Lève ValueError si le dossier est introuvable ou sans PDF.
    """
    declarant = str(declarant)
    if declarant not in CASES_DECLARANT:
        raise ValueError("Le déclarant doit être 1, 2, 3 ou 4.")
    libelle, case = CASES_DECLARANT[declarant]

    charger_patterns_perso(log)
    src = Path(dossier)
    out = Path(sortie)
    if not src.is_dir():
        raise ValueError(f"Dossier introuvable : {src}")
    out.mkdir(parents=True, exist_ok=True)
    dbg = None
    if debug:
        dbg = out / "debug"
        dbg.mkdir(exist_ok=True)

    out_res = out.resolve()
    pdfs = sorted(
        p for p in src.rglob("*")
        if p.is_file() and p.suffix.lower() == ".pdf" and out_res not in p.resolve().parents
    )
    if not pdfs:
        raise ValueError("Aucun PDF trouvé dans ce dossier.")
    log(f"{len(pdfs)} PDF trouvés, lecture en cours…")

    vus = {}
    resultats, doublons = [], []
    for p in pdfs:
        h = hash_fichier(p)
        if h in vus:
            doublons.append((p, vus[h]))
            continue
        vus[h] = p
        try:
            resultats.append(analyser(p, dbg))
        except Exception as e:
            resultats.append({"fichier": p, "avert": [f"erreur de lecture : {e}"]})

    # ----- rangement (copie) -----
    utilises = set()
    for r in resultats:
        ok = r.get("annee") and r.get("mois")
        if ok:
            dest_dir = out / str(r["annee"])
            base = f"{r['annee']}-{r['mois']:02d}_{nom_sur(r.get('employeur') or 'Inconnu')}"
        else:
            dest_dir = out / "A_verifier"
            base = r["fichier"].stem
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / f"{base}.pdf"
        n = 2
        while dest in utilises or dest.exists():
            dest = dest_dir / f"{base}_{n}.pdf"
            n += 1
        utilises.add(dest)
        shutil.copy2(r["fichier"], dest)
        r["classe_dans"] = dest

    # ----- agrégats -----
    par_cle = defaultdict(list)
    for r in resultats:
        if r.get("annee"):
            par_cle[(r["annee"], r.get("employeur") or "Inconnu")].append(r)
    manques = calculer_mois_manquants(par_cle)

    annees = {}          # année -> infos pour la fiche de report
    lignes_recap = []
    for (an, emp) in sorted(par_cle):
        fiches = par_cle[(an, emp)]
        somme = sum(f["net_imposable"] or 0 for f in fiches)
        sans_montant = sum(1 for f in fiches if f["net_imposable"] is None)
        cumuls = [(f["mois"], f["cumul"]) for f in fiches if f.get("cumul") is not None]
        cumul = max(cumuls)[1] if cumuls else None
        ecart = round(cumul - somme, 2) if cumul is not None else None
        rem = []
        d = annees.setdefault(an, {"total": 0.0, "nb": 0, "employeurs": [], "raisons": []})
        d["total"] += somme
        d["nb"] += len(fiches)
        d["employeurs"].append((emp, somme, len(fiches)))
        if sans_montant:
            rem.append(f"{sans_montant} fiche(s) sans net imposable lu")
            d["raisons"].append(f"{emp} : {sans_montant} fiche(s) sans net imposable lu")
        if ecart is not None and abs(ecart) > 1:
            rem.append("écart avec le cumul : fiches manquantes ou lecture à vérifier")
            d["raisons"].append(f"{emp} : écart de {format_euro(ecart, 2)} € avec le cumul "
                                "de la dernière fiche")
        if manques.get((emp, an)):
            d["raisons"].append(f"{emp} : mois sans fiche ({', '.join(manques[(emp, an)])})")
        lignes_recap.append([an, emp, len(fiches), round(somme, 2), cumul, ecart, " ; ".join(rem)])

    nb_non_classees = sum(1 for r in resultats if not (r.get("annee") and r.get("mois")))
    avert_globaux = []
    if nb_non_classees:
        avert_globaux.append(
            f"{nb_non_classees} fiche(s) n'ont pas pu être classées (dossier « A_verifier ») : "
            "leurs montants ne sont pas dans les totaux ci-dessous.")
    rappels_txt = rappels(case)

    wb = Workbook()
    gras = Font(bold=True)

    # --- onglet 1 : aide à la déclaration ---
    wa = wb.active
    wa.title = "Aide déclaration"
    wa.append(["Aide au remplissage de la déclaration de revenus (à vérifier avant de recopier)"])
    wa.cell(wa.max_row, 1).font = Font(bold=True, size=13)
    wa.append([f"Personne concernée : {libelle} — case « traitements et salaires » : {case}"])
    for a in avert_globaux:
        wa.append([f"ATTENTION : {a}"])
        wa.cell(wa.max_row, 1).font = gras
    wa.append([])
    wa.append(["Année des revenus", "Case", "Montant à reporter (€, arrondi)",
               "Montant exact (€)", "Nb de fiches", "Fiabilité", "À vérifier"])
    for c in wa[wa.max_row]:
        c.font = gras
    for an in sorted(annees):
        d = annees[an]
        wa.append([an, case, arrondi_euro(d["total"]), round(d["total"], 2), d["nb"],
                   "OK" if not d["raisons"] else "À vérifier", " ; ".join(d["raisons"])])
    wa.append([])
    wa.append(["Détail par employeur"])
    wa.cell(wa.max_row, 1).font = gras
    wa.append(["Année", "Employeur", "Nb de fiches", "Net imposable (€)"])
    for c in wa[wa.max_row]:
        c.font = gras
    for an in sorted(annees):
        for emp, somme, nb in annees[an]["employeurs"]:
            wa.append([an, emp, nb, round(somme, 2)])
    wa.append([])
    wa.append(["À savoir"])
    wa.cell(wa.max_row, 1).font = gras
    for r in rappels_txt:
        wa.append(["• " + r])

    # --- onglet 2 : récap annuel ---
    ws = wb.create_sheet("Récap annuel")
    ws.append(["Année", "Employeur", "Nb fiches", "Net imposable (somme des fiches)",
               "Cumul lu sur la dernière fiche", "Écart", "Remarque"])
    for ligne in lignes_recap:
        ws.append(ligne)
    ws.append([])
    ws.append(["TOTAL PAR ANNÉE"])
    ws.cell(ws.max_row, 1).font = gras
    totaux_annee = {an: d["total"] for an, d in annees.items()}
    for an in sorted(totaux_annee):
        ws.append([an, "Tous employeurs", None, round(totaux_annee[an], 2)])

    wd = wb.create_sheet("Détail")
    wd.append(["Fichier d'origine", "Classé dans", "Année", "Mois", "Employeur", "Brut",
               "Net imposable", "Net à payer", "Cumul net imposable", "Avertissements"])
    for r in sorted(resultats, key=lambda x: (x.get("annee") or 0, x.get("mois") or 0)):
        wd.append([r["fichier"].name, str(r["classe_dans"].relative_to(out)), r.get("annee"),
                   r.get("mois"), r.get("employeur"), r.get("brut"), r.get("net_imposable"),
                   r.get("net_a_payer"), r.get("cumul"), " ; ".join(r["avert"])])

    wm = wb.create_sheet("Mois manquants")
    wm.append(["Employeur", "Année", "Mois sans fiche (entre la 1re et la dernière fiche)"])
    for (emp, an) in sorted(manques):
        wm.append([emp, an, ", ".join(manques[(emp, an)])])
    wm.append([])
    wm.append(["Remarque : seuls les mois situés entre la première et la dernière fiche d'un employeur "
               "sont signalés. Pour de l'intérim à la semaine, plusieurs fiches par mois sont normales."])

    wv = wb.create_sheet("À vérifier")
    wv.append(["Fichier", "Problème"])
    nb_verif = 0
    for r in resultats:
        if r["avert"]:
            wv.append([r["fichier"].name, " ; ".join(r["avert"])])
            nb_verif += 1
    for p, orig in doublons:
        wv.append([p.name, f"doublon exact de {orig.name} (ignoré)"])

    for feuille in wb.worksheets:
        if feuille.title != "Aide déclaration":   # cet onglet a sa propre mise en forme
            for c in feuille[1]:
                c.font = gras
        for col in feuille.columns:
            largeur = max((len(str(c.value)) for c in col if c.value is not None and len(str(c.value)) < 70),
                          default=8)
            feuille.column_dimensions[col[0].column_letter].width = min(largeur + 2, 60)

    chemin_xlsx = out / "recapitulatif_fiches_paie.xlsx"
    wb.save(chemin_xlsx)
    chemin_html = out / "fiche_de_report_declaration.html"
    ecrire_html(chemin_html, libelle, case, annees, avert_globaux, rappels_txt)

    nb_ok = sum(1 for r in resultats if r.get("annee") and r.get("net_imposable") is not None)
    log(f"{len(pdfs)} PDF trouvés, {len(doublons)} doublon(s) ignoré(s), "
        f"{nb_ok}/{len(resultats)} fiches lues complètement.")
    log(f"À reporter dans la case {case} ({libelle}) :")
    for an in sorted(annees):
        d = annees[an]
        etat = "" if not d["raisons"] else "  ← À VÉRIFIER"
        log(f"  {an} : {format_euro(arrondi_euro(d['total']))} €{etat}")
    log(f"Fiche de report : {chemin_html}")
    log(f"Tableau complet : {chemin_xlsx}")
    log("IMPORTANT : vérifie ces montants (cumul de la dernière fiche de l'année, revenus déjà "
        "préremplis) avant de les reporter sur la déclaration.")

    return {
        "xlsx": chemin_xlsx, "html": chemin_html, "sortie": out, "nb_pdf": len(pdfs),
        "doublons": len(doublons), "lues": nb_ok, "a_verifier": nb_verif + len(doublons),
        "totaux": totaux_annee, "case": case, "aide": annees,
    }


def main():
    ap = argparse.ArgumentParser(description="Trie des fiches de paie PDF, fait un récap annuel "
                                             "et une fiche de report pour la déclaration de revenus.")
    ap.add_argument("dossier", help="dossier contenant les PDF (sous-dossiers inclus)")
    ap.add_argument("-o", "--sortie", default="tri_fiches", help="dossier de sortie")
    ap.add_argument("--declarant", choices=sorted(CASES_DECLARANT), default="1",
                    help="1=déclarant 1 (1AJ), 2=conjoint (1BJ), 3/4=personnes à charge (1CJ/1DJ)")
    ap.add_argument("--debug", action="store_true", help="écrit le texte brut extrait de chaque PDF")
    args = ap.parse_args()
    try:
        traiter(args.dossier, args.sortie, args.debug, declarant=args.declarant)
    except ValueError as e:
        sys.exit(str(e))


if __name__ == "__main__":
    main()
