# -*- coding: utf-8 -*-
"""
Analyse cadastrale par CODE POSTAL (France entière)
===================================================
Toute l'analyse (stats, histogrammes, tableaux, carte) est limitée aux
communes correspondant au code postal fourni — et non au département entier.

1. Résout le code postal via geo.api.gouv.fr (noms, codes INSEE, départements)
2. Télécharge les parcelles et bâtiments du Cadastre Etalab pour ces communes
   uniquement (fichiers par commune : rapides)
3. Croise avec les Demandes de Valeurs Foncières (DVF) pour les prix au m²
4. Croise avec le PLU (Géoportail de l'urbanisme) : zones constructibles
   (U, AU) et emprise au sol maximale par parcelle — depuis les
   prescriptions structurées (WFS), ou à défaut depuis le règlement
   écrit du PLU (PDF, nécessite pypdf)
5. Exports (dossier ./resultats_<code_postal>/) :
   - stats_synthese.txt        : synthèse (nombre, moyenne, médiane, bâti, prix, PLU)
   - tableau_parcelles.xlsx    : tableau détaillé par parcelle (onglets)
   - tableau_constructibilite.xlsx : zonage PLU + emprise au sol max par parcelle
   - stats_par_communes.csv    : stats par commune du code postal
   - histogramme_tailles.png   : distribution des tailles
   - histogramme_log.png       : distribution (échelle log)
   - emprise_batie_communes.png : surfaces bâties par commune
   - carte_parcelles.html      : carte interactive des parcelles
   - carte_zones_constructibles.html : carte des zones constructibles (PLU)

Prérequis :
    pip install geopandas matplotlib folium openpyxl
    pip install pypdf        # optionnel : emprise au sol depuis le règlement PDF

Usage :
    python analyse-cadastrale.py 10000      # Troyes
    python analyse-cadastrale.py 75001     # Paris 1er
    Sans argument : 10000 (Troyes) par défaut.
    [!] Argument = CODE POSTAL uniquement (pas le code INSEE) : un code INSEE
    (ex. 10387) est rejeté avec un message d'aide.
"""

import gzip
import hashlib
import io
import json
import os
import re
import shutil
import sys
import time
import urllib.parse
import urllib.request
import zipfile

import geopandas as gpd
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# Console Windows (cp1252) : un caractère non représentable ne doit pas
# faire planter le script (UnicodeEncodeError) — repli sur « ? ».
if sys.platform == "win32":
    sys.stdout.reconfigure(errors="replace")
    sys.stderr.reconfigure(errors="replace")

BASE_URL = "https://cadastre.data.gouv.fr/data/etalab-cadastre/latest/geojson/communes"

# DVF : demandes de valeurs foncières géolocalisées (data.gouv.fr).
# Fichier unique couvrant janvier 2021 - décembre 2025 (CSV gzippé).
# [!] L'identifiant de ressource DVF change à chaque millésime : si l'URL
# fixe échoue, le script interroge l'API data.gouv.fr pour trouver le
# lien le plus récent (voir trouver_url_dvf).
DVF_URL_FIXE = ("https://static.data.gouv.fr/resources/"
                "demandes-de-valeurs-foncieres-geolocalisees/"
                "20260717-142041/dvf.csv.gz")
DVF_API_DATASET = "demandes-de-valeurs-foncieres-geolocalisees"
DVF_ANNEE = "2021-2025"
DVF_CACHE_LEGACY = f"dvf_{DVF_ANNEE}.csv"  # ancien cache intégral (racine)

# PLU : Géoportail de l'urbanisme (WFS IGN + API du GPU)
GPU_WFS = "https://data.geopf.fr/wfs/ows"
GPU_ZONE_URBA = "wfs_du:zone_urba"                  # zonage du PLU
GPU_PRESCRIPTION_SURF = "wfs_du:prescription_surf"  # prescriptions surfaciques
GPU_API = "https://www.geoportail-urbanisme.gouv.fr/api"
CACHE_DIR = ".cache_cadastre"

ARG_CP = sys.argv[1] if len(sys.argv) > 1 else "10000"

# Certains serveurs rejettent le User-Agent « Python-urllib » par défaut.
USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
              "AppleWebKit/537.36 (KHTML, like Gecko) "
              "Chrome/126.0.0.0 Safari/537.36")


def _ouvrir_url(url, timeout=60):
    """Ouvre une URL avec un User-Agent de navigateur."""
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    return urllib.request.urlopen(req, timeout=timeout)


_dernier_telechargement = 0.0  # throttling entre deux requêtes (rate-limiting)


def _telecharger_fichier(url, dest, tentatives=6, timeout=300):
    """Télécharge url vers dest avec plusieurs tentatives espacées
    (les serveurs ferment parfois la connexion : WinError 10054,
    timeouts…). Les requêtes sont espacées d'au moins 1 s pour éviter
    le rate-limiting. Renvoie True si réussi, False sinon."""
    global _dernier_telechargement
    for essai in range(1, tentatives + 1):
        reste = 1.0 - (time.monotonic() - _dernier_telechargement)
        if reste > 0:
            time.sleep(reste)
        _dernier_telechargement = time.monotonic()
        try:
            with _ouvrir_url(url, timeout=timeout) as rep, \
                    open(dest, "wb") as f:
                shutil.copyfileobj(rep, f)
            return True
        except Exception as exc:
            if os.path.exists(dest):
                os.remove(dest)
            if essai < tentatives:
                delai = 5 * essai  # 5 s, 10 s, 15 s, 20 s, 25 s
                print(f"[!] {os.path.basename(dest)} : essai {essai}/"
                      f"{tentatives} ({exc}), nouvel essai dans {delai} s…",
                      file=sys.stderr)
                time.sleep(delai)
            else:
                print(f"[!] {os.path.basename(dest)} indisponible ({exc})",
                      file=sys.stderr)
    return False


def resoudre_code_postal(cp):
    """Résout un CODE POSTAL via geo.api.gouv.fr (codes INSEE refusés).
    Renvoie (codes_insee, df_communes) ; df_communes contient code_insee,
    nom_commune, code_postal et code_departement pour chaque commune du
    code postal (un code postal peut couvrir plusieurs départements)."""
    cp = str(cp).strip()
    if not cp.isdigit():
        raise SystemExit(f"[!] {cp} : merci d'indiquer un CODE POSTAL à 5 chiffres "
                         f"(ex. 10000 pour Troyes).")

    url = (f"https://geo.api.gouv.fr/communes?codePostal={urllib.parse.quote(cp)}"
           f"&fields=code,nom,codeDepartement&format=json")
    with _ouvrir_url(url, timeout=30) as rep:
        communes = json.load(rep)
    if not communes:
        # Le paramètre répondrait aussi à un code INSEE : vérifions ce cas
        url2 = (f"https://geo.api.gouv.fr/communes?code={urllib.parse.quote(cp)}"
                f"&fields=code,nom&format=json")
        with _ouvrir_url(url2, timeout=30) as rep:
            par_insee = json.load(rep)
        if par_insee:
            raise SystemExit(
                f"[!] {cp} semble être un CODE INSEE (commune de "
                f"{par_insee[0]['nom']}), pas un code postal. Les codes INSEE "
                f"ne sont pas acceptés. Utilisez son code postal (ex. 10000 "
                f"pour Troyes, 75001 pour Paris 1er).")
        raise SystemExit(f"[!] {cp} : code postal inconnu.")

    df = pd.DataFrame([{"code_insee": c["code"], "nom_commune": c["nom"],
                        "code_postal": cp,
                        "code_departement": c["codeDepartement"]}
                       for c in communes])
    libelle = ", ".join(c["nom"] for c in communes[:5]) + \
              ("…" if len(communes) > 5 else "")
    deps = sorted(set(df["code_departement"]))
    print(f"[i] Code postal {cp} -> {libelle} "
          f"(département{'s' if len(deps) > 1 else ''} {'/'.join(deps)})")
    return df["code_insee"].tolist(), df


def telecharger_commune(dep, code_insee, couche):
    """Télécharge et décompresse un fichier communal du Cadastre Etalab.
    Structure réelle des URL : communes/{département}/{commune}/…
    Les fichiers sont mis en cache dans .cache_cadastre/ (les anciens
    fichiers téléchargés à la racine sont réutilisés).
    Renvoie le chemin du .json, ou None si indisponible pour la commune."""
    gz = f"cadastre-{code_insee}-{couche}.json.gz"
    js = gz[:-3]
    for dossier in (CACHE_DIR, "."):
        chemin = os.path.join(dossier, js)
        if os.path.exists(chemin):
            return chemin
    os.makedirs(CACHE_DIR, exist_ok=True)
    dest_gz = os.path.join(CACHE_DIR, gz)
    dest_js = os.path.join(CACHE_DIR, js)
    url = f"{BASE_URL}/{dep}/{code_insee}/{gz}"
    print(f"[>>] Téléchargement : {gz}")
    if not _telecharger_fichier(url, dest_gz, timeout=300):
        return None
    try:
        with gzip.open(dest_gz, "rb") as f_in, open(dest_js, "wb") as f_out:
            shutil.copyfileobj(f_in, f_out)
        os.remove(dest_gz)
        return dest_js
    except Exception as exc:
        if os.path.exists(dest_gz):
            os.remove(dest_gz)
        print(f"[!] {gz} illisible après téléchargement ({exc})",
              file=sys.stderr)
        return None


def reparer_geometries(gdf):
    """Répare les géométries invalides (polygones auto-intersectés, etc.)."""
    nb = int((~gdf.geometry.is_valid).sum())
    if nb:
        print(f"[i] Réparation de {nb:,} géométries invalides...")
        gdf = gdf.copy()
        try:
            gdf["geometry"] = gdf.geometry.make_valid()
        except AttributeError:
            gdf["geometry"] = gdf.geometry.buffer(0)
        gdf = gdf[~gdf.geometry.isna() & ~gdf.geometry.is_empty]
    return gdf


def _cle_cache(url):
    """Empreinte courte d'une URL (nom de fichier de cache)."""
    return hashlib.md5(url.encode("utf-8")).hexdigest()[:10]


def trouver_url_dvf():
    """URL du CSV géolocalisé DVF le plus récent : l'URL fixe connue
    d'abord, sinon interrogation de l'API data.gouv.fr (dernier millésime).
    Renvoie None si indisponible."""
    try:
        req = urllib.request.Request(DVF_URL_FIXE, method="HEAD",
                                     headers={"User-Agent": USER_AGENT})
        with urllib.request.urlopen(req, timeout=30) as rep:
            if rep.status == 200:
                return DVF_URL_FIXE
    except Exception:
        pass
    print("[...] Recherche du DVF géolocalisé via l'API data.gouv.fr...")
    try:
        url = f"https://www.data.gouv.fr/api/1/datasets/{DVF_API_DATASET}/"
        with _ouvrir_url(url, timeout=30) as rep:
            dataset = json.load(rep)
        candidats = []
        for res in dataset.get("resources", []):
            url_res = res.get("url") or ""
            titre = f"{res.get('title') or ''} {url_res}"
            if url_res.lower().endswith(".csv.gz") and \
                    re.search(r"dvf|g[ée]olocalis", titre, re.I):
                candidats.append((res.get("last_modified") or "", url_res))
        if candidats:
            candidats.sort()
            return candidats[-1][1]
    except Exception as exc:
        print(f"[!] API data.gouv.fr indisponible ({exc})", file=sys.stderr)
    return None


def charger_dvf(deps):
    """Ventes DVF géolocalisées, filtrées sur les départements du code postal.
    Le fichier national (plusieurs Go) est lu par morceaux ; le résultat
    filtré est mis en cache (.cache_cadastre/). Renvoie un DataFrame ou
    None si indisponible."""
    deps = [str(d) for d in deps]
    usecols = ["id_parcelle", "valeur_fonciere", "surface_reelle_bati",
               "nature_mutation", "code_departement"]
    os.makedirs(CACHE_DIR, exist_ok=True)

    # Source : cache national existant (téléchargement intégral déjà fait),
    # sinon téléchargement du millésime le plus récent.
    source = None
    cle = None
    national = os.path.join(CACHE_DIR, "dvf_national.csv")
    for candidat, cle_c in ((DVF_CACHE_LEGACY, "legacy"), (national, "cache")):
        if os.path.exists(candidat) and os.path.getsize(candidat) > 1e9:
            source, cle = candidat, cle_c
            break
    if source is None:
        url = trouver_url_dvf()
        if url is None:
            print("[!] DVF indisponible : colonnes de prix vides.",
                  file=sys.stderr)
            return None
        cle = _cle_cache(url)
        gz = os.path.join(CACHE_DIR, "dvf.csv.gz")
        print(f"[>>] Téléchargement DVF {DVF_ANNEE} (France entière, "
              f"~500 Mo) : {url}")
        if not _telecharger_fichier(url, gz, tentatives=6, timeout=1800):
            print("[!] DVF indisponible : colonnes de prix vides.",
                  file=sys.stderr)
            return None
        try:
            with gzip.open(gz, "rb") as f_in, open(national, "wb") as f_out:
                shutil.copyfileobj(f_in, f_out)
            os.remove(gz)
            source = national
        except Exception as exc:
            print(f"[!] DVF illisible après téléchargement ({exc}) : "
                  f"colonnes de prix vides.", file=sys.stderr)
            if os.path.exists(gz):
                os.remove(gz)
            return None

    cache_filtre = os.path.join(
        CACHE_DIR, f"dvf_filtre_{cle}_dep{'+'.join(deps)}.csv")
    if os.path.exists(cache_filtre):
        print(f"[...] Lecture du cache DVF filtré : {cache_filtre}")
        dvf = pd.read_csv(cache_filtre, dtype={"id_parcelle": str,
                                               "code_departement": str},
                          low_memory=False)
    else:
        print(f"[...] Lecture du fichier DVF et filtrage sur le(s) "
              f"département(s) {'/'.join(deps)}...")
        try:
            morceaux = []
            for chunk in pd.read_csv(source, sep=",", usecols=usecols,
                                     dtype={"id_parcelle": str,
                                            "code_departement": str},
                                     low_memory=False, chunksize=2_000_000):
                chunk = chunk[chunk["code_departement"].isin(deps)]
                chunk = chunk[chunk["nature_mutation"] == "Vente"]
                if len(chunk):
                    morceaux.append(chunk)
            dvf = pd.concat(morceaux, ignore_index=True) if morceaux \
                else pd.DataFrame(columns=usecols)
        except ValueError as exc:
            print(f"[!] Colonnes DVF inattendues ({exc}).", file=sys.stderr)
            return None
        dvf.to_csv(cache_filtre, index=False)
        print(f"[OK] Cache DVF filtré : {cache_filtre} ({len(dvf):,} lignes)")

    dvf = dvf.dropna(subset=["id_parcelle", "valeur_fonciere"])
    # Identifiants de parcelles : 14 caractères (commune 5 + préfixe 3 +
    # section 2 + numéro 4) dans le DVF comme dans le cadastre Etalab.
    dvf["id_parcelle"] = (dvf["id_parcelle"].astype(str)
                          .str.replace("_", "", regex=False).str.zfill(14))
    print(f"[OK] {len(dvf):,} mutations retenues pour le(s) département(s) "
          f"{'/'.join(deps)}.")
    return dvf


def requete_wfs_gpu(couche, cql, page=5000):
    """Requête WFS paginée sur le Géoportail de l'urbanisme (GeoServer IGN).
    Renvoie un GeoDataFrame reprojeté en EPSG:2154, ou None si vide."""
    params = {"service": "WFS", "version": "2.0.0", "request": "GetFeature",
              "typeNames": couche, "outputFormat": "application/json",
              "CQL_FILTER": cql, "COUNT": str(page)}
    features = []
    debut = 0
    while True:
        params["STARTINDEX"] = str(debut)
        url = GPU_WFS + "?" + urllib.parse.urlencode(params)
        data = None
        for essai in range(1, 5):
            try:
                with _ouvrir_url(url, timeout=300) as rep:
                    data = json.load(rep)
                break
            except Exception as exc:
                if essai < 4:
                    print(f"[!] WFS {couche} (page {debut}) : essai "
                          f"{essai}/4 ({exc}), nouvel essai dans "
                          f"{5 * essai} s…", file=sys.stderr)
                    time.sleep(5 * essai)
        if data is None:
            raise RuntimeError(f"WFS {couche} indisponible après 4 essais")
        lot = data.get("features", [])
        features.extend(lot)
        if len(lot) < page:
            break
        debut += page
    if not features:
        return None
    gdf = gpd.GeoDataFrame.from_features(features, crs="EPSG:4326")
    if gdf.crs is None or gdf.crs.to_epsg() != 2154:
        gdf = gdf.to_crs(2154)
    return gdf


def _taux_emprise(texte):
    """Taux d'emprise au sol (en %) dans un libellé de prescription PLU.
    « Taux d'emprise au sol : 40 % » -> 40.0 ; « CES 0,4 » -> 40.0.
    Renvoie None si aucun taux n'est trouvé."""
    texte = str(texte)
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*(?:%|pour cent)", texte)
    if m:
        taux = float(m.group(1).replace(",", "."))
        return taux * 100 if taux <= 1 else taux
    # Ratio décimal sans % (ex. « CES 0,4 »)
    m = re.search(r"(?:CES|taux|emprise)[^0-9]{0,20}(\d+(?:[.,]\d+)?)",
                  texte, re.I)
    if m:
        taux = float(m.group(1).replace(",", "."))
        if 0 < taux <= 1:
            return taux * 100
    return None


def telecharger_reglement_gpu(gpu_doc_id):
    """Télécharge le règlement (PDF) d'un document d'urbanisme via l'API
    du Géoportail de l'urbanisme. Renvoie le chemin du fichier, ou None
    si indisponible."""
    if not gpu_doc_id:
        return None
    os.makedirs(CACHE_DIR, exist_ok=True)
    for f in os.listdir(CACHE_DIR):
        if f.startswith(f"reglement_{gpu_doc_id}") and f.endswith(".pdf"):
            return os.path.join(CACHE_DIR, f)
    try:
        url = f"{GPU_API}/document/{gpu_doc_id}/files"
        with _ouvrir_url(url, timeout=60) as rep:
            fichiers = json.load(rep)
    except Exception as exc:
        print(f"[!] API du Géoportail de l'urbanisme indisponible ({exc})",
              file=sys.stderr)
        return None
    if isinstance(fichiers, dict):
        fichiers = fichiers.get("files", [])
    choix = None
    for f in fichiers:
        nom = str(f.get("name") or f.get("fileName") or "")
        titre = str(f.get("title") or "")
        if nom.lower().endswith(".pdf") and \
                re.search(r"r[eè]glement", f"{nom} {titre}", re.I):
            choix = f
            break
    if choix is None:
        print("[!] Aucun règlement (PDF) trouvé pour ce document d'urbanisme.",
              file=sys.stderr)
        return None
    nom = choix.get("name") or choix.get("fileName")
    dest = os.path.join(CACHE_DIR, f"reglement_{gpu_doc_id}_{nom}")
    url = f"{GPU_API}/document/{gpu_doc_id}/files/{urllib.parse.quote(nom)}"
    print(f"[>>] Téléchargement du règlement : {nom}")
    if _telecharger_fichier(url, dest, timeout=600):
        return dest
    return None


def _lire_texte_pdf(pdf_path):
    """Texte brut d'un PDF (pypdf, optionnel). Renvoie None si indisponible."""
    try:
        from pypdf import PdfReader
    except ImportError:
        print("[!] pypdf non installé : lecture du règlement ignorée "
              "(pip install pypdf pour l'emprise au sol).", file=sys.stderr)
        return None
    try:
        reader = PdfReader(pdf_path)
        pages = []
        for page in reader.pages:
            try:
                pages.append(page.extract_text() or "")
            except Exception:
                pass
        texte = "\n".join(pages)
    except Exception as exc:
        print(f"[!] Lecture du règlement impossible ({exc})", file=sys.stderr)
        return None
    # Caractères non imprimables (polices PDF exotiques) -> espaces
    return "".join(c if c.isprintable() else " " for c in texte)


def _est_clause_ppri(texte):
    """Vrai si le texte relève d'une clause PPRI (plan de prévention des
    risques d'inondation) : limitation d'emprise liée au risque
    (« sous réserve d'en limiter la vulnérabilité »), à ne pas confondre
    avec une règle d'emprise au sol du PLU."""
    return bool(re.search(r"(?i)vuln[ée]rab|ne pas augmenter|aggraver|"
                          r"al[ée]a|inondab|crue", texte))


def _taux_emprise_dans_bloc(bloc):
    """Premier pourcentage d'un bloc de texte rattaché à une emprise au
    sol : la phrase précédant le % doit mentionner l'emprise (ou le CES).
    Un article intitulé « emprise au sol » peut ne contenir en réalité
    qu'un taux d'imperméabilisation (ex. Creney, zone UG : « La surface
    du sol imperméabilisée ne doit pas dépasser 60 % », à ne pas confondre
    avec l'emprise au sol des constructions) : ces pourcentages sont
    ignorés, de même que les clauses PPRI (risque inondation). En second
    recours, un pourcentage directement accolé à la locution
    (« EMPRISE AU SOL : 40 % ») ou un coefficient décimal
    (« coefficient d'emprise au sol : 0,4 » -> 40 %) est retenu. Enfin,
    « il n'est pas fixé de règles » ou « non réglementée » dans une
    sous-section « emprise au sol » signifie l'absence de limite :
    100 %. Renvoie None si aucun taux n'est trouvé."""
    for m in re.finditer(r"(\d+(?:[.,]\d+)?)(?:\s+(\d{1,2}))?\s*%", bloc):
        # La fenêtre est la phrase précédant le % (les tirets ne sont pas
        # des séparateurs : les césures PDF produisent « encorbel - lements »).
        debut = max(bloc.rfind(".", 0, m.start()),
                    bloc.rfind(";", 0, m.start()),
                    bloc.rfind("\u2014", 0, m.start()))
        fenetre = bloc[debut + 1:m.start()].lower()
        if "imperméabilis" in fenetre or "impermeabilis" in fenetre:
            continue  # taux d'imperméabilisation, pas une emprise
        if not re.search(r"emprise|coefficient|C\.?\s*E\.?\s*S", fenetre):
            continue
        if _est_clause_ppri(bloc[m.end():m.end() + 200]):
            continue  # clause PPRI (risque inondation), pas une règle PLU
        # « 6 0 % » : artefact d'extraction PDF pour « 60 % »
        valeur = m.group(1) + (m.group(2) or "")
        taux = float(valeur.replace(",", "."))
        if taux > 0:
            return taux
    m = re.search(r"(?i)emprise\s+au\s+sol\s*[-:\u2013\u2014]?\s*"
                  r"(\d+(?:[.,]\d+)?)\s*%", bloc)
    if m and not _est_clause_ppri(bloc[m.end():m.end() + 200]):
        return float(m.group(1).replace(",", "."))
    # Coefficient décimal : « coefficient d'emprise au sol : 0,4 » -> 40 %
    m = re.search(r"(?i)(?:coefficient|taux)\s+(?:d.?emprise|de l.?emprise)"
                  r"[^0-9]{0,60}?(\d+(?:[.,]\d+)?)", bloc)
    if not m:
        m = re.search(r"(?i)\bC\.?\s*E\.?\s*S\.?\b[^0-9]{0,30}?"
                      r"(\d+(?:[.,]\d+)?)", bloc)
    if m:
        v = float(m.group(1).replace(",", "."))
        return v * 100 if v <= 1 else v
    # « Il n'est pas fixé de règles » dans la sous-section « emprise au
    # sol » : aucune limite -> 100 % (valable uniquement dans ce contexte)
    if re.search(r"(?i)emprise\s+au\s+sol.{0,120}?"
                 r"(pas\s+fix|pas\s+de\s+r[eè]gle)", bloc):
        return 100.0
    # « Non réglementée » (même sens : aucune limite -> 100 %). Le
    # contexte est la locution « emprise au sol » ou, dans les
    # règlements au titre du RNU, « possibilités maximales d'occupation
    # du sol » (article 14 ; ex. Montsuzain : « Article non réglementé »).
    if re.search(r"(?i)(?:emprise\s+au\s+sol|possibilit[eé]s?\s+maximales?"
                 r"\s+d.?occupation\s+du\s+sol).{0,120}?"
                 r"(?:non|n'est\s+pas|pas)\s+r[eéè]glement", bloc):
        return 100.0
    return None


def _parse_article_emprise(section):
    """Extrait les règles d'emprise au sol d'un article de règlement PLU.
    Renvoie {"taux": pct ou None, "tranches": [(surface_min_m2, pct)],
    "alias": {zone_cible: zone_source}}. Le taux peut dépendre de la
    surface du terrain (paliers), ex. UAA à Troyes : < 100 m² non
    réglementé, 100-500 m² 90 %, 500-1000 m² 70 %, >= 1000 m² 50 %."""
    regles = {"taux": None, "tranches": None, "alias": {}}
    # Renvois : « les dispositions prévues pour le secteur UAA sont
    # applicables aux secteurs 1AUAA », « ... dans la zone UCB sont
    # applicables au secteur 2AUCB ». Le mot-clé (secteur/zone) est exigé
    # devant la source et le texte entre les deux ne doit contenir ni
    # point ni deux-points : cela exclut les pieds de page du type
    # « ... Règlement : pièces écrites - Règles applicables aux zones UB »
    # (fausse correspondance « PLU ... applicables aux zones UB »).
    # « sect eur » / « z one » : artefacts d'extraction PDF (saut de ligne).
    for m in re.finditer(r"(?:sect\s*eur|z\s*one)\s+([A-Z0-9]{1,8})\s+"
                         r"[^.:]{0,120}?"
                         r"(?:applicables?|s'appliquent)\s+aux?\s+"
                         r"(?:sect\s*eurs?|z\s*ones?)\s+([A-Z0-9]{1,8})",
                         section):
        regles["alias"][m.group(2)] = m.group(1)
    # Paliers selon la surface du terrain :
    # « Terrain de moins de 100 m² : non réglementé »,
    # « Terrain de 100 m² à moins de 500 m² : 90% »,
    # « Terrains de 500 à moins de 1000m² : 70% »,
    # « Terrains à partir de 1000m² : 50% »
    tranches = []
    for m in re.finditer(
            r"[Tt]errains?\s+(?:de\s+|à partir de\s+|de plus de\s+)?"
            r"(moins de\s+)?(\d+)\s*(?:m²)?"
            r"(?:\s+à\s+(?:moins de\s+)?(\d+)\s*m²)?"
            r"\s*:\s*(non r[eé]glement[ée]|interdit\w*|(\d+(?:[.,]\d+)?)\s*%?)",
            section):
        moins_de, bas, haut, valeur, pct = m.groups()
        if valeur.startswith("non") or valeur.startswith("interdit"):
            taux = None
        else:
            taux = float(pct.replace(",", "."))
        tranches.append((0 if moins_de else int(bas), taux))
    if tranches:
        regles["tranches"] = sorted(tranches)
    else:
        # Format tableau : « Terrain : < 100 m² 100 à - de 500 m² ...
        # > 1000 m² CES : 80% 70% 50% 40% » (chaque valeur a son %)
        m = re.search(r"[Tt]errain\s*:\s*([^:]{0,200}?)\s*CES\s*:\s*"
                      r"((?:\d+(?:[.,]\d+)?\s*%\s*)+)", section) or \
            re.search(r"[Tt]errain\s*:\s*([^:]{0,200}?)\s*CES\s*:\s*"
                      r"((?:\d+(?:[.,]\d+)?\s*)+)", section)
        if m:
            bornes = []
            for b in re.finditer(r"<\s*(\d+)\s*m²"
                                 r"|(\d+)\s*à\s*[-\u2013]?\s*de\s*(\d+)\s*m²"
                                 r"|>\s*(\d+)\s*m²", m.group(1)):
                if b.group(1):
                    bornes.append(0)
                elif b.group(2):
                    bornes.append(int(b.group(2)))
                else:
                    bornes.append(int(b.group(4)))
            valeurs = [float(v.replace(",", "."))
                       for v in re.findall(r"\d+(?:[.,]\d+)?", m.group(2))]
            if bornes and len(bornes) == len(valeurs):
                regles["tranches"] = sorted(zip(bornes, valeurs))
    if regles["tranches"] is None and not regles["alias"]:
        # Taux unique : « limitée à 40% de la superficie de la parcelle ».
        taux = _taux_emprise_dans_bloc(section)
        if taux is not None:
            regles["taux"] = taux
    return regles


def extraire_taux_emprise_reglement(pdf_path):
    """Taux d'emprise au sol par zone, extraits du règlement (PDF) du PLU.
    Renvoie {code_zone: règles} (voir _parse_article_emprise), ou None
    si le règlement est illisible.

    Deux repérages complémentaires, car les mises en page sont variées :
    1. les articles dont l'intitulé porte sur l'emprise au sol — ou sur
       la « volumétrie et implantation » et les « possibilités maximales
       d'occupation du sol » (PMOS), où la règle figure en sous-section :
       « ARTICLE UB 10 : EMPRISE AU SOL ... » (Troyes),
       « ARTICLE 9 - EMPRISE AU SOL ... » (Creney),
       « ARTICLE U4 : VOLUMETRIE ET IMPLANTATION ... » (Vailly) ;
       l'analyse article par article gère aussi les paliers selon la
       surface du terrain et les renvois entre secteurs ;
    2. la recherche directe de la locution « emprise au sol » : le
       paragraphe qui suit contient le pourcentage. La zone est alors
       celle de l'article ou de l'en-tête de section en cours
       (« 8 UC PERSPECTIVES ... », « UG ARTICLE 6 - ... »,
       « ARTICLE U4 : ... »)."""
    texte = _lire_texte_pdf(pdf_path)
    if not texte:
        return None
    if len(texte) < 200:
        print("[!] Règlement illisible (PDF probablement scanné, sans "
              "couche de texte).", file=sys.stderr)
        return None
    texte = re.sub(r"\s+", " ", texte)
    tiret = r"[-:\u2013\u2014]"
    # Code de zone : lettres et chiffres, la lettre en première ou
    # deuxième position (« UB », « 1AU »), collée ou non au numéro
    # d'article (« ARTICLE U4 : », « ARTICLE 1AU 9 - »).
    zone_code = r"(?:[A-Z][A-Z0-9]{0,7}|\d[A-Z][A-Z0-9]{0,6})"
    num = r"\d{1,2}"
    # Marqueurs de zone (position, zone) : en-têtes de section et
    # intitulés d'articles portant le code de zone (« 8 UC PERSPECTIVES »,
    # « 12 UC ARTICLE 9 - », « ARTICLE U4 : »,
    # « DISPOSITIONS APPLICABLES A LA ZONE UCA »).
    entetes = []
    for m in re.finditer(r"\b\d{1,2}\s+([A-Z0-9]{1,6})\s+PERSPECTIVES", texte):
        entetes.append((m.start(), m.group(1).upper()))
    for m in re.finditer(r"\b\d{1,2}\s+([A-Z0-9]{1,6})\s+ARTICLE\s+" + num
                         + r"\s*" + tiret, texte):
        entetes.append((m.start(), m.group(1).upper()))
    for m in re.finditer(r"ARTICLE\s+(" + zone_code + r")\s*" + num + r"\s*"
                         + tiret, texte):
        entetes.append((m.start(), m.group(1).upper()))
    for m in re.finditer(r"(?i)DISPOSITIONS APPLICABLES A LA ZONE\s+"
                         r"([A-Z0-9]{1,6})(?![A-Z0-9'])", texte):
        entetes.append((m.start(), m.group(1).upper()))
    # Ne garder que les codes de zone vraisemblables : au moins une lettre,
    # pas un numéro de partie en chiffres romains (I, II, III...) ni une
    # référence légale d'article (« ARTICLE L130-1 » -> « L13 »).
    entetes = [(pos, z) for pos, z in entetes
               if re.search(r"[A-Z]", z) and not re.fullmatch(r"[IVX]+", z)
               and not re.fullmatch(r"[LR]\d+", z)]
    entetes.sort()

    def zone_du_contexte(pos):
        """Dernière zone des marqueurs précédant la position."""
        zone = ""
        for debut, z in entetes:
            if debut >= pos:
                break
            zone = z
        return zone

    # 1. Articles portant sur l'emprise au sol
    titre_emprise = re.compile(
        r"EMPRISE AU SOL|COEFFICIENT D.?EMPRISE|TAUX D.?EMPRISE|"
        r"POSSIBILITES MAXIMALES D.?OCCUPATION DU SOL|"
        r"VOLUMETRIE.{0,8}IMPLANTATION", re.I)
    # L'intitulé ne doit pas avaler l'article suivant (« ARTICLE ... »)
    article_re = re.compile(r"ARTICLE\s+(?:(" + zone_code + r")\s*)?(" + num
                            + r")\s*" + tiret
                            + r"\s*((?:(?!ARTICLE\b)[^:-]){0,80})")
    article_suivant = re.compile(r"ARTICLE\s+(?:" + zone_code + r"\s*)?"
                                 + num + r"\s*" + tiret)
    candidats = []  # (position, zone, bloc, article_complet)
    for m in article_re.finditer(texte):
        if not titre_emprise.search(m.group(3)):
            continue
        zone = (m.group(1) or zone_du_contexte(m.start()) or "").upper()
        if not zone:
            continue
        nxt = article_suivant.search(texte, m.end())
        fin = nxt.start() if nxt else min(m.end() + 3000, len(texte))
        candidats.append((m.start(), zone, texte[m.start():fin], True))
    # 2. Locution « emprise au sol » : le paragraphe qui suit contient le
    #    taux (le bloc reprend un peu de contexte avant la locution)
    for m in re.finditer(r"(?i)emprise\s+au\s+sol", texte):
        zone = zone_du_contexte(m.start()).upper()
        if not zone:
            continue
        debut = max(0, m.start() - 80)
        fin = min(m.end() + 250, len(texte))
        candidats.append((m.start(), zone, texte[debut:fin], False))
    candidats.sort()

    articles = {}
    for _, zone, bloc, article_complet in candidats:
        precedent = articles.get(zone)
        if precedent is not None and (precedent["taux"] is not None
                                     or precedent["tranches"]
                                     or precedent["alias"]):
            continue  # une règle est déjà connue pour cette zone
        if article_complet:
            regles = _parse_article_emprise(bloc)
        else:
            # Bloc court (locution seule) : taux simple, sans paliers
            regles = {"taux": _taux_emprise_dans_bloc(bloc),
                      "tranches": None, "alias": {}}
        if precedent is None or regles["taux"] is not None \
                or regles["tranches"] or regles["alias"]:
            articles[zone] = regles
    if not articles:
        print("[!] Aucune règle d'emprise au sol trouvée dans le règlement.",
              file=sys.stderr)
        return None
    print(f"[i] Règlement : règles d'emprise au sol pour {len(articles)} "
          f"zones ({', '.join(sorted(articles))}).")
    return articles


def regles_zone(code_zone, articles):
    """Règles d'emprise au sol d'un code de zone du PLU, en suivant les
    renvois d'articles (« les dispositions du secteur UAA s'appliquent au
    secteur 1AUAA ») et les préfixes (zone UCAY -> article UCA si UCAY
    absent). Renvoie None si la zone n'a pas de règle d'emprise."""
    code = str(code_zone).strip().upper()
    vus = set()
    while code and code not in vus:
        vus.add(code)
        regles = articles.get(code)
        if regles and (regles["taux"] is not None or regles["tranches"]):
            return regles
        # Renvoi vers un autre secteur
        renvoi = None
        for art in articles.values():
            if code in art["alias"]:
                renvoi = art["alias"][code]
                break
        if renvoi:
            code = renvoi
            continue
        # Article de la zone mère (1AUAA -> AUAA -> AU)
        prefixes = [a for a in articles if code.startswith(a)]
        if prefixes:
            code = max(prefixes, key=len)
            continue
        # Chiffre initial (1AU -> AU)
        sans_chiffre = re.sub(r"^\d+", "", code)
        if sans_chiffre != code:
            code = sans_chiffre
            continue
        return None
    return None


def taux_emprise_zone(code_zone, surface_m2, articles):
    """Taux d'emprise au sol (en %) pour une parcelle : la règle dépend
    parfois de la surface du terrain (paliers). None si non réglementé
    ou si la zone n'a pas de règle d'emprise."""
    regles = regles_zone(code_zone, articles)
    if not regles:
        return None
    if regles["tranches"]:
        taux = None
        for bas, t in regles["tranches"]:
            if surface_m2 >= bas:
                taux = t
            else:
                break
        return taux
    return regles["taux"]


def main():
    codes_insee, communes = resoudre_code_postal(ARG_CP)
    deps = sorted(set(communes["code_departement"]))
    OUT_DIR = f"resultats_{ARG_CP}"
    os.makedirs(OUT_DIR, exist_ok=True)

    # -------------------------------------------------- 1. Parcelles des communes du code postal
    fichiers_parcelles = []
    for _, com in communes.iterrows():
        f = telecharger_commune(com["code_departement"], com["code_insee"],
                                "parcelles")
        if f:
            fichiers_parcelles.append(f)
    if not fichiers_parcelles:
        raise SystemExit("[!] Aucune donnée cadastrale récupérable pour ce code postal.")

    print("[...] Chargement des parcelles...")
    morceaux = []
    for f in fichiers_parcelles:
        gdf = gpd.read_file(f)
        if gdf.crs is None or gdf.crs.to_epsg() != 2154:
            gdf = gdf.to_crs(2154)
        morceaux.append(gdf)
    parcelles = pd.concat(morceaux, ignore_index=True)
    parcelles = reparer_geometries(parcelles)
    print(f"[OK] {len(parcelles):,} parcelles chargées.")

    parcelles["surface_m2"] = parcelles.geometry.area
    parcelles = parcelles[~parcelles.geometry.isna() & ~parcelles.geometry.is_empty]
    parcelles = parcelles[parcelles["surface_m2"] > 0]
    parcelles["code_commune"] = parcelles["id"].str[:5]

    # -------------------------------------------------- 2. Synthèse du code postal
    surf = parcelles["surface_m2"]
    nb_baties = 0
    emprise_totale = 0.0
    stats = (
        f"CODE POSTAL {ARG_CP} — {', '.join(communes['nom_commune'])}\n"
        f"(département{'s' if len(deps) > 1 else ''} {'/'.join(deps)}) "
        f"— CADASTRE ETALAB + DVF + PLU\n"
        f"{'=' * 55}\n"
        f"Nombre de parcelles      : {len(parcelles):>15,}\n"
        f"Surface totale           : {surf.sum() / 1e6:>13,.3f} km²\n"
        f"Taille moyenne           : {surf.mean():>13,.0f} m² "
        f"({surf.mean() / 10000:.2f} ha)\n"
        f"Taille médiane           : {surf.median():>13,.0f} m² "
        f"({surf.median() / 10000:.2f} ha)\n"
        f"Écart-type               : {surf.std():>13,.0f} m²\n"
        f"Plus petite parcelle    : {surf.min():>13,.0f} m²\n"
        f"Plus grande parcelle    : {surf.max():>13,.0f} m²\n"
    )
    print(stats)

    # -------------------------------------------------- 3. Histogrammes (code postal uniquement)
    seuil = surf.quantile(0.99)
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.hist(surf[surf <= seuil] / 10000, bins=60, color="#2e7d32", edgecolor="white")
    ax.set_xlabel("Surface de la parcelle (ha)")
    ax.set_ylabel("Nombre de parcelles")
    ax.set_title(f"Tailles des parcelles — code postal {ARG_CP}\n"
                 f"({', '.join(communes['nom_commune'][:3])}"
                 f"{'…' if len(communes) > 3 else ''} ; 99 % ≤ {seuil / 10000:.1f} ha)")
    ax.axvline(surf.mean() / 10000, color="#e65100", linestyle="--",
               label=f"Moyenne : {surf.mean() / 10000:.2f} ha")
    ax.axvline(surf.median() / 10000, color="#1565c0", linestyle="-.",
               label=f"Médiane : {surf.median() / 10000:.2f} ha")
    ax.legend()
    fig.tight_layout()
    fig.savefig(os.path.join(OUT_DIR, "histogramme_tailles.png"), dpi=150)
    plt.close(fig)

    fig2, ax2 = plt.subplots(figsize=(10, 6))
    # Bornes dynamiques : de la plus petite à la plus grande parcelle
    # (des bornes fixes coupaient les grandes parcelles en zone rurale).
    bins_log = 10 ** np.arange(np.floor(np.log10(surf.min())),
                               np.ceil(np.log10(surf.max())) + 0.1, 0.1)
    ax2.hist(surf, bins=bins_log, color="#2e7d32", edgecolor="white")
    ax2.set_xscale("log")
    ax2.set_xlabel("Surface de la parcelle (m², échelle log)")
    ax2.set_ylabel("Nombre de parcelles")
    ax2.set_title(f"Tailles des parcelles — code postal {ARG_CP}, échelle log")
    fig2.tight_layout()
    fig2.savefig(os.path.join(OUT_DIR, "histogramme_log.png"), dpi=150)
    plt.close(fig2)
    print("[OK] Histogrammes exportés.")

    # -------------------------------------------------- 4. Emprise bâtie par parcelle
    parcelles["surface_batie_m2"] = 0.0
    try:
        morceaux_b = []
        for _, com in communes.iterrows():
            f = telecharger_commune(com["code_departement"], com["code_insee"],
                                    "batiments")
            if f:
                morceaux_b.append(f)
        if morceaux_b:
            batiments = pd.concat([gpd.read_file(f) for f in morceaux_b],
                                  ignore_index=True)
            if batiments.crs is None or batiments.crs.to_epsg() != 2154:
                batiments = batiments.to_crs(2154)
            batiments = batiments[~batiments.geometry.isna()
                                  & ~batiments.geometry.is_empty]
            batiments = reparer_geometries(batiments)
            print(f"[OK] {len(batiments):,} bâtiments chargés.")

            print("[...] Attribution bâtiments -> parcelles...")
            batiments_k = batiments[["geometry"]].copy()
            batiments_k.insert(0, "id_batiment", range(len(batiments_k)))
            parcelles_k = parcelles[["id", "geometry"]].rename(
                columns={"id": "id_parcelle"})

            try:
                paires = gpd.sjoin(batiments_k, parcelles_k, predicate="intersects")
            except TypeError:
                paires = gpd.sjoin(batiments_k, parcelles_k, op="intersects")

            geo_bat = batiments_k.geometry
            geo_par = parcelles_k.set_index("id_parcelle").geometry

            # Intersection vectorisée (nettement plus rapide qu'un apply
            # ligne à ligne) : les deux GeoSeries sont recréées sur un index
            # positionnel (0..n) pour s'aligner paire par paire.
            inter_b = gpd.GeoSeries(
                geo_bat.loc[paires["id_batiment"]].to_numpy(),
                crs=batiments_k.crs).intersection(gpd.GeoSeries(
                    geo_par.loc[paires["id_parcelle"]].to_numpy(),
                    crs=batiments_k.crs))
            paires["emprise_m2"] = inter_b.area.to_numpy()
            paires = paires[paires["emprise_m2"] > 0]

            emprise = (paires.groupby("id_parcelle", as_index=False)
                       .agg(surface_batie_m2=("emprise_m2", "sum")))
            parcelles = parcelles.merge(emprise, left_on="id",
                                        right_on="id_parcelle", how="left")
            parcelles["surface_batie_m2"] = parcelles["surface_batie_m2_y"].fillna(0)
            parcelles = parcelles.drop(columns=["surface_batie_m2_x",
                                                "surface_batie_m2_y"],
                                       errors="ignore")
            nb_baties = len(emprise)
            emprise_totale = emprise["surface_batie_m2"].sum()
            print(f"[OK] {nb_baties:,} parcelles bâties identifiées.")
            stats += (f"{'=' * 55}\n"
                      f"EMPRISE AU SOL BÂTIE\n{'=' * 55}\n"
                      f"Emprise bâtie totale     : "
                      f"{emprise_totale / 1e6:>13,.3f} km²\n"
                      f"Part bâtie du code postal : "
                      f"{100 * emprise_totale / surf.sum():>11,.2f} %\n")
        else:
            print("[!] Données bâtiments indisponibles pour ce code postal.",
                  file=sys.stderr)
    except Exception:
        import traceback
        print(traceback.format_exc(), file=sys.stderr)
        print("[!] Analyse du bâti ignorée.", file=sys.stderr)

    # -------------------------------------------------- 5. Prix au m² (DVF)
    tableau = parcelles[["id", "code_commune", "surface_m2",
                        "surface_batie_m2", "geometry"]].copy()
    tableau["nb_ventes"] = float("nan")
    tableau["prix_m2_bati"] = float("nan")
    dvf = charger_dvf(deps)
    if dvf is not None and len(dvf):
        # Identifiants de parcelles : 14 caractères concaténés des deux
        # côtés (commune 5 + préfixe 3 + section 2 + numéro 4). La
        # normalisation ci-dessous est conservée par sécurité.
        tableau = tableau.drop(columns=["nb_ventes"], errors="ignore")
        tableau["id_dvf"] = (tableau["id"].str.replace("_", "", regex=False)
                             .str.zfill(14))
        # Prix au m² BÂTI : valeur foncière ÷ surface bâtie vendue (DVF).
        # Plus parlant que valeur ÷ surface de parcelle. Pour les ventes de
        # terrain nu (surface bâtie = 0 ou absente), le prix au m² bâti n'a
        # pas de sens : ces ventes sont exclues, la colonne reste vide.
        prix = (dvf.dropna(subset=["surface_reelle_bati"])
               .query("surface_reelle_bati > 0")
               .groupby("id_parcelle", as_index=False)
               .agg(nb_ventes=("valeur_fonciere", "size"),
                    valeur_fonciere_totale=("valeur_fonciere", "sum"),
                    surface_batie_vendue_m2=("surface_reelle_bati", "sum")))
        tableau = tableau.merge(prix, left_on="id_dvf", right_on="id_parcelle",
                                how="left")
        tableau = tableau.drop(columns=["id_parcelle", "id_dvf"], errors="ignore")
        tableau["prix_m2_bati"] = (tableau["valeur_fonciere_totale"]
                                  / tableau["surface_batie_vendue_m2"]).round(1)
        tableau["prix_m2_bati"] = tableau["prix_m2_bati"].where(
            tableau["nb_ventes"].notna())
        nb_match = int(tableau["nb_ventes"].notna().sum())
        print(f"[i] Croisement DVF : {nb_match:,} parcelles avec vente(s) "
              f"trouvée(s) sur {len(tableau):,}.")
        stats += (f"{'=' * 55}\n"
                  f"VALEURS FONCIÈRES (DVF {DVF_ANNEE})\n{'=' * 55}\n"
                  f"Parcelles avec vente bâtie : "
                  f"{int(tableau['prix_m2_bati'].notna().sum()):>13,}\n"
                  f"Prix moyen au m² bâti      : "
                  f"{tableau['prix_m2_bati'].mean():>13,.0f} €/m²\n"
                  f"Prix médian au m² bâti      : "
                  f"{tableau['prix_m2_bati'].median():>13,.0f} €/m²\n")

    # -------------------------------------------------- 5b. PLU : zones constructibles + emprise au sol
    print("[...] Interrogation du Géoportail de l'urbanisme (PLU)...")
    try:
        partitions = {f"DU_{c}" for c in codes_insee}
        minx, miny, maxx, maxy = parcelles.to_crs(4326).total_bounds
        marge = 0.02  # ~2 km : récupère aussi les zones des communes limitrophes
        cql = (f"BBOX(the_geom,{minx - marge},{miny - marge},"
               f"{maxx + marge},{maxy + marge},'EPSG:4326') "
               f"AND partition LIKE 'DU_%'")

        zones = requete_wfs_gpu(GPU_ZONE_URBA, cql)
        if zones is None or not len(zones):
            raise RuntimeError("aucune zone d'urbanisme reçue")
        for col in ("libelle", "typezone", "geometry"):
            if col not in zones.columns:
                raise RuntimeError(f"colonne '{col}' absente du zonage PLU")
        zones = zones[zones["partition"].isin(partitions)]
        if not len(zones):
            raise RuntimeError("aucune zone d'urbanisme pour ce code postal")
        # Zones constructibles : U (urbaines) et AU (à urbaniser)
        zones["constructible"] = zones["typezone"].astype(str).str.upper() \
            .str.startswith(("U", "AU"))
        print(f"[OK] {len(zones):,} zones PLU reçues "
              f"({int(zones['constructible'].sum()):,} constructibles U/AU).")

        # Communes du code postal sans zonage reçu (document non publié
        # ou hors périmètre) : leurs parcelles resteront « non constructibles ».
        parts_recues = set(zones["partition"].astype(str).unique())
        sans_zonage = [c for c in communes["code_insee"]
                       if f"DU_{c}" not in parts_recues]
        if sans_zonage:
            noms = ", ".join(communes.loc[communes["code_insee"]
                                           .isin(sans_zonage), "nom_commune"])
            print(f"[!] Communes sans zonage PLU reçu : {noms}",
                  file=sys.stderr)

        # Documents d'urbanisme en vigueur : un code postal peut couvrir
        # plusieurs communes, donc plusieurs PLU (un gpu_doc_id par
        # document). Le taux d'emprise au sol est extrait du règlement
        # écrit (PDF) de CHACUN d'eux (boucle plus bas).
        if "gpu_doc_id" in zones.columns:
            gpu_doc_ids = [str(d).strip() for d in
                           zones["gpu_doc_id"].dropna().unique().tolist()
                           if str(d).strip()]
        else:
            gpu_doc_ids = []
        nb_docs = len(gpu_doc_ids)
        # Document le plus représenté (affiché dans la synthèse)
        try:
            idurba, datvalid, _ = \
                zones.groupby(["idurba", "datvalid", "gpu_doc_id"]) \
                .size().idxmax()
        except (KeyError, ValueError):
            idurba, datvalid = "inconnu", ""
        if isinstance(datvalid, str) and len(datvalid) == 8:
            datvalid = f"{datvalid[:4]}-{datvalid[4:6]}-{datvalid[6:]}"

        # Jointure spatiale parcelles x zones : la zone retenue est la
        # zone dominante (plus grande intersection). Les deux couches sont
        # en EPSG:2154 (mètres).
        par_k = parcelles[["id", "geometry"]].rename(columns={"id": "id_parcelle"})
        cols_z = ["libelle", "typezone"]
        if "libelong" in zones.columns:
            cols_z.append("libelong")
        # gpu_doc_id : document d'urbanisme de la zone — indispensable quand
        # le code postal couvre plusieurs PLU (un règlement par document).
        for col in ("gpu_doc_id", "partition"):
            if col in zones.columns:
                cols_z.append(col)
        zr = zones[cols_z + ["constructible", "geometry"]].copy()
        zr.insert(0, "id_zone", range(len(zr)))
        # reset_index : une parcelle intersectant plusieurs zones apparaît
        # plusieurs fois avec le même index ; l'index doit rester unique
        # pour la sélection de la zone dominante (idxmax + .loc).
        paires_z = gpd.sjoin(par_k, zr, predicate="intersects") \
            .reset_index(drop=True)
        inter_z = gpd.GeoSeries(
            par_k.set_index("id_parcelle").geometry
            .loc[paires_z["id_parcelle"]].to_numpy(), crs=parcelles.crs
        ).intersection(gpd.GeoSeries(
            zr.set_index("id_zone").geometry
            .loc[paires_z["id_zone"]].to_numpy(), crs=parcelles.crs))
        paires_z["surface_zone_m2"] = inter_z.area.to_numpy()
        paires_z = paires_z[paires_z["surface_zone_m2"] > 0]
        # Surface totale en zone constructible (une parcelle peut être à cheval)
        surf_zc = (paires_z[paires_z["constructible"]]
                   .groupby("id_parcelle")["surface_zone_m2"].sum())
        dom_z = (paires_z.loc[paires_z.groupby("id_parcelle")
                             ["surface_zone_m2"].idxmax()]
                 .set_index("id_parcelle"))

        tableau["code_zone"] = tableau["id"].map(dom_z["libelle"])
        tableau["typezone"] = tableau["id"].map(dom_z["typezone"])
        if "gpu_doc_id" in dom_z.columns:
            tableau["gpu_doc_id"] = tableau["id"].map(dom_z["gpu_doc_id"])
        else:
            tableau["gpu_doc_id"] = float("nan")
        if "libelong" in dom_z.columns:
            tableau["libelle_zone"] = tableau["id"].map(dom_z["libelong"])
        tableau["surface_en_zone_m2"] = tableau["id"].map(
            dom_z["surface_zone_m2"]).round(0)
        tableau["en_zone_constructible"] = tableau["id"].map(
            dom_z["constructible"]).map({True: "Oui", False: "Non"}) \
            .fillna("Non")
        tableau["surface_en_zone_constructible_m2"] = tableau["id"].map(
            surf_zc).fillna(0).round(0)

        nb_en_zc = int((tableau["en_zone_constructible"] == "Oui").sum())
        print(f"[OK] {nb_en_zc:,} parcelles en zone constructible "
              f"(U/AU) sur {len(tableau):,}.")

        # Taux d'emprise au sol : d'abord les prescriptions surfaciques
        # structurées du WFS (sous-secteurs), sinon le règlement écrit (PDF).
        tableau["taux_emprise_pct"] = float("nan")
        source_taux = "aucune"
        presc = requete_wfs_gpu(GPU_PRESCRIPTION_SURF, cql)
        if presc is not None and len(presc):
            presc = presc[presc["partition"].isin(partitions)]
            lib = presc["libelle"].fillna("").astype(str)
            if "txt" in presc.columns:
                lib = lib + " " + presc["txt"].fillna("").astype(str)
            masque = (lib.str.contains(r"emprise|C\.?\s*E\.?\s*S\.?",
                                       regex=True, case=False) &
                      ~lib.str.contains(r"protection|[eé]largissement|"
                                        r"cheminement|voirie|monument|"
                                        r"hors emprise", regex=True,
                                        case=False))
            presc_e = presc[masque].copy()
            presc_e["taux"] = (presc_e["libelle"].fillna("").astype(str) +
                               (" " + presc_e["txt"].fillna("").astype(str)
                                if "txt" in presc_e.columns else "")
                               ).apply(_taux_emprise)
            presc_e = presc_e[presc_e["taux"].notna()]
            if len(presc_e):
                zr_e = presc_e[["taux", "geometry"]].copy()
                zr_e.insert(0, "id_psc", range(len(zr_e)))
                paires_p = gpd.sjoin(par_k, zr_e, predicate="intersects") \
                    .reset_index(drop=True)
                inter_p = gpd.GeoSeries(
                    par_k.set_index("id_parcelle").geometry
                    .loc[paires_p["id_parcelle"]].to_numpy(),
                    crs=parcelles.crs
                ).intersection(gpd.GeoSeries(
                    zr_e.set_index("id_psc").geometry
                    .loc[paires_p["id_psc"]].to_numpy(), crs=parcelles.crs))
                paires_p["surface_psc_m2"] = inter_p.area.to_numpy()
                paires_p = paires_p[paires_p["surface_psc_m2"] > 0]
                dom_p = (paires_p.loc[paires_p.groupby("id_parcelle")
                                     ["surface_psc_m2"].idxmax()]
                         .set_index("id_parcelle"))
                tableau["taux_emprise_pct"] = tableau["id"].map(dom_p["taux"])
                source_taux = "prescriptions PLU (WFS)"
                print(f"[OK] Taux d'emprise structurés : "
                      f"{int(tableau['taux_emprise_pct'].notna().sum()):,} "
                      f"parcelles couvertes.")

        # Compléter par le règlement écrit (PDF) : les prescriptions WFS
        # ne couvrent pas toujours toutes les zones, et un code postal peut
        # rassembler plusieurs PLU — on lit le règlement de chaque document
        # pour les parcelles qui n'ont pas encore de taux.
        nb_avant_pdf = int(tableau["taux_emprise_pct"].notna().sum())
        for doc in gpu_doc_ids:
            reste = (tableau["taux_emprise_pct"].isna() &
                     (tableau["gpu_doc_id"] == doc))
            if not reste.any():
                continue
            pdf = telecharger_reglement_gpu(doc)
            articles = extraire_taux_emprise_reglement(pdf) if pdf else None
            if not articles:
                continue
            sel = tableau[reste]
            taux_pdf = pd.Series(
                [taux_emprise_zone(z, s, articles)
                 for z, s in zip(sel["code_zone"], sel["surface_m2"])],
                index=sel.index, dtype="float64")
            tableau.loc[reste, "taux_emprise_pct"] = taux_pdf
            print(f"[OK] Règlement {doc[:8]}… : "
                  f"{int(taux_pdf.notna().sum()):,} taux extraits "
                  f"({int(reste.sum()):,} parcelles sans prescription WFS).")
        if int(tableau["taux_emprise_pct"].notna().sum()) > nb_avant_pdf:
            source_taux = ("règlement écrit (PDF)"
                           if source_taux == "aucune"
                           else source_taux + " + règlement écrit (PDF)")

        # Taux théorique par défaut : 35 % pour les parcelles en zone
        # constructible dont aucune règle d'emprise au sol n'a été trouvée
        # (la colonne « taux_emprise_source » indique l'origine du taux).
        tableau["taux_emprise_source"] = ""
        tableau.loc[tableau["taux_emprise_pct"].notna(),
                    "taux_emprise_source"] = "PLU"
        masque_theo = (tableau["taux_emprise_pct"].isna() &
                       (tableau["en_zone_constructible"] == "Oui"))
        if masque_theo.any():
            tableau.loc[masque_theo, "taux_emprise_pct"] = 35.0
            tableau.loc[masque_theo, "taux_emprise_source"] = \
                "théorique (35 % par défaut)"
            print(f"[i] {int(masque_theo.sum()):,} parcelles en zone "
                  f"constructible sans règle d'emprise : taux théorique "
                  f"de 35 % appliqué.")

        # Emprise au sol maximale = taux x surface de la parcelle.
        # Hypothèse : la base de calcul est la parcelle cadastrale entière
        # (le règlement raisonne en « unité foncière », le plus souvent
        # confondue avec la parcelle cadastrale).
        # La colonne affichée est arrondie, mais le test d'atteinte se fait
        # sur la valeur exacte (évite les faux « atteints » sur les
        # parcelles minuscules où l'arrondi donnait un maximum de 0).
        emprise_max_calc = (tableau["taux_emprise_pct"] / 100
                            * tableau["surface_m2"])
        tableau["emprise_max_m2"] = emprise_max_calc.round(1)
        bati = tableau["surface_batie_m2"]
        avec_taux = (tableau["taux_emprise_pct"].notna() &
                     (tableau["en_zone_constructible"] == "Oui"))
        atteinte = pd.Series("", index=tableau.index, dtype=object)
        atteinte[avec_taux & (bati > 0) & (bati >= emprise_max_calc)] = "Oui"
        atteinte[avec_taux & (bati < emprise_max_calc)] = "Non"
        tableau["emprise_max_atteinte"] = atteinte

        nb_bati_zc = int(((tableau["en_zone_constructible"] == "Oui") &
                          (bati > 0)).sum())
        nb_att = int((tableau["emprise_max_atteinte"] == "Oui").sum())
        nb_taux = int(tableau["taux_emprise_pct"].notna().sum())
        nb_theo = int(tableau["taux_emprise_source"]
                      .str.startswith("théorique").sum())
        surf_tot_zones = float(zones.geometry.area.sum())
        surf_zc_tot = float(zones.loc[zones["constructible"], "geometry"]
                            .area.sum())
        lib_nb_docs = f"{nb_docs:,}" if nb_docs else "inconnu"
        bloc_plu = (
            f"{'=' * 55}\n"
            f"PLU — ZONES CONSTRUCTIBLES (Géoportail de l'urbanisme)\n"
            f"{'=' * 55}\n"
            f"Documents d'urbanisme    : {lib_nb_docs:>15}\n")
        if idurba != "inconnu":
            bloc_plu += (f"  doc. le plus fréquent  : {idurba} "
                         f"(approuvé le {datvalid})\n")
        bloc_plu += (
            f"Surface zones U + AU     : {surf_zc_tot / 1e6:>13,.3f} km²"
            f"  ({100 * surf_zc_tot / max(surf_tot_zones, 1):>6.2f} %"
            f" du territoire)\n"
            f"Parcelles en zone constr.: {nb_en_zc:>15,}\n"
            f"  dont déjà bâties       : {nb_bati_zc:>15,}\n"
            f"Taux d'emprise au sol    : {source_taux}\n"
            f"Parcelles avec taux connu: {nb_taux:>15,}\n"
            f"  dont taux théoriques   : {nb_theo:>15,}\n"
            f"Emprise max atteinte     : {nb_att:>15,}\n")
        stats += bloc_plu
        print(bloc_plu)

        # Tableau constructibilité (zonage PLU + emprise au sol max)
        tab_plu = tableau.merge(communes, left_on="code_commune",
                                right_on="code_insee", how="left")
        tab_plu["numero_parcelle"] = tab_plu["id"].str[5:]
        tab_plu["taux_bati_pct"] = (100 * tab_plu["surface_batie_m2"]
                                    / tab_plu["surface_m2"]).round(1)
        cols_plu = ["code_postal", "nom_commune", "numero_parcelle",
                    "surface_m2", "code_zone", "typezone", "libelle_zone",
                    "en_zone_constructible", "surface_en_zone_m2",
                    "surface_en_zone_constructible_m2", "surface_batie_m2",
                    "taux_bati_pct", "taux_emprise_pct", "taux_emprise_source",
                    "emprise_max_m2", "emprise_max_atteinte", "prix_m2_bati",
                    "nb_ventes", "id"]
        vue_plu = tab_plu[[c for c in cols_plu if c in tab_plu.columns]] \
            .rename(columns={"id": "id_parcelle"}) \
            .sort_values(["nom_commune", "numero_parcelle"])
        zones_tab = zones.copy()
        zones_tab["surface_m2"] = zones_tab.geometry.area.round(0)
        if "libelong" not in zones_tab.columns:
            zones_tab["libelong"] = zones_tab["typezone"]
        # partition + gpu_doc_id : un code postal peut couvrir plusieurs
        # PLU ; les inclure dans les clés évite de fusionner des zones
        # homonymes appartenant à des règlements différents.
        cles_zones = [c for c in ("partition", "gpu_doc_id")
                      if c in zones_tab.columns]
        cles_zones += ["libelle", "typezone"]
        agg_zones = {"libelle_zone": ("libelong", "first"),
                     "constructible": ("constructible", "first"),
                     "surface_m2": ("surface_m2", "sum"),
                     "nb_secteurs": ("libelle", "size")}
        for col in ("partition", "gpu_doc_id"):
            if col in zones_tab.columns:
                agg_zones[col] = (col, "first")
        zones_tab = (zones_tab.groupby(cles_zones, as_index=False)
                     .agg(**agg_zones)
                     .rename(columns={"libelle": "code_zone"}))
        zones_tab["constructible"] = zones_tab["constructible"] \
            .map({True: "Oui", False: "Non"})
        zones_tab = zones_tab.sort_values(["constructible", "code_zone"],
                                          ascending=[False, True])
        xlsx_plu = os.path.join(OUT_DIR, "tableau_constructibilite.xlsx")
        try:
            with pd.ExcelWriter(xlsx_plu, engine="openpyxl") as writer:
                (vue_plu[vue_plu["en_zone_constructible"] == "Oui"]
                 .to_excel(writer, sheet_name="En zone constructible",
                           index=False))
                vue_plu.to_excel(writer, sheet_name="Toutes les parcelles",
                                 index=False)
                zones_tab.to_excel(writer, sheet_name="Zones du PLU",
                                   index=False)
            print(f"[OK] Excel constructibilité : {xlsx_plu}")
        except Exception as exc:
            csv_plu = os.path.join(OUT_DIR, "tableau_constructibilite.csv")
            vue_plu.to_csv(csv_plu, index=False, sep=";")
            print(f"[!] Excel impossible ({exc}), CSV exporté : {csv_plu}",
                  file=sys.stderr)

        # Carte des zones constructibles (fond IGN, comme la carte parcelles)
        try:
            import folium
            print("[...] Carte des zones constructibles...")
            tab_carte = tab_plu.to_crs(4326)
            zones_carte = zones.to_crs(4326)
            bounds = tab_carte.total_bounds
            carte_z = folium.Map(
                location=[(bounds[1] + bounds[3]) / 2,
                          (bounds[0] + bounds[2]) / 2],
                zoom_start=14, tiles=None)
            carte_z.fit_bounds([[bounds[1], bounds[0]],
                                [bounds[3], bounds[2]]])
            folium.TileLayer(
                tiles=("https://data.geopf.fr/wmts?service=WMTS&request=GetTile"
                       "&version=1.0.0&layer=GEOGRAPHICALGRIDSYSTEMS.PLANIGNV2"
                       "&style=normal&tilematrixset=PM&format=image/png"
                       "&tilematrix={z}&tilerow={y}&tilecol={x}"),
                attr="IGN", name="Plan IGN", max_native_zoom=18,
                control=False).add_to(carte_z)

            def _style_zone(feat):
                p = feat["properties"]
                if str(p.get("typezone") or "").upper() \
                        .startswith(("U", "AU")):
                    return {"fillColor": "#ff9800", "color": "#e65100",
                            "weight": 1, "fillOpacity": 0.30}
                return {"fillColor": "#9e9e9e", "color": "#616161",
                        "weight": 1, "fillOpacity": 0.15}

            def _style_parcelle_zc(feat):
                p = feat["properties"]
                if p.get("emprise_max_atteinte") == "Oui":
                    return {"fillColor": "#c62828", "color": "#7f0000",
                            "weight": 1, "fillOpacity": 0.65}
                if p.get("en_zone_constructible") == "Oui":
                    if (p.get("surface_batie_m2") or 0) > 0:
                        return {"fillColor": "#ef6c00", "color": "#e65100",
                                "weight": 0.5, "fillOpacity": 0.55}
                    return {"fillColor": "#2e7d32", "color": "#1b5e20",
                            "weight": 0.5, "fillOpacity": 0.45}
                return {"fillColor": "#bdbdbd", "color": "#757575",
                        "weight": 0.5, "fillOpacity": 0.35}

            legende_z = (
                '<div style="position: fixed; bottom: 30px; left: 30px;'
                ' z-index: 9999; background: white; padding: 10px 14px;'
                ' border: 2px solid grey; border-radius: 5px;'
                ' font-family: sans-serif; font-size: 13px;">'
                '<b>Zones constructibles — code postal ' + str(ARG_CP) +
                '</b><br>'
                '<span style="color:#ff9800;">&#9632;</span> Zone U / AU '
                '(constructible)<br>'
                '<span style="color:#9e9e9e;">&#9632;</span> Zone N '
                '(non constructible)<br>'
                '<span style="color:#c62828;">&#9632;</span> Parcelle : '
                'emprise max atteinte<br>'
                '<span style="color:#ef6c00;">&#9632;</span> Parcelle bâtie '
                '(sous le max)<br>'
                '<span style="color:#2e7d32;">&#9632;</span> Parcelle non '
                'bâtie (constructible)<br>'
                '<span style="color:#bdbdbd;">&#9632;</span> Hors zone '
                'constructible</div>')
            carte_z.get_root().html.add_child(folium.Element(legende_z))

            folium.GeoJson(
                json.loads(zones_carte.to_json(na="null")),
                name="Zones PLU",
                style_function=_style_zone,
                tooltip=folium.GeoJsonTooltip(
                    fields=[c for c in ["libelle", "typezone", "libelong"]
                            if c in zones_carte.columns],
                    aliases=["Zone", "Type", "Destination"],
                    localize=True, sticky=False),
            ).add_to(carte_z)
            folium.GeoJson(
                json.loads(tab_carte.to_json(na="null")),
                name="Parcelles",
                style_function=_style_parcelle_zc,
                tooltip=folium.GeoJsonTooltip(
                    fields=[c for c in ["nom_commune", "numero_parcelle",
                                        "code_zone", "surface_m2",
                                        "surface_batie_m2",
                                        "taux_emprise_pct", "emprise_max_m2",
                                        "emprise_max_atteinte"]
                            if c in tab_carte.columns],
                    aliases=["Commune", "Parcelle n°", "Zone",
                             "Surface parcelle (m²)",
                             "Surface bâtie (m²)", "Taux emprise max (%)",
                             "Emprise max (m²)", "Emprise max atteinte"],
                    localize=True, sticky=False),
            ).add_to(carte_z)
            folium.LayerControl().add_to(carte_z)
            carte_z.save(os.path.join(OUT_DIR,
                                      "carte_zones_constructibles.html"))
            print(f"[OK] Carte zones constructibles : "
                  f"{OUT_DIR}/carte_zones_constructibles.html")
        except Exception:
            import traceback
            print(traceback.format_exc(), file=sys.stderr)
            print("[!] Carte des zones ignorée (folium installé ?).",
                  file=sys.stderr)
    except Exception:
        import traceback
        print(traceback.format_exc(), file=sys.stderr)
        print("[!] Analyse PLU ignorée (Géoportail de l'urbanisme "
              "indisponible ?).", file=sys.stderr)

    # -------------------------------------------------- 6. Tableau final + Excel
    tableau = tableau.merge(communes, left_on="code_commune",
                            right_on="code_insee", how="left")
    tableau["numero_parcelle"] = tableau["id"].str[5:]
    # Pourcentage de surface bâtie par rapport à la surface de la parcelle
    tableau["taux_bati_pct"] = (100 * tableau["surface_batie_m2"]
                                / tableau["surface_m2"]).round(1)
    tableau = tableau.rename(columns={"id": "id_parcelle"})

    colonnes = ["code_postal", "nom_commune", "numero_parcelle", "surface_m2",
                "surface_batie_m2", "taux_bati_pct", "prix_m2_bati",
                "nb_ventes", "code_zone", "libelle_zone",
                "en_zone_constructible", "surface_en_zone_m2",
                "taux_emprise_pct", "emprise_max_m2", "emprise_max_atteinte",
                "id_parcelle"]
    export = tableau[[c for c in colonnes if c in tableau.columns] +
                     ["geometry"]].sort_values(["code_postal", "nom_commune",
                                                "numero_parcelle"])

    xlsx_path = os.path.join(OUT_DIR, "tableau_parcelles.xlsx")
    print("[...] Export Excel...")
    try:
        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
            sans_geom = export.drop(columns="geometry")
            vendues = sans_geom[sans_geom["prix_m2_bati"].notna()] \
                if "prix_m2_bati" in sans_geom.columns else sans_geom.iloc[0:0]
            vendues.to_excel(writer, sheet_name="Parcelles vendues (DVF)", index=False)
            sans_geom.to_excel(writer, sheet_name="Toutes les parcelles", index=False)
            # Onglet par commune
            for _, com in communes.iterrows():
                onglet = sans_geom[sans_geom["code_postal"] == com["code_postal"]]
                # Un onglet par commune (utile si plusieurs communes)
                onglet = sans_geom[sans_geom.get("nom_commune", pd.Series()) ==
                                   com["nom_commune"]] if "nom_commune" in sans_geom \
                    else onglet
                nom_feuille = str(com["nom_commune"])[:28]
                if len(onglet) and nom_feuille not in writer.sheets:
                    onglet.to_excel(writer, sheet_name=nom_feuille, index=False)
        print(f"[OK] Excel exporté : {xlsx_path}")
    except Exception as exc:
        csv_path = os.path.join(OUT_DIR, "tableau_parcelles.csv")
        export.drop(columns="geometry").to_csv(csv_path, index=False, sep=";")
        print(f"[!] Excel impossible ({exc}), CSV exporté : {csv_path}",
              file=sys.stderr)
        print("    -> Pour l'Excel : pip install openpyxl", file=sys.stderr)

    # -------------------------------------------------- 7. Stats par commune du code postal
    par_commune = (export.drop(columns="geometry")
                  .groupby(["code_postal", "nom_commune"], as_index=False)
                  .agg(nb_parcelles=("surface_m2", "size"),
                       surface_totale_m2=("surface_m2", "sum"),
                       surface_moyenne_m2=("surface_m2", "mean"),
                       surface_mediane_m2=("surface_m2", "median"),
                       surface_batie_totale_m2=("surface_batie_m2", "sum")))
    par_commune["taux_bati_pct"] = (100 * par_commune["surface_batie_totale_m2"]
                                    / par_commune["surface_totale_m2"])
    par_commune.to_csv(os.path.join(OUT_DIR, "stats_par_communes.csv"), index=False)

    # Graphique emprise bâtie par commune (si plusieurs communes)
    if len(par_commune) > 1:
        fig3, ax3 = plt.subplots(figsize=(10, 6))
        ax3.barh(par_commune["nom_commune"][::-1],
                 par_commune["surface_batie_totale_m2"][::-1] / 1e4, color="#e65100")
        ax3.set_xlabel("Emprise au sol bâtie (ha)")
        ax3.set_title(f"Surface bâtie par commune — code postal {ARG_CP}")
        fig3.tight_layout()
        fig3.savefig(os.path.join(OUT_DIR, "emprise_batie_communes.png"), dpi=150)
        plt.close(fig3)
        print("[OK] Graphique d'emprise bâtie exporté.")

    # -------------------------------------------------- 8. Carte interactive
    try:
        import folium
        print("[...] Carte interactive...")
        # Folium/Leaflet attend des coordonnées WGS84 (lat/lon) : les
        # géométries cadastrales sont en Lambert-93 (mètres), il faut donc
        # reprojeter avant l'affichage, sans quoi la carte est centrée
        # n'importe où et les polygones restent invisibles.
        carte_gdf = export.to_crs(4326).copy()
        bounds = carte_gdf.total_bounds  # [minx, miny, maxx, maxy] en WGS84
        centre_lat = (bounds[1] + bounds[3]) / 2
        centre_lon = (bounds[0] + bounds[2]) / 2
        # Fond Plan IGN (Géoportail) : gratuit, sans clé, sans blocage.
        # OSM bloque certaines applications, CartoDB exige une clé, et le
        # fond Esri "Light Gray" ne renvoie pas de tuiles ("map data not
        # yet available"). Le Plan IGN est le fond officiel français.
        carte = folium.Map(location=[centre_lat, centre_lon], zoom_start=14,
                           tiles=None)
        # Cadrage automatique sur l'étendue exacte des parcelles du code postal
        carte.fit_bounds([[bounds[1], bounds[0]], [bounds[3], bounds[2]]])
        folium.TileLayer(
            tiles=("https://data.geopf.fr/wmts?service=WMTS&request=GetTile"
                   "&version=1.0.0&layer=GEOGRAPHICALGRIDSYSTEMS.PLANIGNV2"
                   "&style=normal&tilematrixset=PM&format=image/png"
                   "&tilematrix={z}&tilerow={y}&tilecol={x}"),
            attr="IGN", name="Plan IGN", max_native_zoom=18,
            control=False).add_to(carte)

        def _style(feat):
            p = feat["properties"]
            bati = p.get("surface_batie_m2") or 0
            prix = p.get("prix_m2_bati")
            # Trois couleurs : bleu = parcelle vendue (prix au m² bâti connu),
            # orange = bâtie sans vente, vert = non bâtie.
            if prix is not None:
                couleur = "#1565c0"
            elif bati > 0:
                couleur = "#e65100"
            else:
                couleur = "#2e7d32"
            return {"fillColor": couleur, "color": "#333333",
                    "weight": 0.5, "fillOpacity": 0.5}

        # Légende des couleurs
        legende_html = (
            '<div style="position: fixed; bottom: 30px; left: 30px; z-index: 9999;'
            ' background: white; padding: 10px 14px; border: 2px solid grey;'
            ' border-radius: 5px; font-family: sans-serif; font-size: 13px;">'
            '<b>Parcelles — code postal ' + str(ARG_CP) + '</b><br>'
            '<span style="color:#1565c0;">&#9632;</span> Vendue '
            '(prix au m² bâti)<br>'
            '<span style="color:#e65100;">&#9632;</span> Bâtie (sans vente)<br>'
            '<span style="color:#2e7d32;">&#9632;</span> Non bâtie'
            '</div>')
        carte.get_root().html.add_child(folium.Element(legende_html))

        # Nettoyage des valeurs manquantes : les parcelles sans vente DVF
        # ont prix/nb_ventes à NaN, que GeoPandas sérialise par défaut en
        # "NaN" — du JSON invalide que le navigateur rejette en silence,
        # ce qui fait disparaître TOUTE la couche de parcelles.
        # na="null" sérialise correctement les valeurs manquantes.
        # Diagnostic : la couche GeoJSON est écrite dans un fichier séparé
        # pour vérifier son contenu en cas de problème d'affichage.
        geojson_str = carte_gdf.to_json(na="null")
        geojson_data = json.loads(geojson_str)
        print(f"[i] Couche parcelles : {len(geojson_data['features']):,} "
              f"entités GeoJSON générées.")
        if not geojson_data["features"]:
            raise RuntimeError("Aucune entité GeoJSON générée !")

        folium.GeoJson(
            geojson_data,
            style_function=_style,
            tooltip=folium.GeoJsonTooltip(
                fields=[c for c in ["code_postal", "nom_commune", "numero_parcelle",
                                   "surface_m2", "surface_batie_m2",
                                   "taux_bati_pct", "prix_m2_bati"]
                       if c in carte_gdf.columns],
                aliases=["Code postal", "Commune", "Parcelle n°",
                         "Surface parcelle (m²)", "Surface bâtie (m²)",
                         "Taux bâti (%)", "Prix au m² bâti (€, DVF)"],
                localize=True, sticky=False),
        ).add_to(carte)
        # Sauvegarde du GeoJSON brut à côté de la carte (diagnostic)
        with open(os.path.join(OUT_DIR, "parcelles.geojson"), "w",
                  encoding="utf-8") as f:
            f.write(geojson_str)
        carte.save(os.path.join(OUT_DIR, "carte_parcelles.html"))
        print(f"[OK] Carte interactive : {OUT_DIR}/carte_parcelles.html "
              f"({len(carte_gdf):,} parcelles)")
    except Exception:
        import traceback
        print(traceback.format_exc(), file=sys.stderr)
        print("[!] Carte ignorée (folium installé ?).", file=sys.stderr)

    with open(os.path.join(OUT_DIR, "stats_synthese.txt"), "w", encoding="utf-8") as f:
        f.write(stats)
    print(stats)
    print(f"[OK] Terminé. Résultats dans ./{OUT_DIR}/")


if __name__ == "__main__":
    main()
    main()