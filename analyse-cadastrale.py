# -*- coding: utf-8 -*-
"""
Analyse cadastrale par CODE POSTAL (France entière)
===================================================
Toute l'analyse (stats, histogrammes, tableaux, carte) est limitée aux
communes correspondant au code postal fourni — et non au département entier.

1. Résout le code postal via geo.api.gouv.fr (noms, codes INSEE, département)
2. Télécharge les parcelles et bâtiments du Cadastre Etalab pour ces communes
   uniquement (fichiers par commune : rapides)
3. Croise avec les Demandes de Valeurs Foncières (DVF) pour les prix au m²
4. Exports (dossier ./resultats_<code_postal>/) :
   - stats_synthese.txt        : synthèse (nombre, moyenne, médiane, bâti, prix)
   - tableau_parcelles.xlsx    : tableau détaillé par parcelle (onglets)
   - stats_par_communes.csv    : stats par commune du code postal
   - histogramme_tailles.png   : distribution des tailles
   - histogramme_log.png       : distribution (échelle log)
   - emprise_batie_communes.png : surfaces bâties par commune
   - carte_parcelles.html      : carte interactive des parcelles

Prérequis :
    pip install geopandas matplotlib folium openpyxl

Usage :
    python analyse_cadastre_aube.py 10000      # Troyes
    python analyse_cadastre_aube.py 75001     # Paris 1er
    Sans argument : 10000 (Troyes) par défaut.
    ⚠ Argument = CODE POSTAL uniquement (pas le code INSEE) : un code INSEE
    (ex. 10387) est rejeté avec un message d'aide.
"""

import gzip
import io
import json
import os
import shutil
import sys
import urllib.parse
import urllib.request
import zipfile

import geopandas as gpd
import matplotlib.pyplot as plt
import pandas as pd

BASE_URL = "https://cadastre.data.gouv.fr/data/etalab-cadastre/latest/geojson/communes"

# DVF : demandes de valeurs foncières géolocalisées (data.gouv.fr).
# Fichier unique couvrant janvier 2021 - décembre 2025 (CSV gzippé).
DVF_URL = ("https://static.data.gouv.fr/resources/"
           "demandes-de-valeurs-foncieres-geolocalisees/"
           "20260717-142041/dvf.csv.gz")
# ⚠ L'identifiant de ressource DVF change à chaque millésime : si le
# téléchargement échoue, récupérer le lien le plus récent sur
# https://www.data.gouv.fr/fr/datasets/demandes-de-valeurs-foncieres-geolocalisees/
DVF_ANNEE = "2021-2025"

ARG_CP = sys.argv[1] if len(sys.argv) > 1 else "10000"


def resoudre_code_postal(cp):
    """Résout un CODE POSTAL via geo.api.gouv.fr (codes INSEE refusés).
    Renvoie (departement, codes_insee, df_communes) ; df_communes contient
    code_insee, nom_commune, code_postal pour chaque commune du code postal."""
    cp = str(cp).strip()
    if not cp.isdigit():
        raise SystemExit(f"[!] {cp} : merci d'indiquer un CODE POSTAL à 5 chiffres "
                         f"(ex. 10000 pour Troyes).")

    url = (f"https://geo.api.gouv.fr/communes?codePostal={urllib.parse.quote(cp)}"
           f"&fields=code,nom,codeDepartement&format=json")
    with urllib.request.urlopen(url, timeout=30) as rep:
        communes = json.load(rep)
    if not communes:
        # Le paramètre répondrait aussi à un code INSEE : vérifions ce cas
        url2 = (f"https://geo.api.gouv.fr/communes?code={urllib.parse.quote(cp)}"
                f"&fields=code,nom&format=json")
        with urllib.request.urlopen(url2, timeout=30) as rep:
            par_insee = json.load(rep)
        if par_insee:
            raise SystemExit(
                f"[!] {cp} semble être un CODE INSEE (commune de "
                f"{par_insee[0]['nom']}), pas un code postal. Les codes INSEE "
                f"ne sont pas acceptés. Utilisez son code postal (ex. 10000 "
                f"pour Troyes, 75001 pour Paris 1er).")
        raise SystemExit(f"[!] {cp} : code postal inconnu.")

    dep = communes[0]["codeDepartement"]
    df = pd.DataFrame([{"code_insee": c["code"], "nom_commune": c["nom"],
                        "code_postal": cp} for c in communes])
    libelle = ", ".join(c["nom"] for c in communes[:5]) + \
              ("…" if len(communes) > 5 else "")
    print(f"[i] Code postal {cp} -> {libelle} (département {dep})")
    return dep, df["code_insee"].tolist(), df


def telecharger_commune(dep, code_insee, couche):
    """Télécharge et décompresse un fichier communal du Cadastre Etalab.
    Structure réelle des URL : communes/{département}/{commune}/…
    Renvoie le chemin du .json, ou None si indisponible pour la commune."""
    gz = f"cadastre-{code_insee}-{couche}.json.gz"
    js = gz[:-3]
    if os.path.exists(js):
        return js
    url = f"{BASE_URL}/{dep}/{code_insee}/{gz}"
    try:
        print(f"[↓] Téléchargement : {gz}")
        with urllib.request.urlopen(url, timeout=300) as rep, open(gz, "wb") as f:
            shutil.copyfileobj(rep, f)
        with gzip.open(gz, "rb") as f_in, open(js, "wb") as f_out:
            shutil.copyfileobj(f_in, f_out)
        os.remove(gz)
        return js
    except Exception as exc:
        print(f"[!] {gz} indisponible ({exc})", file=sys.stderr)
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


def charger_dvf(dep):
    """Ventes DVF géolocalisées, filtrées sur le département.
    Renvoie un DataFrame ou None si indisponible."""
    dest = f"dvf_{DVF_ANNEE}.csv"
    if not os.path.exists(dest):
        print(f"[↓] Téléchargement DVF {DVF_ANNEE} (France entière, ~500 Mo)...")
        try:
            gz = dest + ".gz"
            with urllib.request.urlopen(DVF_URL, timeout=1800) as rep, \
                    open(gz, "wb") as f:
                shutil.copyfileobj(rep, f)
            with gzip.open(gz, "rb") as f_in, open(dest, "wb") as f_out:
                shutil.copyfileobj(f_in, f_out)
            os.remove(gz)
        except Exception as exc:
            print(f"[!] DVF indisponible ({exc}) : colonnes de prix vides.",
                  file=sys.stderr)
            return None
    print("[⋯] Lecture du fichier DVF et filtrage sur le département...")
    usecols = ["id_parcelle", "valeur_fonciere", "surface_reelle_bati",
               "nature_mutation", "code_departement"]
    try:
        dvf = pd.read_csv(dest, sep=",", usecols=usecols, low_memory=False)
    except ValueError as exc:
        print(f"[!] Colonnes DVF inattendues ({exc}).", file=sys.stderr)
        return None
    dvf = dvf[dvf["code_departement"].astype(str) == str(dep)]
    dvf = dvf[dvf["nature_mutation"] == "Vente"]
    dvf = dvf.dropna(subset=["id_parcelle", "valeur_fonciere"])
    dvf["id_parcelle"] = dvf["id_parcelle"].astype(str).str.zfill(14)
    print(f"[✓] {len(dvf):,} mutations retenues dans le département.")
    return dvf


def main():
    dep, codes_insee, communes = resoudre_code_postal(ARG_CP)
    OUT_DIR = f"resultats_{ARG_CP}"
    os.makedirs(OUT_DIR, exist_ok=True)

    # -------------------------------------------------- 1. Parcelles des communes du code postal
    fichiers_parcelles = []
    for code in codes_insee:
        f = telecharger_commune(dep, code, "parcelles")
        if f:
            fichiers_parcelles.append(f)
    if not fichiers_parcelles:
        raise SystemExit("[!] Aucune donnée cadastrale récupérable pour ce code postal.")

    print("[⋯] Chargement des parcelles...")
    morceaux = []
    for f in fichiers_parcelles:
        gdf = gpd.read_file(f)
        if gdf.crs is None or gdf.crs.to_epsg() != 2154:
            gdf = gdf.to_crs(2154)
        morceaux.append(gdf)
    parcelles = pd.concat(morceaux, ignore_index=True)
    parcelles = reparer_geometries(parcelles)
    print(f"[✓] {len(parcelles):,} parcelles chargées.")

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
        f"(département {dep}) — CADASTRE ETALAB + DVF\n"
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

    import numpy as np
    fig2, ax2 = plt.subplots(figsize=(10, 6))
    ax2.hist(surf, bins=10 ** np.arange(-1, 3.2, 0.1), color="#2e7d32", edgecolor="white")
    ax2.set_xscale("log")
    ax2.set_xlabel("Surface de la parcelle (m², échelle log)")
    ax2.set_ylabel("Nombre de parcelles")
    ax2.set_title(f"Tailles des parcelles — code postal {ARG_CP}, échelle log")
    fig2.tight_layout()
    fig2.savefig(os.path.join(OUT_DIR, "histogramme_log.png"), dpi=150)
    print("[✓] Histogrammes exportés.")

    # -------------------------------------------------- 4. Emprise bâtie par parcelle
    parcelles["surface_batie_m2"] = 0.0
    try:
        morceaux_b = []
        for code in codes_insee:
            f = telecharger_commune(dep, code, "batiments")
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
            print(f"[✓] {len(batiments):,} bâtiments chargés.")

            print("[⋯] Attribution bâtiments → parcelles...")
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

            def _emprise(ligne):
                geom = geo_bat.loc[ligne["id_batiment"]].intersection(
                    geo_par.loc[ligne["id_parcelle"]])
                return 0.0 if geom.is_empty else geom.area

            paires["emprise_m2"] = paires.apply(_emprise, axis=1)
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
            print(f"[✓] {nb_baties:,} parcelles bâties identifiées.")
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
    tableau["nb_ventes"] = pd.NA
    tableau["prix_m2_bati"] = pd.NA
    dvf = charger_dvf(dep)
    if dvf is not None and len(dvf):
        # Les identifiants de parcelles diffèrent : cadastre Etalab utilise
        # "10015_000A_0001" (avec underscores), le DVF "10015000A0001".
        # On normalise en supprimant les underscores pour le croisement.
        tableau = tableau.drop(columns=["nb_ventes"], errors="ignore")
        tableau["id_dvf"] = tableau["id"].str.replace("_", "", regex=False)
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
                "nb_ventes", "id_parcelle"]
    export = tableau[[c for c in colonnes if c in tableau.columns] +
                     ["geometry"]].sort_values(["code_postal", "nom_commune",
                                                "numero_parcelle"])

    xlsx_path = os.path.join(OUT_DIR, "tableau_parcelles.xlsx")
    print("[⋯] Export Excel...")
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
        print(f"[✓] Excel exporté : {xlsx_path}")
    except Exception as exc:
        csv_path = os.path.join(OUT_DIR, "tableau_parcelles.csv")
        export.drop(columns="geometry").to_csv(csv_path, index=False, sep=";")
        print(f"[!] Excel impossible ({exc}), CSV exporté : {csv_path}",
              file=sys.stderr)
        print("    → Pour l'Excel : pip install openpyxl", file=sys.stderr)

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
        print("[✓] Graphique d'emprise bâtie exporté.")

    # -------------------------------------------------- 8. Carte interactive
    try:
        import folium
        print("[⋯] Carte interactive...")
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
        print(f"[✓] Carte interactive : {OUT_DIR}/carte_parcelles.html "
              f"({len(carte_gdf):,} parcelles)")
    except Exception:
        import traceback
        print(traceback.format_exc(), file=sys.stderr)
        print("[!] Carte ignorée (folium installé ?).", file=sys.stderr)

    with open(os.path.join(OUT_DIR, "stats_synthese.txt"), "w", encoding="utf-8") as f:
        f.write(stats)
    print(stats)
    print(f"[✓] Terminé. Résultats dans ./{OUT_DIR}/")


if __name__ == "__main__":
    main()