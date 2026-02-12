#!/usr/bin/env python3
"""
Outil d'import Compulocks → PrestaShop 1.7.4 pour Elockstore.com

Transforme le fichier catalogue Compulocks/Maclocks (TSV) en fichier
d'import CSV compatible PrestaShop 1.7.4.

Transformations effectuées:
  - UPC notation scientifique → EAN13
  - Poids livres → kilogrammes
  - Dimensions pouces → centimètres
  - Noms produits EN → FR
  - Descriptions EN → FR
  - Bullet points → liste HTML
  - Images 1-4 → URLs séparées par virgules
  - Catégorisation automatique par mots-clés
  - Génération meta title / meta description / URL slug

Usage:
    python prestashop_import.py --input catalogue_compulocks.tsv --output import_prestashop.csv
    python prestashop_import.py --input catalogue.xlsx --prices prix.csv --output import.csv
    python prestashop_import.py --input catalogue.tsv --dry-run
"""

import argparse
import csv
import json
import math
import os
import re
import sys
import unicodedata
from pathlib import Path

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

SCRIPT_DIR = Path(__file__).parent
DEFAULT_CONFIG = SCRIPT_DIR.parent / "config" / "field_mapping.json"

# Toutes les colonnes PrestaShop 1.7.4 dans l'ordre officiel
PRESTASHOP_COLUMNS = [
    "ID", "Active (0/1)", "Name *", "Categories (x,y,z...)",
    "Price tax excluded", "Price tax included", "Tax rule ID",
    "Wholesale price", "On sale (0/1)", "Discount amount",
    "Discount percent", "Discount from (yyyy-mm-dd)",
    "Discount to (yyyy-mm-dd)", "Reference", "Supplier reference",
    "Supplier", "Manufacturer", "EAN13", "UPC", "MPN", "Ecotax",
    "Width", "Height", "Depth", "Weight",
    "Delivery time of in-stock products",
    "Delivery time of out-of-stock products",
    "Quantity", "Minimal quantity", "Low stock level",
    "Send me an email when the quantity is under this level (0/1)",
    "Visibility", "Additional shipping cost", "Unity", "Unit price",
    "Short description", "Description", "Tags (x,y,z...)",
    "Meta title", "Meta keywords", "Meta description", "URL rewritten",
    "Text when in stock", "Text when backorder allowed",
    "Available for order (0/1)", "Product availability date",
    "Product creation date", "Show price (0/1)",
    "Image URLs (x,y,z...)", "Image alt texts (x,y,z...)",
    "Delete existing images (0/1)",
    "Feature (Name:Value:Position:Customized)",
    "Available online only (0/1)", "Condition",
    "Customizable (0/1)", "Uploadable files (0/1)",
    "Text fields (0/1)", "Out of stock action",
    "Virtual product (0/1)", "File URL",
    "Number of allowed downloads", "Expiration date (yyyy-mm-dd)",
    "Number of days", "ID / Name of shop",
    "Advanced Stock Management", "Depends on stock",
    "Warehouse", "Accessories (x,y,z...)",
]

# =====================================================================
# DICTIONNAIRE DE TRADUCTION EN → FR (termes produits Compulocks)
# =====================================================================
TRANSLATION_TERMS = {
    # Types de produits
    "Space Enclosure": "Boîtier Space",
    "Apex Secured Enclosure": "Boîtier Sécurisé Apex",
    "Enclosure": "Boîtier",
    "Secured Kickstand": "Support Chevalet Sécurisé",
    "Kickstand": "Support Chevalet",
    "Wall Mount": "Fixation Murale",
    "Floor Stand": "Totem Sur Pied",
    "Desk Mount": "Support Bureau",
    "Desk Stand": "Support Bureau",
    "Swing Arm": "Bras Pivotant",
    "Articulating Arm": "Bras Articulé",
    "Reach Arm": "Bras Extensible Reach",
    "Extra Long Arm": "Bras Extra Long",
    "Cable Lock": "Câble Antivol",
    "Lock Slot": "Slot de Verrouillage",
    "Edge Case": "Coque Edge",
    "Rugged Case": "Coque Renforcée",
    "Charging Station": "Station de Recharge",
    "Charging Dock": "Dock de Recharge",
    "Rise Stand": "Support Rise",
    "Cling Mount": "Support Adhésif Cling",
    "POS Stand": "Support Point de Vente",
    "Kiosk": "Borne Interactive",
    # Appareils
    "iPad": "iPad",
    "Galaxy Tab": "Galaxy Tab",
    "Surface": "Surface",
    "MacBook": "MacBook",
    "Laptop": "Ordinateur Portable",
    "Tablet": "Tablette",
    # Termes techniques
    "High-grade aluminum": "Aluminium haute qualité",
    "High -grade aluminum": "Aluminium haute qualité",
    "Conceals charging cable": "Dissimule le câble de charge",
    "Quick keyed access": "Accès rapide par clé",
    "VESA pattern": "Standard VESA",
    "enclosure keys": "clés pour boîtier",
    "Metal kickstand": "Support chevalet en métal",
    "extra functionality": "fonctionnalité supplémentaire",
    "Lock slot and cable lock": "Emplacement verrou et câble antivol",
    "Mobile protective": "Protection mobile",
    "Rubber Edge case": "Coque Edge en caoutchouc",
    "Landscape orientation only": "Orientation paysage uniquement",
    "Sleek metal frame": "Cadre métallique élégant",
    "Exposed front and back cameras": "Caméras avant et arrière accessibles",
    "sensors": "capteurs",
    "Cooling fan": "Ventilateur de refroidissement",
    "Multi-protection Safety System": "Système de sécurité multi-protection",
    "Multi-protection safety system": "Système de sécurité multi-protection",
    "High Efficiency Electric rectification": "Rectification électrique haute efficacité",
    "High efficiency electric rectification": "Rectification électrique haute efficacité",
    "USB ports": "ports USB",
    "Ergonomic Monitor Arm Wall Mount": "Fixation murale bras moniteur ergonomique",
    "Maximum weight": "Poids maximum",
    # Couleurs dans les noms
    "Black": "Noir",
    "White": "Blanc",
    "Silver": "Argent",
}

# Traductions des descriptions longues (phrases courantes)
DESCRIPTION_TRANSLATIONS = {
    "boasts rounded edges and open corners, enhancing ventilation and allowing limited access to buttons and inputs":
        "présente des bords arrondis et des coins ouverts, améliorant la ventilation et permettant un accès limité aux boutons et entrées",
    "Like all our tablet designs, the iPad can be continuously charged while enclosed, and peripheral cables can remain plugged in while mounted":
        "Comme tous nos designs tablettes, l'iPad peut être chargé en continu dans le boîtier, et les câbles périphériques restent branchés une fois monté",
    "You can mount this iPad enclosure flush to the wall or surface using the standard VESA pattern":
        "Vous pouvez monter ce boîtier iPad au ras du mur ou d'une surface grâce au standard VESA",
    "or pair it with one of our stands for added functionality":
        "ou l'associer à l'un de nos supports pour plus de fonctionnalités",
    "making it suitable for any workspace":
        "ce qui le rend adapté à tout espace de travail",
    "ensuring maximum security in high-traffic environments":
        "garantissant une sécurité maximale dans les environnements à fort passage",
    "it is compatible with our T-bar cable locks":
        "il est compatible avec nos câbles antivol T-bar",
    "This wall mount is designed to offer flexibility":
        "Cette fixation murale est conçue pour offrir de la flexibilité",
    "Experience the perfect blend of mobility, security, and functionality":
        "Découvrez le mélange parfait de mobilité, sécurité et fonctionnalité",
    "offers an intuitive and fully functional way to display your company's devices":
        "offre un moyen intuitif et entièrement fonctionnel d'exposer les appareils de votre entreprise",
    "seamlessly integrates into each of our KickStand enclosures for an added layer of protection":
        "s'intègre parfaitement dans chacun de nos boîtiers KickStand pour une couche de protection supplémentaire",
    "may safeguard your device from accidental scratches and bumps":
        "peut protéger votre appareil contre les rayures et chocs accidentels",
    "These kickstands can also enhance comfort for you, your employees, and your customers":
        "Ces supports chevalet améliorent également le confort pour vous, vos employés et vos clients",
    "Many users enjoy the secure grip our enclosures provide, which helps minimize user fatigue":
        "De nombreux utilisateurs apprécient la prise en main sécurisée de nos boîtiers, qui aide à réduire la fatigue",
    "do not block speakers or ports, so you continue to have easy access":
        "ne bloquent pas les haut-parleurs ni les ports, vous conservez un accès facile",
    "All bundles include":
        "Tous les packs incluent",
    "and a locking cable":
        "et un câble antivol",
    "Safe to Use":
        "Utilisation sûre",
    "Multi protection safety system ensures complete protection":
        "Le système de sécurité multi-protection assure une protection complète",
    "from electrical short circuit, over-heat, electric surge, over-charging, over current":
        "contre les courts-circuits, la surchauffe, les surtensions, la surcharge et le surcourant",
    "blends a sleek metal frame with practical features tailored for business use":
        "allie un cadre métallique élégant à des fonctionnalités pratiques adaptées à un usage professionnel",
    "quick keyed access and charging cable access":
        "accès rapide par clé et accès au câble de charge",
    "front and back cameras, as well as sensors, remain exposed":
        "les caméras avant et arrière, ainsi que les capteurs, restent accessibles",
    "offers versatility and security":
        "offre polyvalence et sécurité",
    "making it an ideal solution for displaying tablets in any business setting":
        "en faisant une solution idéale pour exposer des tablettes dans tout environnement professionnel",
}


def load_config(config_path):
    """Charge la configuration de mapping."""
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


# =====================================================================
# FONCTIONS DE TRANSFORMATION
# =====================================================================

def fix_scientific_upc(value):
    """Convertit un UPC en notation scientifique (8,19472E+11) en nombre entier."""
    if not value:
        return ""
    value = str(value).strip()
    # Notation scientifique format FR: 8,19472E+11
    value = value.replace(",", ".")
    try:
        num = float(value)
        result = str(int(num))
        # Pad to 12 digits (UPC) or 13 (EAN13)
        if len(result) < 12:
            result = result.zfill(12)
        return result
    except (ValueError, OverflowError):
        # Déjà un nombre normal ?
        clean = re.sub(r"[^\d]", "", value)
        return clean if clean else ""


def lbs_to_kg(value):
    """Convertit les livres en kilogrammes."""
    if not value:
        return ""
    value = str(value).strip().replace(",", ".")
    try:
        lbs = float(value)
        kg = lbs * 0.453592
        return f"{kg:.3f}"
    except ValueError:
        return ""


def inches_to_cm(value):
    """Convertit les pouces en centimètres."""
    if not value:
        return ""
    value = str(value).strip().replace(",", ".")
    try:
        inches = float(value)
        cm = inches * 2.54
        return f"{cm:.2f}"
    except ValueError:
        return ""


def parse_dimensions(dim_string):
    """Parse '13x10x2' (pouces) et retourne (largeur_cm, hauteur_cm, profondeur_cm)."""
    if not dim_string:
        return ("", "", "")
    dim_string = str(dim_string).strip()
    # Formats: "13x10x2", "13 x 10 x 2", "13X10X2"
    parts = re.split(r"[xX×\s]+", dim_string)
    parts = [p for p in parts if p]
    result = []
    for p in parts[:3]:
        p_clean = p.replace(",", ".")
        try:
            cm = float(p_clean) * 2.54
            result.append(f"{cm:.2f}")
        except ValueError:
            result.append("")
    while len(result) < 3:
        result.append("")
    return tuple(result[:3])


def translate_name(name):
    """Traduit un nom de produit de l'anglais au français."""
    if not name:
        return ""
    result = str(name).strip()
    # Appliquer les traductions de termes connus
    for en, fr in sorted(TRANSLATION_TERMS.items(), key=lambda x: -len(x[0])):
        result = result.replace(en, fr)
    return result


def translate_description(desc):
    """Traduit une description de l'anglais au français (approximation par remplacement)."""
    if not desc:
        return ""
    result = str(desc).strip()
    # Appliquer les traductions de phrases connues
    for en, fr in sorted(DESCRIPTION_TRANSLATIONS.items(), key=lambda x: -len(x[0])):
        result = result.replace(en, fr)
    # Appliquer les termes individuels
    for en, fr in sorted(TRANSLATION_TERMS.items(), key=lambda x: -len(x[0])):
        result = result.replace(en, fr)
    # Convertir les sauts de ligne en HTML
    result = result.replace("\r\n", "\n")
    paragraphs = [p.strip() for p in result.split("\n\n") if p.strip()]
    if len(paragraphs) > 1:
        result = "".join(f"<p>{p}</p>" for p in paragraphs)
    else:
        result = result.replace("\n", "<br>")
    return result


def bullets_to_html(bullet_text):
    """Convertit les bullet points en liste HTML."""
    if not bullet_text:
        return ""
    text = str(bullet_text).strip()
    # Format: "- item1\n- item2" ou "item1; item2" ou "item1, item2"
    items = []
    if "\n" in text:
        for line in text.split("\n"):
            line = line.strip()
            line = re.sub(r"^[-•*]\s*", "", line)
            if line:
                items.append(line)
    elif ";" in text:
        items = [i.strip() for i in text.split(";") if i.strip()]
    elif ", " in text and text.count(",") >= 2:
        items = [i.strip() for i in text.split(",") if i.strip()]
    else:
        return f"<p>{text}</p>"

    if not items:
        return ""

    # Traduire chaque bullet
    translated_items = []
    for item in items:
        translated = item
        for en, fr in sorted(TRANSLATION_TERMS.items(), key=lambda x: -len(x[0])):
            translated = translated.replace(en, fr)
        translated_items.append(translated)

    html = "<ul>\n"
    for item in translated_items:
        html += f"  <li>{item}</li>\n"
    html += "</ul>"
    return html


def slugify(text):
    """Convertit un texte en slug URL."""
    if not text:
        return ""
    text = unicodedata.normalize("NFKD", str(text))
    text = text.encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^\w\s-]", "", text.lower())
    text = re.sub(r"[-\s]+", "-", text).strip("-")
    return text[:128]


def categorize_product(name, config):
    """Détermine la catégorie PrestaShop d'après le nom du produit."""
    name_lower = str(name).lower()
    for rule in config.get("category_rules", []):
        pattern = rule.get("pattern", "")
        if pattern == "default":
            continue
        if re.search(pattern, name_lower):
            return rule["category"]
    # Catégorie par défaut
    for rule in config.get("category_rules", []):
        if rule.get("pattern") == "default":
            return rule["category"]
    return "Accueil > Supports et Fixations"


def build_features(row, config):
    """Construit la chaîne de features PrestaShop."""
    features = []
    color_map = config.get("color_translation", {})
    material_map = config.get("material_translation", {})
    country_map = config.get("country_translation", {})

    # Couleur
    color = str(row.get("Color", "")).strip()
    if color:
        color_fr = color_map.get(color, color)
        features.append(f"Couleur:{color_fr}:0:0")

    # Matériau
    material = str(row.get("Material", "")).strip()
    if material:
        material_fr = material_map.get(material, material)
        features.append(f"Matériau:{material_fr}:1:0")

    # Pays d'origine
    coo = str(row.get("COO", "")).strip()
    if coo:
        coo_fr = country_map.get(coo, coo)
        features.append(f"Pays d'origine:{coo_fr}:2:0")

    # Code douanier
    hs = str(row.get("HS Code", "")).strip()
    if hs:
        features.append(f"Code douanier:{hs}:3:0")

    # Compatibilité
    compat = str(row.get("Compatibility", "")).strip()
    if compat and compat != "None":
        # Nettoyer les sauts de ligne
        compat_clean = compat.replace("\n", " | ").replace("\r", "")
        features.append(f"Compatibilité:{compat_clean}:4:0")

    return ",".join(features)


def build_image_urls(row):
    """Construit la liste des URLs d'images séparées par virgules."""
    urls = []
    for key in ["Picture 1", "Picture 2", "Picture 3", "Picture 4"]:
        url = str(row.get(key, "")).strip()
        if url and url != "None":
            urls.append(url)
    return ",".join(urls)


def build_compatibility_html(compat_text):
    """Construit un tableau HTML de compatibilité."""
    if not compat_text or compat_text == "None":
        return ""
    lines = [l.strip() for l in str(compat_text).split("\n") if l.strip()]
    if not lines:
        return ""
    html = '<div class="compatibility"><h4>Compatibilité</h4><table class="table table-striped">'
    html += "<thead><tr><th>Appareil</th><th>Taille</th><th>Année</th><th>Modèles</th></tr></thead><tbody>"
    for line in lines:
        # Format: "iPad (7th gen.) | 10.2" | 2019 | A2197, A2200"
        parts = [p.strip() for p in line.split("|")]
        if len(parts) >= 4:
            html += f"<tr><td>{parts[0]}</td><td>{parts[1]}</td><td>{parts[2]}</td><td>{parts[3]}</td></tr>"
        elif len(parts) >= 2:
            html += f'<tr><td colspan="4">{" | ".join(parts)}</td></tr>'
        else:
            html += f'<tr><td colspan="4">{line}</td></tr>'
    html += "</tbody></table></div>"
    return html


# =====================================================================
# LECTURE DES FICHIERS
# =====================================================================

def read_input_file(filepath):
    """Lit le fichier catalogue Compulocks (TSV, CSV, ou Excel)."""
    filepath = Path(filepath)

    if filepath.suffix.lower() in (".xlsx", ".xls"):
        if not HAS_OPENPYXL:
            print("ERREUR: Le module 'openpyxl' est requis pour les fichiers Excel.")
            print("  pip install openpyxl")
            sys.exit(1)
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        ws = wb.active
        rows_raw = list(ws.iter_rows(values_only=True))
        wb.close()
        if not rows_raw:
            return [], []
        headers = [str(h).strip() if h else f"col_{i}" for i, h in enumerate(rows_raw[0])]
        data = []
        for row in rows_raw[1:]:
            row_dict = {}
            for i, val in enumerate(row):
                if i < len(headers):
                    row_dict[headers[i]] = str(val) if val is not None else ""
            data.append(row_dict)
        return headers, data

    # Fichier texte (TSV/CSV)
    encodings = ["utf-8-sig", "utf-8", "latin-1", "cp1252", "iso-8859-1"]

    for enc in encodings:
        try:
            with open(filepath, "r", encoding=enc) as f:
                sample = f.read(8192)

            # Détecter le délimiteur
            tab_count = sample.count("\t")
            semi_count = sample.count(";")
            comma_count = sample.count(",")

            if tab_count > semi_count and tab_count > comma_count:
                delimiter = "\t"
            elif semi_count > comma_count:
                delimiter = ";"
            else:
                delimiter = ","

            with open(filepath, "r", encoding=enc) as f:
                reader = csv.DictReader(f, delimiter=delimiter)
                headers = reader.fieldnames or []
                data = list(reader)

            if headers and data:
                print(f"  Encodage: {enc}")
                print(f"  Délimiteur: {'TAB' if delimiter == chr(9) else delimiter}")
                return headers, data

        except (UnicodeDecodeError, csv.Error):
            continue

    print(f"ERREUR: Impossible de lire {filepath}")
    sys.exit(1)


def read_price_file(filepath):
    """Lit un fichier de prix séparé et retourne un dict {reference: {prix}}."""
    if not filepath or not Path(filepath).exists():
        return {}
    _, data = read_input_file(filepath)
    prices = {}
    for row in data:
        ref = ""
        for key in ["PN", "Reference", "reference", "ref", "SKU"]:
            if key in row and row[key]:
                ref = str(row[key]).strip()
                break
        if not ref:
            continue
        price_info = {}
        for key in ["Cost", "prix_achat_ht", "PA_HT", "Wholesale"]:
            if key in row and row[key]:
                price_info["wholesale"] = str(row[key]).strip().replace(",", ".")
                break
        for key in ["Price_HT", "prix_vente_ht", "PV_HT", "Price"]:
            if key in row and row[key]:
                price_info["price_ht"] = str(row[key]).strip().replace(",", ".")
                break
        for key in ["Price_TTC", "prix_vente_ttc", "PV_TTC"]:
            if key in row and row[key]:
                price_info["price_ttc"] = str(row[key]).strip().replace(",", ".")
                break
        if price_info:
            prices[ref] = price_info
    return prices


# =====================================================================
# TRANSFORMATION PRINCIPALE
# =====================================================================

def transform_row(row, config, prices):
    """Transforme une ligne Compulocks en ligne PrestaShop."""
    ps = {col: "" for col in PRESTASHOP_COLUMNS}

    # Valeurs par défaut
    for key, val in config.get("default_values", {}).items():
        if key in ps:
            ps[key] = val

    # --- Référence ---
    pn = str(row.get("PN", "")).strip()
    if not pn or pn == "None":
        return None
    ps["Reference"] = pn
    ps["Supplier reference"] = pn

    # --- UPC / EAN ---
    upc_raw = row.get("UPC", "")
    ean = fix_scientific_upc(upc_raw)
    if len(ean) == 12:
        ps["UPC"] = ean
    elif len(ean) >= 13:
        ps["EAN13"] = ean[:13]
    else:
        ps["UPC"] = ean

    # --- Nom du produit ---
    marketing_name = str(row.get("Marketing Name", "")).strip()
    if not marketing_name or marketing_name == "None":
        return None
    ps["Name *"] = translate_name(marketing_name)

    # --- Catégorie ---
    ps["Categories (x,y,z...)"] = categorize_product(marketing_name, config)

    # --- Poids (lb → kg) ---
    ps["Weight"] = lbs_to_kg(row.get("Weight (lb)", ""))

    # --- Dimensions (in → cm) ---
    width, height, depth = parse_dimensions(row.get("Dimensions (In)", ""))
    ps["Width"] = width
    ps["Height"] = height
    ps["Depth"] = depth

    # --- Description courte (bullet points → HTML) ---
    bullets = str(row.get("Marketing Bullet Points", "")).strip()
    if bullets and bullets != "None":
        ps["Short description"] = bullets_to_html(bullets)
    else:
        # Essayer la version "as list"
        bullets_list = str(row.get("Marketing Bullet Points as list", "")).strip()
        if bullets_list and bullets_list != "None":
            ps["Short description"] = bullets_to_html(bullets_list)

    # --- Description longue ---
    long_desc = str(row.get("Long description", "")).strip()
    if long_desc and long_desc != "None":
        desc_html = translate_description(long_desc)
        # Ajouter le tableau de compatibilité
        compat = str(row.get("Compatibility", "")).strip()
        if compat and compat != "None":
            desc_html += build_compatibility_html(compat)
        ps["Description"] = desc_html
    elif ps["Short description"]:
        ps["Description"] = ps["Short description"]

    # --- Images ---
    image_urls = build_image_urls(row)
    ps["Image URLs (x,y,z...)"] = image_urls
    if image_urls:
        ps["Image alt texts (x,y,z...)"] = ",".join(
            [ps["Name *"]] * len(image_urls.split(","))
        )

    # --- Features ---
    ps["Feature (Name:Value:Position:Customized)"] = build_features(row, config)

    # --- Prix (depuis fichier séparé si disponible) ---
    if pn in prices:
        p = prices[pn]
        if "wholesale" in p:
            try:
                ps["Wholesale price"] = f"{float(p['wholesale']):.6f}"
            except ValueError:
                pass
        if "price_ht" in p:
            try:
                ps["Price tax excluded"] = f"{float(p['price_ht']):.6f}"
            except ValueError:
                pass
        if "price_ttc" in p:
            try:
                ps["Price tax included"] = f"{float(p['price_ttc']):.6f}"
            except ValueError:
                pass
        # Si on a le HT mais pas le TTC, calculer avec TVA 20%
        if ps["Price tax excluded"] and not ps["Price tax included"]:
            try:
                ht = float(ps["Price tax excluded"])
                tva_rate = config.get("price_config", {}).get("tva_rate", 20)
                ttc = ht * (1 + tva_rate / 100)
                ps["Price tax included"] = f"{ttc:.6f}"
            except ValueError:
                pass

    # --- Meta / SEO ---
    name_fr = ps["Name *"]
    ps["Meta title"] = (name_fr + " - Elockstore")[:70]
    short_clean = re.sub(r"<[^>]+>", "", ps.get("Short description", ""))
    if short_clean:
        ps["Meta description"] = short_clean[:160]
    else:
        ps["Meta description"] = f"Achetez {name_fr} sur Elockstore.com. Livraison rapide, prix compétitifs."[:160]
    ps["URL rewritten"] = slugify(name_fr)

    # --- Tags ---
    tags = set()
    name_lower = marketing_name.lower()
    tag_keywords = ["ipad", "galaxy", "surface", "macbook", "enclosure", "mount",
                     "stand", "kickstand", "charging", "lock", "arm", "kiosk"]
    for kw in tag_keywords:
        if kw in name_lower:
            tags.add(kw)
    color = str(row.get("Color", "")).strip().lower()
    if color and color != "none":
        tags.add(color)
    ps["Tags (x,y,z...)"] = ",".join(sorted(tags))

    return ps


# =====================================================================
# SORTIE
# =====================================================================

def write_output(rows, output_path, delimiter=";"):
    """Écrit le fichier CSV au format PrestaShop."""
    with open(output_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(
            f, fieldnames=PRESTASHOP_COLUMNS, delimiter=delimiter,
            quoting=csv.QUOTE_ALL, extrasaction="ignore"
        )
        writer.writeheader()
        writer.writerows(rows)


def print_summary(input_count, output_count, skipped, has_prices, categories):
    """Affiche un résumé."""
    print("\n" + "=" * 65)
    print("  RÉSUMÉ DE LA TRANSFORMATION COMPULOCKS → PRESTASHOP")
    print("=" * 65)
    print(f"  Produits en entrée   : {input_count}")
    print(f"  Produits exportés    : {output_count}")
    print(f"  Produits ignorés     : {skipped}")
    print(f"  Fichier prix chargé  : {'Oui' if has_prices else 'Non'}")
    print()
    if categories:
        print("  Catégories détectées:")
        for cat, count in sorted(categories.items(), key=lambda x: -x[1]):
            print(f"    {cat:50s} ({count})")
    print("=" * 65)


# =====================================================================
# MAIN
# =====================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Import Compulocks → PrestaShop 1.7.4 (Elockstore.com)"
    )
    parser.add_argument(
        "--input", "-i", required=True,
        help="Fichier catalogue Compulocks (TSV, CSV, XLSX)"
    )
    parser.add_argument(
        "--output", "-o", default="import_prestashop.csv",
        help="Fichier CSV de sortie pour PrestaShop (défaut: import_prestashop.csv)"
    )
    parser.add_argument(
        "--prices", "-p", default=None,
        help="Fichier de prix séparé (CSV) avec colonnes: PN, Cost, Price_HT, Price_TTC"
    )
    parser.add_argument(
        "--config", "-c", default=str(DEFAULT_CONFIG),
        help="Fichier de configuration (défaut: config/field_mapping.json)"
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Aperçu sans générer de fichier"
    )

    args = parser.parse_args()

    # Charger la config
    config_path = Path(args.config)
    if not config_path.exists():
        print(f"ERREUR: Config introuvable: {config_path}")
        sys.exit(1)
    config = load_config(config_path)
    print(f"Configuration: {config_path}")

    # Lire le catalogue Compulocks
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"ERREUR: Fichier introuvable: {input_path}")
        sys.exit(1)
    print(f"\nLecture du catalogue: {input_path}")
    headers, data = read_input_file(input_path)
    print(f"  {len(headers)} colonnes, {len(data)} produits")
    print(f"  Colonnes: {', '.join(headers[:8])}...")

    # Vérifier que c'est bien un fichier Compulocks
    expected = {"PN", "Marketing Name"}
    found = set(headers)
    if not expected.issubset(found):
        print(f"\nATTENTION: Colonnes attendues manquantes: {expected - found}")
        print("Ce fichier ne semble pas être un catalogue Compulocks standard.")
        print(f"Colonnes trouvées: {', '.join(headers)}")

    # Charger les prix si disponible
    prices = {}
    if args.prices:
        print(f"\nLecture des prix: {args.prices}")
        prices = read_price_file(args.prices)
        print(f"  {len(prices)} prix chargés")
    else:
        print("\nAucun fichier de prix fourni (--prices).")
        print("Les prix devront être ajoutés manuellement dans PrestaShop.")

    if args.dry_run:
        print("\n--- MODE APERÇU ---")
        for row in data[:3]:
            pn = row.get("PN", "?")
            name = row.get("Marketing Name", "?")
            upc = fix_scientific_upc(row.get("UPC", ""))
            weight_kg = lbs_to_kg(row.get("Weight (lb)", ""))
            dims = parse_dimensions(row.get("Dimensions (In)", ""))
            cat = categorize_product(name, config)
            print(f"\n  [{pn}] {name}")
            print(f"    UPC: {upc} | Poids: {weight_kg} kg | Dim: {dims[0]}x{dims[1]}x{dims[2]} cm")
            print(f"    Catégorie: {cat}")
            print(f"    Nom FR: {translate_name(name)}")
        print(f"\n  ... et {max(0, len(data) - 3)} autres produits")
        print("\n(Mode aperçu: aucun fichier généré)")
        return

    # Transformer
    print("\nTransformation en cours...")
    output_rows = []
    skipped = 0
    categories = {}

    for row in data:
        ps_row = transform_row(row, config, prices)
        if ps_row is None:
            skipped += 1
            continue
        output_rows.append(ps_row)
        cat = ps_row.get("Categories (x,y,z...)", "")
        categories[cat] = categories.get(cat, 0) + 1

    # Écrire
    output_path = Path(args.output)
    print(f"\nÉcriture: {output_path}")
    write_output(output_rows, output_path)

    print_summary(len(data), len(output_rows), skipped, bool(prices), categories)

    if not prices:
        print("\n  RAPPEL: Les prix ne sont pas inclus.")
        print("  Créez un fichier prix_elockstore.csv avec les colonnes:")
        print("    PN;Cost;Price_HT;Price_TTC")
        print("  Puis relancez avec: --prices prix_elockstore.csv")

    print(f"\nFichier prêt: {output_path}")
    print("Import via: Back Office > Paramètres avancés > Import > Produits")


if __name__ == "__main__":
    main()
