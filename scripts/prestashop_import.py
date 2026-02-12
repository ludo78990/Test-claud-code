#!/usr/bin/env python3
"""
Outil d'import simplifié pour PrestaShop 1.7.4 - Elockstore.com
Transforme un fichier fournisseur (CSV/Excel) en format d'import PrestaShop.

Usage:
    python prestashop_import.py --input fichier_fournisseur.csv --output import_prestashop.csv
    python prestashop_import.py --input fichier_fournisseur.xlsx --output import_prestashop.csv
    python prestashop_import.py --input fichier.csv --config ../config/field_mapping.json --output import.csv
"""

import argparse
import csv
import json
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

# Colonnes d'export PrestaShop 1.7.4 (ordre officiel)
PRESTASHOP_COLUMNS = [
    "ID",
    "Active (0/1)",
    "Name *",
    "Categories (x,y,z...)",
    "Price tax excluded",
    "Price tax included",
    "Tax rule ID",
    "Wholesale price",
    "On sale (0/1)",
    "Discount amount",
    "Discount percent",
    "Discount from (yyyy-mm-dd)",
    "Discount to (yyyy-mm-dd)",
    "Reference",
    "Supplier reference",
    "Supplier",
    "Manufacturer",
    "EAN13",
    "UPC",
    "MPN",
    "Ecotax",
    "Width",
    "Height",
    "Depth",
    "Weight",
    "Delivery time of in-stock products",
    "Delivery time of out-of-stock products",
    "Quantity",
    "Minimal quantity",
    "Low stock level",
    "Send me an email when the quantity is under this level (0/1)",
    "Visibility",
    "Additional shipping cost",
    "Unity",
    "Unit price",
    "Short description",
    "Description",
    "Tags (x,y,z...)",
    "Meta title",
    "Meta keywords",
    "Meta description",
    "URL rewritten",
    "Text when in stock",
    "Text when backorder allowed",
    "Available for order (0/1)",
    "Product availability date",
    "Product creation date",
    "Show price (0/1)",
    "Image URLs (x,y,z...)",
    "Image alt texts (x,y,z...)",
    "Delete existing images (0/1)",
    "Feature (Name:Value:Position:Customized)",
    "Available online only (0/1)",
    "Condition",
    "Customizable (0/1)",
    "Uploadable files (0/1)",
    "Text fields (0/1)",
    "Out of stock action",
    "Virtual product (0/1)",
    "File URL",
    "Number of allowed downloads",
    "Expiration date (yyyy-mm-dd)",
    "Number of days",
    "ID / Name of shop",
    "Advanced Stock Management",
    "Depends on stock",
    "Warehouse",
    "Accessories (x,y,z...)",
]


def load_config(config_path):
    """Charge la configuration de mapping."""
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def slugify(text):
    """Convertit un texte en slug URL compatible PrestaShop."""
    if not text:
        return ""
    text = unicodedata.normalize("NFKD", str(text))
    text = text.encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^\w\s-]", "", text.lower())
    text = re.sub(r"[-\s]+", "-", text).strip("-")
    return text[:128]


def clean_html(text):
    """Nettoie le HTML basique tout en gardant la structure."""
    if not text:
        return ""
    text = str(text).strip()
    # Remplacer les sauts de ligne par des <br>
    text = text.replace("\r\n", "<br>").replace("\n", "<br>")
    return text


def parse_price(value):
    """Parse un prix depuis différents formats (1.234,56 ou 1,234.56 ou 1234.56)."""
    if not value:
        return ""
    value = str(value).strip()
    # Supprimer les symboles monétaires
    value = re.sub(r"[€$£\s]", "", value)
    if not value:
        return ""
    # Détection du format: si la virgule est le dernier séparateur => format FR
    # Si le point est le dernier séparateur => format EN
    last_comma = value.rfind(",")
    last_dot = value.rfind(".")

    if last_comma > last_dot:
        # Format FR: 1.234,56
        value = value.replace(".", "").replace(",", ".")
    elif last_dot > last_comma:
        # Format EN: 1,234.56
        value = value.replace(",", "")
    else:
        # Pas de séparateur décimal ou un seul type
        value = value.replace(",", ".")

    try:
        return f"{float(value):.6f}"
    except ValueError:
        return ""


def parse_int(value):
    """Parse un entier."""
    if not value:
        return "0"
    try:
        return str(int(float(str(value).strip().replace(",", "."))))
    except (ValueError, TypeError):
        return "0"


def parse_float(value):
    """Parse un flottant."""
    if not value:
        return ""
    try:
        v = str(value).strip().replace(",", ".")
        return f"{float(v):.2f}"
    except (ValueError, TypeError):
        return ""


def parse_weight(value):
    """Parse un poids (en kg)."""
    if not value:
        return ""
    value = str(value).strip().lower()
    value = re.sub(r"[a-z\s]", "", value).replace(",", ".")
    try:
        return f"{float(value):.3f}"
    except (ValueError, TypeError):
        return ""


def parse_bool(value):
    """Parse une valeur booléenne vers 0/1."""
    if not value:
        return "1"
    value = str(value).strip().lower()
    if value in ("0", "non", "no", "false", "inactif", "désactivé"):
        return "0"
    return "1"


def strip_ean(value):
    """Nettoie un code EAN13."""
    if not value:
        return ""
    value = re.sub(r"[^\d]", "", str(value).strip())
    if len(value) == 13:
        return value
    if len(value) == 12:
        return "0" + value
    return value if len(value) <= 13 else value[:13]


def title_case_fr(text):
    """Met en forme le titre en respectant le français."""
    if not text:
        return ""
    text = str(text).strip()
    # Ne pas capitaliser les mots courts (articles, etc.)
    small_words = {"de", "du", "des", "le", "la", "les", "un", "une", "et", "ou", "en", "à", "au", "aux", "pour", "par", "sur", "avec", "sans"}
    words = text.split()
    result = []
    for i, word in enumerate(words):
        if i == 0 or word.lower() not in small_words:
            result.append(word.capitalize())
        else:
            result.append(word.lower())
    return " ".join(result)


def truncate(text, max_len):
    """Tronque un texte à la longueur max."""
    if not text:
        return ""
    text = str(text).strip()
    if len(text) <= max_len:
        return text
    return text[: max_len - 3] + "..."


def generate_meta_title(name):
    """Génère un meta title à partir du nom produit."""
    if not name:
        return ""
    base = f"{name} - Elockstore.com"
    return truncate(base, 70)


def generate_meta_description(name, short_desc=""):
    """Génère une meta description."""
    if short_desc:
        # Supprimer les tags HTML
        clean = re.sub(r"<[^>]+>", "", str(short_desc))
        return truncate(clean, 160)
    if name:
        return truncate(f"Achetez {name} sur Elockstore.com. Livraison rapide et prix compétitifs.", 160)
    return ""


def auto_detect_columns(headers, config):
    """Détecte automatiquement le mapping des colonnes."""
    patterns = config.get("auto_detect_columns", {}).get("patterns", {})
    detected = {}

    for header in headers:
        header_lower = header.lower().strip()
        header_normalized = re.sub(r"[^a-z0-9]", "", header_lower)

        for field_key, field_patterns in patterns.items():
            for pattern in field_patterns:
                pattern_normalized = re.sub(r"[^a-z0-9]", "", pattern.lower())
                if pattern_normalized == header_normalized or pattern_normalized in header_normalized:
                    detected[field_key] = header
                    break
            if field_key in detected:
                break

    return detected


TRANSFORM_FUNCTIONS = {
    "strip": lambda v: str(v).strip() if v else "",
    "strip_ean": strip_ean,
    "title_case": title_case_fr,
    "clean_html": clean_html,
    "parse_price": parse_price,
    "parse_int": parse_int,
    "parse_float": parse_float,
    "parse_weight": parse_weight,
    "parse_bool": parse_bool,
    "slugify": slugify,
    "truncate_70": lambda v: truncate(v, 70),
    "truncate_160": lambda v: truncate(v, 160),
    "map_category": lambda v: v,  # Sera remplacé dynamiquement
    "map_tax_rule": lambda v: v,  # Sera remplacé dynamiquement
}


def read_input_file(filepath, delimiter=";", encoding="utf-8-sig"):
    """Lit un fichier CSV ou Excel et retourne les headers et les lignes."""
    filepath = Path(filepath)

    if filepath.suffix.lower() in (".xlsx", ".xls"):
        if not HAS_OPENPYXL:
            print("ERREUR: Le module 'openpyxl' est requis pour lire les fichiers Excel.")
            print("Installez-le avec: pip install openpyxl")
            sys.exit(1)
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        if not rows:
            return [], []
        headers = [str(h) if h else f"col_{i}" for i, h in enumerate(rows[0])]
        data = []
        for row in rows[1:]:
            row_dict = {}
            for i, val in enumerate(row):
                if i < len(headers):
                    row_dict[headers[i]] = str(val) if val is not None else ""
            data.append(row_dict)
        return headers, data

    # CSV
    encodings_to_try = [encoding, "utf-8-sig", "utf-8", "latin-1", "cp1252", "iso-8859-1"]
    delimiters_to_try = [delimiter, ";", ",", "\t", "|"]

    for enc in encodings_to_try:
        try:
            with open(filepath, "r", encoding=enc) as f:
                sample = f.read(4096)
                # Détecter le délimiteur
                best_delim = delimiter
                max_count = 0
                for d in delimiters_to_try:
                    count = sample.count(d)
                    if count > max_count:
                        max_count = count
                        best_delim = d
            with open(filepath, "r", encoding=enc) as f:
                reader = csv.DictReader(f, delimiter=best_delim)
                headers = reader.fieldnames or []
                data = list(reader)
            if headers and data:
                print(f"  Encodage détecté: {enc}")
                print(f"  Délimiteur détecté: '{best_delim}'")
                return headers, data
        except (UnicodeDecodeError, csv.Error):
            continue

    print(f"ERREUR: Impossible de lire le fichier {filepath}")
    sys.exit(1)


def transform_row(row, column_mapping, field_mapping, config):
    """Transforme une ligne du fichier source vers le format PrestaShop."""
    ps_row = {col: "" for col in PRESTASHOP_COLUMNS}

    # Appliquer les valeurs par défaut
    defaults = config.get("default_values", {})
    for key, value in defaults.items():
        if key in ps_row:
            ps_row[key] = value

    # Mapping des catégories
    category_map = config.get("category_mapping", {})
    tax_map = config.get("tax_rule_mapping", {})

    # Transformer chaque champ mappé
    for internal_key, source_column in column_mapping.items():
        value = row.get(source_column, "")
        if not value or value == "None":
            continue

        mapping_info = field_mapping.get(internal_key, {})
        ps_field = mapping_info.get("prestashop_field", "")
        transform_name = mapping_info.get("transform", "strip")

        if not ps_field:
            continue

        # Transformations spéciales
        if transform_name == "map_category":
            value_lower = str(value).lower().strip()
            matched = False
            for cat_key, cat_value in category_map.items():
                if cat_key in value_lower:
                    value = cat_value
                    matched = True
                    break
            if not matched:
                value = category_map.get("default", "Accueil") + " > " + str(value).strip()
        elif transform_name == "map_tax_rule":
            value = tax_map.get(str(value).strip(), tax_map.get("20", "1"))
        else:
            transform_fn = TRANSFORM_FUNCTIONS.get(transform_name, TRANSFORM_FUNCTIONS["strip"])
            value = transform_fn(value)

        if ps_field in ps_row:
            ps_row[ps_field] = value

    # Post-processing: générer les champs dérivés si manquants
    name = ps_row.get("Name *", "")
    short_desc = ps_row.get("Short description", "")

    if name and not ps_row.get("Meta title"):
        ps_row["Meta title"] = generate_meta_title(name)

    if name and not ps_row.get("Meta description"):
        ps_row["Meta description"] = generate_meta_description(name, short_desc)

    if name and not ps_row.get("URL rewritten"):
        ps_row["URL rewritten"] = slugify(name)

    if name and not ps_row.get("Image alt texts (x,y,z...)") and ps_row.get("Image URLs (x,y,z...)"):
        ps_row["Image alt texts (x,y,z...)"] = name

    return ps_row


def write_output(rows, output_path, delimiter=";"):
    """Écrit le fichier CSV au format PrestaShop."""
    with open(output_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=PRESTASHOP_COLUMNS, delimiter=delimiter,
                                quoting=csv.QUOTE_ALL, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def print_summary(input_count, output_count, skipped, column_mapping):
    """Affiche un résumé de l'import."""
    print("\n" + "=" * 60)
    print("  RÉSUMÉ DE LA TRANSFORMATION")
    print("=" * 60)
    print(f"  Lignes en entrée  : {input_count}")
    print(f"  Lignes exportées  : {output_count}")
    print(f"  Lignes ignorées   : {skipped}")
    print(f"  Colonnes mappées  : {len(column_mapping)}")
    print()
    print("  Mapping détecté:")
    for internal, source in column_mapping.items():
        print(f"    {source:30s} -> {internal}")
    print("=" * 60)


def main():
    parser = argparse.ArgumentParser(
        description="Outil d'import PrestaShop 1.7.4 pour Elockstore.com"
    )
    parser.add_argument(
        "--input", "-i", required=True,
        help="Fichier d'entrée (CSV, XLS, XLSX)"
    )
    parser.add_argument(
        "--output", "-o", default="import_prestashop.csv",
        help="Fichier de sortie CSV pour PrestaShop (défaut: import_prestashop.csv)"
    )
    parser.add_argument(
        "--config", "-c", default=str(DEFAULT_CONFIG),
        help="Fichier de configuration du mapping (défaut: config/field_mapping.json)"
    )
    parser.add_argument(
        "--delimiter", "-d", default=";",
        help="Délimiteur du fichier de sortie (défaut: ;)"
    )
    parser.add_argument(
        "--skip-empty-name", action="store_true", default=True,
        help="Ignorer les lignes sans nom de produit (défaut: True)"
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Afficher le mapping sans générer le fichier"
    )

    args = parser.parse_args()

    # Charger la configuration
    config_path = Path(args.config)
    if not config_path.exists():
        print(f"ERREUR: Fichier de configuration introuvable: {config_path}")
        sys.exit(1)

    print(f"Chargement de la configuration: {config_path}")
    config = load_config(config_path)

    # Lire le fichier d'entrée
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"ERREUR: Fichier d'entrée introuvable: {input_path}")
        sys.exit(1)

    csv_settings = config.get("csv_settings", {})
    print(f"\nLecture du fichier: {input_path}")
    headers, data = read_input_file(
        input_path,
        delimiter=csv_settings.get("input_delimiter", ";"),
        encoding=csv_settings.get("input_encoding", "utf-8-sig"),
    )

    print(f"  {len(headers)} colonnes détectées")
    print(f"  {len(data)} lignes de données")
    print(f"  Colonnes: {', '.join(headers[:10])}{'...' if len(headers) > 10 else ''}")

    # Auto-détection des colonnes
    print("\nDétection automatique des colonnes...")
    column_mapping = auto_detect_columns(headers, config)
    field_mapping = config.get("field_mapping", {})

    if not column_mapping:
        print("ATTENTION: Aucune colonne détectée automatiquement.")
        print("Colonnes disponibles dans le fichier source:")
        for i, h in enumerate(headers):
            print(f"  [{i}] {h}")
        print("\nVeuillez configurer le mapping dans le fichier de configuration.")
        sys.exit(1)

    print(f"  {len(column_mapping)} colonnes mappées automatiquement")

    if args.dry_run:
        print_summary(len(data), 0, 0, column_mapping)
        print("\n(Mode dry-run: aucun fichier généré)")
        return

    # Transformer les données
    print("\nTransformation des données...")
    output_rows = []
    skipped = 0

    for i, row in enumerate(data):
        ps_row = transform_row(row, column_mapping, field_mapping, config)

        # Vérifier que le produit a un nom
        if args.skip_empty_name and not ps_row.get("Name *", "").strip():
            skipped += 1
            continue

        output_rows.append(ps_row)

    # Écrire le fichier de sortie
    output_path = Path(args.output)
    print(f"\nÉcriture du fichier: {output_path}")
    write_output(output_rows, output_path, delimiter=args.delimiter)

    print_summary(len(data), len(output_rows), skipped, column_mapping)
    print(f"\nFichier prêt pour import PrestaShop: {output_path}")
    print("Importez-le via: Back Office > Paramètres avancés > Import")


if __name__ == "__main__":
    main()
