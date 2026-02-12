# Import Compulocks → PrestaShop 1.7.4 | Elockstore.com

Outil de transformation du catalogue Compulocks/Maclocks vers le format d'import PrestaShop 1.7.4.

## Transformations effectuées

| Champ source (Compulocks) | Transformation | Champ PrestaShop |
|---|---|---|
| PN | Copie directe | Reference + Supplier reference |
| UPC (8,19472E+11) | Conversion notation scientifique | EAN13 / UPC |
| Weight (lb) | Livres → Kilogrammes (×0.4536) | Weight |
| Dimensions (In) | Pouces → Centimètres (×2.54), split LxHxP | Width, Height, Depth |
| Marketing Name | Traduction EN→FR | Name * |
| Marketing Bullet Points | Traduction + conversion liste HTML | Short description |
| Long description | Traduction + mise en forme HTML | Description |
| Compatibility | Tableau HTML | Ajouté à Description |
| Picture 1-4 | Concaténation par virgules | Image URLs |
| Color / Material / COO | Traduction + Features PrestaShop | Feature |
| HS Code | Feature | Feature (Code douanier) |

## Structure du projet

```
├── scripts/
│   ├── prestashop_import.py          # Script Python principal
│   └── requirements.txt              # Dépendances (openpyxl)
├── n8n-workflow/
│   └── prestashop_import_workflow.json  # Workflow N8N importable
├── config/
│   └── field_mapping.json            # Configuration mapping + traductions
├── templates/
│   ├── exemple_fichier_fournisseur.csv  # Exemple catalogue Compulocks (TSV)
│   ├── prix_elockstore_exemple.csv      # Exemple fichier de prix
│   └── prestashop_import_template.csv   # Template CSV PrestaShop vide
└── README.md
```

## Option 1 : Script Python

### Installation

```bash
pip install openpyxl   # uniquement si fichier Excel
```

### Utilisation

```bash
# Import basique (sans prix)
python scripts/prestashop_import.py -i catalogue_compulocks.tsv -o import_prestashop.csv

# Import avec fichier de prix
python scripts/prestashop_import.py -i catalogue.tsv -p prix_elockstore.csv -o import.csv

# Aperçu sans générer de fichier
python scripts/prestashop_import.py -i catalogue.tsv --dry-run
```

### Point important : les prix

Le catalogue Compulocks **ne contient pas de prix**. Deux options :

1. **Fichier prix séparé** : Créez un CSV avec les colonnes `PN;Cost;Price_HT;Price_TTC` et utilisez `--prices`
2. **Import manuel** : Importez d'abord les produits, puis ajoutez les prix dans le back office PrestaShop

Exemple de fichier prix (`prix_elockstore.csv`) :
```
PN;Cost;Price_HT;Price_TTC
102IPDSB;45.00;89.00;106.80
1050MAAW;120.00;229.00;274.80
```

## Option 2 : Workflow N8N (VPS Hostinger)

### Installation

1. Ouvrir N8N sur votre VPS Hostinger
2. **Workflows** → **Import from File**
3. Importer `n8n-workflow/prestashop_import_workflow.json`
4. Configurer le chemin du fichier source dans "Lire Catalogue Compulocks"
5. (Optionnel) Activer le cron hebdomadaire

### Pipeline du workflow

```
Catalogue TSV → Parser → Filtrer → Transformer EN→FR → CSV PrestaShop → Fichier
                                        ↓
                                  Contrôle Qualité (logs)
```

## Import dans PrestaShop

1. Générer le fichier CSV avec l'un des outils ci-dessus
2. Back Office → **Paramètres avancés** → **Import**
3. Type d'entité : **Produits**
4. Charger le CSV généré (délimiteur `;`)
5. Vérifier le mapping → Lancer l'import

## Configuration

Le fichier `config/field_mapping.json` contient :

- **category_rules** : règles de catégorisation automatique par mots-clés
- **color_translation** / **material_translation** / **country_translation** : traductions
- **default_values** : valeurs par défaut (TVA, délais livraison, etc.)
- **price_config** : configuration du fichier prix et taux de TVA
