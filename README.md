# Outil d'Import PrestaShop 1.7.4 - Elockstore.com

Outil simplifié de transformation et d'import de données fournisseur vers PrestaShop 1.7.4.

## Structure du projet

```
├── scripts/
│   ├── prestashop_import.py     # Script Python de transformation
│   └── requirements.txt         # Dépendances Python
├── n8n-workflow/
│   └── prestashop_import_workflow.json  # Workflow N8N importable
├── config/
│   └── field_mapping.json       # Configuration du mapping des colonnes
├── templates/
│   ├── prestashop_import_template.csv   # Template CSV PrestaShop vide
│   └── exemple_fichier_fournisseur.csv  # Exemple de fichier fournisseur
└── README.md
```

## Option 1 : Script Python (usage local)

### Installation

```bash
cd scripts/
pip install -r requirements.txt
```

### Utilisation

```bash
# Import basique (CSV)
python scripts/prestashop_import.py --input fichier_fournisseur.csv --output import_prestashop.csv

# Import depuis Excel
python scripts/prestashop_import.py --input fichier_fournisseur.xlsx --output import_prestashop.csv

# Mode aperçu (sans générer de fichier)
python scripts/prestashop_import.py --input fichier_fournisseur.csv --dry-run

# Avec configuration personnalisée
python scripts/prestashop_import.py --input fichier.csv --config config/field_mapping.json --output import.csv
```

### Fonctionnalités

- **Auto-détection des colonnes** : le script identifie automatiquement les colonnes du fichier source (référence, nom, prix, EAN, etc.)
- **Formats supportés** : CSV (tout délimiteur), Excel (.xlsx, .xls)
- **Encodages supportés** : UTF-8, Latin-1, CP1252, ISO-8859-1
- **Transformations automatiques** :
  - Nettoyage des prix (formats FR et EN)
  - Validation des codes EAN13
  - Génération automatique des URL rewritten (slugs)
  - Génération des meta title / meta description
  - Mapping des catégories
  - Nettoyage HTML des descriptions

## Option 2 : Workflow N8N (automatisé sur VPS)

### Installation sur N8N (Hostinger VPS)

1. **Ouvrir N8N** sur votre VPS
2. Aller dans **Workflows** > **Import from File**
3. Importer le fichier `n8n-workflow/prestashop_import_workflow.json`
4. **Configurer** :
   - Mettre à jour le chemin du fichier source dans le noeud "Lire Fichier Fournisseur"
   - Mettre à jour le chemin de sortie dans le noeud "Sauvegarder Fichier"
   - (Optionnel) Activer le noeud "Import Quotidien" pour un import automatique
   - (Optionnel) Configurer les credentials SMTP pour les notifications email

### Noeuds du workflow

| Noeud | Fonction |
|-------|----------|
| Start / Cron | Déclenchement manuel ou programmé |
| Lire Fichier Fournisseur | Lecture du CSV/Excel source |
| Parser CSV | Parsing du fichier en données structurées |
| Filtrer Lignes Vides | Exclusion des lignes sans nom produit |
| Transformer Données | Conversion au format PrestaShop 1.7.4 |
| Contrôle Qualité | Vérification des données (ref, prix, EAN) |
| Générer CSV PrestaShop | Export CSV au format PrestaShop |
| Sauvegarder Fichier | Écriture du fichier sur le VPS |
| Notification Email | (Optionnel) Rapport par email |

## Import dans PrestaShop

1. Générer le fichier CSV avec l'un des deux outils
2. Se connecter au **Back Office PrestaShop**
3. Aller dans **Paramètres avancés** > **Import**
4. Sélectionner **Produits** comme type d'entité
5. Charger le fichier CSV généré
6. Vérifier le mapping des colonnes (devrait être automatique avec le format standard)
7. Lancer l'import

## Configuration du mapping

Le fichier `config/field_mapping.json` permet de personnaliser :

- **field_mapping** : correspondance entre les colonnes source et les champs PrestaShop
- **default_values** : valeurs par défaut pour les champs non présents dans la source
- **category_mapping** : traduction des catégories fournisseur vers les catégories PrestaShop
- **tax_rule_mapping** : correspondance des taux de TVA vers les règles de taxe PrestaShop
- **auto_detect_columns** : patterns de détection automatique des noms de colonnes

### Adapter à votre fichier fournisseur

Si les colonnes de votre fichier ne sont pas détectées automatiquement, modifiez la section `auto_detect_columns.patterns` dans `field_mapping.json` pour ajouter les noms exacts de vos colonnes.

Exemple : si votre fichier a une colonne "Désignation article" pour le nom du produit, ajoutez :

```json
"nom_produit": ["designation article", "désignation", ...]
```
