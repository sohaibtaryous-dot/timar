# TIMAR Tanger – Import & Export ALSTOM (Ventilation groupage · BL & Factures)

Application web (Streamlit) qui ventile automatiquement, **ligne par ligne**, les remorques de groupage multi-fournisseurs à l'import (client ALSTOM Fez). Elle tourne dans le navigateur, en local sur le PC, **sans droits administrateur**.

---

## 🚀 Lancement en une seule commande

Il faut Python 3.9 ou plus récent (la version « pour l'utilisateur » de python.org s'installe sans droits admin : décochez *Install for all users*).

```bash
pip install --user -r requirements.txt && python -m streamlit run app.py
```

Le navigateur s'ouvre sur **http://localhost:8501**.

Sous Windows, vous pouvez aussi **double-cliquer sur `lancer_windows.bat`**, qui installe les dépendances puis lance l'application.

> Astuce : pour isoler l'installation, `python -m venv .venv` puis `.venv\Scripts\activate` (Windows) ou `source .venv/bin/activate` (Linux/Mac), avant la commande ci-dessus. Aucun droit admin n'est requis non plus.

---

## 🎨 Interface

Thème sombre professionnel, en-tête animé avec camion TIMAR en 3D (suit la souris), étapes visuelles
Importer → Analyser → Exporter, cartes d'indicateurs 3D avec compteurs animés, graphiques interactifs
(valeur HT par fournisseur, statut du contrôle NGP) et statut général coloré.

> Le dossier caché `.streamlit/` doit rester à côté de `app.py` : il contient le thème.
> Les polices viennent de Google Fonts si Internet est disponible ; sinon une police système est utilisée.

## 🧭 Utilisation

1. **Importer le dossier PDF complet** (factures, packing lists, CMR, manifeste). Vous pouvez déposer un seul PDF qui regroupe tout, ou plusieurs PDF.
2. **Importer le fichier Excel de référence** (références et NGP ; formats `.xlsx`, `.xls`, `.xlsm` ou `.csv`).
3. Cliquer sur **« Lancer la ventilation du dossier »**. La barre de progression affiche les étapes :
   - Étape 1 : lecture OCR et extraction des données PDF
   - Étape 2 : réconciliation et matching NGP avec l'Excel
   - Étape 3 : génération du classeur Excel à 3 feuilles
4. Consulter le **résumé exécutif** : lignes extraites, factures, fournisseurs, NON TROUVÉ, écarts NGP, totaux HT par devise et poids.
5. Si besoin, **corriger ou compléter** les données dans les onglets « Lignes extraites » et « Manifeste ». Le matching, les contrôles et le fichier Excel sont alors recalculés.
6. **Télécharger le fichier Excel ventilé (.xlsx)**.

Des fichiers de démonstration se trouvent dans `exemples/` : un dossier PDF numérique, le même dossier scanné, et un référentiel NGP.

---

## ⚙️ Règles de gestion appliquées

| Règle | Mise en œuvre |
|---|---|
| **Individualisation stricte** | Chaque ligne de facture devient une ligne Excel. Une référence présente 3 fois donne 3 lignes, sans aucun regroupement ni somme. |
| **Segmentation du dossier** | Chaque page est classée (Facture / Packing List / CMR / Manifeste). Les pages de suite sont rattachées au bon document, et une packing list est liée à sa facture par n° de facture, puis par fournisseur, puis par position dans le dossier. |
| **Extraction** | On essaie trois méthodes et on garde celle qui donne le plus de lignes valorisées : 1) tableaux à filets (pdfplumber), 2) tableaux reconstitués d'après la position des mots (PDF sans filets et scans), 3) analyse ligne à ligne avec la vérification **Qté × PU = Montant** (prix aux 1, 100 ou 1000 unités). |
| **Formats numériques** | `1.234,56`, `1,234.56` et `1 234,56` sont reconnus. Le séparateur décimal est détecté pour chaque document. |
| **Poids** | Le poids par ligne (facture ou packing list) est repris tel quel. Sinon, le poids global de la facture ou de la packing list est ventilé **au prorata de la valeur HT** (ou de la quantité, au choix dans la barre latérale), avec un arrondi au gramme qui conserve le total exact. |
| **Colis** | Les colis par ligne sont repris s'ils existent. Sinon, le total de la facture est porté sur la 1ʳᵉ ligne, pour ne jamais le compter en double. |
| **Normalisation NGP** | Espaces, points et tirets sont supprimés. On garde 8 ou 10 chiffres, et le zéro de tête perdu par Excel est restitué (7→8, 9→10 chiffres). |
| **Matching** | Recherche de la référence article, puis de la référence Alstom, d'abord en correspondance exacte, puis sans séparateurs, puis avec une tolérance aux confusions OCR (O/0, I/1…). Toute correspondance non exacte est signalée dans les Remarques. |
| **NON TROUVÉ** | « NON TROUVÉ » est écrit en rouge, le NGP de l'en-tête de facture est appliqué, et la remarque « Nouveau Réf - NGP Facture appliqué » est ajoutée. |
| **Lignes de frais** | Fret, emballage, assurance… restent des lignes distinctes, avec le Type Ligne « Frais annexes ». |

---

## 📊 Classeur Excel généré

Mise en forme : police Arial, en-têtes sur fond bleu foncé `#1F3864` en texte blanc gras, filtres activés, volets figés, mise en forme conditionnelle. Les formules sont recalculées à l'ouverture.

1. **TOUS_CLIENTS** : détail ligne par ligne, 24 colonnes (Fournisseur → Remarques).
   La colonne Contrôle NGP utilise la formule :
   `=IF(OR(U2="NON TROUVÉ",T2=""),"À vérifier",IF(LEFT(U2,8)=LEFT(T2,8),"OK","ÉCART"))`
2. **CONTROLE** : récapitulatif par fournisseur et par devise. Il contient Total Valeur HT (`SUMIFS`), Goods Value du manifeste (cellule jaune modifiable), Écart, poids net et brut (`SUMIFS`), poids brut du manifeste, colis, nombre de lignes (`COUNTIFS`), nombre de NON TROUVÉ, nombre d'ÉCART et Statut.
3. **RECAP_GLOBAL** : synthèse de la remorque. On y trouve le nombre de fournisseurs, de factures et de lignes, le total HT **par devise**, les poids net et brut, les colis, les NON TROUVÉ, les écarts NGP et le **Statut Général** (CONFORME / À VÉRIFIER).

---

## 📤 Module EXPORT — BL & Factures

Dans la barre latérale, choisissez le module **EXPORT — BL & Factures**. Déposez ensuite tous les PDF du dossier « extrait » : Packing Lists, Bons de Livraison et Factures commerciales. Les scans sont lus par OCR.

Le site génère **deux fichiers séparés** :

| Fichier | Contenu |
|---|---|
| `FACTURES.xlsx` | Une seule feuille `Factures` de 5 colonnes, au format strict d'import pour « Créer une déclaration ». N° de facture en **texte**, **vraie date Excel** `DD/MM/YYYY`, devise sur 3 lettres, **montant HT numérique** au format `0.00` (sans séparateur de milliers ni symbole). Une ligne par facture, sans ligne vide ni total. |
| `BL_REFERENCES_SUMMARY.xlsx` | Trois feuilles : `BL References` (11 colonnes, **une ligne source = une ligne Excel**, références répétées jamais fusionnées, poids et colis **uniquement sur la 1re ligne de chaque BL**, une couleur par BL), `BL SUMMARY` (une ligne par BL, formules `SUMIFS`) et `GLOBAL TOTAL` (COLIS, POIDS BRUT, POIDS NET, VALEUR DEVISE, QUANTITE). Si les devises diffèrent, la valeur est présentée séparément pour chaque devise. |

Ce que fait l'application :

- **Rapprochement BL ↔ facture :** il se fait d'abord par les numéros croisés (n° de BL cité dans la facture, ou n° de facture cité dans le BL), puis par les références, puis par la position dans le dossier. Chaque cas incertain est signalé.
- **Correspondance des lignes :** chaque ligne de BL est reliée à sa ligne de facture par la référence et la quantité. Les occurrences répétées sont rapprochées une par une, jamais sur la seule référence.
- **Données reprises :** la description et le montant HT viennent de la facture, la quantité et l'unité du BL. Le Customer PO est pris sur la ligne si elle en a un, sinon dans l'en-tête.
- **Montant de la facture :** c'est le **Total HT**. Le « Net à payer » n'est retenu qu'avec une TVA à 0 % et sans TTC distinct. Le montant TTC n'est jamais utilisé.
- **Vérification après création :** les deux fichiers sont relus pour contrôler les feuilles, les colonnes, les types de cellules (texte, date, nombre) et les formats `DD/MM/YYYY` et `0.00`. On vérifie aussi que poids et colis ne figurent que sur la 1re ligne de chaque BL.
- **Compte rendu :** nombre de factures, BL et lignes, références répétées, colis, poids, valeur HT par devise, quantité, et anomalies.
- **Téléchargement :** chaque fichier séparément, ou un `.zip` avec les 2 fichiers et `COMPTE_RENDU.txt`.
- **Correction avant export :** les onglets « FACTURES » et « BL References » sont modifiables. Les fichiers sont régénérés et revérifiés automatiquement.

> Si votre système refuse `FACTURES.xlsx` (« Erreur à la ligne X, colonne Y »), ouvrez l'onglet « FACTURES (modifiable) ». Corrigez la cellule indiquée, puis retéléchargez le fichier.

## 🔍 OCR des documents scannés (optionnel)

Une page est traitée par OCR automatiquement quand elle ne contient pas de texte. Avant la reconnaissance, les filets des tableaux sont effacés pour améliorer la lecture. Deux moteurs sont possibles, et aucun ne demande de droits admin :

- **Tesseract en version portable** (recommandé, léger) : décompressez une version portable de Tesseract-OCR dans votre dossier utilisateur. Indiquez ensuite le chemin de `tesseract.exe` dans la barre latérale, par exemple `C:\Users\<vous>\Tesseract-OCR\tesseract.exe`. Ajoutez `fra.traineddata` et `deu.traineddata` dans `tessdata` pour le français et l'allemand.
- **EasyOCR** (100 % Python, sans binaire externe, mais environ 1 Go avec PyTorch) : `pip install --user easyocr`.

Le moteur détecté s'affiche dans la barre latérale. Sans moteur OCR, les PDF numériques sont traités normalement et les pages scannées sont signalées dans le journal.

---

## 📁 Structure

```
app.py               Application Streamlit + moteur de ventilation + génération Excel
export_engine.py     Module EXPORT : BL / Packing Lists + Factures -> FACTURES.xlsx + BL_REFERENCES_SUMMARY.xlsx
ui.py                Design de l'interface (thème, en-tête 3D animé, cartes KPI 3D, graphiques)
.streamlit/config.toml  Thème sombre, barre d'outils minimale (bouton « Deploy » masqué)
requirements.txt     Dépendances Python
lancer_windows.bat   Lancement en un double-clic (Windows)
exemples/            Dossiers PDF de démonstration Import et Export (numériques + scannés) et référentiel NGP
```

## ❓ Dépannage

- **`streamlit` n'est pas reconnu** : utilisez `python -m streamlit run app.py`, car le dossier des scripts `--user` n'est pas toujours dans le PATH.
- **Port déjà utilisé** : `python -m streamlit run app.py --server.port 8502`.
- **Facture sans ligne détectée** : le journal (onglet « Documents & journal ») l'indique. Vous pouvez ajouter les lignes à la main dans l'onglet « Lignes extraites », et l'export les prend en compte.
- **Fournisseur affiché « À compléter »** : corrigez-le dans le tableau, la feuille CONTROLE se met à jour.
