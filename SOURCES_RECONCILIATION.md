# Sources utilisées par chaque réconciliation

Pour chaque type de réconciliation, voici **quel fichier** et **quelle colonne (ou feuille/plage)** sont lus.  
*(Hors OS_Customers, comme demandé.)*

**Dossiers datés (DDMM)** : chaque semaine un nouveau dossier est utilisé. Ex. `2302` = 23 février, `0203` = 2 mars. Les sources sont lues dans `STIBO/{date}/`, `CT/{date}/`, `JEEVES/{date}/`. Lancement : `python run_reconciliation.py --market ekofisk --date 2302` (défaut : `--date 2302`). Les noms de fichiers dans chaque dossier suivent le suffixe de la date (ex. `Products_0203.xlsx`, `Invoice_Vendors_0203.xlsx`).

---

## 1. Product (Range Reconciliation)

| Source  | Fichier utilisé | Feuille / emplacement | Colonne / donnée |
|---------|------------------|------------------------|-------------------|
| **JEEVES** | 1er fichier contenant « Product » dans `JEEVES/{date}/` ou `JEEVES/` | `2-EXCELMASTER` | **Colonne A** à partir de **A3** ; code produit = **SUPC**. |
| **CT**     | 1er fichier contenant « Product » dans `CT/{date}/` ou `CT/` | `Item` (ou 1ère feuille) | En-têtes **ligne 6**, données à partir de la **ligne 7** ; **colonne B** = SUPC. |
| **STIBO**  | `STIBO/{date}/Products_{date}.xlsx` (ex. `Products_0203.xlsx`), sinon `STIBO/extract_stibo_all_products.xlsx` | Feuille active | En-têtes **ligne 1** ; **colonne C** = SUPC ; données à partir de **C2**. |

---

## 2. Vendor Invoice (onglet Vendor Invoice)

| Source  | Fichier utilisé | Feuille / emplacement | Colonne / donnée |
|---------|------------------|------------------------|-------------------|
| **STIBO** | `STIBO/{date}/Invoice_Vendors_{date}.xlsx` (ex. 0203) ; sinon extract root | Feuille active | 1ère colonne ou **« SUVC Invoice »** ; données à partir de la ligne 2. |
| **CT**     | 1er fichier contenant « Vendor » dans `CT/{date}/` ou `CT/` (et market si précisé) | **Invoice** | **Colonne C**, à partir de la **ligne 8**. |
| **JEEVES** | 1er fichier contenant « Vendor » dans `JEEVES/{date}/` ou `JEEVES/` | Feuille active | Colonne **« SUVC - Invoice »** ; données à partir de la **ligne 2**. |

---

## 3. Vendor OS (Ordering-Shipping)

| Source  | Fichier utilisé | Feuille / emplacement | Colonne / donnée |
|---------|------------------|------------------------|-------------------|
| **STIBO** | `STIBO/{date}/OS_Vendors_{date}.xlsx` ; sinon extract root ou vide | Feuille active | **Colonne « SUVC Ordering/Shipping »** ; données à partir de la ligne 2. |
| **CT**     | *Même fichier CT que Vendor Invoice* (dans `CT/{date}/` ou `CT/`) | **OrderingShipping** | **Colonne C**, à partir de la **ligne 8**. |
| **JEEVES** | *Même fichier JEEVES Vendor que pour Invoice* (dans `JEEVES/{date}/` ou `JEEVES/`) | **ORDERSHIPPING** | **Colonne A** ; données à partir de la **ligne 2**. |

---

## 4. Customer Invoice (onglet Customer Invoice)

| Source  | Fichier utilisé | Feuille / emplacement | Colonne / donnée |
|---------|------------------|------------------------|-------------------|
| **STIBO** | `STIBO/{date}/Invoice_Customers_{date}.xlsx` ou `Invoice_Customer_{date}.xlsx` ; sinon extract root | Feuille active | **Colonne « Invoice Customer Code »** ; données à partir de la ligne 2. |
| **CT**     | 1er fichier contenant « Customer » dans `CT/{date}/` ou `CT/` | **Invoice** | **Colonne C**, à partir de la **ligne 8**. |
| **JEEVES** | 1er fichier contenant « Customer » dans `JEEVES/{date}/` ou `JEEVES/` | **INVOICECUSTOMER** | **Colonne A**, à partir de la **ligne 3**. |

---

## 5. Customer OS (Ordering-Shipping) — *sauf OS_Customers*

| Source  | Fichier utilisé | Feuille / emplacement | Colonne / donnée |
|---------|------------------|------------------------|-------------------|
| **STIBO** | `STIBO/{date}/OS_Customers_{date}.xlsx` quand disponible ; sinon extract ou vide | Feuille active | À définir. |
| **CT**     | *Même fichier CT que Customer Invoice* (dans `CT/{date}/` ou `CT/`) | **OrderingShipping** | **Colonne C**, à partir de la **ligne 8** (C8 et en dessous). |
| **JEEVES** | *Même fichier JEEVES Customer que pour Invoice* (dans `JEEVES/{date}/` ou `JEEVES/`) | **ORDERSHIPPING** | **Colonne A**, à partir de la **ligne 3**. |

---

## Récap par réconciliation (sauf OS_Customers)

| Réconciliation   | STIBO | CT | JEEVES |
|-----------------|-------|----|--------|
| **Product**     | `STIBO/{date}/Products_{date}.xlsx` – **col. C** SUPC, L2+ | 1er *Product* dans `CT/{date}/` ou `CT/` – feuille Item, B, L6/L7+ | 1er *Product* dans `JEEVES/{date}/` ou `JEEVES/` – feuille `2-EXCELMASTER`, **A3+** |
| **Vendor Invoice**  | `STIBO/{date}/Invoice_Vendors_{date}.xlsx` (1ère col. ou SUVC Invoice) | 1er *Vendor* dans `CT/{date}/` ou `CT/` – feuille **Invoice**, **C L8+** | 1er *Vendor* dans `JEEVES/{date}/` ou `JEEVES/` – col. **SUVC - Invoice**, L2+ |
| **Vendor OS**   | `STIBO/{date}/OS_Vendors_{date}.xlsx` – col. **SUVC Ordering/Shipping** | Même fichier CT Vendor – feuille **OrderingShipping**, **C L8+** | Même fichier JEEVES Vendor – feuille **ORDERSHIPPING**, **col. A** L2+ |
| **Customer Invoice** | `STIBO/{date}/Invoice_Customers_{date}.xlsx` – col. **Invoice Customer Code** | 1er *Customer* dans `CT/{date}/` ou `CT/` – feuille **Invoice**, **C L8+** | 1er *Customer* dans `JEEVES/{date}/` ou `JEEVES/` – feuille **INVOICECUSTOMER**, **col. A** L3+ |
| **Customer OS** | `STIBO/{date}/OS_Customers_{date}.xlsx` (quand disponible) ou vide | Même fichier CT Customer – feuille **OrderingShipping**, **C8+** | Même fichier JEEVES Customer – feuille **ORDERSHIPPING**, **col. A** L3+ |

---

*Toutes les sources sont lues dans les dossiers datés `STIBO/{date}/`, `CT/{date}/`, `JEEVES/{date}/` (ex. date=2302, 0203). Si le dossier daté n’existe pas, repli sur la racine (`CT/`, `JEEVES/`). Fichiers STIBO : `Products_{date}.xlsx`, `Invoice_Vendors_{date}.xlsx`, `OS_Vendors_{date}.xlsx`, `Invoice_Customers_{date}.xlsx`, etc.*
