# Outils d'import OpenProd

Deux outils pour éviter la ressaisie dans OpenProd, côté Wandercraft, Inc.

| Fichier | Ce qu'il fait |
|---|---|
| `index.html` | Page web. Convertit un export OpenProd en CSV Xero, ou un rapport Amazon en classeur d'import OpenProd. |
| `Import-AmazonPO.ps1` | Script PowerShell. Crée directement les commandes d'achat Amazon dans OpenProd, sans passer par Excel. |

La page est publiée sur https://sylina6.github.io/xero-bills-tool/

## Lequel utiliser

| Situation | Outil |
|---|---|
| Créer les commandes d'achat Amazon dans OpenProd | `Import-AmazonPO.ps1` |
| Pas sur le réseau interne, ou l'API est indisponible | La page web, onglet *Amazon → OpenProd PO*, puis import Excel manuel |
| Passer des factures fournisseur d'OpenProd vers Xero | La page web, onglet *OpenProd → Xero* |

---

# 1. Import-AmazonPO.ps1

Lit le rapport de rapprochement Amazon Business, regroupe les lignes d'article par commande, et crée une commande d'achat OpenProd par commande Amazon, avec toutes ses lignes.

## Prérequis

- Être sur le réseau Wandercraft. Le serveur `erp.local.wandercraft.eu` ne résout que sur le DNS interne. À distance, il faut le VPN.
- Un compte OpenProd sur la base `Wandercraft_US`.
- Windows. PowerShell est installé d'office, rien à installer.

## La première fois

**1. Récupérer le script.** Sur GitHub, ouvre `Import-AmazonPO.ps1` et clique sur l'icône de téléchargement (*Download raw file*). Ne fais pas de copier-coller depuis l'affichage du navigateur.

**2. Le débloquer.** Windows marque tout fichier venant d'internet et PowerShell refuse alors de l'exécuter. Une seule fois :

```powershell
cd "$HOME\Downloads"
Unblock-File .\Import-AmazonPO.ps1
```

**3. Vérifier qu'il existe une commande d'achat Amazon dans OpenProd.** Menu **Purchase > Purchase order**, cherche `Amazon` dans *Partner Name*.

S'il n'y en a aucune, crées-en une à la main : **Create**, choisis `[PA000016] Amazon` dans *Partner Name*, **Save**. Pas besoin de ligne d'article.

Le script s'en sert de modèle : il y lit la devise, les adresses, les conditions de paiement, la méthode de facturation et le système comptable. Ces valeurs ne sont pas devinables, et l'API ne les remplit pas toute seule. **Ne supprime pas cette commande.**

## Le déroulé, à chaque lot

**1. Télécharger le rapport depuis Amazon Business.** C'est le fichier `reconciliation_from_AAAAMMJJ_to_AAAAMMJJ_....csv`. Le `.xlsx` marche aussi.

Ne l'ouvre pas dans Excel pour ensuite l'enregistrer : Excel réécrit le fichier et peut en perdre des lignes.

**2. Ouvrir PowerShell.** Touche Windows, tape `PowerShell`, ouvre **Windows PowerShell**. Pas besoin d'être administrateur. Si tu tombes dans l'invite de commandes classique, tape `powershell` puis Entrée.

```powershell
cd "$HOME\Downloads"
```

**3. Vérifier le fichier.** Ne contacte même pas OpenProd.

```powershell
.\Import-AmazonPO.ps1 -Path .\reconciliation_from_20260901_to_20260914.csv -ParseOnly
```

Contrôle le nombre de commandes, le total hors taxes, et les lignes écartées. Une ligne écartée est normale pour un remboursement ou une quantité nulle ; elles sont toutes listées.

**4. Simuler.** Demande le mot de passe, se connecte, résout tous les identifiants, repère les commandes déjà importées, et affiche ce qu'il créerait. N'écrit rien.

```powershell
.\Import-AmazonPO.ps1 -Path .\reconciliation_from_20260901_to_20260914.csv
```

Le mot de passe ne s'affiche pas pendant la frappe, pas même des étoiles. C'est normal.

**5. Créer une seule commande, pour contrôler.**

```powershell
.\Import-AmazonPO.ps1 -Path .\reconciliation_from_20260901_to_20260914.csv -Apply -Limit 1
```

Ouvre la commande dans OpenProd et vérifie le prix unitaire, le sous-total, le nombre de lignes et l'affaire.

**6. Passer tout le lot.**

```powershell
.\Import-AmazonPO.ps1 -Path .\reconciliation_from_20260901_to_20260914.csv -Apply
```

Les commandes déjà créées sont reconnues par leur référence Amazon et sautées. Relancer le même fichier ne crée pas de doublon.

**7. Contrôler dans OpenProd.** Les commandes arrivent **en brouillon**. Elles ne partent pas en validation toutes seules.

## Les paramètres

| Paramètre | Par défaut | À quoi ça sert |
|---|---|---|
| `-Path` | obligatoire | Le rapport Amazon, `.csv` ou `.xlsx` |
| `-Apply` | absent | Écrit vraiment. Sans lui, le script ne fait que simuler |
| `-Limit` | 0 | Ne traite que les N premières commandes |
| `-ParseOnly` | absent | Vérifie le fichier sans contacter OpenProd |
| `-SupplierCode` | `PA000016` | Code de la fiche fournisseur Amazon |
| `-ProductCode` | `US - Office Expenses` | Code produit porté par toutes les lignes |
| `-AffairCode` | `AF0006` | Code affaire. Vide pour ne pas en mettre |
| `-TaxDescription` | vide | Description de la taxe, par exemple `NYC Purchase Tax 8.875`. Vide, la commande sort sans taxe |
| `-HeaderDescription` | `Ref` | `Ref` donne `date - fournisseur - référence`. Sinon `FirstItem` ou `Count` |
| `-TemplateOrder` | la plus récente | Référence de la commande servant de modèle, par exemple `INC-PO032795` |
| `-ShowRequired` | absent | Liste les champs obligatoires des deux modèles, puis s'arrête |
| `-InspectOrder` | absent | Affiche une commande existante champ par champ, puis s'arrête |

## Si ça bloque

| Message | Quoi faire |
|---|---|
| `cannot be loaded because running scripts is disabled` | `Unblock-File .\Import-AmazonPO.ps1` |
| `Authentication refused` | Mot de passe mal saisi. Il ne s'affiche pas, c'est vite arrivé |
| `No purchase order to copy the supplier settings from` | Créer une commande Amazon à la main, voir *La première fois* |
| `Supplier : nothing found with reference = ...` | Le code fournisseur n'existe pas. Vérifier sur la fiche partenaire |
| `Product : nothing found with code = ...` | Le code produit n'existe pas. C'est le code, pas le libellé |
| `Purchaser : nothing found with name = ...` | L'acheteur du rapport Amazon n'existe pas comme utilisateur OpenProd sous ce nom exact |
| `a mandatory field is not correctly set` | Le script affichera les champs manquants juste en dessous. Un champ obligatoire a été ajouté au modèle |
| Un montant à 0,00 sur la commande créée | Ouvrir la ligne et regarder *Quantity in price unit* |

Rien n'est jamais écrit tant que `-Apply` n'est pas passé. En cas de doute, relancer sans lui.

## Ce que le script ne fait pas

- **Le département d'autorisation reste FIN.** Ce n'est pas un champ de la commande d'achat, il appartient au circuit de validation. L'API ne peut pas l'écrire.
- **Aucune taxe n'est posée** si `-TaxDescription` n'est pas fourni. L'API utilisée ne déclenche pas les automatismes du formulaire, ce qui est voulu : c'est ce qui garantit que le prix Amazon n'est pas recalculé.
- **Les commandes restent en brouillon.** La demande de validation est manuelle.
- **Tous les articles vont sur le même produit générique.** Amazon n'est pas référencé article par article.

---

# 2. index.html

Page autonome, aucune installation, rien n'est envoyé nulle part : tout est traité dans le navigateur.

## Onglet OpenProd → Xero

Dépose l'export Excel d'OpenProd, choisis le format de date et le TaxType par défaut, télécharge le CSV prêt pour Xero.

## Onglet Amazon → OpenProd PO

Dépose le rapport de rapprochement Amazon, vérifie l'aperçu, télécharge le classeur.

Puis dans OpenProd, base **Wandercraft_US** : Settings > Import and export > Excel import > Batch > Import Batch. Charge le classeur sur la ligne de traitement, choisis l'importeur `Amazon - Purchase order`, puis Draft > Import.

Attention : la feuille `Lignes` n'a pas de clé de modification. Si tu relances un import sur une commande déjà créée, ses lignes s'ajoutent au lieu de se remplacer. Supprime la commande avant de relancer.

C'est précisément ce que le script PowerShell évite, d'où sa préférence.
