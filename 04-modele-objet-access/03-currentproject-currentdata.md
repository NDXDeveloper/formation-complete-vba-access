🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 4.3. CurrentProject et CurrentData

La section 4.2 a montré que `Application` expose deux propriétés jumelles : **`CurrentProject`** et **`CurrentData`**. Toutes deux décrivent la **base de données courante**, mais sous deux angles différents — l'une regroupe les objets de l'**application**, l'autre ceux des **données**. Cette section les détaille, et présente un objet précieux qu'elles offrent : une connexion ADO prête à l'emploi.

## Deux conteneurs pour une même base

`CurrentProject` et `CurrentData` pointent vers la **même** base de données, mais répartissent ses objets en deux familles — exactement la distinction « objets d'application / objets de données » vue en section 1.5 :

- **`CurrentProject`** rassemble les objets de l'**application** : formulaires, états, macros, modules. Il porte aussi les informations de **fichier** et de **connexion**.
- **`CurrentData`** rassemble les objets de **données** : tables et requêtes.

Un moyen mnémotechnique simple : *Project* pour le versant **applicatif** (ce qui constitue l'application), *Data* pour le versant **données** (ce que l'application manipule).

```vba
' CurrentProject : objets de l'application
Debug.Print CurrentProject.AllForms.Count
Debug.Print CurrentProject.AllReports.Count

' CurrentData : objets de données
Debug.Print CurrentData.AllTables.Count
Debug.Print CurrentData.AllQueries.Count
```

## CurrentProject : les objets de l'application

### Les collections All… du projet

`CurrentProject` expose une collection par type d'objet applicatif :

- **`AllForms`** — tous les formulaires ;
- **`AllReports`** — tous les états ;
- **`AllMacros`** — toutes les macros ;
- **`AllModules`** — tous les modules.

Ces collections sont détaillées à la section suivante (4.4). Retenez dès maintenant un point essentiel, développé plus bas : elles contiennent des **descripteurs** d'objets, et non les objets eux-mêmes.

### Les informations de fichier et de connexion

Au-delà des objets, `CurrentProject` renseigne sur la **base elle-même** — son emplacement, son état, sa connexion :

```vba
Debug.Print CurrentProject.Name       ' "Gestion.accdb"
Debug.Print CurrentProject.Path       ' "C:\Apps"
Debug.Print CurrentProject.FullName   ' "C:\Apps\Gestion.accdb"
```

Ces propriétés sont précieuses pour gérer des **chemins relatifs**, localiser le back-end, ou construire des fichiers à côté de la base (un sujet récurrent au déploiement, chapitre 21).

La propriété **`IsTrusted`** indique si la base est **approuvée** (emplacement de confiance, contenu activé). Elle permet de vérifier, par code, si le code et les macros pourront s'exécuter — un écho direct aux questions de sécurité des sections 1.3, 1.4 et 20.5.

```vba
If Not CurrentProject.IsTrusted Then
    MsgBox "La base n'est pas approuvée : le code pourrait ne pas s'exécuter."
End If
```

## CurrentData : les objets de données

De façon symétrique, `CurrentData` expose les collections d'objets de données. Pour une base `.accdb` classique, deux comptent :

- **`AllTables`** — toutes les tables ;
- **`AllQueries`** — toutes les requêtes.

```vba
Dim obj As AccessObject
For Each obj In CurrentData.AllTables
    Debug.Print obj.Name
Next obj
```

> ℹ️ `CurrentData` expose aussi des collections comme `AllViews`, `AllStoredProcedures` ou `AllFunctions`, qui ne concernent que les anciens projets Access connectés à SQL Server (ADP) et sont sans objet dans une base `.accdb` standard.

## Les collections contiennent des descripteurs (AccessObject)

Voici le point à bien comprendre. Les collections `All…` ne contiennent **pas** les formulaires, états ou tables eux-mêmes, mais des objets **`AccessObject`** — des **descripteurs** porteurs de métadonnées : le nom, l'état (ouvert ou non), le type, les dates de création et de modification…

```vba
Dim obj As AccessObject
For Each obj In CurrentProject.AllForms
    Debug.Print obj.Name, obj.IsLoaded   ' nom + état (chargé/ouvert ou non)
Next obj
```

Il faut donc bien distinguer deux choses :

- **`CurrentProject.AllForms("frmClients")`** renvoie un **descripteur** (`AccessObject`) — il existe que le formulaire soit ouvert ou non ;
- **`Forms("frmClients")`** renvoie le **formulaire réel et ouvert** (`Form`) — et n'existe que si le formulaire est ouvert.

C'est exactement la distinction « objets existants / objets ouverts » introduite en section 4.1, et approfondie en section 4.4.

## CurrentProject.Connection : une connexion ADO prête à l'emploi

`CurrentProject` rend un dernier service très commode : sa propriété **`Connection`** fournit une **connexion ADO** (`ADODB.Connection`) déjà ouverte vers la base courante — **sans avoir à construire de chaîne de connexion**.

```vba
' Une connexion ADO vers la base courante, sans chaîne à écrire
Dim cn As Object            ' ADODB.Connection
Set cn = CurrentProject.Connection
' ... usage ADO (chapitre 10) ...
```

C'est l'occasion de souligner une **symétrie** importante avec ce qui vient ensuite. Pour atteindre les données de la base courante, deux chemins coexistent :

- **`CurrentProject.Connection`** → une connexion **ADO** (chapitre 10) ;
- **`CurrentDb`** → une base **DAO** (sections 4.5 et 4.6).

Le choix entre les deux technologies est discuté en section 10.1 ; on retiendra ici que `CurrentProject` est le point d'entrée ADO « tout prêt », tandis que `CurrentDb` sera le point d'entrée DAO.

## Récapitulatif

| Conteneur | Versant | Collections principales | Connexion fournie |
|---|---|---|---|
| **`CurrentProject`** | application | `AllForms`, `AllReports`, `AllMacros`, `AllModules` | ADO (`.Connection`) |
| **`CurrentData`** | données | `AllTables`, `AllQueries` | ADO (`.Connection`) |

## À retenir

- **`CurrentProject`** et **`CurrentData`** décrivent la **même base courante** sous deux angles : objets d'**application** d'un côté, objets de **données** de l'autre.
- `CurrentProject` expose `AllForms`, `AllReports`, `AllMacros`, `AllModules` ; `CurrentData` expose `AllTables` et `AllQueries`.
- Ces collections contiennent des **descripteurs** (`AccessObject` : nom, `IsLoaded`, type…), **pas** les objets eux-mêmes — à distinguer de `Forms("x")` qui renvoie le formulaire **ouvert** (section 4.4).
- `CurrentProject` fournit aussi des infos de fichier (`Name`, `Path`, `FullName`) et l'état d'approbation (**`IsTrusted`**, lien sécurité section 1.4).
- **`CurrentProject.Connection`** offre une connexion **ADO** prête à l'emploi (chapitre 10) — pendant du **`CurrentDb`** côté **DAO** (sections 4.5 et 4.6).

---


⏭️ [4.4. Collections AllForms, AllReports, AllTables, AllQueries...](/04-modele-objet-access/04-collections-allobjects.md)
