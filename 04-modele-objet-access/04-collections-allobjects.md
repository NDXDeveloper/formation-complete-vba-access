🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 4.4. Collections AllForms, AllReports, AllTables, AllQueries...

Les sections 4.1 et 4.3 l'ont annoncé : les collections **`All…`** énumèrent **tous** les objets d'un type donné, qu'ils soient ouverts ou non. Cette section les détaille — leur contenu, l'objet `AccessObject` qu'elles renvoient, et leurs usages les plus courants : inventorier les objets, tester leur existence, et savoir s'ils sont ouverts.

## Énumérer tous les objets d'un type

Comme vu en section 4.3, ces collections se répartissent entre les deux conteneurs de l'application courante :

- sous **`CurrentProject`** : `AllForms`, `AllReports`, `AllMacros`, `AllModules` ;
- sous **`CurrentData`** : `AllTables`, `AllQueries`.

Chacune liste **l'intégralité** des objets de son type. On les parcourt avec `For Each` (section 3.2) :

```vba
Dim obj As AccessObject
For Each obj In CurrentProject.AllForms
    Debug.Print obj.Name
Next obj

For Each obj In CurrentData.AllTables
    Debug.Print obj.Name
Next obj
```

## Forms vs AllForms : la distinction clé

C'est le cœur de cette section, déjà esquissé en 4.1. Il faut distinguer deux familles de collections :

- les collections **`Forms`** et **`Reports`** (sous `Application`) ne contiennent que les objets **ouverts**, sous forme d'objets **réels** (`Form`, `Report`) que l'on peut piloter directement ;
- les collections **`All…`** contiennent **tous** les objets, sous forme de **descripteurs** (`AccessObject`), ouverts ou non.

| | `Forms` | `CurrentProject.AllForms` |
|---|---|---|
| Contenu | formulaires **ouverts** | **tous** les formulaires |
| Type des éléments | `Form` (objet réel) | `AccessObject` (descripteur) |
| Présent si fermé ? | non | oui |
| Accès | indice (base 0) ou nom | par nom (et indice) |
| Usage typique | manipuler un formulaire ouvert | inventorier, tester l'existence ou l'état |

En clair : pour **agir** sur un formulaire ouvert, on passe par `Forms` ; pour **savoir ce qui existe** dans la base ou **interroger l'état** d'un objet, on passe par `AllForms`.

## L'objet AccessObject : un descripteur

Chaque élément d'une collection `All…` est un **`AccessObject`** : non pas l'objet lui-même, mais un **descripteur** porteur de métadonnées. Ses propriétés les plus utiles :

- **`Name`** — le nom de l'objet ;
- **`IsLoaded`** — booléen : l'objet est-il actuellement ouvert ?
- **`CurrentView`** — la vue dans laquelle il est ouvert (s'il l'est) ;
- **`Type`** — le type d'objet (`acForm`, `acTable`…) ;
- **`DateCreated`** / **`DateModified`** — dates de création et de dernière modification ;
- **`Properties`** — ses propriétés personnalisées (section 12.8).

```vba
Dim obj As AccessObject
Set obj = CurrentProject.AllForms("frmClients")
Debug.Print obj.Name           ' "frmClients"
Debug.Print obj.IsLoaded       ' True si le formulaire est ouvert
Debug.Print obj.DateModified   ' date de dernière modification
```

Parce qu'il ne s'agit que d'un descripteur, l'`AccessObject` ne permet pas de manipuler directement le contenu de l'objet : pour cela, on **ouvre** le formulaire (via `DoCmd`, chapitre 5) ou l'on accède à la table en DAO (chapitre 9), puis on travaille avec l'objet réel.

## Usages courants

### Lister les objets (documentation, maintenance)

L'usage le plus immédiat : dresser l'**inventaire** des objets de la base — pour générer une documentation, alimenter un menu dynamique, ou faire le ménage. La génération de documentation par code est abordée en section 24.3.

### Savoir si un objet existe

Les collections `All…`, comme l'objet `Collection` vu en section 3.7, ne possèdent **pas** de méthode `Exists`. Pour vérifier qu'un objet existe avant de l'ouvrir ou de le créer, on parcourt la collection :

```vba
Function FormulaireExiste(nom As String) As Boolean
    Dim obj As AccessObject
    For Each obj In CurrentProject.AllForms
        If obj.Name = nom Then
            FormulaireExiste = True
            Exit Function
        End If
    Next obj
End Function
```

### Savoir si un objet est ouvert (IsLoaded)

C'est le **test canonique** « ce formulaire est-il ouvert ? », et la propriété **`IsLoaded`** est la façon propre de le faire — bien plus fiable que de tenter d'accéder à `Forms("frmClients")` et d'intercepter l'erreur.

```vba
' Ce formulaire est-il ouvert ?
If CurrentProject.AllForms("frmClients").IsLoaded Then
    Debug.Print "frmClients est ouvert"
End If
```

Si l'on a besoin de connaître la **vue** dans laquelle l'objet est ouvert (création, formulaire, aperçu…), la propriété `CurrentView` la renvoie :

```vba
Dim obj As AccessObject
Set obj = CurrentProject.AllForms("frmClients")
If obj.IsLoaded Then
    Debug.Print obj.CurrentView   ' ex. acCurViewFormBrowse (mode Formulaire)
End If
```

> ℹ️ Il existe une méthode plus ancienne pour tester l'état d'un objet, `SysCmd(acSysCmdGetObjectState, …)`, présentée en section 4.7. `IsLoaded` lui est aujourd'hui préférée pour sa simplicité.

## Les constantes de type (acForm, acTable...)

La propriété `Type` d'un `AccessObject` renvoie une constante indiquant la nature de l'objet. Ces mêmes constantes servent aussi avec `DoCmd` (chapitre 5).

| Constante | Type d'objet |
|---|---|
| `acTable` | Table |
| `acQuery` | Requête |
| `acForm` | Formulaire |
| `acReport` | État |
| `acMacro` | Macro |
| `acModule` | Module |

## Récapitulatif

| Collection | Conteneur | Contenu | Élément |
|---|---|---|---|
| `Forms` / `Reports` | `Application` | objets **ouverts** | objet réel (`Form`/`Report`) |
| `AllForms` / `AllReports` / `AllMacros` / `AllModules` | `CurrentProject` | **tous** les objets applicatifs | descripteur (`AccessObject`) |
| `AllTables` / `AllQueries` | `CurrentData` | **tous** les objets de données | descripteur (`AccessObject`) |

## À retenir

- Les collections **`All…`** (`AllForms`, `AllReports`, `AllTables`, `AllQueries`, `AllMacros`, `AllModules`) énumèrent **tous** les objets d'un type, ouverts **ou non**.
- Elles contiennent des **`AccessObject`** — de simples **descripteurs** (`Name`, `IsLoaded`, `CurrentView`, `Type`, dates…), pas les objets eux-mêmes : pour agir, il faut ouvrir l'objet réel.
- À distinguer de **`Forms`**/**`Reports`**, qui ne listent que les objets **ouverts** sous forme d'objets réels manipulables.
- Comme `Collection`, ces collections n'ont **pas de `Exists`** : on teste l'existence en les parcourant.
- **`IsLoaded`** est le **test canonique** pour savoir si un objet est ouvert (préférable à l'interception d'erreur ; alternative plus ancienne : `SysCmd`, section 4.7).

---


⏭️ [4.5. DBEngine, Workspace et Database (DAO)](/04-modele-objet-access/05-dbengine-workspace-database.md)
