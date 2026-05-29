🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 3.1. Variables, types de données et constantes

Premier des rappels : les **variables**, qui stockent les données manipulées par le code, leurs **types**, qui déterminent ce qu'elles peuvent contenir, et les **constantes**, valeurs fixes nommées. Fidèle à l'esprit de ce chapitre — un condensé de référence —, cette section va à l'essentiel, en signalant les points qui méritent une attention particulière dans le contexte Access.

## Déclarer une variable

Une variable se déclare avec l'instruction `Dim`, suivie de son nom, du mot-clé `As` et de son type :

```vba
Dim compteur As Long
Dim nomClient As String
Dim estActif As Boolean
```

Avec l'option **Déclaration des variables obligatoire** (`Option Explicit`, vivement recommandée — section 1.4), toute variable doit être déclarée avant usage, ce qui évite quantité d'erreurs. Une variable déclarée **sans type** est de type `Variant` (voir plus bas).

Attention à un **piège classique** lors des déclarations multiples sur une même ligne : chaque variable a besoin de **son propre** `As Type`.

```vba
Dim x, y As Integer              ' ⚠️ x est un Variant, seul y est un Integer !
Dim a As Integer, b As Integer   ' correct : a et b sont tous deux des Integer
```

## Les types de données

VBA propose un éventail de types. Le tableau suivant les récapitule.

| Type | Contenu | Taille | Remarque |
|---|---|---|---|
| **Boolean** | `True` / `False` | 2 o | |
| **Byte** | entier 0 à 255 | 1 o | |
| **Integer** | entier −32 768 à 32 767 | 2 o | préférer `Long` |
| **Long** | entier ≈ ±2,1 milliards | 4 o | type entier **par défaut recommandé** |
| **LongLong** | entier 64 bits | 8 o | VBA 64 bits uniquement |
| **LongPtr** | entier taille pointeur | 4/8 o | pour les **API** (sections 21.7, 22.1) |
| **Single** | réel simple précision | 4 o | |
| **Double** | réel double précision | 8 o | type **décimal par défaut** |
| **Currency** | monétaire, 4 décimales | 8 o | pour les **montants** (évite les arrondis) |
| **Date** | date et heure | 8 o | stocké comme un `Double` |
| **String** | texte | variable | concaténation avec `&` |
| **Object** | référence d'objet | 4/8 o | liaison tardive (section 2.6) |
| **Variant** | tout type | ≥ 16 o | par défaut si non typé ; peut valoir `Null`/`Empty` |

Quelques points méritent d'être soulignés.

**`Long` plutôt qu'`Integer`.** Sur les processeurs actuels, `Long` n'est pas plus coûteux qu'`Integer` et évite les débordements ; c'est le type entier à privilégier par défaut.

**`Currency` pour les montants.** Le type `Currency` stocke les valeurs sous forme d'entier mis à l'échelle (4 décimales), ce qui **élimine les erreurs d'arrondi** propres aux `Double`. C'est le choix indiqué pour les valeurs financières — et il correspond naturellement au type de champ « Monétaire » d'Access.

```vba
Dim prix As Currency
prix = 19.99      ' un montant : Currency, pas Double
```

**`Date` et le format des littéraux.** Une date est, en interne, un `Double` (partie entière = le jour, partie décimale = l'heure). Les **littéraux de date** s'entourent de `#` et s'écrivent **toujours au format américain** (MM/JJ/AAAA), quelle que soit la langue du poste — un point crucial, qui se répercute sur le SQL dynamique (section 11.7).

```vba
Dim d As Date
d = #12/31/2025#   ' 31 décembre 2025 — littéral toujours au format US
d = Now            ' date + heure courantes
```

**`Variant`, `Null` et `Empty`.** Le `Variant` peut tout contenir, mais aussi deux valeurs particulières : `Empty` (variable jamais initialisée) et `Null` (absence de donnée valide). Ce `Null` est omniprésent sous Access, car un **champ de table peut être Null**. Il se distingue de la chaîne vide `""` et de `0`. Pour le neutraliser, Access fournit la fonction `Nz()`, détaillée avec l'accès aux données (section 9.5).

**Le type `Decimal`.** Il existe pour les besoins de très haute précision, mais ne peut pas être déclaré directement avec `Dim` : on l'obtient comme sous-type d'un `Variant`, via la fonction de conversion `CDec`. Usage rare.

## Les variables objet et le mot-clé Set

Les variables d'objet ne contiennent pas une valeur, mais une **référence** vers un objet (une base, un recordset, un formulaire…). Leur affectation **exige** le mot-clé `Set` — son oubli est l'une des erreurs les plus fréquentes des débutants.

```vba
Dim db As DAO.Database
Set db = CurrentDb        ' Set obligatoire pour un objet
Debug.Print db.Name

Dim n As Long
n = 42                    ' pas de Set pour un type valeur
```

Ce point est central sous Access, où l'on manipule en permanence des objets DAO, ADO, des formulaires et des contrôles. Le type `Object` générique sert, lui, à la liaison tardive (section 2.6).

## Valeurs par défaut à l'initialisation

À la déclaration, chaque type reçoit une valeur initiale :

- types numériques → `0` ;
- `Boolean` → `False` ;
- `String` (longueur variable) → `""` (chaîne vide) ;
- `Variant` → `Empty` ;
- `Object` → `Nothing` ;
- `Date` → `0`, soit le 30 décembre 1899.

## La conversion de types

VBA convertit souvent les types **implicitement**, mais il est plus sûr de convertir **explicitement** avec les fonctions dédiées : `CLng`, `CInt`, `CDbl`, `CSng`, `CCur`, `CDate`, `CStr`, `CBool`, `CByte`, `CDec`, `CVar`.

```vba
Dim s As String, n As Long
s = "150"
n = CLng(s)               ' conversion explicite chaîne -> Long
```

Avant de convertir, on peut **tester** la valeur avec les fonctions `IsNumeric`, `IsDate`, `IsNull`, `IsEmpty`, `IsObject`, `IsArray`. À noter : certaines conversions (comme `CDate` ou `CDbl`) tiennent compte des **paramètres régionaux** du poste, ce qui a son importance lors de la construction de SQL — sujet repris en section 11.7.

## Les constantes

Une **constante** est une valeur fixe nommée, définie une fois pour toutes et impossible à modifier en cours d'exécution :

```vba
Const TVA As Double = 0.2
Const APP_NOM As String = "Gestion Ventes"
```

Leur intérêt est triple : **lisibilité** (un nom parlant à la place d'un « nombre magique »), **maintenance** (un seul point à modifier), et **fiabilité** (impossible de l'altérer par erreur). Comme les variables, leur **portée** dépend de l'endroit où on les déclare (locale dans une procédure, `Private` ou `Public` au niveau module) — la portée est traitée en section 3.4.

VBA fournit en outre de nombreuses **constantes intégrées** : celles du langage (`vbCrLf`, `vbYes`, `vbTab`…) et celles des **bibliothèques référencées** (section 2.5), comme `acViewPreview` (Access) ou `dbOpenDynaset` (DAO).

```vba
MsgBox "Ligne 1" & vbCrLf & "Ligne 2"
```

## Les énumérations (Enum)

Pour regrouper un ensemble de constantes liées, on utilise une **énumération**. Ses membres sont des valeurs `Long`, et le type ainsi défini sert à typer des variables, ce qui rend le code plus clair et plus sûr.

```vba
Public Enum StatutCommande
    scBrouillon = 0
    scValidee = 1
    scExpediee = 2
    scAnnulee = 3
End Enum

Dim s As StatutCommande
s = scValidee
```

## Conventions de nommage

Un nom de variable doit commencer par une lettre, ne contenir ni espace ni caractère interdit, et ne pas être un mot réservé. Au-delà de ces règles, l'adoption d'une **convention de nommage** cohérente (préfixes indiquant le type ou la nature, style Leszynski/Reddick) améliore nettement la lisibilité. Ces conventions, spécifiques à Access, sont détaillées en section 24.1 et en annexe G.

## À retenir

- Déclarez vos variables avec `Dim … As Type` ; avec `Option Explicit`, c'est obligatoire — et chaque variable d'une ligne multiple a besoin de **son propre** `As Type`.
- Préférez **`Long`** à `Integer`, **`Double`** pour les décimaux, et surtout **`Currency`** pour les **montants** (pas de `Double` pour l'argent).
- Une **`Date`** est un `Double` en interne ; ses **littéraux** (`#…#`) sont **toujours au format US** — point clé pour le SQL (section 11.7).
- Le **`Variant`** peut valoir `Null` (absence de donnée, fréquent pour les **champs Access**, géré par `Nz()`) ou `Empty` ; à distinguer de `""` et de `0`.
- Les **variables objet** s'affectent avec **`Set`** : son oubli est une erreur classique, et le sujet est omniprésent sous Access (DAO, ADO, formulaires).
- Utilisez des **constantes** et des **énumérations** plutôt que des « nombres magiques » ; VBA et les bibliothèques en fournissent déjà beaucoup (`vbCrLf`, `acViewPreview`, `dbOpenDynaset`…).

---


⏭️ [3.2. Structures de contrôle (If, Select Case, boucles)](/03-rappels-fondamentaux/02-structures-controle.md)
