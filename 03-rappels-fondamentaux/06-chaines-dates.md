🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 3.6. Gestion des chaînes de caractères et des dates

Sixième rappel : la manipulation des **chaînes de caractères** et des **dates**. Ces deux sujets, banals en apparence, portent en réalité des **enjeux particuliers sous Access** — la concaténation de chaînes sert à construire du SQL, et le formatage des dates dans le SQL est l'une des sources de bugs les plus classiques. Au-delà des fonctions de base, cette section insiste donc sur ces pièges.

## Les chaînes de caractères

### Concaténer : toujours avec &

Pour assembler des chaînes, VBA offre deux opérateurs, mais un seul est recommandé : **`&`**. L'opérateur `+` fonctionne aussi, mais se comporte mal avec les nombres et surtout avec `Null`, qu'il **propage** au lieu de l'ignorer.

```vba
Dim nom As String, prenom As String
nom = "Dupont" : prenom = "Marie"
Debug.Print prenom & " " & nom        ' "Marie Dupont" — toujours &

' & traite Null comme une chaîne vide ; + propage Null
Debug.Print "a" & Null                ' "a"
' Debug.Print "a" + Null              ' Null !
```

Cette différence est cruciale sous Access, où l'on concatène souvent des **champs susceptibles d'être `Null`** : `&` permet de le faire sans accident.

### Les fonctions de chaînes essentielles

Le tableau suivant récapitule les fonctions les plus courantes.

| Fonction | Rôle | Exemple → résultat |
|---|---|---|
| `Len(s)` | longueur | `Len("abc")` → 3 |
| `Left(s, n)` | n premiers caractères | `Left("Bonjour", 3)` → "Bon" |
| `Right(s, n)` | n derniers caractères | `Right("Bonjour", 4)` → "jour" |
| `Mid(s, d, n)` | n caractères à partir de d | `Mid("Bonjour", 2, 3)` → "onj" |
| `InStr(s, x)` | position de x (0 si absent) | `InStr("abc", "b")` → 2 |
| `InStrRev(s, x)` | position en partant de la fin | `InStrRev("a.b.c", ".")` → 4 |
| `Replace(s, a, b)` | remplace a par b | `Replace("a-b", "-", "+")` → "a+b" |
| `UCase` / `LCase` | majuscules / minuscules | `UCase("abc")` → "ABC" |
| `Trim(s)` | supprime les espaces aux extrémités | `Trim("  x  ")` → "x" |
| `Format(v, f)` | met en forme | `Format(1234.5, "#,##0.00")` → "1 234,50" |
| `Chr(c)` / `Asc(x)` | code ↔ caractère | `Chr(65)` → "A" |

On y ajoute `LTrim`/`RTrim`, `Space(n)`, `String(n, c)`, `StrComp`, ainsi que `Split` et `Join` vues en section 3.5.

### Les variantes en $

Beaucoup de ces fonctions existent en une version suffixée de **`$`** (`Left$`, `Right$`, `Mid$`, `Trim$`, `UCase$`…). Ces variantes renvoient directement un **`String`** (et non un `Variant`), ce qui est légèrement plus rapide et plus rigoureux en typage. À utiliser quand on travaille à coup sûr sur des chaînes.

### Comparaison et Option Compare

La façon dont VBA **compare** les chaînes dépend de l'instruction `Option Compare` placée en tête de module (section 2.3) :

- **`Option Compare Database`** (le réglage **par défaut** sous Access) : comparaison selon l'ordre de tri de la base, généralement **insensible à la casse** ;
- **`Option Compare Binary`** : comparaison binaire, **sensible à la casse** ;
- **`Option Compare Text`** : comparaison textuelle, **insensible à la casse**.

Ce réglage influence les opérateurs de comparaison (`=`, `<`…) **et** des fonctions comme `InStr`. Le connaître évite des surprises sur la casse.

### Constantes utiles

Pour les caractères spéciaux, VBA fournit des constantes : `vbCrLf` (retour à la ligne), `vbTab` (tabulation), `vbNullString` (chaîne vide), ainsi que `vbCr` et `vbLf`.

```vba
MsgBox "Ligne 1" & vbCrLf & "Ligne 2"
```

## Les dates

### Le type Date (rappel)

Comme vu en section 3.1, une `Date` est, en interne, un **`Double`** (partie entière = le jour, partie décimale = l'heure), et les **littéraux** s'écrivent entre `#`, **toujours au format américain** (`#12/31/2025#`). Ce dernier point est au cœur du piège SQL traité plus bas.

### Obtenir la date et l'heure

```vba
Debug.Print Now      ' date + heure courantes
Debug.Print Date     ' date seule
Debug.Print Time     ' heure seule
```

### Construire et décomposer une date

Pour fabriquer une date **sans aucune ambiguïté de format**, `DateSerial` est préférable aux littéraux, surtout quand les composantes viennent de variables :

```vba
Dim d As Date
d = DateSerial(2025, 12, 31)        ' 31 décembre 2025, sans ambiguïté
```

À l'inverse, on en extrait les composantes avec `Year`, `Month`, `Day`, `Hour`, `Minute`, `Second`, `Weekday` (ou `DatePart`) :

```vba
Debug.Print Year(d), Month(d), Day(d)   ' 2025  12  31
```

### Calculer avec les dates

Deux fonctions clés : `DateAdd` (ajouter un intervalle) et `DateDiff` (calculer un écart). L'intervalle est désigné par un code (`"d"` jour, `"m"` mois, `"yyyy"` année, `"h"` heure…).

```vba
Debug.Print DateAdd("d", 7, d)      ' une semaine plus tard
Debug.Print DateDiff("d", Date, d)  ' nombre de jours d'aujourd'hui jusqu'à d
```

Puisqu'une date est un `Double`, l'arithmétique directe fonctionne aussi (`d + 1` donne le lendemain), mais `DateAdd`/`DateDiff` sont plus clairs et gèrent correctement les mois et les années.

### Mettre en forme et analyser

`Format` produit une représentation **lisible** d'une date (formats personnalisés ou nommés comme « Date, abrégé ») ; il respecte les **paramètres régionaux** pour les formats nommés — donc l'affichage. À l'inverse, `CDate` et `DateValue` **analysent** une chaîne pour la convertir en date, en se fiant eux aussi aux **paramètres régionaux** : c'est commode pour l'affichage, mais une source de bugs dès qu'on bascule dans le SQL. `IsDate` permet de tester au préalable.

## Le point critique : chaînes et dates dans le SQL

C'est ici que se concentre l'essentiel des difficultés sous Access.

### Les dates dans le SQL Access

Le moteur SQL d'Access n'interprète correctement les dates que dans un format **non ambigu** : le format **américain** (`MM/DD/YYYY`) ou le format **ISO** (`YYYY-MM-DD`), entourés de `#`. Si l'on insère dans une requête une date au **format local** (par exemple `03/04/2025` interprété à la française), elle sera **mal comprise** — `3 avril` devenant `4 mars`, ou la requête échouant.

La technique robuste consiste donc à **formater explicitement** la date au moment de bâtir le SQL :

```vba
Dim sql As String
sql = "SELECT * FROM tblCommandes WHERE DateCmd >= #" _
    & Format(d, "mm/dd/yyyy") & "#"
```

C'est l'un des pièges les plus fréquents du développement Access. Il est traité en détail — avec ses bonnes pratiques — en **section 11.7**.

### Les chaînes dans le SQL

De même, une valeur texte insérée dans une requête doit être **délimitée par des guillemets**, et une **apostrophe** présente dans la donnée (comme dans « O'Brien ») rompt la requête. Plutôt que de jongler manuellement avec l'échappement, la solution propre est le **SQL paramétré**, qui élimine à la fois ce problème et les risques d'injection. Le sujet est développé aux **sections 11.5 et 11.6**.

### Les Null avec Nz

Enfin, en concaténant des champs susceptibles d'être `Null` (section 3.1), la fonction **`Nz`** d'Access permet de substituer une valeur de remplacement :

```vba
Debug.Print "Tél : " & Nz(rs!Telephone, "(non renseigné)")
```

`Nz` est détaillée avec l'accès aux données en section 9.5.

## À retenir

- Concaténez **toujours avec `&`** (et non `+`) : `&` traite `Null` comme une chaîne vide, ce qui compte pour les **champs Access**.
- Maîtrisez les fonctions de base (`Len`, `Left`/`Right`/`Mid`, `InStr`, `Replace`, `Trim`, `Format`…) ; leurs variantes en **`$`** renvoient un `String` et sont un peu plus rapides.
- La comparaison de chaînes dépend d'**`Option Compare`** (Database par défaut sous Access, généralement insensible à la casse).
- Construisez les dates dynamiques avec **`DateSerial`** (sans ambiguïté) et calculez avec **`DateAdd`/`DateDiff`** ; `Format` et `CDate` dépendent des **paramètres régionaux**.
- **Piège majeur** : dans le **SQL Access**, les dates doivent être au format **US ou ISO entre `#`** (`Format(d, "mm/dd/yyyy")`) — voir section 11.7 ; pour les **chaînes**, préférez le **SQL paramétré** (sections 11.5/11.6) ; utilisez **`Nz`** pour les `Null` (section 9.5).

---


⏭️ [3.7. Collections et Dictionary (Scripting.Dictionary)](/03-rappels-fondamentaux/07-collections-dictionary.md)
