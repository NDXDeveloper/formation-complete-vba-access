🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 3.5. Tableaux statiques et dynamiques

Cinquième rappel : les **tableaux**, qui regroupent plusieurs valeurs de même type sous un seul nom, accessibles par un **indice**. VBA distingue les tableaux **statiques** (taille fixe) et **dynamiques** (taille ajustable). Cette section revoit les deux, ainsi que les fonctions associées, en signalant leurs usages sous Access.

## Qu'est-ce qu'un tableau ?

Plutôt que de déclarer douze variables `mois1`, `mois2`, … `mois12`, on déclare un **tableau** `mois` à douze cases, que l'on adresse par leur **indice** : `mois(1)`, `mois(2)`, etc. C'est la structure de base pour manipuler des collections de valeurs homogènes.

## Les tableaux statiques

Un tableau **statique** a une taille **fixée à la déclaration**. Le nombre entre parenthèses indique l'indice **maximal** ; par défaut, les indices commencent à **0**.

```vba
Dim notes(9) As Long          ' 10 éléments, indices 0 à 9
Dim mois(1 To 12) As String   ' 12 éléments, indices 1 à 12 (bornes explicites)

notes(0) = 15
mois(1) = "Janvier"
```

La forme `(1 To 12)` précise explicitement les bornes inférieure et supérieure. Elle est **préférable** au réglage `Option Base 1` (qui modifie la borne inférieure par défaut au niveau du module) : déclarer les bornes en clair rend le code sans ambiguïté.

## Les tableaux dynamiques

Un tableau **dynamique** se déclare **sans taille**, puis se dimensionne (et se redimensionne) avec `ReDim` :

```vba
Dim valeurs() As Long          ' taille inconnue à la déclaration
ReDim valeurs(4)               ' dimensionné : indices 0 à 4
```

Pour **changer la taille en conservant** le contenu existant, on ajoute `Preserve` :

```vba
ReDim Preserve valeurs(9)      ' agrandi à 10 cases, les valeurs déjà présentes sont gardées
```

Deux limites à connaître pour `ReDim Preserve` : il ne peut redimensionner que la **dernière dimension** d'un tableau multidimensionnel, et il a un **coût** (il recopie les données). On évite donc de l'appeler dans une boucle serrée case par case.

Enfin, `Erase` réinitialise un tableau — il **libère** la mémoire d'un tableau dynamique, et remet à leur valeur par défaut les éléments d'un tableau statique :

```vba
Erase valeurs                  ' libère le tableau dynamique
```

## Accéder aux éléments et connaître les bornes

On lit ou écrit un élément par son indice : `notes(0) = 15`. Un indice **hors limites** déclenche l'erreur d'exécution n° 9 (« L'indice n'appartient pas à la sélection »).

Pour parcourir un tableau **sans coder ses bornes en dur**, deux fonctions essentielles : `LBound` (borne inférieure) et `UBound` (borne supérieure).

```vba
Dim i As Long
For i = LBound(notes) To UBound(notes)
    Debug.Print notes(i)
Next i
```

C'est le motif de parcours à privilégier : il reste correct même si la taille du tableau change.

## Les tableaux multidimensionnels

Un tableau peut avoir plusieurs dimensions (jusqu'à 60, mais on dépasse rarement deux ou trois).

```vba
Dim grille(1 To 3, 1 To 4) As Long
grille(2, 3) = 7
Debug.Print UBound(grille, 1)   ' 3 (1re dimension)
Debug.Print UBound(grille, 2)   ' 4 (2e dimension)
```

`LBound` et `UBound` acceptent un second argument pour cibler une **dimension** précise.

## Parcourir un tableau

Outre `For…Next` avec `LBound`/`UBound`, on peut utiliser `For Each` — mais avec une **nuance importante** : la variable de boucle est une **copie** des éléments (pour un tableau de valeurs). On peut donc **lire** les éléments, mais **pas les modifier** à travers elle.

```vba
Dim v As Variant
For Each v In notes
    Debug.Print v               ' lecture OK ; v est une copie -> modifier v ne change pas le tableau
Next v
```

Pour **modifier** les éléments, il faut donc revenir à `For…Next` avec l'indice.

## Fonctions utiles

Quelques fonctions facilitent grandement le travail avec les tableaux :

```vba
Dim v As Variant
v = Array("Lun", "Mar", "Mer")        ' crée un tableau Variant à partir d'une liste

Dim parts() As String
parts = Split("a,b,c", ",")           ' découpe une chaîne -> "a","b","c" (indices 0 à 2)

Dim s As String
s = Join(parts, ";")                  ' recompose une chaîne -> "a;b;c"
```

- **`Array`** crée un tableau `Variant` à partir d'une liste de valeurs.
- **`Split`** découpe une chaîne en tableau selon un séparateur (résultat indexé à partir de 0).
- **`Join`** réalise l'opération inverse.
- **`Filter`** extrait d'un tableau de chaînes celles qui correspondent à un motif.
- **`IsArray`** indique si une variable est un tableau.

`Split` et `Join` sont étudiées plus en détail avec les chaînes de caractères (section 3.6).

## Passer et renvoyer des tableaux

Un tableau se passe en argument **par référence** ; le paramètre se déclare avec des parenthèses vides : `arr() As Type`. Une **fonction** peut aussi **renvoyer** un tableau, en typant son retour avec des parenthèses :

```vba
Function ListeMois() As String()
    Dim t(1 To 3) As String
    t(1) = "Jan" : t(2) = "Fév" : t(3) = "Mar"
    ListeMois = t
End Function
```

## Côté Access : GetRows et Split

Dans une application Access, on manipule généralement les données via des **recordsets** plutôt que des tableaux. Deux ponts méritent néanmoins d'être signalés.

D'abord, la méthode **`GetRows`** d'un recordset (DAO ou ADO) charge d'un coup plusieurs enregistrements dans un **tableau à deux dimensions** — une technique rapide pour traiter des données en mémoire, utile pour les performances (sections 9 et 18.2). Attention à l'ordre des indices : le tableau est organisé en **`(champ, ligne)`**.

```vba
Dim donnees As Variant
donnees = rs.GetRows(100)             ' jusqu'à 100 enregistrements
Debug.Print donnees(0, 0)            ' 1er champ, 1re ligne
```

Ensuite, **`Split`** est précieux pour analyser des chaînes structurées que l'on rencontre couramment sous Access : valeurs délimitées, ou la chaîne passée à un formulaire via `OpenArgs` (section 6.6).

## À retenir

- Un tableau **statique** a une taille fixe (`Dim notes(9)` ou `Dim mois(1 To 12)`) ; les indices commencent à **0** par défaut — préférez des **bornes explicites** à `Option Base`.
- Un tableau **dynamique** (`Dim arr()`) se dimensionne avec **`ReDim`**, et `ReDim Preserve` redimensionne **en conservant** les valeurs (uniquement la **dernière dimension**, à un certain coût).
- Parcourez avec **`LBound`/`UBound`** plutôt qu'avec des tailles codées en dur ; un indice hors limites lève l'**erreur n° 9**.
- `For Each` lit les éléments via une **copie** : pour les **modifier**, utilisez `For…Next` avec l'indice.
- Pensez à **`Array`**, **`Split`** et **`Join`** ; sous Access, **`GetRows`** remplit un tableau 2D `(champ, ligne)` depuis un recordset (sections 9 et 18.2).

---


⏭️ [3.6. Gestion des chaînes de caractères et des dates](/03-rappels-fondamentaux/06-chaines-dates.md)
