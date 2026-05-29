🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.10 Tri et filtrage côté `Recordset`

## Introduction

Après la recherche d'un enregistrement précis (section [9.9](/09-dao-data-access-objects/09-recherche-recordset.md)), voici le tri et le filtrage d'un jeu d'enregistrements dans son ensemble. DAO propose pour cela deux propriétés du `Recordset`, `Filter` et `Sort`. Leur fonctionnement réserve toutefois une surprise de taille, qui est la première chose à comprendre : contrairement à ce que leur nom suggère, ces propriétés ne modifient **pas** le `Recordset` sur lequel on les applique. Cette particularité, et l'alternative souvent préférable du tri/filtrage à la source, font l'objet de cette section.

## Le comportement à bien comprendre

Affecter une valeur à `Filter` ou à `Sort` ne filtre ni ne trie le `Recordset` courant. Ces propriétés **définissent un critère** qui ne s'appliquera qu'à un **nouveau `Recordset` dérivé** du courant, créé par la méthode `OpenRecordset` appelée sur ce dernier. Autrement dit, on règle `Filter` ou `Sort`, puis on ouvre un `Recordset` enfant qui, lui, reflète le filtre ou le tri demandé.

C'est la source de confusion la plus fréquente sur ces propriétés : on les affecte, on parcourt le `Recordset` d'origine, et l'on constate qu'il n'a pas changé — ce qui est le comportement attendu. Ces propriétés s'appliquent aux types **Dynaset** et **Snapshot**.

## La propriété `Filter`

La propriété `Filter` reçoit un critère qui s'écrit comme une clause `WHERE` sans le mot-clé : `"Ville = 'Rouen'"`. Le motif d'utilisation correct consiste à régler `Filter`, puis à dériver un nouveau `Recordset` :

```vba
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim rsFiltre As DAO.Recordset
Set db = CurrentDb()
Set rs = db.OpenRecordset("tblClients", dbOpenSnapshot)

rs.Filter = "Ville = 'Rouen'"      ' définit le critère…
Set rsFiltre = rs.OpenRecordset()  ' …appliqué au Recordset dérivé

Do While Not rsFiltre.EOF
    Debug.Print rsFiltre!NomClient
    rsFiltre.MoveNext
Loop
```

C'est `rsFiltre`, et non `rs`, qui ne contient que les clients de Rouen.

## La propriété `Sort`

La propriété `Sort` fonctionne sur le même principe. Elle reçoit un critère de tri qui s'écrit comme une clause `ORDER BY` sans le mot-clé : `"NomClient ASC"` ou `"Ville, NomClient DESC"`. L'ordre n'est appliqué qu'au `Recordset` dérivé :

```vba
rs.Sort = "NomClient ASC"
Set rsTrie = rs.OpenRecordset()
```

On peut bien sûr combiner les deux propriétés avant de dériver le `Recordset` :

```vba
rs.Filter = "Ville = 'Rouen'"
rs.Sort = "NomClient ASC"
Set rsResultat = rs.OpenRecordset()   ' filtré ET trié
```

## Le cas du type Table : la propriété `Index`

Pour un `Recordset` de type **Table**, le tri ne passe pas par la propriété `Sort` mais par la propriété **`Index`** : désigner un index détermine l'ordre dans lequel les enregistrements sont parcourus. À la différence de `Sort`, il s'agit ici d'un ordonnancement **du `Recordset` courant lui-même**, et non d'un jeu dérivé.

```vba
Set rs = db.OpenRecordset("tblClients", dbOpenTable)
rs.Index = "PrimaryKey"            ' parcours dans l'ordre de cet index
```

Cette même propriété `Index` est, on l'a vu, indispensable à la méthode `Seek` (section 9.9). L'index doit exister sur la table ; sa création par code relève de la section [12.5](/12-querydefs-tabledefs/05-creer-modifier-tables.md).

## Côté `Recordset` ou à la source ?

Comme pour la recherche, il faut prendre du recul. Dans la grande majorité des cas, il est **préférable de filtrer et trier dès la source**, en exprimant `WHERE` et `ORDER BY` directement dans le SQL au moment d'ouvrir le `Recordset` :

```vba
Set rs = db.OpenRecordset( _
    "SELECT * FROM tblClients WHERE Ville='Rouen' ORDER BY NomClient", _
    dbOpenSnapshot)
```

Cette approche est plus efficace — le moteur exploite ses index et ne rapatrie que les données utiles — plus lisible, et elle évite le comportement déroutant du `Recordset` dérivé. Les propriétés `Filter` et `Sort` gardent néanmoins leur utilité lorsqu'on dispose **déjà** d'un `Recordset` ouvert et qu'on souhaite en tirer plusieurs vues filtrées ou triées sans relancer de requête vers la source — par exemple sur un jeu déjà rapatrié, ou dans le contexte du `RecordsetClone` d'un formulaire (section [9.12](/09-dao-data-access-objects/12-recordsetclone.md)). Le filtrage à la source par `SELECT` est traité à la section [11.2](/11-sql-access-vba/02-select-recordset.md).

## Ne pas confondre avec le filtrage des formulaires

Il existe un mécanisme **distinct** et souvent confondu : les formulaires possèdent leurs propres propriétés `Filter` et `OrderBy`, accompagnées de `FilterOn` et `OrderByOn`. Celles-ci, contrairement aux propriétés homonymes du `Recordset`, filtrent et trient **en place** les enregistrements affichés dans le formulaire. Ce sont deux choses différentes : les propriétés du `Recordset`, objet de cette section, dérivent un nouveau jeu ; les propriétés du formulaire agissent sur l'affichage courant. Le filtrage des formulaires par code est traité à la section [6.5](/06-formulaires/05-recordsource-filtres-dynamiques.md).

## Format des critères

Le critère de `Filter` obéit aux mêmes règles que celui des méthodes `Find` (section 9.9) : valeurs texte entre apostrophes ou guillemets, dates encadrées de dièses et au **format américain**, vigilance sur les chaînes contenant une apostrophe. Le critère de `Sort`, lui, liste des champs suivis de `ASC` ou `DESC`, séparés par des virgules. La localisation des dates dans ce type d'expression est détaillée à la section [11.7](/11-sql-access-vba/07-localisation-formatage-sql.md).

## Effacer un filtre ou un tri

Pour annuler un filtre, on réaffecte une chaîne vide à la propriété, avant de dériver à nouveau un `Recordset` :

```vba
rs.Filter = ""                     ' supprime le filtre
Set rsComplet = rs.OpenRecordset() ' jeu dérivé, sans restriction
```

Le même principe vaut pour `Sort`.

## Pièges courants

Les erreurs récurrentes sont les suivantes. Croire que `rs.Filter = …` ou `rs.Sort = …` modifie le `Recordset` courant, alors qu'il faut dériver un nouveau jeu par `OpenRecordset` pour voir l'effet. Confondre les propriétés `Filter`/`Sort` du `Recordset` avec les propriétés homonymes des formulaires, qui agissent en place. Tenter d'utiliser `Sort` sur un type Table (où c'est `Index` qui ordonne). Mal délimiter un critère de `Filter`, comme pour `Find`. Et, plus largement, filtrer en VBA un jeu complet là où une clause `WHERE` à la source aurait été plus efficace.

## Points clés à retenir

Les propriétés `Filter` et `Sort` d'un `Recordset` ne modifient pas le jeu courant : elles définissent un critère appliqué à un **nouveau `Recordset` dérivé** par `OpenRecordset`, sur les types Dynaset et Snapshot. Pour un type Table, l'ordre se règle via la propriété `Index`, qui agit cette fois en place. Dans la plupart des situations, filtrer et trier dès la source par `WHERE` et `ORDER BY` est préférable — plus efficace et plus clair ; `Filter` et `Sort` s'avèrent surtout utiles pour dériver des vues d'un `Recordset` déjà ouvert, comme celui d'un formulaire. Enfin, on prendra garde à ne pas confondre ces propriétés avec les propriétés de filtrage et de tri propres aux formulaires, qui obéissent à une tout autre logique.

⏭️ [9.11. Clôture et libération des objets DAO](/09-dao-data-access-objects/11-cloture-liberation-dao.md)
