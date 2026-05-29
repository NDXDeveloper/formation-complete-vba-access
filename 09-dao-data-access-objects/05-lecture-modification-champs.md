🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.5 Lecture et modification des champs (`Fields`)

## Introduction

Naviguer dans un `Recordset` (section [9.4](/09-dao-data-access-objects/04-navigation-recordset.md)) positionne le curseur sur un enregistrement ; encore faut-il accéder à ses **valeurs**. C'est le rôle de la collection `Fields` : chacun de ses éléments représente une colonne de l'enregistrement courant, dont on lit ou modifie le contenu. Cette section détaille les différentes façons d'accéder à un champ, la lecture de sa valeur, le traitement crucial des valeurs `Null`, et le principe de l'écriture — étant entendu que le déroulé complet de la modification (`Edit`/`Update`) et de l'ajout (`AddNew`/`Update`) fait l'objet des deux sections suivantes.

## La collection `Fields`

Tout `Recordset` expose une collection `Fields` qui regroupe ses colonnes. Chaque `Field` correspond à une colonne et porte, outre sa valeur, un ensemble de propriétés décrivant cette colonne (nom, type, taille…). La propriété `rs.Fields.Count` donne le nombre de champs. On accède à un champ donné par son nom ou par son indice, exactement comme pour les autres collections DAO décrites à la section [9.1](/09-dao-data-access-objects/01-architecture-dao.md).

## Les façons d'accéder à un champ

DAO offre plusieurs syntaxes équivalentes pour désigner la valeur d'un champ. Les trois écritures suivantes renvoient la même chose, car `Value` est la propriété par défaut de l'objet `Field` :

```vba
x = rs!NomClient                  ' bang : la plus concise et la plus lisible
x = rs.Fields("NomClient")        ' accès explicite par nom
x = rs.Fields("NomClient").Value  ' forme entièrement explicite
```

Quelques cas particuliers imposent une syntaxe précise. Lorsque le nom du champ **contient un espace** ou un caractère spécial, on encadre ce nom de crochets avec le bang, ou l'on passe par `Fields` :

```vba
x = rs![Code Postal]
x = rs.Fields("Code Postal")
```

Surtout, lorsque le nom du champ est **contenu dans une variable**, le bang n'est pas utilisable : il faut alors obligatoirement recourir à `Fields` :

```vba
Dim nomChamp As String
nomChamp = "NomClient"
x = rs.Fields(nomChamp)           ' correct
' x = rs!nomChamp                 ' incorrect : chercherait un champ nommé "nomChamp"
```

L'accès par **indice** (`rs.Fields(0)`) existe également, mais il est déconseillé dans la plupart des cas : l'ordre des colonnes peut changer, ce qui rend le code fragile. On privilégie donc l'accès **par nom** — `rs!Champ` pour sa lisibilité dans le cas courant, `rs.Fields("...")` pour les noms à espaces ou dynamiques.

## Lire la valeur d'un champ

La lecture d'un champ se réduit à une affectation :

```vba
Dim montant As Currency
montant = rs!Montant
```

Les valeurs renvoyées par un champ sont de type `Variant` et sont converties à l'affectation vers le type de la variable réceptrice. Pour des conversions explicites ou délicates, on recourt aux fonctions usuelles (`CStr`, `CLng`, `CDate`…), abordées à la section [3.6](/03-rappels-fondamentaux/06-chaines-dates.md). Mais avant toute conversion, un cas mérite la plus grande attention : celui des valeurs absentes.

## Le cas des valeurs `Null`

Un champ peut ne contenir **aucune valeur** : il vaut alors `Null`. C'est sans doute le point qui cause le plus d'erreurs lors de la lecture des champs, car `Null` se propage : toute opération impliquant `Null` produit `Null`, et l'affectation d'un champ `Null` à une variable de type **non-Variant** (`String`, `Long`, `Date`…) déclenche l'erreur d'exécution **94** (« Utilisation incorrecte de Null »).

Première règle : on teste la présence d'une valeur avec la fonction `IsNull`, jamais avec une comparaison `= Null` (qui renvoie toujours `Null`, donc jamais `True`).

```vba
If IsNull(rs!DateNaissance) Then
    Debug.Print "Date inconnue"
Else
    Debug.Print rs!DateNaissance
End If
```

Deuxième outil : la fonction `Nz`, qui substitue une valeur de remplacement lorsque le champ est `Null`.

```vba
Dim solde As Currency
solde = Nz(rs!Solde, 0)           ' 0 si le champ est Null
```

Troisième astuce, propre aux chaînes : l'opérateur de concaténation `&` convertit `Null` en chaîne vide (contrairement à l'opérateur `+`, qui propage `Null`). Affecter directement un champ potentiellement `Null` à une variable `String` lève une erreur, mais lui adjoindre `& ""` la neutralise :

```vba
Dim commentaire As String
' commentaire = rs!Commentaire        ' ERREUR 94 si le champ est Null
commentaire = rs!Commentaire & ""     ' OK : Null converti en chaîne vide
commentaire = Nz(rs!Commentaire, "")  ' équivalent, plus explicite
```

En pratique, dès qu'un champ est susceptible d'être `Null` — ce qui est le cas de tout champ non obligatoire — on protège systématiquement sa lecture par `IsNull` ou `Nz`.

## Écrire dans un champ

Modifier la valeur d'un champ s'écrit par une simple affectation :

```vba
rs!Telephone = "0235123456"
```

Mais cette affectation n'est autorisée **que si le `Recordset` est en mode édition ou ajout**. On y entre par `Edit` (pour modifier l'enregistrement courant) ou `AddNew` (pour en créer un nouveau), puis l'on valide par `Update`. Voici le motif minimal d'une modification :

```vba
rs.Edit                       ' entre en mode modification
rs!Telephone = "0235123456"
rs!DateMAJ = Now()
rs.Update                     ' valide les changements
```

Toute écriture hors de ce cadre échoue par une erreur d'exécution ; en particulier, appeler `Update` sans `Edit` ni `AddNew` préalable déclenche l'erreur **3020**. Le déroulé complet, les pièges et les bonnes pratiques de ces opérations sont traités aux sections [9.6](/09-dao-data-access-objects/06-ajout-enregistrements.md) (ajout, `AddNew`/`Update`) et [9.7](/09-dao-data-access-objects/07-modification-enregistrements.md) (modification, `Edit`/`Update`). La présente section se concentre sur l'**accès aux champs** lui-même.

## Les propriétés du champ

Au-delà de sa valeur, un objet `Field` porte des **métadonnées** souvent utiles. Les plus courantes sont `Name` (le nom du champ), `Type` (le type de données, exprimé par une constante DAO comme `dbText`, `dbLong`, `dbDate`, `dbCurrency`…) et `Size` (la taille, pour les champs texte). On peut parcourir l'ensemble des champs d'un enregistrement :

```vba
Dim fld As DAO.Field
For Each fld In rs.Fields
    Debug.Print fld.Name, fld.Type, fld.Size
Next fld
```

La propriété `OldValue` mérite une mention particulière : elle renvoie la valeur du champ **avant** la modification en cours, ce qui est précieux pour la détection des conflits de mise à jour en environnement multi-utilisateur (sujet traité à la section [14.5](/14-transactions/05-conflits-mise-a-jour.md)). D'autres propriétés (`Required`, `AllowZeroLength`, `DefaultValue`, `ValidationRule`) décrivent les contraintes de la colonne ; elles relèvent davantage de la définition de la table, abordée au chapitre 12, mais restent lisibles depuis le `Recordset`.

## Respecter le type et les contraintes

L'écriture d'un champ doit être cohérente avec son type et ses contraintes, faute de quoi DAO lève une erreur au moment de l'`Update`. Affecter une chaîne plus longue que la taille du champ, écrire une valeur `Null` dans un champ obligatoire, ou affecter une valeur incompatible avec le type attendu provoquent des erreurs qu'il convient d'anticiper. Pour les dates, on affecte une valeur de type `Date` ; pour un champ booléen, une valeur `True`/`False` ; et ainsi de suite. Lorsque les données proviennent d'une source dont la forme est incertaine (saisie utilisateur, import), une validation préalable et, au besoin, une conversion explicite évitent ces écueils.

## Pièges courants

Les principales sources d'erreur tiennent à trois points. Le premier, et de loin le plus fréquent, est la **lecture d'un champ `Null`** dans une variable typée, qui provoque l'erreur 94 ; on s'en prémunit par `IsNull` ou `Nz`. Le deuxième est l'usage du **bang avec une variable** (`rs!nomChamp` au lieu de `rs.Fields(nomChamp)`), qui cherche un champ portant littéralement ce nom. Le troisième est l'**écriture hors mode édition**, qui exige `Edit` ou `AddNew` au préalable. L'accès par indice plutôt que par nom, enfin, fragilise le code face aux changements d'ordre des colonnes.

## Points clés à retenir

La collection `Fields` donne accès aux colonnes de l'enregistrement courant. On lit et écrit la valeur d'un champ via le bang (`rs!Champ`) ou via `rs.Fields("Champ")`, cette seconde forme étant requise pour les noms à espaces ou contenus dans une variable ; `Value` étant la propriété par défaut, ces écritures sont équivalentes. Le traitement des valeurs `Null` est incontournable : on les teste par `IsNull` et on les neutralise par `Nz` ou par l'opérateur `&`, sous peine d'erreur 94. L'écriture d'un champ n'est possible qu'après `Edit` ou `AddNew`, suivie d'`Update` — motif dont le détail est l'objet des sections 9.6 et 9.7. Enfin, l'objet `Field` expose, au-delà de sa valeur, des métadonnées utiles telles que `Name`, `Type`, `Size` et `OldValue`.

⏭️ [9.6. Ajout d'enregistrements (AddNew/Update)](/09-dao-data-access-objects/06-ajout-enregistrements.md)
