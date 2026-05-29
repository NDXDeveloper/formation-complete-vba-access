🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.13 Champs multivalués (MVF) et champs Pièce jointe — `Recordset2` et `Field2`

## Introduction

Cette dernière section du chapitre aborde deux types de champs particuliers, introduits avec le format `.accdb` et le moteur ACE (voir l'historique du [README](/09-dao-data-access-objects/README.md) du chapitre) : les **champs multivalués** (MVF, *Multi-Valued Fields*) et les **champs Pièce jointe**. Ces types sortent du cadre classique « une valeur par champ » et nécessitent, pour être manipulés en DAO, des objets étendus spécifiques : `Recordset2` et `Field2`. C'est un domaine plus avancé et relativement spécialisé, qui clôt naturellement le tour des manipulations de données du chapitre.

## Des champs qui sortent du modèle relationnel

Un champ classique contient une valeur unique. Un **champ multivalué** peut, lui, en contenir plusieurs : par exemple un champ « Catégories » stockant simultanément plusieurs étiquettes pour un même enregistrement. Un **champ Pièce jointe** va plus loin encore : il stocke un ou plusieurs fichiers, avec leur contenu binaire et leur nom.

Techniquement, Access implémente ces champs au moyen de **tables enfants masquées** : sous l'apparence d'un champ unique se cache en réalité une relation un-à-plusieurs. C'est ce qui explique que leur valeur ne soit pas un scalaire, mais un véritable **jeu d'enregistrements**. Cette entorse au modèle relationnel a des conséquences pratiques que l'on évoquera en fin de section.

## `Recordset2` et `Field2` : les objets étendus d'ACE

Les objets `Field` et `Recordset` hérités de l'époque Jet ne savent pas traiter ces types complexes. ACE a donc introduit deux objets qui les étendent : `Field2`, qui prolonge `Field`, et `Recordset2`, qui prolonge `Recordset`.

L'idée centrale est la suivante : la propriété `Value` d'un champ complexe ne renvoie pas une valeur simple, mais un **`Recordset2` enfant** — le jeu d'enregistrements de la table masquée. `Field2` ajoute par ailleurs des membres absents de `Field` : la propriété `IsComplex` (pour détecter ces champs) et, pour les pièces jointes, les méthodes `LoadFromFile` et `SaveToFile`.

## Détecter un champ complexe

Avant de traiter un champ, on peut déterminer s'il est complexe via la propriété `IsComplex` de `Field2`, qui vaut `True` pour un MVF comme pour un champ Pièce jointe :

```vba
Dim fld As DAO.Field2
For Each fld In rs.Fields
    If fld.IsComplex Then
        Debug.Print fld.Name & " est un champ complexe (MVF ou pièce jointe)"
    End If
Next fld
```

Le type précis se lit dans la propriété `Type` : la famille `dbComplexText`, `dbComplexLong`, `dbComplexDouble`, etc. désigne les champs multivalués selon leur type sous-jacent, tandis que `dbAttachment` désigne les champs Pièce jointe.

## Lire un champ multivalué

Pour lire un MVF, on récupère son `Recordset2` enfant via `.Value`, puis on le parcourt comme n'importe quel `Recordset`. Pour un MVF simple, chaque valeur se trouve dans un champ enfant nommé `Value` :

```vba
Dim rsChild As DAO.Recordset2
Set rsChild = rs!Categories.Value      ' le champ MVF expose un Recordset2 enfant

Do While Not rsChild.EOF
    Debug.Print rsChild!Value           ' chaque valeur du champ multivalué
    rsChild.MoveNext
Loop

rsChild.Close
Set rsChild = Nothing
```

Toutes les techniques de navigation de la section [9.4](/09-dao-data-access-objects/04-navigation-recordset.md) s'appliquent à ce jeu enfant, et il convient de le libérer comme tout `Recordset` (section [9.11](/09-dao-data-access-objects/11-cloture-liberation-dao.md)).

## Lire un champ Pièce jointe

Le principe est identique, mais le `Recordset2` enfant d'un champ Pièce jointe possède des champs spécifiques, notamment `FileName` (le nom du fichier), `FileType` et `FileData` (le contenu binaire). Ce dernier, un `Field2`, expose la méthode `SaveToFile` qui extrait le fichier vers le disque :

```vba
Dim rsPieces As DAO.Recordset2
Set rsPieces = rs!Documents.Value      ' le champ Pièce jointe expose un Recordset2 enfant

Do While Not rsPieces.EOF
    Debug.Print rsPieces!FileName
    rsPieces!FileData.SaveToFile "C:\Export\" & rsPieces!FileName
    rsPieces.MoveNext
Loop

rsPieces.Close
Set rsPieces = Nothing
```

Un détail pratique : `SaveToFile` échoue si le fichier de destination **existe déjà**. On veillera donc à choisir un chemin libre, ou à supprimer le fichier existant au préalable.

## Modifier un champ complexe : le modèle imbriqué

C'est la partie la plus délicate. Modifier un champ complexe suppose des opérations **imbriquées** : on édite l'enregistrement parent, on manipule le `Recordset2` enfant (par `AddNew`/`Update`), puis on valide le parent. L'ajout d'une valeur à un MVF illustre ce déroulé en quatre temps :

```vba
Dim rsChild As DAO.Recordset2
rs.Edit                                 ' 1. édition de l'enregistrement parent
Set rsChild = rs!Categories.Value       ' 2. recordset enfant du MVF
rsChild.AddNew                          ' 3a. ajout dans le jeu enfant
rsChild!Value = "Urgent"
rsChild.Update                          ' 3b. validation de l'enfant
rs.Update                               ' 4. validation de l'enregistrement parent
```

L'`Update` de l'enfant inscrit la nouvelle valeur dans le champ ; l'`Update` du parent persiste l'enregistrement. Oublier l'un ou l'autre laisse la modification incomplète. Les opérations d'écriture des sections [9.6](/09-dao-data-access-objects/06-ajout-enregistrements.md) à [9.8](/09-dao-data-access-objects/08-suppression-enregistrements.md) s'appliquent ici au jeu enfant.

## Ajouter une pièce jointe

L'ajout d'une pièce jointe suit le même modèle imbriqué, en utilisant cette fois `LoadFromFile`, qui charge un fichier dans le champ `FileData` — et renseigne automatiquement `FileName` :

```vba
Dim rsPieces As DAO.Recordset2
rs.Edit
Set rsPieces = rs!Documents.Value
rsPieces.AddNew
rsPieces!FileData.LoadFromFile "C:\Docs\contrat.pdf"   ' FileName renseigné automatiquement
rsPieces.Update
rs.Update
```

## Déclarer en `Recordset2` et `Field2`

Pour exploiter les membres propres à ces objets — `IsComplex`, `LoadFromFile`, `SaveToFile` — il est recommandé de déclarer les variables concernées explicitement en `DAO.Recordset2` et `DAO.Field2`, et non en `Recordset`/`Field`. La liaison précoce (section [2.6](/02-interface-environnement/06-early-vs-late-binding.md)) sur ces types étendus garantit l'accès à ces membres et la vérification à la compilation.

## Faut-il utiliser ces types ?

Au-delà de la technique, une question de conception se pose, et la réponse penche souvent vers la prudence. Les champs multivalués et les pièces jointes présentent en effet plusieurs inconvénients sérieux.

Ils **compliquent le SQL** et brisent la normalisation relationnelle, en masquant une relation un-à-plusieurs derrière un faux champ unique. Ils **résistent mal à la migration** vers SQL Server, qui ne connaît pas de type multivalué natif : lors d'une migration (chapitre [23](/23-migration-interoperabilite/README.md)), ces champs doivent être convertis en tables de jonction, ce qui complexifie l'opération. Les pièces jointes, enfin, **alourdissent la base** : leur contenu binaire s'impute sur la limite de 2 Go du fichier (section [18.9](/18-optimisation-performance/09-limites-taille-2go.md)).

En pratique, pour une nouvelle application — et a fortiori si une migration vers SQL Server est envisageable — on préférera généralement modéliser explicitement une **table enfant** (vraie relation un-à-plusieurs ou plusieurs-à-plusieurs) plutôt qu'un champ multivalué, et stocker des **chemins de fichiers** vers des documents externes plutôt que des pièces jointes embarquées. Ces types restent néanmoins utiles à connaître, ne serait-ce que pour maintenir des applications existantes qui les emploient.

## Pièges courants

Les erreurs récurrentes sont les suivantes. Tenter de lire un champ complexe comme une valeur simple, alors que `.Value` renvoie un `Recordset2` à parcourir. Oublier l'une des deux validations imbriquées (`Update` de l'enfant, puis du parent) lors d'une modification. Déclarer les variables en `Recordset`/`Field` et ne pas accéder aux membres étendus (`IsComplex`, `SaveToFile`, `LoadFromFile`). Appeler `SaveToFile` sur un chemin déjà occupé, ce qui échoue. Et, plus largement, adopter ces types par facilité sans mesurer leurs conséquences sur le SQL, la migration et la taille de la base.

## Points clés à retenir

Les champs multivalués et les champs Pièce jointe, propres au moteur ACE, stockent plusieurs valeurs sous un champ unique, au moyen de tables enfants masquées. Leur manipulation passe par les objets étendus `Recordset2` et `Field2` : la propriété `.Value` d'un tel champ renvoie un `Recordset2` enfant que l'on parcourt comme un jeu ordinaire, et `Field2` ajoute `IsComplex`, `LoadFromFile` et `SaveToFile`. Toute modification suit un modèle imbriqué — éditer le parent, modifier l'enfant et le valider, puis valider le parent. On déclare ces variables en `Recordset2`/`Field2` pour accéder à leurs membres. Enfin, ces types étant délicats pour le SQL, la migration et la taille de la base, on privilégie le plus souvent, en conception, une table enfant normalisée et le stockage de chemins de fichiers.

⏭️ [10. ADO – ActiveX Data Objects pour Access](/10-ado-access/README.md)
