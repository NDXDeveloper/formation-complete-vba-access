🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.2 Ouverture et gestion d'une Database avec `CurrentDb()`

## Introduction

L'objet `Database` est, comme l'a montré la section précédente, le point de départ de presque tout traitement DAO : c'est à partir de lui que l'on ouvre des `Recordset`, exécute des requêtes ou parcourt la structure de la base. Encore faut-il obtenir une référence vers la base courante. Dans une application Access, le moyen standard pour cela est la fonction `CurrentDb()`. Cette section explique ce qu'elle renvoie, comment l'utiliser correctement, comment gérer la durée de vie de la référence obtenue, et comment, à l'inverse, ouvrir une base *autre* que la base courante.

## Qu'est-ce que `CurrentDb()` ?

`CurrentDb()` est une méthode du modèle objet Access — plus précisément de l'objet `Application` — qui renvoie un objet `DAO.Database` représentant la base de données actuellement ouverte dans la fenêtre Access. C'est ce rattachement à l'objet `Application` qui explique qu'on puisse l'appeler directement, sans passer par `DBEngine`.

Le point essentiel à retenir est le suivant : **chaque appel à `CurrentDb()` crée une nouvelle instance** d'objet `Database` pointant sur la base courante. Cette instance neuve présente un avantage déterminant : ses collections (`TableDefs`, `QueryDefs`, `Relations`…) sont à jour. Si une table ou une requête vient d'être créée par code, elle sera bien visible dans l'objet renvoyé par `CurrentDb()`.

On peut écrire indifféremment `CurrentDb` ou `CurrentDb()` ; les deux formes sont acceptées par VBA. L'usage des parenthèses est une convention de lisibilité qui rappelle qu'il s'agit d'un appel de fonction.

## Le schéma d'utilisation de base

Le modèle d'usage d'un objet `Database` suit toujours trois temps : **déclarer**, **affecter**, puis **libérer** une fois le travail terminé.

```vba
Sub ExempleCurrentDb()
    Dim db As DAO.Database

    Set db = CurrentDb()        ' affectation de la référence

    ' Utilisation de l'objet
    Debug.Print db.Name              ' chemin complet du fichier de base
    Debug.Print db.TableDefs.Count   ' nombre de définitions de tables

    Set db = Nothing            ' libération de la référence
End Sub
```

On notera deux propriétés utiles de l'objet : `Name`, qui renvoie le chemin complet du fichier de la base, et `Version`, qui indique la version du moteur. L'instruction `Set db = Nothing` en fin de procédure libère la référence : sa nécessité et les bonnes pratiques de libération sont détaillées à la section [9.11](/09-dao-data-access-objects/11-cloture-liberation-dao.md).

## `CurrentDb()` ou `DBEngine(0)(0)` ?

Il existe une seconde manière d'atteindre la base courante, `DBEngine(0)(0)` — c'est-à-dire `DBEngine.Workspaces(0).Databases(0)` — déjà rencontrée à la section 9.1. La différence de comportement entre les deux est importante en pratique.

`DBEngine(0)(0)` renvoie une **référence à l'instance déjà ouverte** par Access. Cette forme est marginalement plus rapide, puisqu'elle ne crée pas de nouvel objet. En revanche, ses collections ne sont pas systématiquement rafraîchies : si l'on crée un objet (table, requête) par code puis qu'on tente de le référencer aussitôt via `DBEngine(0)(0)`, on risque une erreur indiquant que l'élément est introuvable, tant qu'aucun `Refresh` n'a été appelé.

`CurrentDb()`, en créant une instance neuve aux collections à jour, évite ce piège — c'est la raison principale pour laquelle il constitue le choix recommandé par défaut. La comparaison exhaustive des deux approches, avec leurs cas d'usage respectifs, fait l'objet de la section [4.6](/04-modele-objet-access/06-currentdb-vs-dbengine.md). Pour ce chapitre, on retiendra simplement : **`CurrentDb()` dans le doute**.

## Ne pas appeler `CurrentDb()` à répétition

Le fait que chaque appel crée un nouvel objet a une conséquence directe sur les performances : appeler `CurrentDb()` de façon répétée, en particulier à l'intérieur d'une boucle, est inutilement coûteux. La bonne pratique consiste à obtenir la référence **une seule fois** et à la conserver dans une variable.

```vba
' À éviter : un nouvel objet Database est créé à chaque tour de boucle
Dim i As Long
For i = 1 To 1000
    Debug.Print CurrentDb().TableDefs.Count
Next i

' Préférable : la référence est obtenue une seule fois
Dim db As DAO.Database
Set db = CurrentDb()
For i = 1 To 1000
    Debug.Print db.TableDefs.Count
Next i
Set db = Nothing
```

Le même principe vaut au-delà des boucles : dès qu'une procédure utilise la base à plusieurs reprises, on l'affecte une fois en début de procédure plutôt que de multiplier les appels.

## Portée et durée de vie de la référence

La question de la **durée de vie** d'une référence `Database` mérite réflexion. Le schéma le plus simple et le plus sûr consiste à déclarer la variable au niveau de la **procédure** : la référence naît au début du traitement et disparaît à sa fin. C'est l'approche à privilégier dans la majorité des cas, car elle évite tout risque de référence obsolète.

Pour du code très sollicité, on peut envisager de mettre en cache une référence au niveau d'un **module** afin d'éviter de la recréer à chaque appel. Cette optimisation est réelle mais demande de la rigueur : une référence conservée trop longtemps peut devenir obsolète ou maintenir des ressources mobilisées. Les stratégies de mise en cache des variables de module sont abordées à la section [18.4](/18-optimisation-performance/04-cache-variables-module.md). En l'absence de besoin avéré de performance, la portée procédure reste le choix de référence.

## Ouvrir une base *autre* que la base courante

`CurrentDb()` ne donne accès qu'à la base actuellement ouverte dans Access. Pour ouvrir un **fichier de base externe** — par exemple lire des données dans une autre application `.accdb` — on utilise la méthode `OpenDatabase` de `DBEngine` (ou d'un `Workspace`).

```vba
Dim dbExterne As DAO.Database
Set dbExterne = DBEngine.OpenDatabase("C:\Donnees\Archive.accdb")

' ... traitement sur dbExterne ...

dbExterne.Close            ' obligatoire pour une base ouverte explicitement
Set dbExterne = Nothing
```

La méthode `OpenDatabase` accepte des arguments optionnels permettant notamment d'ouvrir la base en mode exclusif ou en lecture seule. Une différence de gestion essentielle distingue ce cas du précédent : une base que l'on a ouverte soi-même avec `OpenDatabase` doit impérativement être **fermée** par `Close` avant d'être libérée, contrairement à la référence renvoyée par `CurrentDb()`.

## Faut-il appeler `Close` sur l'objet de `CurrentDb()` ?

Non. Sur l'objet renvoyé par `CurrentDb()`, on ne fait que `Set db = Nothing` : il ne faut pas appeler `Close`. La méthode `Close` est réservée aux bases que l'on a ouvertes explicitement avec `OpenDatabase`. Cette distinction est importante et constitue une source fréquente de confusion : on ferme ce qu'on a ouvert soi-même, on se contente de libérer la référence pour la base courante.

## Pièges courants

Quelques erreurs reviennent régulièrement. Conserver une seule référence `Database` partagée pendant toute la durée de vie de l'application sans précaution peut conduire à des comportements imprévisibles si la structure de la base évolue ; en cas de doute, mieux vaut une référence fraîche par traitement. À l'inverse, appeler `CurrentDb()` des dizaines de fois dans une même procédure dégrade les performances sans aucun bénéfice. Enfin, appeler `Close` sur l'objet issu de `CurrentDb()`, ou oublier de fermer une base ouverte par `OpenDatabase`, sont deux erreurs symétriques à éviter.

Il faut aussi garder à l'esprit que `CurrentDb()` suppose l'existence d'une base courante : la fonction n'a de sens qu'au sein d'une session Access disposant d'une base ouverte.

## Points clés à retenir

`CurrentDb()` est le moyen standard d'obtenir une référence vers la base courante en VBA Access. Il renvoie à chaque appel une instance neuve, aux collections à jour, ce qui en fait un choix plus sûr que `DBEngine(0)(0)`. La bonne pratique consiste à l'affecter une seule fois à une variable, à utiliser cette variable pour l'ensemble du traitement, puis à la libérer par `Set db = Nothing` — sans jamais appeler `Close` sur cet objet précis. Pour accéder à une base externe, on recourt à `OpenDatabase`, en n'oubliant pas, cette fois, de fermer la base par `Close`. La discipline de libération des objets, évoquée ici, est traitée systématiquement à la section [9.11](/09-dao-data-access-objects/11-cloture-liberation-dao.md).

⏭️ [9.3. Recordset DAO — types (Table, Dynaset, Snapshot, Forward-only)](/09-dao-data-access-objects/03-recordset-types.md)
