🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.4 Navigation dans un Recordset (`MoveFirst`, `MoveNext`, `EOF`, `BOF`)

## Introduction

Une fois un `Recordset` ouvert (section [9.3](/09-dao-data-access-objects/03-recordset-types.md)), tout l'enjeu consiste à s'y déplacer pour atteindre, lire ou modifier les enregistrements. C'est l'objet de cette section : comprendre la notion d'enregistrement courant, maîtriser les méthodes de déplacement, et surtout savoir reconnaître les **bornes** du jeu grâce aux propriétés `EOF` et `BOF`. Le schéma de parcours qui en découle — la boucle de lecture — est l'un des motifs les plus utilisés de toute la programmation DAO ; bien le poser évite quantité d'erreurs classiques, de la boucle infinie à l'enregistrement oublié.

## L'enregistrement courant et le curseur

Un `Recordset` n'expose jamais tous ses enregistrements simultanément : à un instant donné, il est **positionné sur un seul enregistrement**, dit *enregistrement courant*. On peut se représenter cela comme un curseur qui pointe sur une ligne du jeu. Toutes les opérations de lecture et d'écriture des champs (`rs!NomChamp`, abordées à la section [9.5](/09-dao-data-access-objects/05-lecture-modification-champs.md)) s'appliquent à cet enregistrement courant. Naviguer dans un `Recordset`, c'est donc déplacer ce curseur d'un enregistrement à un autre.

À l'ouverture, un jeu **non vide** est automatiquement positionné sur son **premier enregistrement**. Un jeu **vide**, lui, n'a aucun enregistrement courant — situation que `BOF` et `EOF` permettent précisément de détecter.

## Les méthodes de déplacement

DAO fournit un ensemble de méthodes pour déplacer le curseur :

- `MoveFirst` — se positionne sur le premier enregistrement ;
- `MoveLast` — se positionne sur le dernier enregistrement ;
- `MoveNext` — avance d'un enregistrement ;
- `MovePrevious` — recule d'un enregistrement ;
- `Move n` — se déplace de `n` enregistrements (valeur positive vers l'avant, négative vers l'arrière).

Ces méthodes sont disponibles dans leur intégralité pour les types Table, Dynaset et Snapshot. Le type **Forward-only** fait exception : il n'autorise que `MoveNext`, ce qui correspond à sa vocation de parcours unique vers l'avant (voir plus bas).

## `EOF` et `BOF` : les bornes du jeu

Deux propriétés booléennes signalent que le curseur a atteint les extrémités du jeu d'enregistrements :

`EOF` (*End Of File*) vaut `True` lorsque le curseur se trouve **au-delà** du dernier enregistrement. `BOF` (*Beginning Of File*) vaut `True` lorsqu'il se trouve **avant** le premier. Tant que le curseur est positionné sur un enregistrement réel, les deux valent `False`.

Trois situations sont à connaître. Lorsqu'on appelle `MoveNext` alors qu'on était sur le dernier enregistrement, le curseur passe « après la fin » et `EOF` devient `True`. Symétriquement, `MovePrevious` depuis le premier enregistrement fait passer `BOF` à `True`. Enfin, sur un jeu **vide**, `EOF` et `BOF` valent **simultanément `True`** dès l'ouverture.

Un point de vigilance important : tenter de se déplacer alors qu'on est déjà sur une borne — par exemple appeler `MoveNext` alors que `EOF` est déjà `True` — déclenche l'erreur **3021** (« Aucun enregistrement courant »). C'est l'une des erreurs les plus fréquentes en DAO, et la structure de boucle présentée ci-dessous est conçue pour ne jamais la provoquer.

## Le schéma de parcours canonique

Le parcours d'un `Recordset` du début à la fin suit toujours le même motif : on boucle **tant que** `EOF` est faux, on traite l'enregistrement courant, puis on avance.

```vba
Dim db As DAO.Database
Dim rs As DAO.Recordset
Set db = CurrentDb()
Set rs = db.OpenRecordset("SELECT NomClient FROM tblClients", dbOpenForwardOnly)

Do While Not rs.EOF
    Debug.Print rs!NomClient   ' traitement de l'enregistrement courant
    rs.MoveNext                ' passage au suivant
Loop

rs.Close
Set rs = Nothing
Set db = Nothing
```

L'ordre des opérations est essentiel. La condition `Do While Not rs.EOF` est évaluée **avant** chaque tour : sur un jeu vide, la boucle ne s'exécute donc tout simplement pas, sans erreur. Le `MoveNext` est placé **en fin** de boucle, après le traitement, de sorte que le premier enregistrement est bien traité et que le dernier `MoveNext` amène proprement à `EOF`, ce qui termine la boucle.

Deux erreurs guettent les débutants. **Oublier le `MoveNext`** crée une boucle infinie : le curseur ne bouge jamais, `EOF` reste faux indéfiniment. Placer le `MoveNext` **avant** le traitement, à l'inverse, fait sauter le premier enregistrement. Le motif ci-dessus — tester, traiter, avancer — évite ces deux pièges.

## Gérer un jeu vide

La boucle `Do While Not rs.EOF` est sûre sur un jeu vide. En revanche, certaines opérations ne le sont pas : appeler `MoveFirst` ou `MoveLast` sur un `Recordset` vide déclenche l'erreur 3021. Dès qu'un traitement commence par un déplacement explicite, il faut donc le protéger en testant d'abord la vacuité du jeu, ce qui se fait en vérifiant que `BOF` et `EOF` ne sont pas tous deux vrais :

```vba
If Not (rs.BOF And rs.EOF) Then
    rs.MoveFirst
    Do While Not rs.EOF
        ' ... traitement ...
        rs.MoveNext
    Loop
End If
```

Ce test `Not (rs.BOF And rs.EOF)` est la manière fiable de répondre à la question « le jeu contient-il au moins un enregistrement ? ».

## Compter les enregistrements : `MoveLast` et `RecordCount`

Comme indiqué à la section 9.3, la propriété `RecordCount` ne donne un total exact immédiat que pour les `Recordset` de type **Table**. Pour les types **Dynaset** et **Snapshot**, `RecordCount` ne reflète, à l'ouverture, que le nombre d'enregistrements **déjà parcourus** — souvent 1. Pour obtenir le total réel, il faut forcer DAO à parcourir l'ensemble du jeu en appelant `MoveLast`, puis revenir au début :

```vba
If Not (rs.BOF And rs.EOF) Then
    rs.MoveLast                  ' force le parcours complet du jeu
    Debug.Print rs.RecordCount   ' total désormais exact
    rs.MoveFirst                 ' retour au premier enregistrement
End If
```

Il en découle une règle pratique à ne pas oublier : **`RecordCount` ne doit pas servir à tester si un jeu est vide.** Sur un Dynaset non encore parcouru, sa valeur peut être trompeuse. Pour déterminer la vacuité, on s'appuie toujours sur `EOF` et `BOF`, jamais sur `RecordCount`. Notons enfin que le type Forward-only ne prenant pas en charge `MoveLast`, cette technique de comptage ne lui est pas applicable.

## Mémoriser une position : les signets

Il est parfois utile de noter une position pour y revenir plus tard, après s'être déplacé ailleurs dans le jeu. C'est le rôle de la propriété `Bookmark` : elle renvoie une valeur (de type `Variant`) identifiant l'enregistrement courant, que l'on peut conserver puis réaffecter pour repositionner le curseur.

```vba
Dim position As Variant
position = rs.Bookmark      ' mémorise la position courante
' ... déplacements dans le jeu ...
rs.Bookmark = position      ' revient à la position mémorisée
```

Les signets sont pris en charge par les types Table, Dynaset et Snapshot, mais **pas** par le type Forward-only. En cas de doute sur une source particulière, la propriété `Bookmarkable` du `Recordset` indique si les signets sont disponibles. Les signets sont également au cœur de la manipulation des enregistrements affichés dans un formulaire, via le `RecordsetClone` (section [9.12](/09-dao-data-access-objects/12-recordsetclone.md)).

## Position relative : `AbsolutePosition` et `PercentPosition`

Pour les types Dynaset et Snapshot, deux propriétés renseignent sur la position *relative* du curseur dans le jeu. `AbsolutePosition` donne le rang (à partir de 0) de l'enregistrement courant, et `PercentPosition` une estimation en pourcentage de la progression dans le jeu. Ces propriétés sont surtout utiles pour informer l'utilisateur de l'avancement d'un traitement — par exemple alimenter une barre de progression (voir [17.5](/17-interface-utilisateur-avancee/05-barre-progression.md)).

Une mise en garde s'impose toutefois : `AbsolutePosition` **n'est pas un identifiant stable** d'enregistrement. Sa valeur peut changer lorsque des enregistrements sont ajoutés ou supprimés. Pour repérer durablement un enregistrement, on utilise sa clé primaire ou un signet, jamais `AbsolutePosition`.

## Le cas particulier du Forward-only

Le type Forward-only impose ses contraintes en matière de navigation : seul `MoveNext` est disponible. `MoveFirst`, `MoveLast`, `MovePrevious`, les déplacements vers l'arrière, les signets et `AbsolutePosition` ne lui sont pas accessibles. Cela découle directement de sa nature : un flux que l'on lit une fois, du début à la fin. Tant que le traitement se limite à un parcours séquentiel unique, ces restrictions n'ont aucune importance — et l'on bénéficie en échange du type le plus rapide. Mais dès qu'un retour en arrière, un comptage par `MoveLast` ou une mémorisation de position devient nécessaire, il faut choisir un autre type (Dynaset ou Snapshot).

## Pièges courants

Les principales sources d'erreur en navigation tiennent en quelques points. Oublier le `MoveNext` provoque une boucle infinie ; le placer avant le traitement fait sauter le premier enregistrement. Appeler `MoveFirst` ou `MoveLast` sur un jeu potentiellement vide sans test préalable déclenche l'erreur 3021. Se déplacer au-delà d'une borne déjà atteinte produit la même erreur. Enfin, se fier à `RecordCount` pour tester la vacuité, ou à `AbsolutePosition` pour identifier un enregistrement, conduit à des comportements erronés. Le motif « tester `EOF`, traiter, `MoveNext` », encadré au besoin par un test `Not (rs.BOF And rs.EOF)`, met à l'abri de la plupart de ces écueils.

## Points clés à retenir

Un `Recordset` est toujours positionné sur un enregistrement courant, que les méthodes `Move*` permettent de déplacer. Les propriétés `EOF` et `BOF` signalent les bornes du jeu — toutes deux vraies sur un jeu vide. Le schéma de parcours de référence consiste à boucler tant que `EOF` est faux, à traiter l'enregistrement courant, puis à appeler `MoveNext`. Pour obtenir un comptage exact sur un Dynaset ou un Snapshot, on appelle `MoveLast` avant de lire `RecordCount`, mais la vacuité d'un jeu se teste toujours via `EOF`/`BOF`. Les signets permettent de mémoriser et retrouver une position, hors type Forward-only. Ce dernier, enfin, ne connaît que `MoveNext`, en cohérence avec sa vocation de parcours unique.

⏭️ [9.5. Lecture et modification des champs (Fields)](/09-dao-data-access-objects/05-lecture-modification-champs.md)
