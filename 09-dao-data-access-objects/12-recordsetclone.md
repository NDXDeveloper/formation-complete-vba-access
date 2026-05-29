🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.12 `RecordsetClone` — manipuler les enregistrements d'un formulaire

## Introduction

Les sections précédentes ont traité le `Recordset` comme un objet ouvert directement sur une table ou une requête. Mais un formulaire lié possède lui aussi un jeu d'enregistrements sous-jacent — ceux qu'il affiche — et l'on a souvent besoin d'agir sur ces enregistrements par code : retrouver un client, compter les lignes, parcourir les données. Le faire en déplaçant directement le formulaire perturberait l'affichage de l'utilisateur. La propriété `RecordsetClone` apporte la solution. Cette section, qui fait converger la navigation (section [9.4](/09-dao-data-access-objects/04-navigation-recordset.md)), la recherche (section [9.9](/09-dao-data-access-objects/09-recherche-recordset.md)) et les signets, expose son fonctionnement et son idiome central : rechercher dans le clone, puis synchroniser l'affichage du formulaire.

## Qu'est-ce que le `RecordsetClone` ?

`Form.RecordsetClone` est un `Recordset` DAO qui constitue un **clone** du jeu d'enregistrements du formulaire. Clone signifie ici qu'il partage exactement les **mêmes données** et le même ensemble d'enregistrements que le formulaire, mais qu'il dispose de son **propre pointeur d'enregistrement courant**, indépendant de celui du formulaire.

Cette indépendance est tout l'intérêt : on peut naviguer, rechercher ou compter dans le clone **sans déplacer l'enregistrement affiché** à l'écran. S'agissant généralement d'un `Recordset` de type Dynaset, le clone prend en charge les méthodes `Find` et les signets, sur lesquels repose la synchronisation.

## `RecordsetClone` ou `Recordset` du formulaire ?

Un formulaire expose en réalité **deux** propriétés liées à son jeu d'enregistrements, qu'il importe de distinguer.

`RecordsetClone` possède un pointeur **indépendant** : s'y déplacer ne bouge pas le formulaire.

`Recordset` (la propriété `Form.Recordset`) donne accès au jeu d'enregistrements **propre** du formulaire : s'y déplacer **déplace** l'enregistrement affiché, puisqu'il s'agit du même curseur.

Le choix en découle. Pour rechercher ou compter sans perturber l'affichage, on utilise le **clone**. Pour piloter directement la position du formulaire par des méthodes de `Recordset`, on utilise `Form.Recordset`. L'idiome le plus courant — chercher d'abord, n'afficher qu'en cas de succès — passe par le clone, car il permet de tester le résultat avant de déplacer le formulaire.

## Le motif fondamental : rechercher puis synchroniser

C'est le cas d'usage emblématique du `RecordsetClone`, typiquement derrière une zone de recherche : on localise un enregistrement dans le clone, puis on positionne le formulaire dessus en affectant le signet du clone à la propriété `Bookmark` du formulaire.

```vba
With Me.RecordsetClone
    .FindFirst "IDClient = " & Me!txtRecherche
    If Not .NoMatch Then
        Me.Bookmark = .Bookmark      ' synchronise l'affichage du formulaire
    Else
        MsgBox "Client introuvable."
    End If
End With
```

La recherche s'effectue dans le clone (sans bouger le formulaire), on teste `NoMatch` comme vu à la section 9.9, et ce n'est qu'en cas de succès que l'on déplace effectivement le formulaire via son signet. L'enregistrement recherché s'affiche alors à l'écran.

## La compatibilité des signets

La clé de ce mécanisme est que le signet du formulaire (`Me.Bookmark`) et celui du clone (`.Bookmark`) sont **interchangeables**. C'est précisément parce que le clone est un clone du jeu du formulaire — mêmes données, même ensemble d'enregistrements — que l'on peut affecter l'un à l'autre. Cette compatibilité ne vaut qu'entre un formulaire et son propre clone : on ne mélange jamais des signets issus de `Recordset` sans rapport, ce qui n'aurait aucun sens.

## Compter les enregistrements d'un formulaire

Le clone fournit aussi un moyen commode de compter les enregistrements affichés. On se souviendra toutefois de l'avertissement de la section 9.4 : `RecordCount` n'est exact qu'après un `MoveLast`.

```vba
With Me.RecordsetClone
    If Not (.BOF And .EOF) Then .MoveLast
    Debug.Print "Nombre d'enregistrements : " & .RecordCount
End With
```

Comme le comptage s'effectue dans le clone, le `MoveLast` ne déplace pas l'enregistrement affiché par le formulaire.

## Parcourir et modifier via le clone

Le clone étant un `Recordset` Dynaset modifiable, on peut l'utiliser pour parcourir et modifier les enregistrements du formulaire avec les méthodes vues aux sections 9.6 à 9.8. Les modifications portant sur les mêmes données, elles concernent bien le jeu du formulaire.

```vba
Dim rs As DAO.Recordset
Set rs = Me.RecordsetClone

Do While Not rs.EOF
    rs.Edit
    rs!Statut = "Vérifié"
    rs.Update
    rs.MoveNext
Loop

Me.Refresh                   ' met à jour les valeurs affichées
Set rs = Nothing             ' libère la référence — sans Close (voir plus bas)
```

Une nuance de synchronisation est à connaître. Après des **modifications de valeurs** via le clone, un `Me.Refresh` rafraîchit l'affichage. Si l'on a **ajouté ou supprimé** des enregistrements, c'est un `Me.Requery` qu'il faut, car l'ensemble des enregistrements a changé (sachant que `Requery` repositionne le formulaire sur le premier enregistrement). En cas de suppression de l'enregistrement courant du formulaire, on veillera en outre à repositionner ce dernier, selon les principes de la section 9.8.

## Ne pas fermer le `RecordsetClone`

Un point important distingue le clone des `Recordset` que l'on ouvre soi-même : on ne l'a **pas** ouvert, c'est le formulaire qui le possède. En conséquence, on ne lui applique **pas** `Close`. La discipline de libération de la section [9.11](/09-dao-data-access-objects/11-cloture-liberation-dao.md) s'applique donc avec cette réserve : lorsqu'on a affecté le clone à une variable locale, on libère cette variable par `Set rs = Nothing` quand on a terminé, mais on n'appelle jamais `rs.Close` — fermer un objet appartenant au formulaire est à proscrire.

## Accès depuis l'extérieur du formulaire

Les exemples ci-dessus, écrits dans le module du formulaire, utilisent `Me`. Depuis un autre contexte, on accède au clone par la référence complète au formulaire ouvert :

```vba
Forms!frmClients.RecordsetClone.FindFirst "IDClient = " & lngID
```

Cela suppose, naturellement, que le formulaire soit **ouvert et lié** à une source de données. L'accès aux formulaires ouverts via la collection `Forms` est détaillé à la section [6.2](/06-formulaires/02-acces-formulaires-ouverts.md).

## Pièges courants

Les erreurs récurrentes sont les suivantes. Oublier de synchroniser l'affichage après une recherche dans le clone : la recherche réussit, mais le formulaire ne bouge pas tant qu'on n'a pas affecté `Me.Bookmark`. Confondre `RecordsetClone` (pointeur indépendant) et `Form.Recordset` (qui déplace le formulaire). Lire `RecordCount` sans `MoveLast` préalable. Appeler `Close` sur le clone. Et négliger le `Refresh` ou le `Requery` après modification, ce qui laisse un affichage non synchronisé avec les données.

## Points clés à retenir

`RecordsetClone` est un clone du jeu d'enregistrements d'un formulaire, doté d'un pointeur indépendant : on y navigue, recherche et compte sans perturber l'affichage. Son idiome central consiste à localiser un enregistrement par `FindFirst`, à tester `NoMatch`, puis à synchroniser le formulaire via `Me.Bookmark = .Bookmark`, les signets du formulaire et de son clone étant interchangeables. On distingue ce clone de `Form.Recordset`, dont les déplacements bougent le formulaire. Le clone permet aussi de parcourir et modifier les données affichées, à condition de rafraîchir l'affichage par `Refresh` ou `Requery` selon le cas. Enfin, on ne ferme jamais le `RecordsetClone` : on se contente, le cas échéant, de libérer la variable locale qui le référence.

⏭️ [9.13. Champs multivalués (MVF) et champs Pièce jointe — Recordset2 et Field2](/09-dao-data-access-objects/13-champs-multivalues-pieces-jointes.md)
