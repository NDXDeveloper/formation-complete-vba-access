🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13.5. Erreurs DAO et ADO — interception spécifique

## Introduction

Jusqu'ici, l'objet `Err` a tenu le rôle central : un numéro, une description, une source, et tout le nécessaire pour traiter une erreur. Cette vision suffit pour les erreurs du langage VBA et de l'interface Access. Mais dès que le code dialogue avec le moteur de données — par l'intermédiaire de DAO ou d'ADO —, l'objet `Err` montre ses limites.

La raison tient à une réalité propre aux moteurs de bases de données : **une seule opération peut engendrer plusieurs erreurs simultanées**. Une instruction refusée par un serveur distant, par exemple, produit souvent une cascade d'erreurs — celle du pilote, celle du serveur, puis une erreur générique côté Access. Or l'objet `Err`, on l'a vu, ne conserve qu'une information à la fois, et c'est fréquemment la plus générique et la moins utile qu'il expose. Le détail réellement exploitable — la cause exacte du refus — se trouve ailleurs : dans des **collections d'erreurs** que DAO et ADO maintiennent indépendamment. Savoir interroger ces collections est ce qui distingue un diagnostic précis d'un message d'erreur inexploitable. C'est l'objet de cette section.

## Pourquoi l'objet Err ne suffit pas

Considérons une insertion refusée par un serveur SQL Server via une table liée. Côté VBA, `Err.Number` vaudra typiquement **3146** — « ODBC, échec de l'appel ». Ce message est exact, mais vide de sens pratique : il indique qu'un appel ODBC a échoué, sans dire pourquoi. La vraie cause — violation d'une clé primaire, contrainte non respectée, échec d'authentification, délai dépassé — n'apparaît nulle part dans l'objet `Err`.

Cette information n'est pas perdue : elle a bien été transmise, mais elle réside dans la collection d'erreurs du moteur. Là où `Err` ne retient qu'une ligne du diagnostic, la collection conserve **l'ensemble de la chaîne**, du message le plus précis (celui du serveur) jusqu'au message le plus général (celui d'Access). Pour obtenir un diagnostic utile, il faut donc cesser de se contenter de `Err` et parcourir la collection appropriée — `DBEngine.Errors` en DAO, `Connection.Errors` en ADO.

## Les erreurs DAO : la collection DBEngine.Errors

### La collection et les objets Error

DAO maintient une collection nommée `DBEngine.Errors`, qui contient les erreurs générées par la dernière opération DAO. Chacun de ses éléments est un objet de type `DAO.Error` — à ne pas confondre avec l'objet `Err` du langage. Chaque objet `Error` expose ses propres propriétés `Number`, `Description`, `Source`, `HelpFile` et `HelpContext`.

Lorsqu'une opération DAO échoue, un ou plusieurs objets `Error` sont déposés dans cette collection. Le parcourir révèle l'intégralité du diagnostic :

```vba
Dim errDAO As DAO.Error
For Each errDAO In DBEngine.Errors
    Debug.Print errDAO.Number & " : " & errDAO.Description
Next errDAO
```

### Le plus spécifique d'abord, le plus générique ensuite

L'ordre des éléments dans `DBEngine.Errors` suit une logique précieuse : les premières entrées sont les plus **spécifiques** (l'erreur d'origine, par exemple celle remontée par le pilote ODBC et le serveur), les dernières les plus **générales** (l'erreur Jet/ACE englobante, typiquement le 3146). L'objet `Err`, de son côté, reflète le plus souvent cette dernière entrée — la plus générique.

C'est pourquoi, pour comprendre ce qui s'est réellement passé, il faut regarder le **début** de la collection, et non se fier au numéro exposé par `Err`. La première entrée d'une erreur ODBC contient généralement le message complet du serveur, du type `[Microsoft][ODBC Driver][SQL Server]Violation de la contrainte PRIMARY KEY...` — c'est-à-dire l'information décisive pour le diagnostic.

### Le cas des erreurs ODBC

DAO ne propose pas de propriétés dédiées au code natif du serveur ou à l'état SQL : à la différence d'ADO, l'objet `DAO.Error` ne possède ni `SQLState` ni `NativeError`. Pour les erreurs ODBC, ces informations sont **incluses dans le texte** de la propriété `Description`, sous la forme des préfixes entre crochets indiquant successivement le gestionnaire ODBC, le pilote et le serveur. L'analyse passe donc par la lecture de ce libellé, et non par des propriétés structurées.

### Cycle de vie de la collection

`DBEngine.Errors` est **vidée puis renseignée au début de chaque opération DAO qui génère une erreur**. Entre deux erreurs DAO, elle conserve son contenu précédent. Cette persistance impose une précaution : si l'erreur en cours n'est **pas** une erreur DAO (par exemple une erreur VBA 91 ou une erreur Access 2xxx), la collection peut contenir une ancienne erreur DAO, sans rapport avec l'incident présent. La lire aveuglément conduirait alors à un diagnostic faux.

La parade consiste à ne consulter la collection que lorsque l'on sait être en présence d'une erreur du moteur de données — par exemple en vérifiant que `Err.Number` appartient à la plage 3xxx, ou que la dernière entrée de la collection correspond bien au numéro exposé par `Err`.

### Pattern d'interception DAO

En pratique, le gestionnaire d'erreurs d'une procédure manipulant des données DAO parcourt la collection lorsqu'elle est pertinente, et se rabat sur `Err` dans le cas contraire :

```vba
Public Sub EnregistrerCommande()
    On Error GoTo GestionErreur

    ' dbFailOnError est essentiel : sans lui, Execute peut ignorer
    ' silencieusement les erreurs au lieu de les lever
    CurrentDb.Execute "INSERT INTO T_Commandes (NumCmd, IDClient) " & _
                      "VALUES (1001, 9999)", dbFailOnError

    Exit Sub

GestionErreur:
    Dim errDAO As DAO.Error
    Dim message As String

    ' Le détail réel se trouve dans DBEngine.Errors, pas dans le seul Err
    If DBEngine.Errors.Count > 0 Then
        For Each errDAO In DBEngine.Errors
            message = message & "• " & errDAO.Number & " — " & _
                      errDAO.Description & vbCrLf
        Next errDAO
    Else
        message = Err.Number & " — " & Err.Description
    End If

    MsgBox message, vbCritical, "Enregistrement de la commande"
End Sub
```

On notera l'argument `dbFailOnError` passé à la méthode `Execute`. Il est indispensable : sans lui, une requête action peut échouer **silencieusement**, sans déclencher d'erreur ni alimenter la collection. C'est une bonne pratique systématique pour toute requête action exécutée par code (voir aussi le chapitre 18 sur les performances).

## Les erreurs ADO : la collection Connection.Errors

### La collection sur l'objet Connection

ADO suit un principe comparable, mais sa collection d'erreurs est portée par l'objet **Connection**, et non par un objet global : on y accède via `Connection.Errors`. Chaque élément est un objet `ADODB.Error`, distinct lui aussi de l'objet `Err` du langage.

L'objet `ADODB.Error` est plus riche que son équivalent DAO. Outre `Number`, `Description` et `Source`, il expose deux propriétés particulièrement utiles face à un serveur distant : `SQLState`, le code d'état SQL normalisé sur cinq caractères (par exemple `23000` pour une violation de contrainte d'intégrité), et `NativeError`, le code d'erreur **propre au moteur sous-jacent** (par exemple le numéro d'erreur natif renvoyé par SQL Server).

### Erreurs ADO et erreurs du fournisseur : une distinction cruciale

Le point le plus important — et la principale source de confusion — concerne l'**origine** des erreurs sous ADO, car toutes ne se retrouvent pas dans la collection.

Les erreurs émanant du **fournisseur** (le moteur de base de données, le fournisseur OLE DB) sont déposées dans `Connection.Errors`. En revanche, les erreurs propres à **ADO lui-même** — un argument invalide, une opération impossible au niveau de la couche ADO — sont signalées par le mécanisme classique, c'est-à-dire via l'objet `Err`, et **n'alimentent pas** la collection.

Il en découle qu'un gestionnaire ADO robuste doit interroger **les deux sources** : la collection `Connection.Errors` pour les erreurs du fournisseur, et l'objet `Err` pour les erreurs propres à ADO. Se fier à l'une seule revient à ignorer la moitié des cas possibles.

### Cycle de vie de la collection

`Connection.Errors` est **vidée puis remplie à chaque nouvelle erreur du fournisseur** : lorsqu'une erreur survient, la collection est d'abord effacée, puis les nouveaux objets `Error` y sont ajoutés. Comme en DAO, il est prudent de la vider explicitement après traitement, au moyen de la méthode `Connection.Errors.Clear`, afin d'éviter toute confusion lors d'une opération ultérieure.

### Pattern d'interception ADO

Le gestionnaire combine donc l'examen de la collection et celui de `Err` :

```vba
Public Sub InsererViaADO()
    Dim cnn As ADODB.Connection
    Set cnn = CurrentProject.Connection

    On Error GoTo GestionErreur

    cnn.Execute "INSERT INTO T_Commandes (NumCmd, IDClient) " & _
                "VALUES (1001, 9999)"

    Exit Sub

GestionErreur:
    Dim errADO As ADODB.Error
    Dim message As String

    ' 1. Erreurs du fournisseur : dans la collection
    If Not cnn Is Nothing Then
        If cnn.Errors.Count > 0 Then
            For Each errADO In cnn.Errors
                message = message & "• " & errADO.Number & _
                          " (natif " & errADO.NativeError & _
                          ", état " & errADO.SQLState & ") — " & _
                          errADO.Description & vbCrLf
            Next errADO
            cnn.Errors.Clear
        End If
    End If

    ' 2. Si la collection est vide, l'erreur vient d'ADO lui-même
    If Len(message) = 0 Then
        message = Err.Number & " — " & Err.Description
    End If

    MsgBox message, vbCritical, "Insertion ADO"
End Sub
```

Ce schéma traite d'abord les erreurs du fournisseur — les plus instructives, avec leur code natif et leur état SQL —, puis se rabat sur `Err` lorsque la collection est vide, signe que l'erreur provient de la couche ADO.

## Tableau comparatif

Les deux technologies reposent sur la même idée — une collection d'erreurs distincte de l'objet `Err` — mais diffèrent sur plusieurs points qu'il est utile de garder en vue.

| Aspect | DAO | ADO |
|---|---|---|
| Collection d'erreurs | `DBEngine.Errors` (objet global) | `Connection.Errors` (sur l'objet Connection) |
| Type des éléments | `DAO.Error` | `ADODB.Error` |
| Détail du serveur (ODBC) | inclus dans le texte de `Description` | propriétés dédiées `SQLState` et `NativeError` |
| Erreurs hors moteur | toujours signalées via `Err` | erreurs ADO via `Err`, erreurs du fournisseur via la collection |
| Ordre des entrées | de la plus spécifique à la plus générique | dépend du fournisseur |
| Réinitialisation | au début de chaque opération DAO produisant une erreur | à chaque nouvelle erreur du fournisseur |

## Le cas critique : les tables liées ODBC

S'il ne fallait retenir qu'une situation de cette section, ce serait celle-ci. Les applications Access modernes s'appuient très souvent sur un back-end SQL Server via des tables liées ODBC (architecture hybride détaillée au chapitre 23). Dans ce contexte, **toute** erreur du serveur — contrainte violée, conflit de verrouillage, délai dépassé, droits insuffisants — remonte côté Access sous la forme du même numéro générique : 3146, « ODBC, échec de l'appel ».

Un gestionnaire qui se contenterait d'afficher `Err.Number` et `Err.Description` ne livrerait alors qu'un message identique pour des causes radicalement différentes, rendant le diagnostic impossible. C'est précisément dans ce cas que le parcours de `DBEngine.Errors` (en DAO) ou de `Connection.Errors` (en ADO) devient indispensable : c'est la seule façon d'accéder au message réel du serveur et de comprendre ce qui a échoué. Pour qui développe sur un back-end serveur, la maîtrise des collections d'erreurs n'est pas un raffinement, mais une nécessité quotidienne.

## Bonnes pratiques

L'interception des erreurs du moteur de données obéit à quelques principes constants :

- **Toujours passer `dbFailOnError`** à la méthode `Execute` en DAO, afin que les requêtes action lèvent réellement leurs erreurs au lieu de les ignorer silencieusement.
- **Parcourir la collection dédiée** plutôt que de se fier au seul objet `Err`, et y consigner toutes les entrées : la première est souvent la plus parlante.
- **En ADO, interroger les deux sources** : `Connection.Errors` pour les erreurs du fournisseur, `Err` pour les erreurs propres à ADO.
- **Ne lire `DBEngine.Errors` que dans un contexte DAO avéré**, pour éviter d'exploiter une erreur ancienne et sans rapport, du fait de la persistance de la collection.
- **Vider explicitement la collection** après traitement (`Errors.Clear`), surtout en ADO, pour prévenir toute confusion ultérieure.
- **Conserver une référence accessible** à la Connection (ADO) dans le gestionnaire, faute de quoi la collection serait inatteignable au moment de traiter l'erreur.

## En résumé

DAO et ADO ne se contentent pas de l'objet `Err` : une opération sur le moteur de données peut produire plusieurs erreurs simultanées, et c'est souvent la plus générique — le fameux 3146 pour les accès ODBC — que `Err` expose, tandis que la cause réelle réside dans une collection dédiée. En DAO, cette collection est `DBEngine.Errors`, dont les premières entrées sont les plus spécifiques et où le détail ODBC se lit dans le texte de la description. En ADO, c'est `Connection.Errors`, plus riche grâce aux propriétés `SQLState` et `NativeError`, avec la distinction capitale entre erreurs du fournisseur (dans la collection) et erreurs ADO (via `Err`).

La règle pratique est simple : pour diagnostiquer une erreur de données, parcourir la collection appropriée, en gardant à l'esprit son cycle de vie et sa persistance. Cette discipline devient incontournable dès lors que l'application s'appuie sur un back-end serveur via des tables liées ODBC, où elle constitue le seul moyen d'accéder au message réel du serveur.

Disposer du diagnostic complet n'a toutefois de valeur que si l'on en conserve la trace. La section suivante montre comment journaliser durablement les erreurs dans une table Access, afin de les exploiter au-delà de l'instant où elles surviennent.

---


⏭️ [13.6. Journalisation des erreurs dans une table Access](/13-gestion-erreurs/06-journalisation-erreurs-table.md)
