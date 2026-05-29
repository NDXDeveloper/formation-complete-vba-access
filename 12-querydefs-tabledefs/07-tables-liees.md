🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 12.7. Tables liées (Linked Tables) — gestion programmatique

Toutes les sections précédentes concernaient des tables **locales**. Or une application Access professionnelle est presque toujours **scindée** : une partie *front-end* (formulaires, états, requêtes, code) liée à une partie *back-end* (les tables, dans un fichier séparé ou sur un serveur). Le lien entre les deux passe par des **tables liées** (*linked tables*) — des tables dont les données vivent ailleurs, mais qui apparaissent et se manipulent comme si elles étaient locales.

Gérer ces liens par code est une compétence essentielle, ne serait-ce que pour une raison : lorsque le back-end est **déplacé**, les liens se rompent, et il faut les **rétablir**. Cette section couvre la création, la reliaison, la vérification et la suppression des liens. L'architecture front-end/back-end est détaillée au chapitre 15.1, l'angle connexion distante et multi-utilisateur au chapitre 15.9, et la reliaison automatique au déploiement au chapitre 21.4.

---

## Qu'est-ce qu'une table liée ?

Une table liée est une table dont les **données résident dans un autre fichier ou une autre base** — le back-end —, mais qui figure dans la base courante (le front-end) comme une table ordinaire : on l'interroge et on la modifie de façon transparente. C'est le fondement de l'architecture scindée (chapitre 15.1).

Comme vu à la section 12.4, DAO représente une table liée par un objet `TableDef` doté d'une propriété **`Connect`** non vide (la source) et d'une propriété **`SourceTableName`** (le nom de la table dans la source).

---

## Pourquoi gérer les liens par code ?

Plusieurs besoins justifient la gestion programmatique des liens. Le premier, et de loin le plus fréquent, est la **reliaison** : quand le fichier back-end change d'emplacement ou de nom, les liens du front-end deviennent invalides, et il faut les rétablir vers la nouvelle source — opération centrale au déploiement et à la maintenance (chapitre 21.4). Viennent ensuite la **création de liens** (mettre en place un front-end neuf), la **vérification** des liens (quel back-end ? le lien est-il valide ?), et le **basculement entre environnements** (développement, test, production).

---

## L'anatomie du lien : `Connect` et `SourceTableName`

La propriété **`Connect`** décrit la source du lien, et sa forme dépend du type de back-end. Pour une base **Access**, elle prend la forme d'un point-virgule suivi du mot-clé `DATABASE=` :

```
;DATABASE=C:\Data\backend.accdb
```

Pour un back-end **ODBC** (SQL Server, par exemple), elle commence par `ODBC;` suivi de la chaîne ODBC (dans l'esprit du chapitre 10.9) :

```
ODBC;Driver={ODBC Driver 17 for SQL Server};Server=NOM;Database=MaBase;Trusted_Connection=yes;
```

La propriété **`SourceTableName`** indique le nom de la table *dans la source*. À noter que le **nom local** de la table liée (celui sous lequel elle apparaît dans le front-end) peut **différer** de `SourceTableName`, ce qui permet d'exposer une table sous un autre nom localement.

---

## Créer une table liée par code

La création d'un lien suit le schéma `Append` des sections précédentes, mais **sans définir de champs** : la structure est héritée de la source.

```vba
Dim db As DAO.Database
Dim td As DAO.TableDef
Set db = CurrentDb

Set td = db.CreateTableDef("Clients")              ' nom local du lien
td.Connect = ";DATABASE=C:\Data\backend.accdb"     ' la source (back-end Access)
td.SourceTableName = "Clients"                      ' nom de la table dans la source
db.TableDefs.Append td                              ' crée le lien

Set td = Nothing: Set db = Nothing
```

On ne crée aucun champ : Access récupère la structure depuis le back-end. Pour un lien ODBC, seule la forme de `Connect` change (préfixe `ODBC;`).

---

## Relier à nouveau : `RefreshLink`

C'est l'opération clé. Pour **rétablir un lien** vers une nouvelle source, on modifie la propriété `Connect` de la table liée existante, puis on appelle la méthode **`RefreshLink`** :

```vba
Dim td As DAO.TableDef
Set td = CurrentDb.TableDefs("Clients")
td.Connect = ";DATABASE=C:\NouveauChemin\backend.accdb"
td.RefreshLink                                       ' applique le nouveau lien
```

`RefreshLink` réétablit la connexion à partir de la propriété `Connect` mise à jour. La séquence est invariable : **on change `Connect`, puis on appelle `RefreshLink`**.

---

## Relier toutes les tables : le patron de reliaison

En pratique, on relie rarement une seule table : on relie l'**ensemble** des tables liées vers un nouveau back-end. Le patron consiste à parcourir les `TableDefs`, à repérer les tables liées (`Connect` non vide), et à mettre à jour puis rafraîchir chacune :

```vba
Public Sub ReliaisonTables(ByVal cheminBackend As String)
    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Set db = CurrentDb

    For Each td In db.TableDefs
        If Len(td.Connect) > 0 Then                  ' c'est une table liée
            td.Connect = ";DATABASE=" & cheminBackend
            td.RefreshLink
        End If
    Next td

    Set db = Nothing
End Sub
```

> 📌 Ce patron suppose un back-end **Access**. Dans une application mixte comportant aussi des **tables liées ODBC**, on distinguera les deux en examinant le préfixe de `Connect` (`ODBC;` ou `;DATABASE=`), pour n'appliquer à chaque table que le bon format de chaîne.

---

## Vérifier un lien et réparer un lien rompu

Un lien rompu — back-end introuvable — provoque une **erreur** dès qu'on accède à la table. Pour réagir proprement, le patron habituel, exécuté au démarrage de l'application, consiste à **tester l'accès** à une table liée connue ; en cas d'échec, on déclenche la reliaison (vers un chemin de configuration, ou en demandant l'emplacement à l'utilisateur).

```vba
Public Function LienValide(ByVal nomTable As String) As Boolean
    Dim rs As DAO.Recordset
    On Error Resume Next
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT TOP 1 * FROM [" & nomTable & "];", dbOpenSnapshot)
    LienValide = (Err.Number = 0)
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
End Function
```

Si `LienValide` renvoie `False`, on appelle la reliaison. Ce mécanisme de **vérification et reconnexion au démarrage** est le cœur d'une application scindée robuste, et son automatisation complète est traitée au chapitre 21.4.

---

## Supprimer un lien n'est pas supprimer les données

Distinction capitale : supprimer une table liée **retire le lien, pas les données**.

```vba
CurrentDb.TableDefs.Delete "Clients"   ' supprime LE LIEN, pas les données du back-end
```

Contrairement à la suppression d'une table **locale** — qui détruit définitivement ses données —, supprimer un `TableDef` **lié** ne fait que retirer le lien du front-end : la table et ses données demeurent intactes dans le back-end. On peut donc recréer le lien à tout moment, sans aucune perte.

---

## Performance : garder une connexion ouverte

Pour un back-end **Access** accédé via le réseau, une optimisation bien connue mérite d'être signalée : **maintenir une connexion ouverte** vers le back-end pendant la session — typiquement en gardant ouvert un recordset sur une table du back-end. Sans cela, le front-end ouvre et ferme le fichier back-end à répétition, ce qui est lent sur un réseau. Cette technique de **persistance de connexion**, qui accélère sensiblement les accès, est développée au chapitre 18.8.

---

## Cas particuliers : mot de passe et ODBC

Deux variantes courantes. Si le back-end Access est **protégé par mot de passe**, on l'inclut dans la chaîne `Connect` via le mot-clé `PWD=` (`;DATABASE=…;PWD=…`). Si le back-end est un **serveur** (SQL Server, etc.), les tables sont liées en **ODBC** : la chaîne `Connect` adopte alors la forme `ODBC;…` décrite au chapitre 10.9, avec les mêmes prérequis de pilote installé.

---

## Précautions

Quelques précautions habituelles. La création et la suppression de liens, comme toute opération structurelle, **ne sont pas transactionnelles**. Une table liée **en cours d'utilisation** (formulaire ouvert dessus) peut résister à une reliaison. On **vérifie l'existence** avant de créer un lien. Et l'on retient surtout que **supprimer un lien ne touche jamais les données** du back-end — un filet de sécurité rassurant, mais qui n'autorise pas à supprimer un lien à la légère, sous peine de casser les objets qui en dépendent.

---

## En résumé

Une **table liée** expose dans le front-end des données vivant dans un **back-end** distant, le lien étant porté par les propriétés **`Connect`** (la source, `;DATABASE=…` pour Access, `ODBC;…` pour un serveur) et **`SourceTableName`** d'un `TableDef`. On **crée** un lien par `CreateTableDef` + `Connect` + `SourceTableName` + `Append`, **sans définir de champs** (structure héritée). On **relie à nouveau** en modifiant `Connect` puis en appelant **`RefreshLink`** — le patron de reliaison parcourant toutes les tables liées (`Connect` non vide). On **vérifie** un lien en tentant d'accéder à la table, pour le **réparer** au démarrage si besoin (chapitre 21.4). Enfin, **supprimer un lien retire le lien, jamais les données** du back-end. La persistance d'une connexion ouverte optimise les accès réseau (chapitre 18.8).

Il nous reste un dernier aspect des objets structurels, évoqué à plusieurs reprises sans être détaillé : les **propriétés personnalisées**. Ces propriétés étendent les métadonnées des objets de la base — et expliquent notamment pourquoi des attributs comme `Caption` ou `Format` exigeaient un traitement particulier. C'est l'objet de la dernière section du chapitre.


⏭️ [12.8. Propriétés personnalisées des objets base de données](/12-querydefs-tabledefs/08-proprietes-personnalisees.md)
