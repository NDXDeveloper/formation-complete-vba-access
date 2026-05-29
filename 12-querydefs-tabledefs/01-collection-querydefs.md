🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 12.1. Collection QueryDefs — accès et manipulation des requêtes sauvegardées

La porte d'entrée vers la structure d'une base, annoncée à la fin du README, est sa collection de **requêtes sauvegardées**. Cette section apprend à y **accéder** et à en **explorer** le contenu : identifier les requêtes, lire leur SQL, examiner leur type et la structure qu'elles renvoient, puis les utiliser. La création, la modification et la suppression feront l'objet de la section 12.2 ; le paramétrage, de la section 12.3.

Chaque requête que vous voyez dans le volet de navigation d'Access correspond à un objet **`QueryDef`** (*Query Definition*). Comprendre comment manipuler ces objets par code ouvre la voie à l'inspection, à la documentation et à l'automatisation de la couche requêtes de l'application.

---

## Qu'est-ce qu'un QueryDef ?

Un `QueryDef` est l'objet DAO qui **représente une requête sauvegardée** dans la base. Son cœur est sa propriété **`SQL`**, qui contient le texte SQL de la requête. Tous les `QueryDef` d'une base sont réunis dans la collection **`QueryDefs`** de l'objet `Database`.

Autrement dit, là où le chapitre 11 manipulait du SQL sous forme de texte, on accède ici à l'**objet** qui encapsule ce SQL, avec ses propriétés (nom, type, dates, paramètres) et ses méthodes.

---

## La collection QueryDefs

On accède à la collection via la propriété `QueryDefs` d'un objet `Database`. Plusieurs modes d'accès coexistent :

```vba
Dim db As DAO.Database
Dim qdf As DAO.QueryDef
Set db = CurrentDb

Set qdf = db.QueryDefs("R_ClientsActifs")   ' accès par nom
Set qdf = db.QueryDefs(0)                    ' accès par index
Debug.Print db.QueryDefs.Count               ' nombre de requêtes
```

> 📌 **Rappel important** (chapitres 4.6 et 11.1) : chaque appel à `CurrentDb` renvoie un *nouvel* objet `Database`. Dès qu'on enchaîne plusieurs opérations sur les `QueryDefs`, on **conserve une référence** (`Set db = CurrentDb`) plutôt que de rappeler `CurrentDb` à chaque fois.

---

## Les propriétés essentielles d'un QueryDef

Un `QueryDef` expose un ensemble de propriétés permettant de l'inspecter et de le caractériser.

| Propriété | Contenu |
|---|---|
| `Name` | Nom de la requête |
| `SQL` | Texte SQL de la requête (en lecture **et** en écriture) |
| `Type` | Type de requête (constante `dbQ…`) |
| `DateCreated` / `LastUpdated` | Dates de création / de dernière modification |
| `Connect` | Chaîne ODBC, pour les requêtes Pass-Through (section 11.9) |
| `ReturnsRecords` | Renvoie-t-elle des lignes ? (Pass-Through) |
| `Parameters` | Collection des paramètres (section 12.3) |
| `Fields` | Colonnes renvoyées par la requête |
| `Updatable` | La requête est-elle modifiable ? |
| `RecordsAffected` | Lignes affectées après exécution d'une requête action |

---

## Identifier le type d'une requête

La propriété **`Type`** renvoie une constante indiquant la nature de la requête — précieuse pour écrire du code générique, par exemple pour décider d'ouvrir un recordset (sélection) ou d'exécuter (action).

| Constante | Type de requête |
|---|---|
| `dbQSelect` | Sélection |
| `dbQAction` | Action (générique) |
| `dbQAppend` | Ajout (`INSERT`) |
| `dbQUpdate` | Mise à jour (`UPDATE`) |
| `dbQDelete` | Suppression (`DELETE`) |
| `dbQMakeTable` | Création de table (`SELECT … INTO`) |
| `dbQCrosstab` | Analyse croisée (`TRANSFORM`) |
| `dbQUnion` | Union |
| `dbQDDL` | DDL |
| `dbQSQLPassThrough` | Pass-Through |
| `dbQProcedure` | Procédure SQL |

```vba
If qdf.Type = dbQSelect Then
    ' requête de sélection → on ouvrira un recordset
Else
    ' requête action → on l'exécutera
End If
```

---

## Énumérer les requêtes sauvegardées

On parcourt l'ensemble des requêtes avec une boucle `For Each` :

```vba
Dim qdf As DAO.QueryDef
For Each qdf In CurrentDb.QueryDefs
    Debug.Print qdf.Name & "  [type " & qdf.Type & "]"
Next qdf
```

> ⚠️ **Le piège des requêtes masquées** : la collection `QueryDefs` contient aussi des requêtes **temporaires et masquées** générées automatiquement par Access — typiquement le SQL sous-jacent aux sources d'enregistrements de formulaires, d'états ou de listes déroulantes. Leur nom commence par un **tilde** (`~`), par exemple `~sq_c…`. Pour n'énumérer que les requêtes « réelles », on filtre ces noms :

```vba
For Each qdf In CurrentDb.QueryDefs
    If Left$(qdf.Name, 1) <> "~" Then         ' on ignore les requêtes masquées
        Debug.Print qdf.Name
    End If
Next qdf
```

---

## Lire le SQL d'une requête

La propriété `SQL` donne accès au texte de la requête — fort utile pour la documentation, l'audit ou la mise au point :

```vba
Dim qdf As DAO.QueryDef
Set qdf = CurrentDb.QueryDefs("R_ClientsActifs")
Debug.Print qdf.SQL
```

Cette propriété étant **en lecture et en écriture**, lui affecter une nouvelle valeur **modifie** la requête sauvegardée — opération détaillée à la section 12.2.

---

## Inspecter la structure renvoyée sans exécuter

Atout souvent méconnu : la collection **`Fields`** d'un `QueryDef` expose les **colonnes que la requête renverrait**, sans qu'il soit nécessaire de l'exécuter. C'est précieux pour les outils génériques ou la génération de documentation :

```vba
Dim fld As DAO.Field
For Each fld In qdf.Fields
    Debug.Print fld.Name & " (type " & fld.Type & ")"
Next fld
```

On obtient ainsi le **schéma de sortie** de la requête — noms et types de colonnes — sans en parcourir les données.

---

## Utiliser une requête sauvegardée

Accéder à une requête prend tout son sens lorsqu'on l'**utilise**. Une requête de sélection s'ouvre en recordset, une requête action s'exécute — en passant soit par le **nom** de la requête, soit par l'objet `QueryDef` :

```vba
' Requête de sélection → recordset (par nom)
Set rs = CurrentDb.OpenRecordset("R_ClientsActifs", dbOpenSnapshot)

' Requête action → exécution
CurrentDb.Execute "R_SupprimerInactifs", dbFailOnError

' Ou via l'objet QueryDef directement
qdf.Execute dbFailOnError
```

Recourir à une requête sauvegardée plutôt qu'à du SQL en ligne présente plusieurs avantages, déjà soulignés au chapitre 11.2 : elle est **précompilée** (son plan d'exécution optimisé est conservé), donc plus rapide pour un usage répété ; elle **sort le SQL du code**, ce qui améliore lisibilité et maintenance ; et elle constitue une **brique réutilisable**, notamment pour l'empilement de requêtes vu au chapitre 11.11.

---

## Requêtes sauvegardées et requêtes temporaires

Un `QueryDef` n'est pas nécessairement persistant. On peut créer une `QueryDef` **temporaire**, non enregistrée dans la base, en lui donnant un nom vide (`CreateQueryDef("")`) — technique déjà rencontrée pour le paramétrage (section 11.5) et les requêtes Pass-Through (section 11.9). Une telle requête n'apparaît pas dans la collection `QueryDefs` et disparaît à la fin de son utilisation. La création de requêtes, persistantes comme temporaires, est l'objet de la section suivante.

---

## En résumé

La collection **`QueryDefs`** d'un objet `Database` réunit les **requêtes sauvegardées**, chacune représentée par un objet `QueryDef` dont le SQL réside dans la propriété `SQL`. On y accède par nom ou par index (en conservant la référence `Database`), on l'énumère par `For Each` — **en filtrant les requêtes masquées** au nom préfixé d'un tilde. Les propriétés clés (`Name`, `SQL`, `Type`, `Fields`…) permettent d'inspecter une requête, d'**identifier son type** via les constantes `dbQ…`, et même d'examiner les **colonnes renvoyées** sans l'exécuter. Utiliser une requête sauvegardée — par `OpenRecordset` ou `Execute` — apporte précompilation, lisibilité et réutilisabilité. Enfin, un `QueryDef` peut être **temporaire** (nom vide), auquel cas il n'est pas persisté.

Nous savons désormais accéder aux requêtes existantes et les explorer. L'étape suivante consiste à agir sur elles : les **créer**, modifier leur SQL et les supprimer par code. C'est l'objet de la section suivante.


⏭️ [12.2. Créer, modifier et supprimer des requêtes par code](/12-querydefs-tabledefs/02-creer-modifier-supprimer-requetes.md)
