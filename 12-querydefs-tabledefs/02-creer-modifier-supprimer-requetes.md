🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 12.2. Créer, modifier et supprimer des requêtes par code

La section précédente nous a appris à **accéder** aux requêtes sauvegardées et à les explorer. Nous passons maintenant à l'action : **créer** de nouvelles requêtes, **modifier** leur SQL et les **supprimer**, le tout par code. Ces opérations permettent d'automatiser la gestion de la couche requêtes — lors d'une installation, d'une migration, ou pour maintenir dynamiquement une bibliothèque de requêtes.

Le paramétrage des requêtes, qui mérite un traitement à part, est réservé à la section 12.3. Ici, nous nous concentrons sur le cycle de vie d'une requête : sa naissance, son évolution et sa suppression.

---

## Créer une requête : `CreateQueryDef`

La création repose sur la méthode **`CreateQueryDef`** de l'objet `Database`, dont la signature est :

```vba
db.CreateQueryDef(Nom, [TexteSQL])
```

### Le rôle décisif du nom : persistante ou temporaire

Un détail conditionne tout : la valeur du **nom**. C'est lui qui détermine si la requête est enregistrée ou non.

Avec un **nom non vide**, la requête est **persistante** : elle est enregistrée dans la base, ajoutée à la collection `QueryDefs`, et apparaît dans le volet de navigation. Avec un **nom vide** (`""`), la requête est **temporaire** : elle n'est ni enregistrée ni ajoutée à la collection, et disparaît dès que l'objet est libéré — c'est l'usage rencontré pour le paramétrage (section 11.5) et les requêtes Pass-Through (section 11.9).

> 📌 **Une asymétrie à connaître** : pour un `QueryDef`, **donner un nom suffit à l'enregistrer** — `CreateQueryDef` l'ajoute *automatiquement* à la collection. Inutile d'appeler une méthode `Append`. C'est différent des objets `TableDef`, `Field` ou `Index` (section 12.5), qui exigent un `Append` explicite. Cette particularité surprend souvent.

### Création directe

La forme la plus directe fournit le nom et le SQL en une fois :

```vba
Dim db As DAO.Database
Dim qdf As DAO.QueryDef
Set db = CurrentDb

Set qdf = db.CreateQueryDef("R_ClientsNord", _
    "SELECT * FROM Clients WHERE Region = 'Nord';")
' La requête est immédiatement enregistrée et visible dans le volet de navigation

Set qdf = Nothing
Set db = Nothing
```

### Création en deux temps

On peut aussi créer la requête puis définir son SQL séparément — utile lorsqu'on construit le SQL en plusieurs étapes :

```vba
Set qdf = db.CreateQueryDef("R_NouvelleRequete")
qdf.SQL = "SELECT * FROM Clients;"
```

---

## Modifier une requête : changer son SQL

Modifier une requête existante consiste à accéder à son `QueryDef` et à réaffecter sa propriété **`SQL`**. Le changement est **persisté** dans la base :

```vba
Dim qdf As DAO.QueryDef
Set qdf = CurrentDb.QueryDefs("R_ClientsNord")
qdf.SQL = "SELECT * FROM Clients WHERE Region = 'Sud';"   ' modifie la requête en place
```

D'autres propriétés se modifient de la même façon — par exemple `Connect` et `ReturnsRecords` pour transformer ou ajuster une requête Pass-Through (section 11.9).

> ⚠️ **Requête en cours d'utilisation** : modifier une requête actuellement ouverte (recordset en cours, formulaire lié) peut provoquer une erreur ou rester sans effet sur les instances déjà ouvertes. On s'assure qu'aucun objet ne l'utilise avant de la modifier.

---

## Supprimer une requête : `Delete`

La suppression s'effectue via la méthode `Delete` de la collection `QueryDefs` :

```vba
CurrentDb.QueryDefs.Delete "R_ClientsNord"
```

> ⚠️ **Opération irréversible** : comme le `DROP TABLE` du DDL (section 11.4), la suppression d'une requête ne se laisse pas annuler par une transaction. Et supprimer une requête dont **dépendent** d'autres requêtes, formulaires ou états les casse. On vérifie donc les dépendances avant d'agir.

---

## Vérifier l'existence avant d'agir

Créer une requête portant un nom **déjà existant** déclenche une erreur (objet déjà présent), et en supprimer une **inexistante** échoue également. On teste donc l'existence au préalable, via une fonction utilitaire — analogue à la fonction `TableExiste` de la section 11.4 :

```vba
Public Function RequeteExiste(ByVal nom As String) As Boolean
    Dim qdf As DAO.QueryDef
    On Error Resume Next
    Set qdf = CurrentDb.QueryDefs(nom)
    RequeteExiste = (Err.Number = 0)
    On Error GoTo 0
End Function
```

---

## L'idiome « créer ou remplacer »

Une situation fréquente consiste à vouloir installer une requête, qu'elle existe déjà ou non. L'idiome **« créer ou remplacer »** combine élégamment les opérations précédentes : si la requête existe, on met à jour son SQL ; sinon, on la crée.

```vba
Public Sub CreerOuRemplacerRequete(ByVal nom As String, ByVal sql As String)
    Dim db As DAO.Database
    Set db = CurrentDb

    If RequeteExiste(nom) Then
        db.QueryDefs(nom).SQL = sql              ' modifier l'existante
    Else
        Dim qdf As DAO.QueryDef
        Set qdf = db.CreateQueryDef(nom, sql)    ' créer (et enregistrer)
    End If

    Set db = Nothing
End Sub
```

Ce patron est particulièrement utile dans les scripts d'installation ou de mise à jour, où l'on veut garantir l'état final d'une requête sans présumer de son existence.

---

## Rafraîchir le volet de navigation

Une requête créée par code existe bien dans la base, mais le **volet de navigation** d'Access peut ne pas l'afficher immédiatement. Pour rafraîchir cet affichage, on appelle :

```vba
Application.RefreshDatabaseWindow
```

Cet appel est purement cosmétique — il ne concerne que l'interface, non la base elle-même —, mais il évite la confusion d'une requête créée et pourtant invisible dans le volet.

---

## Persistante ou temporaire : récapitulatif

La distinction mérite d'être ancrée. Un **nom non vide** produit une requête **enregistrée**, présente dans la collection, persistante d'une session à l'autre et visible dans le volet. Un **nom vide** (`""`) produit une requête **temporaire**, absente de la collection, qui s'évanouit avec l'objet. La forme temporaire est le véhicule des exécutions ponctuelles — notamment paramétrées (section 11.5) ou Pass-Through (section 11.9) —, dont l'usage paramétré complet est détaillé à la section suivante.

---

## Construction du SQL : les précautions habituelles

Lorsqu'on **construit le SQL** d'une requête à créer ou à modifier, toutes les précautions du chapitre 11 s'appliquent. La construction dynamique suit les bonnes pratiques de la section 11.6 (espaces, traçabilité) ; le formatage des éventuelles valeurs littérales obéit aux règles de la section 11.7 (dates, nombres, apostrophes) ; et si une saisie utilisateur entre dans le SQL, la vigilance anti-injection de la section 11.5 demeure de rigueur. Un `Debug.Print` du SQL avant de l'affecter reste, ici encore, le meilleur réflexe de mise au point.

---

## Quand créer ou modifier des requêtes par code ?

Cette manipulation programmatique trouve sa place dans plusieurs contextes : **bâtir des requêtes à l'installation** ou lors d'une migration, **générer dynamiquement** des requêtes de reporting, ou **maintenir une bibliothèque** de requêtes par script. En revanche, comme pour les modifications de structure (section 11.4), on reste **prudent quant à la modification de requêtes en production** : altérer une requête dont dépendent des formulaires ou des états en cours d'usage peut surprendre l'utilisateur. La modification de requêtes au fil de l'exécution se justifie, mais s'envisage avec discernement.

---

## En résumé

La méthode **`CreateQueryDef(Nom, [SQL])`** crée une requête, **persistante** si le nom est non vide (enregistrée et ajoutée *automatiquement* à la collection, sans `Append`), **temporaire** si le nom est vide. On **modifie** une requête en réaffectant sa propriété `SQL`, on la **supprime** par `QueryDefs.Delete` — une opération **irréversible**, à éviter sur des requêtes dont dépendent d'autres objets. On **teste l'existence** avant de créer ou supprimer (fonction `RequeteExiste`), l'idiome **« créer ou remplacer »** garantissant un état final déterministe, et `Application.RefreshDatabaseWindow` rafraîchit le volet de navigation. La **construction du SQL** reste soumise aux précautions du chapitre 11 (sécurité, formatage, traçabilité), et la modification de requêtes en production demande du discernement.

Nous avons couvert le cycle de vie d'une requête, mais en laissant de côté un aspect essentiel : les **paramètres**. Définir et exécuter par code une requête paramétrée — la voie DAO du paramétrage évoquée au chapitre 11.5 — fait l'objet de la section suivante.


⏭️ [12.3. Requêtes paramétrées — définition et exécution par code](/12-querydefs-tabledefs/03-requetes-parametrees.md)
