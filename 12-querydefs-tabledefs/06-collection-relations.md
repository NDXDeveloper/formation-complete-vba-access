🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 12.6. Collection Relations — gestion des relations entre tables

La section précédente a créé tables, champs, clés et index — mais sans relier les tables entre elles. C'est le rôle des **relations**, qui matérialisent les liens de la base et, le plus souvent, son **intégrité référentielle**. Ce sont elles que l'on définit visuellement dans la fenêtre « Relations » d'Access ; DAO permet de les gérer par code via la collection **`Relations`** de l'objet `Database`.

Une relation est essentiellement une **clé étrangère** : elle relie un champ d'une table « parent » à un champ d'une table « enfant », et peut imposer des règles garantissant la cohérence des données. Cette section apprend à inspecter, créer et supprimer ces relations par code.

---

## Qu'est-ce qu'une relation ?

Une **relation** définit un lien entre deux tables et peut **appliquer l'intégrité référentielle** (RI) — c'est-à-dire interdire les enregistrements « orphelins ». Il convient de la distinguer d'une simple **jointure de requête** : une jointure est ad hoc, propre à une requête donnée, tandis qu'une relation est un lien **persistant au niveau du schéma**, mémorisé dans la base et assorti, en option, de règles d'intégrité. Les relations visibles dans la fenêtre « Relations » sont précisément ces objets `Relation`.

---

## La collection Relations

On accède à la collection via la propriété `Relations` d'un objet `Database` :

```vba
Dim db As DAO.Database
Dim rel As DAO.Relation
Set db = CurrentDb

Set rel = db.Relations("ClientsCommandes")   ' accès par nom
Debug.Print db.Relations.Count               ' nombre de relations
```

Comme pour les autres collections structurelles, on **conserve la référence** `Database` (rappel des chapitres 4.6 et 12.1).

---

## Les propriétés d'une Relation

Un objet `Relation` expose les propriétés décrivant le lien.

| Propriété | Contenu |
|---|---|
| `Name` | Nom de la relation |
| `Table` | Table **primaire** (côté « un », parent) |
| `ForeignTable` | Table **étrangère** (côté « plusieurs », enfant) |
| `Fields` | Collection des paires de champs liés |
| `Attributes` | Indicateurs binaires (intégrité, cascade, type de jointure…) |

> 📌 **Primaire ou étrangère ?** La propriété `Table` désigne la table **primaire** — celle qui porte la clé référencée, côté « un » de la relation. `ForeignTable` désigne la table **étrangère** — celle qui contient la clé étrangère, côté « plusieurs ». Cette distinction, que les noms ne rendent pas toujours intuitive, est essentielle pour créer correctement une relation.

---

## La collection Fields d'une relation

Le cœur d'une relation est sa collection **`Fields`**, qui définit les **paires de champs liés**. Chaque objet `Field` y porte deux noms : sa propriété **`Name`** désigne le champ dans la table **primaire**, et sa propriété **`ForeignName`** le champ correspondant dans la table **étrangère**.

Pour une relation reposant sur une **clé composite** (plusieurs champs), on ajoute plusieurs objets `Field` à cette collection, un par paire de champs liés.

---

## Les attributs : comportement de l'intégrité référentielle

La propriété **`Attributes`** gouverne le comportement de la relation, notamment l'application de l'intégrité référentielle et les cascades.

| Constante | Effet |
|---|---|
| *(0, par défaut)* | Applique l'intégrité référentielle, **sans** cascade |
| `dbRelationDontEnforce` | **N'applique pas** l'intégrité (lien logique seulement) |
| `dbRelationUnique` | Relation un-à-un |
| `dbRelationUpdateCascade` | **Mise à jour en cascade** |
| `dbRelationDeleteCascade` | **Suppression en cascade** |
| `dbRelationLeft` / `dbRelationRight` | Type de jointure (gauche / droite) dans les requêtes |
| `dbRelationInherited` | Relation héritée d'une base liée |

Par défaut (attributs à `0`), la relation **applique l'intégrité référentielle sans cascade**. Pour un lien purement logique sans contrainte, on précise explicitement `dbRelationDontEnforce`. Les attributs se combinent par `Or` (par exemple `dbRelationUpdateCascade Or dbRelationDeleteCascade`).

---

## Créer une relation par code

La création suit le même schéma `Append` explicite que la création de tables (section 12.5) : on crée la relation, on définit ses champs liants, puis on l'ajoute à la collection.

```vba
Dim db As DAO.Database
Dim rel As DAO.Relation
Dim fld As DAO.Field
Set db = CurrentDb

' 1. Créer la relation (table primaire, table étrangère, attributs)
Set rel = db.CreateRelation("ClientsCommandes", "Clients", "Commandes", _
                            dbRelationUpdateCascade Or dbRelationDeleteCascade)

' 2. Définir la paire de champs liés
Set fld = rel.CreateField("IDClient")     ' champ dans la table PRIMAIRE
fld.ForeignName = "IDClient"              ' champ correspondant dans la table ÉTRANGÈRE
rel.Fields.Append fld

' 3. Ajouter la relation
db.Relations.Append rel

Set rel = Nothing: Set db = Nothing
```

La méthode `CreateRelation(nom, tablePrimaire, tableÉtrangère, [attributs])` admet les attributs en quatrième argument. Ici, la relation applique l'intégrité référentielle avec mise à jour **et** suppression en cascade.

> ⚠️ **Prérequis de l'intégrité référentielle** : pour qu'une relation applique l'intégrité, le champ côté **primaire** doit être **uniquement indexé** (clé primaire ou index unique) dans sa table. Faute de quoi l'`Append` échoue. De plus, les **données existantes doivent être cohérentes** (aucune clé étrangère orpheline), sans quoi l'application de l'intégrité est refusée. Et les types des champs liés doivent être compatibles.

---

## Lire les relations existantes

L'inspection des relations sert à documenter le schéma ou à l'auditer. On parcourt la collection, et pour chaque relation, ses paires de champs :

```vba
Dim rel As DAO.Relation
Dim fld As DAO.Field
For Each rel In CurrentDb.Relations
    Debug.Print rel.Name & " : " & rel.Table & " → " & rel.ForeignTable
    For Each fld In rel.Fields
        Debug.Print "   " & fld.Name & " = " & fld.ForeignName
    Next fld
Next rel
```

Comme pour les `TableDefs`, la collection peut contenir des relations **héritées** d'une base liée (attribut `dbRelationInherited`), que l'on filtre au besoin.

---

## Supprimer une relation

La suppression s'effectue par la méthode `Delete` de la collection :

```vba
CurrentDb.Relations.Delete "ClientsCommandes"
```

Cette opération retire la relation **et la protection d'intégrité** qu'elle assurait. Comme les autres modifications structurelles, elle est **irréversible** et n'est pas annulable par une transaction.

---

## Cascade et intégrité référentielle

Les attributs de cascade méritent une explication, car ils déterminent le comportement de la base face aux modifications. La **mise à jour en cascade** (`dbRelationUpdateCascade`) propage le changement d'une clé primaire vers les clés étrangères des enregistrements enfants. La **suppression en cascade** (`dbRelationDeleteCascade`) supprime automatiquement les enregistrements enfants lorsqu'on supprime leur parent.

En l'**absence de cascade** (le comportement par défaut sous intégrité référentielle), la base **protège** les données : il devient impossible de supprimer un parent qui possède des enfants — ce qui explique précisément l'échec d'un `DELETE` sur un enregistrement parent évoqué à la section 11.3. L'intégrité référentielle gérée par code, et ses implications transactionnelles, sont approfondies au chapitre 14.6.

---

## Précautions

Plusieurs précautions encadrent la gestion des relations. Le **prérequis d'index unique** côté primaire est impératif pour l'intégrité, et les **données doivent être cohérentes** avant son application. Les **tables en cours d'utilisation** compliquent la modification. L'opération **n'est pas transactionnelle** (comme toute modification structurelle). On **vérifie l'existence** avant de créer (un nom de relation en double déclenche une erreur). Et l'on garde à l'esprit que **supprimer une relation retire sa protection** d'intégrité — à ne pas faire à la légère.

---

## En résumé

Les **relations** matérialisent les liens persistants entre tables et, le plus souvent, l'**intégrité référentielle** — à distinguer des jointures de requête, ad hoc et non persistantes. La collection **`Relations`** d'un objet `Database` les réunit, chaque `Relation` exposant `Table` (primaire, côté « un »), `ForeignTable` (étrangère, côté « plusieurs ») et une collection **`Fields`** où chaque champ lie `Name` (primaire) à `ForeignName` (étrangère). La propriété **`Attributes`** gouverne le comportement : intégrité par défaut, `dbRelationDontEnforce` pour un lien logique, `dbRelationUpdateCascade`/`dbRelationDeleteCascade` pour les cascades. On **crée** une relation via `CreateRelation` + `Fields.Append` + `Relations.Append` (le champ primaire devant être **uniquement indexé**), on l'**inspecte** par `For Each`, on la **supprime** par `Relations.Delete`. L'intégrité référentielle, qui protège notamment contre la suppression d'un parent ayant des enfants (section 11.3), est approfondie au chapitre 14.6.

Toutes les sections précédentes concernaient des tables locales. Mais une application Access scindée repose sur des **tables liées** vers un back-end distant. Leur gestion programmatique — création des liens, reliaison après déplacement — est un sujet à part entière, objet de la section suivante.


⏭️ [12.7. Tables liées (Linked Tables) — gestion programmatique](/12-querydefs-tabledefs/07-tables-liees.md)
