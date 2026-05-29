🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 12.5. Créer et modifier des tables par code (champs, index, clés)

Nous savons inspecter la structure d'une table (section 12.4) ; il s'agit maintenant de la **bâtir et de la modifier** par code. C'est l'approche **objet** annoncée au chapitre 11.4 comme alternative au DDL-SQL — et le pendant, côté création, de l'inspection précédente.

Rappelons l'arbitrage posé en 11.4 : le DDL-SQL est concis et portable, tandis que l'approche objet de DAO est plus verbeuse mais offre un **contrôle complet des propriétés propres à Access** (légendes, formats, descriptions) que le DDL ne sait pas définir. Cette section déploie pleinement cette approche objet, depuis la création d'une table jusqu'à la définition de ses propriétés les plus fines.

---

## Créer une table : la séquence DAO

Créer une table en DAO suit une séquence en plusieurs étapes : on crée la définition de table, on lui crée et ajoute des champs, puis on ajoute la table à la base.

```vba
Dim db As DAO.Database
Dim td As DAO.TableDef
Dim fld As DAO.Field
Set db = CurrentDb

' 1. Créer la définition de table
Set td = db.CreateTableDef("Produits")

' 2. Créer et AJOUTER les champs
Set fld = td.CreateField("IDProduit", dbLong)
fld.Attributes = dbAutoIncrField              ' NuméroAuto
td.Fields.Append fld

Set fld = td.CreateField("Designation", dbText, 100)
fld.Required = True
fld.AllowZeroLength = False
td.Fields.Append fld

Set fld = td.CreateField("Prix", dbCurrency)
td.Fields.Append fld

' 3. AJOUTER la table à la base
db.TableDefs.Append td

Set td = Nothing: Set db = Nothing
```

> 📌 **L'asymétrie de l'`Append`** : contrairement à un `QueryDef`, que `CreateQueryDef` enregistre automatiquement dès qu'on lui donne un nom (section 12.2), les objets `TableDef`, `Field` et `Index` exigent un **`Append` explicite**. On crée l'objet, *puis* on l'ajoute à la collection appropriée. Oublier l'`Append` est une erreur fréquente : l'objet existe en mémoire mais n'est jamais enregistré.

Deux points de syntaxe : la méthode `CreateField(nom, type, [taille])` n'utilise la **taille** que pour les champs texte (et binaires) — pour les types numériques, elle est implicite. Et le **NuméroAuto**, conformément à ce que nous avons vu en 12.4, n'est pas un type distinct : c'est un champ `dbLong` doté de l'attribut **`dbAutoIncrField`**.

Les propriétés **intrinsèques** du champ — `Required` (obligatoire), `AllowZeroLength` (chaîne vide autorisée), `DefaultValue`, `ValidationRule`/`ValidationText` — se définissent directement sur l'objet `Field`, comme ci-dessus pour `Required`.

---

## Créer la clé primaire et les index

Un index se crée via `CreateIndex`, se configure (clé primaire ? unique ?), reçoit les champs qui le composent, puis s'ajoute à la collection `Indexes`. Pour la **clé primaire** :

```vba
Dim idx As DAO.Index
Set idx = td.CreateIndex("PrimaryKey")        ' nom conventionnel de la clé primaire
idx.Primary = True
idx.Unique = True

' Ajouter le(s) champ(s) à l'index
Dim fldIdx As DAO.Field
Set fldIdx = idx.CreateField("IDProduit")
idx.Fields.Append fldIdx

td.Indexes.Append idx
```

Un index possède sa **propre collection `Fields`** : on y crée une référence de champ (`idx.CreateField`) que l'on ajoute. Pour un **index multi-colonnes**, on ajoute plusieurs champs. Les propriétés `Primary` et `Unique` caractérisent la clé primaire (unique par nature). On peut créer l'index sur la nouvelle table avant son `Append` final, ou l'ajouter ensuite à une table déjà enregistrée.

> 💡 Les **clés étrangères** et les relations entre tables ne se gèrent pas via les index, mais via la collection `Relations` — objet de la section 12.6.

---

## Modifier une table existante

On modifie la structure d'une table en agissant sur les collections `Fields` et `Indexes` du `TableDef` existant.

**Ajouter un champ** : on le crée et on l'ajoute à la collection `Fields` de la table existante.

```vba
Set td = CurrentDb.TableDefs("Produits")
Set fld = td.CreateField("Categorie", dbText, 50)
td.Fields.Append fld
```

**Supprimer un champ** : on appelle `Delete` sur la collection.

```vba
td.Fields.Delete "Categorie"
```

**Ajouter ou supprimer un index** suit la même logique sur la collection `Indexes` :

```vba
td.Indexes.Delete "idxDesignation"   ' suppression d'un index existant
```

---

## La limite : on ne change pas un type en place

Une limitation importante de l'approche DAO : **on ne peut pas modifier le type ni la taille d'un champ existant**. Les propriétés `Type` et `Size` d'un `Field` deviennent en **lecture seule** une fois le champ créé.

Pour changer le type ou la taille d'un champ d'une table existante, deux voies : recourir au **DDL `ALTER COLUMN`** (section 11.4), ou bien procéder par étapes — créer un nouveau champ, y copier les données, supprimer l'ancien, renommer. La voie DDL est généralement la plus simple pour cette opération précise.

---

## L'atout sur le DDL : les propriétés propres à Access

Voici la raison d'être de l'approche DAO, et son avantage décisif sur le DDL. DAO permet de définir les **propriétés propres à Access** — `Caption` (légende), `Format`, `Description`, `InputMask` (masque de saisie), `DecimalPlaces` — que le DDL-SQL ne sait pas configurer.

Une subtilité doit toutefois être comprise. Certaines propriétés sont **intrinsèques** (présentes d'office sur l'objet `Field`) : `Required`, `AllowZeroLength`, `DefaultValue`, `ValidationRule`, `ValidationText`. On les affecte directement. D'autres sont **non intrinsèques** : `Caption`, `Format`, `Description`, `InputMask`… Ces dernières **n'existent pas tant qu'on ne les a pas créées**, et tenter de les affecter directement déclenche une erreur. Il faut alors les **créer** via `CreateProperty` puis les ajouter à la collection `Properties` du champ.

Le patron habituel encapsule cette logique dans une fonction qui crée la propriété si elle n'existe pas, ou la met à jour sinon :

```vba
Public Sub DefinirProprieteChamp(ByVal fld As DAO.Field, ByVal nom As String, _
                                 ByVal typ As Integer, ByVal valeur As Variant)
    On Error Resume Next
    fld.Properties(nom) = valeur                          ' si elle existe déjà
    If Err.Number <> 0 Then                               ' sinon, on la crée
        Err.Clear
        fld.Properties.Append fld.CreateProperty(nom, typ, valeur)
    End If
    On Error GoTo 0
End Sub

' Usage
DefinirProprieteChamp td.Fields("Designation"), "Caption", dbText, "Désignation du produit"
DefinirProprieteChamp td.Fields("Prix"), "Format", dbText, "Standard"
```

Ce mécanisme de propriétés — central pour exploiter pleinement la richesse d'Access — est approfondi à la section 12.8.

---

## DAO ou DDL ? Rappel de l'arbitrage

Cette section illustre concrètement l'arbitrage du chapitre 11.4. L'approche **DAO** est plus **verbeuse** (création et `Append` en plusieurs étapes), mais offre le **contrôle complet** des propriétés Access (légendes, formats, descriptions, masques) — ce qui la rend incontournable dès qu'une table doit être créée avec ces caractéristiques. L'approche **DDL** (section 11.4) est plus **concise et portable**, idéale pour une structure simple ou destinée à une migration. On choisit selon le besoin : richesse Access et finesse de contrôle pour DAO, concision et portabilité pour le DDL.

---

## Précautions

Plusieurs précautions, dont certaines communes au DDL, encadrent ces opérations. Les modifications de structure **ne sont pas transactionnelles** (section 11.4) : aucun `Rollback` ne les défait. Une table **en cours d'utilisation** ne peut être modifiée. On **vérifie l'existence** avant de créer (créer une table déjà présente déclenche une erreur), à l'aide d'une fonction `TableExiste` comme celle de la section 11.4. Après création, `Application.RefreshDatabaseWindow` rafraîchit le volet de navigation (comme en 12.2). Enfin, rappelons que les **clés étrangères et relations** relèvent de la section suivante, et non des index.

---

## En résumé

Créer une table en DAO suit une **séquence en plusieurs étapes** : `CreateTableDef`, puis `CreateField` et `Fields.Append` pour chaque champ, puis `TableDefs.Append` pour enregistrer la table — avec un **`Append` explicite** obligatoire (contrairement au `QueryDef`). Le **NuméroAuto** est un `dbLong` + attribut `dbAutoIncrField`, et les propriétés intrinsèques (`Required`, `DefaultValue`…) se définissent directement. La **clé primaire** et les **index** se créent via `CreateIndex`, en y ajoutant des champs (`idx.Fields.Append`). On **modifie** une table en agissant sur ses collections `Fields` et `Indexes`, mais on **ne peut pas changer le type d'un champ en place** (recourir au DDL `ALTER COLUMN` ou recréer). L'**atout majeur de DAO** est la définition des **propriétés propres à Access** (`Caption`, `Format`, `Description`…), dont les non-intrinsèques exigent `CreateProperty` (section 12.8). Le choix DAO/DDL suit l'arbitrage du chapitre 11.4 : richesse et contrôle pour DAO, concision et portabilité pour le DDL.

Nous avons créé tables, champs, clés et index — mais sans encore relier les tables entre elles. C'est le rôle des **relations**, qui matérialisent les liens et l'intégrité référentielle de la base. La section suivante leur est consacrée.


⏭️ [12.6. Collection Relations — gestion des relations entre tables](/12-querydefs-tabledefs/06-collection-relations.md)
