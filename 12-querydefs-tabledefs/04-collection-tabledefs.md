🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 12.4. Collection TableDefs — inspection de la structure des tables

Après le tour complet des requêtes (sections 12.1 à 12.3), le chapitre bascule vers l'autre grand objet structurel : les **tables**. Cette section est le pendant, côté tables, de la section 12.1 : elle apprend à **accéder** à la structure des tables et à l'**inspecter** — énumérer les tables, examiner leurs champs, leurs index, distinguer tables locales et liées. La **création** et la **modification** de tables feront l'objet de la section 12.5, les relations de la 12.6, et la gestion des tables liées de la 12.7.

Inspecter la structure par code ouvre la voie à la documentation automatique, aux outils génériques pilotés par les métadonnées, et à l'audit de la base. Tout repose sur la collection **`TableDefs`** et les objets `TableDef` qu'elle contient.

---

## La collection TableDefs

On accède à la collection via la propriété `TableDefs` d'un objet `Database`, selon les modes habituels :

```vba
Dim db As DAO.Database
Dim td As DAO.TableDef
Set db = CurrentDb

Set td = db.TableDefs("Clients")    ' accès par nom
Set td = db.TableDefs(0)            ' accès par index
Debug.Print db.TableDefs.Count      ' nombre de définitions de tables
```

Comme pour les `QueryDefs`, on **conserve la référence** `Database` dès qu'on enchaîne plusieurs opérations (rappel des chapitres 4.6 et 12.1). Point important : la collection `TableDefs` ne contient pas seulement les tables utilisateur, mais aussi les **tables liées** et les **tables système** d'Access — d'où la nécessité de filtrer.

---

## Filtrer les tables système et masquées

La collection inclut les tables système d'Access (au nom préfixé de `MSys`) ainsi que d'éventuelles tables masquées. Pour n'énumérer que les **tables utilisateur**, la méthode robuste s'appuie sur la propriété **`Attributes`**, un indicateur binaire que l'on teste par opération `And` :

```vba
Dim td As DAO.TableDef
For Each td In CurrentDb.TableDefs
    If (td.Attributes And (dbSystemObject Or dbHiddenObject)) = 0 Then
        Debug.Print td.Name        ' tables utilisateur uniquement
    End If
Next td
```

Cette approche par `Attributes` est plus fiable que le filtrage sur le préfixe `MSys` du nom, car elle repose sur la nature réelle de l'objet plutôt que sur une convention de nommage.

---

## Les propriétés essentielles d'un TableDef

Un `TableDef` expose les propriétés permettant d'inspecter et de caractériser une table.

| Propriété | Contenu |
|---|---|
| `Name` | Nom de la table |
| `Fields` | Collection des champs (colonnes) |
| `Indexes` | Collection des index |
| `Connect` | Chaîne de connexion (tables liées) ; vide si locale |
| `SourceTableName` | Nom de la table dans la source (tables liées) |
| `Attributes` | Indicateurs binaires (système, masquée, liée…) |
| `RecordCount` | Nombre d'enregistrements (`-1` pour une table liée non ouverte) |
| `DateCreated` / `LastUpdated` | Dates de création / de modification |
| `ValidationRule` / `ValidationText` | Règle et message de validation au niveau table |

> 📌 **Attention à `RecordCount`** : pour une table **locale**, elle renvoie le nombre exact d'enregistrements ; pour une table **liée**, elle renvoie `-1` tant qu'un recordset n'a pas été ouvert sur la table. Pour compter de façon fiable les lignes d'une table liée, on ouvre un recordset (chapitres 9 et 10).

---

## Locale ou liée ? `Connect` et `Attributes`

Distinguer une table locale d'une table liée est un besoin fréquent. Le test le plus simple porte sur la propriété **`Connect`** : non vide pour une table liée (elle contient la chaîne de connexion vers la source), vide pour une table locale.

```vba
For Each td In CurrentDb.TableDefs
    If (td.Attributes And (dbSystemObject Or dbHiddenObject)) = 0 Then
        If Len(td.Connect) > 0 Then
            Debug.Print td.Name & " → liée (source : " & td.SourceTableName & ")"
        Else
            Debug.Print td.Name & " → locale"
        End If
    End If
Next td
```

La propriété `Attributes` affine la distinction : `dbAttachedTable` désigne une table liée à une source de type ISAM (autre base Access, dBASE…), `dbAttachedODBC` une table liée via ODBC (SQL Server, etc.). La gestion programmatique de ces tables liées est l'objet de la section 12.7.

---

## Inspecter les champs : la collection `Fields`

La collection **`Fields`** d'un `TableDef` expose les colonnes de la table, chacune représentée par un objet `Field`. Ses principales propriétés sont `Name`, `Type`, `Size`, `Required` (champ obligatoire), `AllowZeroLength` (chaîne vide autorisée pour le texte), `DefaultValue`, `Attributes` et `OrdinalPosition`.

```vba
Dim fld As DAO.Field
Set td = CurrentDb.TableDefs("Clients")
For Each fld In td.Fields
    Debug.Print fld.Name & " — type " & fld.Type & _
                ", taille " & fld.Size & _
                IIf(fld.Required, ", obligatoire", "")
Next fld
```

On obtient ainsi le **schéma complet** d'une table — noms, types, tailles, contraintes — sans en parcourir les données.

---

## Les types de champs (constantes DAO)

La propriété `Type` d'un champ renvoie une constante DAO. Le tableau suivant établit la correspondance avec les types Access.

| Type Access | Constante DAO |
|---|---|
| Texte court | `dbText` |
| Texte long (Mémo) | `dbMemo` |
| Octet | `dbByte` |
| Entier | `dbInteger` |
| Entier long | `dbLong` |
| Réel simple | `dbSingle` |
| Réel double | `dbDouble` |
| Monétaire | `dbCurrency` |
| Date/Heure | `dbDate` |
| Oui/Non | `dbBoolean` |
| NuméroAuto | `dbLong` **+ attribut** `dbAutoIncrField` |
| Identifiant réplication | `dbGUID` |
| Pièce jointe | `dbAttachment` |

> 📌 **Le NuméroAuto n'est pas un type distinct** : c'est un champ `dbLong` (entier long) doté de l'attribut **`dbAutoIncrField`**. Pour le détecter, on teste cet attribut :

```vba
If (fld.Attributes And dbAutoIncrField) <> 0 Then
    Debug.Print fld.Name & " est un NuméroAuto"
End If
```

(Les champs Pièce jointe et multivalués, de type `dbAttachment` et `dbComplex`, ont été abordés au chapitre 9.13.)

---

## Inspecter les index : la collection `Indexes`

La collection **`Indexes`** d'un `TableDef` recense les index de la table, chacun étant un objet `Index` doté des propriétés `Name`, `Fields` (les champs indexés), `Primary` (s'agit-il de la clé primaire ?), `Unique`, `Required` et `IgnoreNulls`.

Identifier la **clé primaire** consiste à repérer l'index dont la propriété `Primary` vaut `True` :

```vba
Dim idx As DAO.Index
For Each idx In CurrentDb.TableDefs("Clients").Indexes
    If idx.Primary Then
        Debug.Print "Clé primaire : " & idx.Name
    End If
Next idx
```

La collection `Fields` d'un index énumère elle-même les champs qui le composent — utile pour examiner un index multi-colonnes.

---

## Propriétés propres à Access

Au-delà des propriétés natives ci-dessus, les champs portent des **propriétés propres à Access** — légende (`Caption`), format d'affichage (`Format`), description (`Description`), masque de saisie, règle de validation au niveau champ. Ces propriétés ne sont pas des membres de premier niveau de l'objet `Field` : elles résident dans sa collection **`Properties`**, et tenter d'y accéder lorsqu'elles ne sont pas définies déclenche une erreur. Leur lecture et leur écriture, qui demandent une manipulation particulière, sont détaillées à la section 12.8.

---

## TableDefs (DAO) ou AllTables (Access) ?

Deux collections donnent accès aux tables, avec des vocations distinctes. La collection **`AllTables`** de `CurrentData` (chapitre 4.4) relève du catalogue d'objets d'Access : elle liste les tables et renseigne notamment leur état (ouverte ou non). La collection **`TableDefs`** de DAO, en revanche, donne accès à la **structure détaillée** — champs, index, attributs. Pour inspecter la *structure* d'une table, c'est `TableDefs` qu'on emploie ; `AllTables` répond plutôt à des besoins de catalogue et d'état.

---

## Cas d'usage de l'inspection

Inspecter la structure par code sert plusieurs objectifs concrets : **générer de la documentation** (liste des tables, champs et types), bâtir du **code générique** piloté par les métadonnées (qui s'adapte à la structure sans la coder en dur), **auditer ou valider** la structure d'une base, effectuer des **vérifications préalables** avant une opération structurelle (la table existe-t-elle ? possède-t-elle tel champ ?), ou alimenter une **interface dynamique**. C'est la brique de connaissance sur laquelle s'appuient les outils et les automatisations évoluées.

---

## En résumé

La collection **`TableDefs`** d'un objet `Database` réunit les **définitions de tables** — locales, liées **et système**. On y accède par nom ou par index (en conservant la référence `Database`), et l'on **filtre les tables système et masquées** via la propriété `Attributes` (`dbSystemObject`, `dbHiddenObject`). Les propriétés clés (`Name`, `Fields`, `Indexes`, `Connect`, `Attributes`…) permettent de caractériser une table — notamment de distinguer **locale** (`Connect` vide) et **liée** (`Connect` renseigné, `SourceTableName`). La collection **`Fields`** expose le schéma des colonnes (types via les constantes `dbXxx`, le **NuméroAuto** étant un `dbLong` + attribut `dbAutoIncrField`), et la collection **`Indexes`** les index, dont la clé primaire (`Primary = True`). Les **propriétés propres à Access** résident dans la collection `Properties` (section 12.8). L'inspection alimente documentation, code générique, audit et vérifications préalables.

Nous savons désormais lire la structure d'une table. L'étape suivante consiste à la **bâtir et à la modifier** par code — créer des tables, ajouter des champs, définir des index et des clés. C'est l'approche objet annoncée au chapitre 11.4, qu'aborde la section suivante.


⏭️ [12.5. Créer et modifier des tables par code (champs, index, clés)](/12-querydefs-tabledefs/05-creer-modifier-tables.md)
