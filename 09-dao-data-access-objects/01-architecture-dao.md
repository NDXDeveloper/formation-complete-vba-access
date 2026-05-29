🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.1 Architecture DAO — DBEngine, Workspace, Database, TableDef, QueryDef

## Vue d'ensemble

DAO n'est pas un objet unique mais un **modèle objet** complet : un ensemble d'objets organisés en hiérarchie, du plus général — le moteur de base de données lui-même — au plus précis — un champ donné d'un enregistrement. Comprendre cette architecture, c'est disposer d'une carte qui permet de savoir à tout moment *où l'on se trouve* dans le modèle et *comment atteindre* l'élément que l'on souhaite manipuler.

Cette section décrit d'abord le principe général qui structure l'ensemble du modèle — l'imbrication d'objets et de collections — puis détaille les cinq piliers annoncés dans le titre (`DBEngine`, `Workspace`, `Database`, `TableDef`, `QueryDef`). Elle situe ensuite les objets de bas niveau (`Field`, `Index`, `Parameter`, `Relation`) ainsi que l'objet central de la manipulation des données, le `Recordset`, qui fera l'objet des sections suivantes.

> **Note de périmètre.** Cette section décrit le *rôle* et la *place* de chaque objet dans l'architecture. La création et la modification effectives des objets de structure — `TableDef`, `QueryDef`, `Relation` — par code sont traitées en détail au [chapitre 12](/12-querydefs-tabledefs/README.md). Le présent chapitre se concentre, dès la section 9.2, sur l'exploitation des données via le `Recordset`.

## Le principe directeur : objets et collections

Le modèle DAO repose sur un schéma qui se répète à tous les niveaux : un **objet** contient une ou plusieurs **collections**, et chaque collection regroupe des objets d'un même type. Ainsi un objet `Database` contient une collection `TableDefs` (l'ensemble des définitions de tables), une collection `QueryDefs` (les requêtes sauvegardées), une collection `Relations`, et ainsi de suite. Cette régularité est précieuse : une fois le principe assimilé pour un objet, il s'applique à tous les autres.

On accède à un membre d'une collection de deux manières : par son **indice** (numérique, à partir de 0) ou par son **nom**. Les deux écritures suivantes désignent le même objet :

```vba
db.TableDefs(0)               ' première table, par indice
db.TableDefs("tblClients")    ' table nommée, par son nom
```

Chaque collection expose une propriété `Count` indiquant le nombre d'éléments qu'elle contient. Lorsqu'on a modifié la structure de la base par code, il peut être nécessaire d'appeler la méthode `Refresh` de la collection pour que ces changements y soient reflétés.

Deux opérateurs cohabitent dans le code DAO et il est essentiel de ne pas les confondre. Le **point** (`.`) accède à une *propriété* ou à une *méthode* de l'objet : `rs.RecordCount`, `rs.MoveNext`. Le **point d'exclamation**, ou *bang* (`!`), accède à un *membre nommé de la collection par défaut* de l'objet : `rs!NomClient` est une écriture abrégée de `rs.Fields("NomClient").Value`. Autrement dit, le bang désigne presque toujours des données (un champ), tandis que le point désigne le fonctionnement de l'objet.

Enfin, certains objets possèdent une **collection par défaut**, ce qui autorise une syntaxe condensée par indices successifs. L'expression `DBEngine(0)(0)` se lit en réalité `DBEngine.Workspaces(0).Databases(0)` : `Workspaces` est la collection par défaut de `DBEngine`, et `Databases` celle de `Workspace`.

## DBEngine — la racine du modèle

`DBEngine` représente le moteur de base de données ACE lui-même. C'est l'objet situé au sommet de la hiérarchie : il n'est pas créable et il en existe une seule instance pour toute la session VBA. On l'utilise rarement de façon explicite, car les fonctions de plus haut niveau (comme `CurrentDb()`) dispensent d'y recourir, mais il reste le point d'origine de toute l'arborescence.

`DBEngine` contient principalement deux collections : `Workspaces`, l'ensemble des sessions de travail ouvertes, et `Errors`, qui recense les erreurs renvoyées par le moteur lors de la dernière opération (cette dernière est exploitée au [chapitre 13](/13-gestion-erreurs/05-erreurs-dao-ado.md)). Il porte aussi quelques propriétés globales, comme `Version` (la version du moteur) ou `SystemDB` (le chemin du fichier de groupe de travail), et expose des méthodes utilitaires telles que `CompactDatabase` ou `Idle`.

## Workspace — la session de travail

Un `Workspace` représente une **session** ouverte par un utilisateur sur le moteur. Il joue le rôle de conteneur des bases de données ouvertes au sein de cette session et, surtout, il constitue le périmètre de gestion des **transactions** : les méthodes `BeginTrans`, `CommitTrans` et `Rollback` s'appliquent au niveau du `Workspace` (sujet développé au [chapitre 14](/14-transactions/02-transactions-dao.md)).

Dans la quasi-totalité des cas, on travaille avec l'espace de travail par défaut, créé automatiquement par Access et accessible via `DBEngine.Workspaces(0)`. Il n'est généralement pas nécessaire d'en créer d'autres.

Le `Workspace` hébergeait historiquement les collections `Users` et `Groups`, liées au modèle de sécurité au niveau utilisateur (le fichier de groupe de travail `.mdw`). Ce modèle a été abandonné avec le format `.accdb` : ces collections n'ont plus d'usage pour une application moderne. De même, les *workspaces* de type ODBCDirect ont disparu. En pratique, pour une base `.accdb`, le `Workspace` ne vous intéressera qu'au titre des transactions.

## Database — l'objet central

L'objet `Database` est celui que vous manipulerez le plus. Il représente une base de données ouverte et donne accès à l'ensemble de son contenu, structurel comme transactionnel. C'est à son niveau que se rattachent les collections suivantes :

- `TableDefs` — les définitions de toutes les tables (y compris les tables liées et les tables système) ;
- `QueryDefs` — les requêtes sauvegardées dans la base ;
- `Relations` — les relations définies entre les tables ;
- `Recordsets` — les jeux d'enregistrements actuellement ouverts ;
- `Containers` — les conteneurs d'objets stockés et de leurs propriétés ;
- `Properties` — les propriétés, intégrées et personnalisées, de la base elle-même.

Deux de ses méthodes sont au cœur du travail quotidien : `OpenRecordset`, qui ouvre un jeu d'enregistrements à partir d'une table, d'une requête ou d'une instruction SQL, et `Execute`, qui exécute une requête action (INSERT, UPDATE, DELETE) sans renvoyer de résultat. Il propose également les méthodes de création des objets de structure — `CreateTableDef`, `CreateQueryDef`, `CreateRelation`, `CreateProperty` — qui seront mises en œuvre au chapitre 12.

On obtient un objet `Database` pointant sur la base courante de deux façons principales, `CurrentDb()` et `DBEngine(0)(0)`, dont les différences sont abordées plus loin dans cette section et approfondies aux sections [4.6](/04-modele-objet-access/06-currentdb-vs-dbengine.md) et [9.2](/09-dao-data-access-objects/02-ouverture-database.md).

## TableDef — la définition d'une table

Un `TableDef` décrit la **structure** d'une table : son nom, ses colonnes et ses index. Attention à la nuance de vocabulaire — un `TableDef` n'est pas le *contenu* de la table (les enregistrements relèvent du `Recordset`) mais sa *définition*. Il contient deux collections essentielles : `Fields`, la liste de ses champs, et `Indexes`, la liste de ses index, ainsi qu'une collection `Properties`.

Pour une table liée à une source externe, le `TableDef` porte en outre les propriétés `Connect` (la chaîne de connexion) et `SourceTableName` (le nom de la table dans la source d'origine), ce qui en fait le point d'entrée de la gestion programmatique des tables liées (voir [12.7](/12-querydefs-tabledefs/07-tables-liees.md)).

Parcourir la collection `TableDefs` permet, par exemple, de lister toutes les tables d'une base :

```vba
Dim db As DAO.Database
Dim tdf As DAO.TableDef

Set db = CurrentDb()
For Each tdf In db.TableDefs
    Debug.Print tdf.Name
Next tdf
```

À l'exécution, on remarquera la présence de tables système dont le nom commence par `MSys` : ce sont des objets internes à Access, qu'on identifie par le drapeau `dbSystemObject` de leur propriété `Attributes` et que l'on filtre habituellement lorsqu'on ne souhaite traiter que les tables applicatives.

## QueryDef — la définition d'une requête

Un `QueryDef` représente une **requête sauvegardée** dans la base. Sa propriété la plus parlante est `SQL`, qui contient le texte SQL de la requête, modifiable par code. On lit ainsi le contenu d'une requête existante :

```vba
Debug.Print db.QueryDefs("rqClientsActifs").SQL
```

Le `QueryDef` contient une collection `Fields` (les colonnes que la requête renvoie) et une collection `Parameters`, qui prend tout son sens pour les requêtes paramétrées : on y affecte la valeur de chaque paramètre avant exécution. Comme l'objet `Database`, il expose les méthodes `Execute` (pour une requête action) et `OpenRecordset` (pour une requête SELECT). Pour une requête de type *pass-through* destinée à un serveur distant, il porte aussi une propriété `Connect`.

La définition, la modification et l'exécution paramétrée des `QueryDef` par code sont détaillées aux sections [12.1](/12-querydefs-tabledefs/01-collection-querydefs.md) à [12.3](/12-querydefs-tabledefs/03-requetes-parametrees.md).

## Recordset — le point de contact avec les données

Le `Recordset` est l'objet qui matérialise un **jeu d'enregistrements** : le résultat d'une table, d'une requête ou d'une instruction SQL, que l'on parcourt et que l'on modifie ligne à ligne. C'est l'objet pivot de tout ce chapitre. On l'obtient par la méthode `OpenRecordset` d'un objet `Database`, `TableDef` ou `QueryDef`. Il possède sa propre collection `Fields`, donnant accès aux valeurs de chaque colonne de l'enregistrement courant.

Ses types, sa navigation, la lecture et l'écriture de ses champs, la recherche, le tri, le filtrage et sa libération font l'objet des sections [9.3](/09-dao-data-access-objects/03-recordset-types.md) à [9.13](/09-dao-data-access-objects/13-champs-multivalues-pieces-jointes.md). Il n'est donc présenté ici que pour le situer dans l'architecture d'ensemble.

## Les objets de bas niveau : Field, Index, Parameter

Au niveau le plus fin du modèle se trouvent les objets qui décrivent les éléments atomiques de la structure et des données.

Un `Field` représente une **colonne**. Selon le contexte, il décrit soit la définition d'un champ (lorsqu'il appartient à un `TableDef` : nom, type, taille, valeur par défaut, règle de validation…), soit la donnée d'un champ pour l'enregistrement courant (lorsqu'il appartient à un `Recordset`, via sa propriété `Value`). On accède par exemple aux métadonnées d'une colonne ainsi :

```vba
Dim fld As DAO.Field
Set fld = db.TableDefs("tblClients").Fields("CodePostal")
Debug.Print fld.Name, fld.Type, fld.Size
```

Un `Index` décrit un **index** d'une table : les champs qui le composent (sa propre collection `Fields`) et ses caractéristiques (`Primary`, `Unique`, `Required`). Un `Parameter`, enfin, représente un **paramètre** d'un `QueryDef` : on lui affecte une valeur avant d'exécuter la requête correspondante.

## Relation — les liens entre tables

Un objet `Relation` décrit une **relation** entre deux tables : la table parente (`Table`), la table enfant (`ForeignTable`), les champs mis en correspondance (sa collection `Fields`) et les options d'intégrité référentielle (`Attributes`, qui code notamment les mises à jour et suppressions en cascade). La collection `Relations` de l'objet `Database` regroupe l'ensemble de ces relations. Leur gestion programmatique est traitée à la section [12.6](/12-querydefs-tabledefs/06-collection-relations.md).

## Containers, Documents et Properties

À côté des données et de la structure, DAO expose un volet de **métadonnées**. La collection `Containers` d'une base regroupe des objets `Container` correspondant à des catégories d'objets stockés (tables, relations, bases…), chacun possédant une collection `Documents`. Un `Document` représente un objet sauvegardé et porte ses **propriétés personnalisées**.

Cette mécanique est le support de la collection `Properties` que l'on retrouve sur la plupart des objets DAO. Elle contient à la fois des propriétés intégrées (fournies par le moteur) et des propriétés personnalisées que l'application peut créer et lire pour y stocker ses propres réglages. Ce mécanisme est développé à la section [12.8](/12-querydefs-tabledefs/08-proprietes-personnalisees.md).

## Schéma récapitulatif

L'arborescence complète des objets abordés peut se représenter ainsi :

```
DBEngine                          ' le moteur (instance unique)
├─ Errors                         ' erreurs de la dernière opération
└─ Workspaces
   └─ Workspace                   ' session — gère les transactions
      └─ Databases
         └─ Database              ' base ouverte — objet central
            ├─ TableDefs
            │  └─ TableDef        ' définition d'une table
            │     ├─ Fields → Field
            │     └─ Indexes → Index
            ├─ QueryDefs
            │  └─ QueryDef        ' requête sauvegardée
            │     ├─ Fields → Field
            │     └─ Parameters → Parameter
            ├─ Relations
            │  └─ Relation → Fields → Field
            ├─ Recordsets
            │  └─ Recordset       ' jeu d'enregistrements (cœur du ch. 9)
            │     └─ Fields → Field
            ├─ Containers
            │  └─ Container → Documents → Document
            └─ Properties → Property
```

On retrouve, à chaque niveau, le même couple objet / collection ; c'est ce qui rend la navigation dans le modèle prévisible une fois le principe acquis.

## Deux portes d'entrée vers la Database

Dans le code exécuté à l'intérieur d'Access, on accède à la base courante de deux façons, qu'il convient de ne pas confondre.

`CurrentDb()` renvoie une **nouvelle instance** d'objet `Database` pointant sur la base courante. Cette instance reflète l'état réel et à jour des objets de la base, ce qui en fait le choix recommandé dans la grande majorité des cas.

`DBEngine(0)(0)` — c'est-à-dire `DBEngine.Workspaces(0).Databases(0)` — renvoie une **référence à l'instance existante** déjà ouverte par Access. Cette forme est plus rapide à obtenir mais peut ne pas refléter les modifications de structure récentes et présente quelques effets de bord historiques.

La comparaison détaillée de ces deux approches, avec leurs implications pratiques, est faite à la section [4.6](/04-modele-objet-access/06-currentdb-vs-dbengine.md), et l'ouverture proprement dite d'une base avec `CurrentDb()` est l'objet de la section suivante, [9.2](/09-dao-data-access-objects/02-ouverture-database.md).

## Points clés à retenir

Le modèle DAO s'organise en une hiérarchie régulière d'objets et de collections, du `DBEngine` au sommet jusqu'au `Field` à la base. La compréhension du couple objet / collection, et la distinction entre l'opérateur point (propriétés et méthodes) et l'opérateur bang (membres de la collection par défaut), suffisent à naviguer dans l'ensemble du modèle.

Au quotidien, l'objet `Database` — obtenu via `CurrentDb()` — est le point de départ de presque tout, et le `Recordset` est l'objet par lequel on lit et écrit les données. Les objets `TableDef`, `QueryDef` et `Relation` décrivent la structure de la base ; leur manipulation par code est l'affaire du chapitre 12. Le `Workspace`, enfin, n'interviendra essentiellement que dans le cadre des transactions.

⏭️ [9.2. Ouverture et gestion d'une Database avec CurrentDb()](/09-dao-data-access-objects/02-ouverture-database.md)
