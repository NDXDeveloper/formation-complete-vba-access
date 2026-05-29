🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.3 Recordset DAO — types (Table, Dynaset, Snapshot, Forward-only)

## Introduction

Un `Recordset` n'est pas un objet monolithique : DAO en propose plusieurs **types**, chacun offrant un compromis différent entre capacités, fraîcheur des données, possibilité de modification et performance. Choisir le bon type au moment d'ouvrir un jeu d'enregistrements n'est pas un détail : un type mal adapté peut interdire une opération attendue (impossible de modifier un `Snapshot`), gaspiller de la mémoire (charger un gros jeu en `Snapshot` alors qu'un parcours unique suffisait) ou, à l'inverse, ouvrir la porte à des erreurs subtiles. Cette section présente les quatre types utilisés en pratique avec le moteur ACE — Table, Dynaset, Snapshot et Forward-only — leurs caractéristiques respectives, et la manière de choisir entre eux.

## Spécifier le type à l'ouverture

Le type se choisit au moment de l'appel à `OpenRecordset`, via une constante passée en deuxième argument :

```vba
Dim db As DAO.Database
Dim rs As DAO.Recordset
Set db = CurrentDb()

' Table-type : table locale uniquement
Set rs = db.OpenRecordset("tblClients", dbOpenTable)

' Dynaset : tables, requêtes, SQL, tables liées
Set rs = db.OpenRecordset( _
    "SELECT * FROM tblClients WHERE Ville='Rouen'", dbOpenDynaset)

' Snapshot : copie figée, lecture seule
Set rs = db.OpenRecordset("rqStatistiques", dbOpenSnapshot)

' Forward-only : lecture seule, parcours unique vers l'avant
Set rs = db.OpenRecordset("SELECT Montant FROM tblFactures", dbOpenForwardOnly)
```

Les quatre constantes correspondantes sont `dbOpenTable`, `dbOpenDynaset`, `dbOpenSnapshot` et `dbOpenForwardOnly`.

Lorsque le type est **omis**, DAO applique une règle par défaut : il crée un `Recordset` de type Table si la source est une table locale, et de type Dynaset dans tous les autres cas (requête, table liée, instruction SQL).

```vba
Set rs = db.OpenRecordset("tblClients")               ' → Table-type
Set rs = db.OpenRecordset("SELECT * FROM tblClients") ' → Dynaset
```

Indiquer le type explicitement reste préférable : on documente ainsi son intention et on évite les surprises liées au comportement par défaut.

## Table-type (`dbOpenTable`)

Le type Table donne un accès **direct** à une table de base **locale**. C'est sa contrainte la plus marquante : il ne fonctionne ni sur les requêtes, ni sur les tables liées, ni sur les sources ODBC — uniquement sur une table physiquement présente dans la base courante.

En contrepartie, il offre deux atouts majeurs. D'abord, il est entièrement modifiable et défilable dans les deux sens. Ensuite, et c'est sa particularité décisive, il est le **seul type à prendre en charge la méthode `Seek`** : en s'appuyant sur un index de la table (que l'on choisit via la propriété `Index`), `Seek` réalise une recherche indexée extrêmement rapide, bien plus performante que la méthode `Find` sur de gros volumes. Le détail de `Seek` et de `Find` est traité à la section [9.9](/09-dao-data-access-objects/09-recherche-recordset.md).

Le type Table présente enfin l'avantage d'une propriété `RecordCount` immédiatement exacte : elle reflète le nombre réel d'enregistrements de la table sans avoir à parcourir le jeu au préalable.

## Dynaset-type (`dbOpenDynaset`)

Le type Dynaset est le plus **polyvalent** et, de fait, le plus utilisé. Techniquement, il fonctionne comme un ensemble dynamique de pointeurs (un jeu de clés) vers les enregistrements concernés : les clés sont fixées à l'ouverture, mais les données rattachées à chaque clé sont relues à la demande.

Cette mécanique a une conséquence importante en environnement multi-utilisateur : un Dynaset **reflète les modifications de valeurs** apportées par d'autres utilisateurs sur les enregistrements présents dans le jeu. En revanche, les enregistrements ajoutés par autrui après l'ouverture n'y apparaissent pas tant qu'on n'a pas rafraîchi le jeu (`Requery`).

Le Dynaset est modifiable (lorsque la source l'est) et défilable dans les deux sens, avec prise en charge des signets (`Bookmark`). Il accepte toutes les sources : tables, requêtes, SQL et, point capital, **tables liées**. Il ne prend pas en charge `Seek` ; la recherche s'y fait par `Find`. C'est le type de prédilection dès qu'il faut éditer des données, travailler à travers des jointures ou interroger une table liée.

## Snapshot-type (`dbOpenSnapshot`)

Le type Snapshot est une **copie figée et en lecture seule** des données telles qu'elles existaient au moment de l'ouverture. Il n'est pas modifiable et ne reflète aucune modification ultérieure faite par d'autres utilisateurs : c'est une photographie.

Il charge les données en mémoire, ce qui en fait un choix très efficace pour les jeux de **petite à moyenne taille** consultés en lecture seule : une fois chargé, il n'y a plus de relecture, donc une navigation rapide. Sur de **gros volumes**, en revanche, cette copie intégrale en mémoire devient coûteuse, et un Dynaset (qui ne stocke que des clés) peut alors s'avérer plus économe.

Le Snapshot reste défilable dans les deux sens et prend en charge les signets. C'est le candidat naturel pour alimenter un état, afficher des données de consultation, ou réaliser des lectures qui n'ont pas vocation à modifier la base.

## Forward-only-type (`dbOpenForwardOnly`)

Le type Forward-only est, en quelque sorte, un Snapshot dépouillé : il est en lecture seule, mais on ne peut le parcourir **que vers l'avant** (`MoveNext`). Il n'autorise ni retour en arrière (`MovePrevious`, `MoveFirst`), ni signets.

Cette restriction est aussi sa force : c'est le type le **plus léger et le plus rapide** pour un parcours **séquentiel unique**, où chaque enregistrement est lu une fois et jamais revisité. Il convient idéalement aux traitements en une passe — calcul d'un cumul, export séquentiel, lecture en flux — sur lesquels les capacités de défilement et d'édition seraient inutiles. Quand le besoin se résume à « lire chaque ligne, une fois, du début à la fin », c'est le choix le plus performant.

## Et le type Dynamic ?

DAO définit aussi une constante `dbOpenDynamic`. Elle était liée aux *workspaces* de type ODBCDirect, abandonnés depuis Access 2007, et n'a pas d'usage avec le moteur ACE natif. On ne la rencontre donc pas dans le développement Access courant, et ce chapitre n'en traite pas.

## Tableau comparatif

| Type | Modifiable | Défilement | Sources acceptées | `Seek` | Reflète les modifs d'autrui | Usage typique |
|------|:----------:|:----------:|-------------------|:------:|:---------------------------:|---------------|
| **Table** | Oui | Avant / arrière | Table locale uniquement | Oui (avec index) | Oui | Accès direct et recherche indexée sur une table |
| **Dynaset** | Oui | Avant / arrière | Tables, requêtes, SQL, tables liées | Non (`Find`) | Oui (valeurs) | Édition, jointures, tables liées |
| **Snapshot** | Non | Avant / arrière | Idem Dynaset | Non | Non | Lecture seule, états, consultations |
| **Forward-only** | Non | Avant uniquement | Idem Dynaset | Non | Non | Parcours séquentiel unique, traitements rapides |

Concernant les signets : les types Table, Dynaset et Snapshot les prennent en charge ; le type Forward-only ne les supporte pas. Concernant `RecordCount` : le type Table fournit un comptage exact immédiat, tandis que les autres types ne reflètent que le nombre d'enregistrements déjà parcourus — il faut appeler `MoveLast` pour obtenir le total (voir [9.4](/09-dao-data-access-objects/04-navigation-recordset.md)).

## Comment choisir le bon type

Le raisonnement peut se résumer à quelques questions successives :

- **Faut-il modifier les données ?** Si non, on s'oriente vers Snapshot ou Forward-only. Si oui, vers Dynaset ou Table.
- **Le traitement est-il un parcours unique vers l'avant, sans retour ?** Si oui, Forward-only est le plus performant.
- **A-t-on besoin d'une recherche indexée très rapide sur une table locale ?** Si oui, Table-type avec `Seek` s'impose.
- **Travaille-t-on sur une table liée, une requête ou une jointure, avec édition ?** Alors Dynaset, car Table-type est exclu dans ces cas.
- **S'agit-il d'une consultation à naviguer librement (avant/arrière) sans modification ?** Snapshot convient bien, surtout sur des volumes modérés.

En pratique, deux types couvrent l'immense majorité des besoins : le **Dynaset** pour tout ce qui touche à l'édition et aux sources variées, et le **Forward-only** pour les lectures séquentielles rapides. Les types Table et Snapshot interviennent dans des cas plus spécifiques — respectivement la recherche indexée et la consultation figée.

## Une note sur les tables liées

Le point mérite d'être souligné car il est source d'erreurs : sur une **table liée** (Access liée ou source ODBC), le type Table n'est pas disponible. Toute tentative d'ouverture en `dbOpenTable` sur une telle source échoue. On utilisera donc, selon le besoin, un Dynaset (pour éditer) ou un Snapshot / Forward-only (pour lire). Le contexte des tables liées est développé à la section [12.7](/12-querydefs-tabledefs/07-tables-liees.md).

## Au-delà du type : les options et le verrouillage

Le type n'est que le premier paramètre d'`OpenRecordset`. D'autres arguments permettent d'affiner le comportement du jeu d'enregistrements — options de lecture seule, d'ajout, de cohérence, ou mode de verrouillage. Ces réglages, ainsi que les considérations fines de performance liées au choix du type, sont traités à la section [18.2](/18-optimisation-performance/02-optimisation-recordsets.md), et les stratégies de verrouillage à la section [15.3](/15-multi-utilisateurs/03-strategies-verrouillage.md). La présente section se limite volontairement aux types eux-mêmes.

## Points clés à retenir

DAO propose quatre types de `Recordset` exploitables avec le moteur ACE. Le type **Table** offre un accès direct et la recherche indexée `Seek`, mais uniquement sur une table locale. Le **Dynaset**, le plus polyvalent, est modifiable, accepte toutes les sources et reflète les modifications de valeurs des autres utilisateurs : c'est le choix par défaut pour l'édition. Le **Snapshot** est une copie figée en lecture seule, efficace sur de petits volumes. Le **Forward-only** est le plus léger et le plus rapide, réservé aux parcours uniques vers l'avant. Spécifier explicitement le type à l'ouverture, en fonction de ce que le traitement doit réellement faire, est une bonne habitude qui conditionne à la fois les capacités et les performances du code.

⏭️ [9.4. Navigation dans un Recordset (MoveFirst, MoveNext, EOF, BOF)](/09-dao-data-access-objects/04-navigation-recordset.md)
