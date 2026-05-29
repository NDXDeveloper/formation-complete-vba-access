🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 18.6. Indexation et impact sur les performances SQL

L'indexation est sans doute le levier le plus puissant sur la vitesse des requêtes. Un index bien placé transforme une recherche qui parcourait toute une table en un accès quasi direct ; son absence condamne le moteur à tout lire. Mais l'index n'est pas gratuit : il accélère les lectures au prix des écritures et de l'espace occupé. La bonne démarche n'est donc pas d'indexer partout, mais d'indexer **délibérément** les colonnes qui le justifient — et de vérifier le gain par la mesure.

## Ce qu'est un index et pourquoi il accélère les lectures

Un index est une structure ordonnée (un arbre équilibré) construite sur une ou plusieurs colonnes, qui permet au moteur de localiser des enregistrements sans balayer la table entière. C'est l'équivalent de l'index d'un livre : pour trouver un mot, on consulte l'index trié plutôt que de lire toutes les pages.

Sans index sur la colonne concernée, une condition de recherche oblige le moteur à un **balayage complet** de la table (full scan) : il examine chaque ligne. Sur quelques centaines d'enregistrements, c'est instantané ; sur des centaines de milliers, et a fortiori à travers le réseau, c'est rédhibitoire.

## Où les index font la différence

Les index profitent à tout ce qui suppose de localiser ou d'ordonner des lignes : les critères d'une clause `WHERE` (égalité, plage de valeurs), les colonnes de **jointure**, le `ORDER BY` (un index dont l'ordre correspond évite l'étape de tri), ainsi que `GROUP BY` et `DISTINCT`. Les recherches par fonctions de domaine et la méthode `Seek` d'un Recordset de type table (voir la [section 18.2](02-optimisation-recordsets.md)) s'appuient directement sur les index.

## Le revers : le coût des index

Chaque index doit être tenu à jour : un `INSERT`, un `UPDATE` ou un `DELETE` met à jour la table **et** tous les index concernés, ce qui ralentit les écritures. Les index occupent par ailleurs de l'espace, contribuant à la taille du fichier — et donc à la limite des 2 Go (voir la [section 18.9](09-limites-taille-2go.md)) — comme à son gonflement. Multiplier les index pénalise donc les écritures, alourdit le fichier et donne davantage de travail à l'optimiseur. D'où la règle : on indexe les colonnes réellement utilisées en recherche, jointure ou tri, et celles-là seulement.

## Quoi indexer — et quoi ne pas indexer

### La sélectivité, critère clé

Un index n'est utile que s'il restreint à peu de lignes. Une colonne à forte **cardinalité** — beaucoup de valeurs distinctes, comme une date, un nom ou un code — est un bon candidat. À l'inverse, une colonne à faible cardinalité (un booléen Actif oui/non, une poignée de catégories) profite peu de l'indexation : l'index ne réduit guère l'ensemble, et l'optimiseur peut même choisir de l'ignorer.

### Les clés étrangères

Les colonnes de clé étrangère méritent presque toujours un index, car elles servent aux jointures : joindre sur une colonne non indexée est l'une des causes les plus fréquentes de lenteur. Il ne faut pas compter sur Access pour les indexer automatiquement — la création d'une relation ne crée pas d'index sur le côté clé étrangère. On les ajoute donc explicitement.

> ℹ️ Les relations entre tables sont traitées au [chapitre 12.6](../12-querydefs-tabledefs/06-collection-relations.md).

### Ce qui ne s'indexe pas, ou mal

Les champs Mémo (texte long), OLE et Pièce jointe ne peuvent être indexés. Indexer une table minuscule n'apporte rien (son balayage est immédiat). Et sur une table très sollicitée en écriture, un index dont le bénéfice en lecture est marginal coûte plus qu'il ne rapporte.

## Index simples et composites

Un index peut porter sur une seule colonne ou sur plusieurs (index composite). Un index composite sur `(A, B)` accélère les requêtes filtrant sur `A`, ou sur `A` et `B`, mais **pas** sur `B` seul : seul le préfixe gauche est exploitable. L'ordre des colonnes est donc déterminant. On crée un index composite lorsque les requêtes filtrent ou trient régulièrement sur la même combinaison de colonnes.

## Créer un index par code (DDL)

Les index se créent et se suppriment en SQL DDL, exécuté via `Execute` (voir la [section 18.3](03-execute-vs-runsql.md)).

```sql
-- Index simple
CREATE INDEX idxClientNom ON Clients (Nom);

-- Index unique (interdit aussi les doublons)
CREATE UNIQUE INDEX idxClientEmail ON Clients (Email);

-- Index composite (l'ordre des colonnes compte)
CREATE INDEX idxCmdClientDate ON Commandes (ClientID, DateCmd);

-- Suppression
DROP INDEX idxClientNom ON Clients;
```

Le dialecte Jet accepte des clauses complémentaires comme `WITH PRIMARY` (clé primaire), `WITH DISALLOW NULL` ou `WITH IGNORE NULL`. La création d'index par DAO, sur l'objet `TableDef`, est traitée au [chapitre 12.5](../12-querydefs-tabledefs/05-creer-modifier-tables.md).

> ℹ️ Voir les requêtes DDL ([11.4](../11-sql-access-vba/04-requetes-ddl.md)), le dialecte Jet SQL ([11.8](../11-sql-access-vba/08-dialecte-jet-sql.md)) et la référence Jet SQL (annexe [E](../annexes/e-reference-jet-sql.md)).

## Écrire des conditions « indexables »

Disposer d'un index ne garantit pas qu'il sera utilisé : la façon d'écrire la condition peut le rendre inopérant. Le piège le plus courant est d'appliquer une **fonction à la colonne indexée** : le moteur ne peut alors plus s'appuyer sur l'index et retombe sur un balayage.

```sql
-- Défait l'index sur DateCmd : fonction appliquée à la colonne
WHERE Year([DateCmd]) = 2025

-- Indexable : on exprime une plage de valeurs
WHERE DateCmd >= #2025-01-01# AND DateCmd < #2026-01-01#
```

De même, un joker en tête de `LIKE` empêche l'usage de l'index, alors qu'un préfixe fixe l'autorise.

```sql
-- Non indexable (joker initial)
WHERE Nom LIKE '*Paris'
-- Indexable (préfixe fixe)
WHERE Nom LIKE 'Paris*'
```

(Le moteur Access utilise `*` et `?` comme jokers ; en mode ANSI-92 ou via ADO, ce sont `%` et `_`.) Les conditions de négation (`<>`, `NOT IN`) et les comparaisons entre types différents nuisent également à l'exploitation des index.

> ℹ️ Le format des dates dans le SQL est traité au [chapitre 11.7](../11-sql-access-vba/07-localisation-formatage-sql.md).

## Le rôle de l'optimiseur et des statistiques

Le moteur ACE intègre un optimiseur fondé sur les coûts (technologie Rushmore, capable de combiner plusieurs index). À partir de **statistiques** sur les tables et les index, il décide pour chaque requête s'il vaut mieux utiliser un index ou balayer la table — et peut légitimement ignorer un index s'il l'estime moins rentable.

Ces statistiques sont rafraîchies, et les plans des requêtes sauvegardées recompilés, lors du **compactage** de la base. Des statistiques périmées — après d'importantes modifications de données sans compactage — peuvent conduire à de mauvais choix de plan. Le compactage régulier (voir la [section 18.7](07-compactage-automatique.md)) participe donc directement à la performance des requêtes, en plus de réduire le fichier.

Pour vérifier qu'un index est effectivement utilisé, on consulte le plan d'exécution via *ShowPlan*, présenté à la [section 18.1](01-profilage-mesure-performances.md).

## Tables liées et bases serveur

Pour une table liée à SQL Server (ou un autre serveur), l'indexation se gère **côté serveur** : la table liée hérite des index du serveur. La performance dépend alors de ces index et de la capacité d'Access à transmettre les filtres au serveur plutôt qu'à rapatrier les lignes pour filtrer localement. Les requêtes pass-through et la migration vers un moteur serveur sont traitées aux chapitres [11.9](../11-sql-access-vba/09-requetes-pass-through.md) et [23](../23-migration-interoperabilite/README.md).

## Mesurer avant et après

L'indexation se pilote par la mesure, pas par l'intuition : on relève les performances d'une requête, on ajoute l'index pressenti, puis on mesure de nouveau pour confirmer le gain — en s'aidant des techniques de la [section 18.1](01-profilage-mesure-performances.md) et de *ShowPlan*. On se méfie enfin de l'option « Indexation automatique à l'import/création », qui crée des index sur les champs dont le nom contient certains motifs (ID, clé, code, num…) : pratique parfois, mais source de sur-indexation involontaire.

## Points clés à retenir

- Un index évite le balayage complet : c'est le levier le plus puissant sur la vitesse des lectures, des jointures et des tris.
- Il a un coût : ralentissement des écritures et occupation d'espace (donc fichier plus gros et gonflement). On indexe délibérément, pas systématiquement.
- On indexe les colonnes à forte sélectivité, les clés étrangères (non indexées automatiquement) et les colonnes de recherche, jointure ou tri ; on n'indexe pas les Mémo/OLE, ni les très petites tables.
- Un index composite suit la règle du préfixe gauche : l'ordre des colonnes est décisif.
- Les index se créent en DDL (`CREATE INDEX`, `CREATE UNIQUE INDEX`, `DROP INDEX`) via `Execute`, ou par DAO.
- Une condition mal écrite défait l'index : éviter les fonctions sur la colonne (préférer une plage), les jokers initiaux et les négations.
- L'optimiseur s'appuie sur des statistiques rafraîchies par le compactage ; `ShowPlan` permet de vérifier l'usage réel des index, et la mesure de confirmer le gain.

---


⏭️ [18.7. Compactage automatique de la base par code](/18-optimisation-performance/07-compactage-automatique.md)
