🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 18.2. Optimisation des Recordsets — choix du type et des options

Le Recordset est le principal outil d'accès aux données en VBA, et c'est aussi l'un des plus gros leviers de performance d'une application Access. Son coût dépend fortement de deux décisions prises à l'ouverture : le **type** de jeu d'enregistrements et les **options** associées. Un mauvais choix matérialise inutilement des milliers de lignes en mémoire ou multiplie les allers-retours réseau ; un bon choix rend l'opération quasi instantanée. Cette section traite l'angle performance : le fonctionnement des Recordsets en lui-même est couvert au [chapitre 9](../09-dao-data-access-objects/README.md), et leurs curseurs ADO au [chapitre 10.4](../10-ado-access/04-recordset-ado-curseurs.md).

## La première question : faut-il vraiment un Recordset ?

Avant même de choisir un type, il faut se demander si le Recordset est nécessaire. La plus grande erreur de performance liée aux Recordsets n'est pas le mauvais réglage : c'est le **traitement ligne par ligne** d'une tâche que le moteur réaliserait en une seule passe. Ouvrir un jeu d'enregistrements et boucler pour modifier chaque ligne est presque toujours bien plus lent qu'une unique requête ensembliste.

```vba
Dim db As DAO.Database, rs As DAO.Recordset
Set db = CurrentDb

' À éviter : mise à jour ligne par ligne
Set rs = db.OpenRecordset("SELECT * FROM Articles", dbOpenDynaset)
Do Until rs.EOF
    rs.Edit
    rs!PrixTTC = rs!PrixHT * 1.2
    rs.Update
    rs.MoveNext
Loop
rs.Close

' À préférer : une seule requête ensembliste
db.Execute "UPDATE Articles SET PrixTTC = PrixHT * 1.2", dbFailOnError
```

Le Recordset garde toute sa pertinence lorsqu'on a besoin d'une logique procédurale que le SQL ne sait pas exprimer (appels de fonctions ligne par ligne, traitement conditionnel complexe, alimentation d'un objet métier). Pour tout le reste, on privilégie le SQL.

> ℹ️ Le choix entre `Execute` et `RunSQL` est traité en [section 18.3](03-execute-vs-runsql.md) ; les requêtes action en [chapitre 11.3](../11-sql-access-vba/03-requetes-action-insert-update-delete.md).

## Choisir le bon type de Recordset (DAO)

Lorsqu'un Recordset est justifié, le type conditionne son coût. DAO propose quatre types, du plus léger au plus polyvalent.

| Type | Modifiable | Reflète les modifs d'autrui | Sources autorisées | Atout |
|---|---|---|---|---|
| `dbOpenForwardOnly` | Non | — | Tables, requêtes, liées | Le plus léger : parcours unique |
| `dbOpenSnapshot` | Non | Non | Tables, requêtes, liées | Copie en mémoire, plus d'accès à la source |
| `dbOpenDynaset` | Oui | Oui | Tables, requêtes, liées | Polyvalent, économe sur gros volumes |
| `dbOpenTable` | Oui | Oui | Tables locales uniquement | Recherche indexée par `Seek` |

Si l'on ne précise pas le type, `OpenRecordset` ouvre un **Table-type** sur une table locale, et un **Dynaset** dans tous les autres cas (requête, table liée, chaîne SQL). Préciser explicitement le type évite de subir ce choix par défaut.

### Forward-only — le parcours unique le plus léger

C'est le curseur le plus économe. En lecture seule et sans retour en arrière, il convient parfaitement quand on ne fait que parcourir une fois les enregistrements : agréger, exporter, alimenter une structure. Il ne supporte ni `MovePrevious`, ni `MoveLast`, ni les signets.

```vba
Set rs = db.OpenRecordset( _
    "SELECT Montant FROM Ventes WHERE Annee = 2025", _
    dbOpenForwardOnly)
```

### Snapshot — la copie en lecture seule

Un snapshot charge une copie statique des données en mémoire. Une fois ouvert, il n'effectue plus d'accès à la source, ce qui le rend rapide en navigation… mais coûteux en mémoire dès que le volume grandit, puisqu'il matérialise tout. Il est idéal pour de **petits jeux en lecture seule** (tables de référence, listes de codes), et inadapté aux gros ensembles.

```vba
Set rs = db.OpenRecordset( _
    "SELECT Code, Libelle FROM Pays", _
    dbOpenSnapshot)
```

### Dynaset — le polyvalent modifiable

Le dynaset ne charge pas les données elles-mêmes mais un ensemble de références (clés) vers les enregistrements ; il récupère les valeurs au fur et à mesure des déplacements. Cela le rend **économe sur les gros volumes**, modifiable, et capable de refléter les modifications faites par d'autres utilisateurs. C'est le choix par défaut pour la mise à jour et pour les ensembles importants que l'on ne souhaite pas matérialiser intégralement.

### Table-type — la recherche indexée la plus rapide

Réservé aux tables locales (ni requêtes, ni tables liées), le Table-type donne accès aux index de la table et à la méthode `Seek`, la recherche la plus rapide qui soit sur une colonne indexée — bien plus performante que `FindFirst` sur un dynaset.

```vba
Set rs = CurrentDb.OpenRecordset("Clients", dbOpenTable)
rs.Index = "PrimaryKey"      ' index sur lequel effectuer la recherche
rs.Seek "=", 12345           ' recherche indexée, quasi instantanée
If Not rs.NoMatch Then Debug.Print rs!Nom
rs.Close
```

> ℹ️ La comparaison `Seek` / `FindFirst` est approfondie au [chapitre 9.9](../09-dao-data-access-objects/09-recherche-recordset.md).

## Ne charger que le nécessaire

Le type bien choisi ne suffit pas si l'on demande au moteur trop de données. Sur une base partagée, le volume transféré sur le réseau est souvent le facteur dominant.

### Restreindre colonnes et lignes à la source

Ouvrir un Recordset sur une table entière (`OpenRecordset("Clients")`) ou avec `SELECT *` rapatrie bien plus que nécessaire. On limite les colonnes aux seuls champs utilisés et les lignes par une clause `WHERE`, de façon que la sélection soit faite par le moteur avant le transfert.

```vba
' Coûteux : toute la table, toutes les colonnes
Set rs = db.OpenRecordset("Clients", dbOpenSnapshot)

' Économe : seulement ce qui est utile
Set rs = db.OpenRecordset( _
    "SELECT ClientID, Nom FROM Clients WHERE Ville = 'Paris'", _
    dbOpenSnapshot)
```

### Filtrer et trier à la source, pas après coup

Les propriétés `.Filter` et `.Sort` d'un Recordset s'appliquent à un jeu **déjà ouvert** : les données ont donc déjà été récupérées. La clause `WHERE` (et `ORDER BY`) à la source, elle, est traitée par le moteur et limite réellement le volume transféré — éventuellement en s'appuyant sur un index. À chaque fois que c'est possible, on filtre et l'on trie dans le SQL d'origine plutôt que sur le Recordset.

> ℹ️ Voir le [chapitre 9.10](../09-dao-data-access-objects/10-tri-filtrage-recordset.md) sur le tri et le filtrage côté Recordset.

## Les options d'ouverture qui changent tout

Le troisième argument d'`OpenRecordset` accepte des indicateurs qui modifient sensiblement le comportement et le coût.

| Option | Effet | Quand l'utiliser |
|---|---|---|
| `dbReadOnly` | Ouvre en lecture seule | Dès qu'aucune écriture n'est prévue : moins de surcharge, pas de verrou de mise à jour |
| `dbAppendOnly` | Ne charge aucun enregistrement existant | Ajout pur : ouverture instantanée quelle que soit la taille de la table |
| `dbSeeChanges` | Détecte les modifications concurrentes | Tables liées ODBC (colonnes identité / horodatage) : évite l'erreur 3197 |
| `dbSQLPassThrough` | Envoie le SQL directement au serveur | Sources ODBC : laisse le serveur exécuter la requête sans réinterprétation par ACE |

L'option `dbAppendOnly` mérite une mention particulière, car elle est sous-utilisée. Lorsqu'on n'a qu'à insérer des lignes — alimenter un journal, ajouter des écritures —, elle évite de charger l'existant : l'ouverture devient immédiate, même sur une table volumineuse.

```vba
Set rs = db.OpenRecordset("Journal", dbOpenDynaset, dbAppendOnly)
rs.AddNew
rs!Horodatage = Now
rs!Message = "Connexion"
rs.Update
rs.Close
```

L'option `dbSeeChanges` relève davantage de la correction que de la performance : sans elle, modifier un enregistrement d'une table liée SQL Server comportant une colonne d'horodatage (ou modifiée par ailleurs) peut déclencher l'erreur « les données ont été modifiées ». Elle ajoute un léger surcoût de détection, mais est souvent indispensable dans un contexte ODBC.

## Le piège de `RecordCount`

Sur un dynaset ou un snapshot, la propriété `.RecordCount` ne reflète pas le nombre total d'enregistrements tant qu'on ne les a pas tous parcourus : elle renvoie le nombre de lignes *visitées jusqu'ici* (typiquement 1 juste après l'ouverture d'un jeu non vide). Pour obtenir le total, il faut donc forcer un `MoveLast` — ce qui rapatrie *tous* les enregistrements, opération coûteuse sur un gros volume et sur le réseau.

En conséquence : si l'on veut seulement savoir s'il existe des enregistrements, on teste `.EOF`, jamais `MoveLast`.

```vba
If rs.EOF Then
    ' aucun enregistrement
Else
    ' au moins un enregistrement
End If
```

Si l'on a besoin d'un décompte exact, une fonction de domaine (`DCount`) ou un `SELECT COUNT(*)` laisse le moteur compter sans transférer les lignes — bien plus économe que de matérialiser le Recordset entier.

```vba
Dim n As Long
n = DCount("*", "Commandes", "Statut = 'En attente'")
```

> ℹ️ Les fonctions de domaine sont détaillées au [chapitre 11.10](../11-sql-access-vba/10-fonctions-domaine.md).

## Optimiser les boucles de parcours

Quand un parcours procédural est réellement nécessaire, quelques détails comptent sur de gros volumes.

L'accès répété à un champ par son nom — que ce soit `rs!Montant` ou `rs.Fields("Montant")` — déclenche une recherche dans la collection à *chaque* itération. Mettre en cache la référence du champ avant la boucle évite ce coût répété.

```vba
Dim rs As DAO.Recordset, fldMontant As DAO.Field
Dim total As Currency

Set rs = CurrentDb.OpenRecordset( _
    "SELECT Montant FROM LignesCommande WHERE CommandeID = 42", _
    dbOpenForwardOnly)

Set fldMontant = rs.Fields("Montant")   ' référence mise en cache
Do Until rs.EOF
    total = total + fldMontant.Value
    rs.MoveNext
Loop
rs.Close
```

On évite par ailleurs les `MoveLast` puis `MoveFirst` effectués par habitude à l'ouverture (par exemple « pour charger tous les enregistrements ») : ils n'apportent rien et provoquent un transfert intégral inutile.

## Requêtes sauvegardées plutôt que SQL à la volée

Une requête enregistrée dans la base voit son plan d'exécution optimisé et mémorisé par le moteur. Ouvrir un Recordset à partir d'une telle requête évite l'analyse et l'optimisation que subit, à chaque appel, une chaîne SQL passée en clair. Pour une requête fréquemment exécutée, la version sauvegardée — idéalement paramétrée — est donc plus rapide et plus propre.

```vba
Dim qd As DAO.QueryDef, rs As DAO.Recordset
Set qd = CurrentDb.QueryDefs("qClientsActifs")   ' plan déjà optimisé
Set rs = qd.OpenRecordset(dbOpenSnapshot)
```

> ℹ️ Les requêtes paramétrées exécutées par code sont traitées au [chapitre 12.3](../12-querydefs-tabledefs/03-requetes-parametrees.md).

## Et avec ADO ?

Les mêmes principes s'appliquent en ADO, exprimés différemment via `CursorType` et `LockType`. Le couple le plus rapide en lecture est le curseur *firehose* : `adOpenForwardOnly` associé à `adLockReadOnly`, équivalent du forward-only en lecture seule de DAO. `adOpenStatic` correspond au snapshot. L'emplacement du curseur (`CursorLocation`) joue également : `adUseServer` garde le curseur côté source, tandis que `adUseClient` rapatrie l'ensemble côté client — nécessaire pour les Recordsets déconnectés et pour obtenir un `RecordCount` fiable (lequel vaut −1 avec un curseur forward-only).

> ℹ️ Les curseurs et modes de verrouillage ADO sont détaillés au [chapitre 10.4](../10-ado-access/04-recordset-ado-curseurs.md).

## Libérer les Recordsets

Un Recordset consomme de la mémoire et, sur une base partagée, peut maintenir des ressources côté moteur. On le ferme (`Close`) et on libère la référence (`Set rs = Nothing`) dès qu'il n'est plus utile — point d'autant plus important dans une application qui ouvre de nombreux jeux au cours d'une session.

> ℹ️ La clôture et la libération des objets DAO sont détaillées au [chapitre 9.11](../09-dao-data-access-objects/11-cloture-liberation-dao.md).

## Points clés à retenir

- Avant de choisir un type, vérifier qu'un Recordset est nécessaire : une requête ensembliste bat presque toujours le traitement ligne par ligne.
- Le type se choisit selon l'usage : forward-only pour un parcours unique, snapshot pour de petits jeux en lecture seule, dynaset pour les gros volumes et la mise à jour, Table-type pour la recherche indexée par `Seek`.
- On ne charge que le nécessaire : colonnes et lignes restreintes à la source, filtrage et tri dans le SQL plutôt que via `.Filter` / `.Sort`.
- Les options font la différence : `dbReadOnly` allège la lecture, `dbAppendOnly` rend l'ouverture instantanée pour l'ajout, `dbSeeChanges` est souvent indispensable en ODBC.
- `RecordCount` n'est exact qu'après un `MoveLast` coûteux : tester `.EOF` pour l'existence, `COUNT(*)` ou `DCount` pour un décompte.
- Dans les boucles, mettre en cache les références de champs et éviter les déplacements inutiles.
- Les requêtes sauvegardées, au plan préoptimisé, sont préférables au SQL recompilé à chaque appel.

---


⏭️ [18.3. Execute vs RunSQL — quand et pourquoi](/18-optimisation-performance/03-execute-vs-runsql.md)
