🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11.8. Dialecte SQL Access (Jet SQL) — spécificités et limites

Jusqu'ici, nous avons manié le SQL d'Access en supposant sa syntaxe connue. Il est temps d'en dresser le portrait en tant que **dialecte** à part entière. Le SQL d'Access — le **Jet SQL** (ou ACE SQL selon le moteur) — descend du SQL standard, mais s'en écarte par un ensemble de particularités, d'extensions propres et de limitations. Les ignorer expose à deux types de déconvenues : des requêtes qui échouent inexplicablement, et de mauvaises surprises lors d'une éventuelle migration vers un autre SGBD.

Connaître ces spécificités, c'est écrire un SQL correct sur Access — et savoir ce qui devra être réécrit le jour où l'on portera l'application vers SQL Server (chapitre 23). Pour une consultation exhaustive, l'**annexe E** détaille la syntaxe Jet SQL, et l'**annexe F** établit sa correspondance avec le T-SQL.

---

## Jet SQL : un dialecte à part

Avant tout, une notion structurante : Access connaît **deux modes de syntaxe SQL**. Le mode **ANSI-89** (historique, par défaut) est le mode « Access traditionnel » ; le mode **ANSI-92** est plus conforme au standard. Ce choix de mode n'est pas anecdotique — il modifie le comportement de certaines constructions, à commencer par les caractères génériques, comme nous allons le voir immédiatement.

---

## Le piège des jokers : `*` et `?` ou `%` et `_`

C'est le piège le plus déroutant du dialecte, et celui qui occasionne le plus de bugs en pratique. Les **caractères génériques** utilisés avec l'opérateur `LIKE` **diffèrent selon le mode** — et donc selon la **manière dont la requête est exécutée**.

| Élément recherché | ANSI-89 (DAO / interface Access) | ANSI-92 (ADO / OLEDB) |
|---|---|---|
| Séquence de caractères | `*` | `%` |
| Caractère unique | `?` | `_` |
| Chiffre unique | `#` | (pas d'équivalent direct) |
| Plage de caractères | `[a-z]`, `[!a-z]` | `[a-z]`, `[!a-z]` |

Le point crucial : les requêtes exécutées via le **fournisseur OLEDB (donc en ADO)** tournent en mode **ANSI-92**, quel que soit le réglage de la base — elles utilisent `%` et `_`. Les requêtes lancées via **DAO ou l'interface d'Access** suivent le mode de la base (par défaut ANSI-89) — elles utilisent `*` et `?`.

```sql
-- La MÊME recherche, selon le mode d'exécution :
SELECT * FROM Clients WHERE Nom LIKE 'Dur*';   -- ANSI-89 (DAO / interface)
SELECT * FROM Clients WHERE Nom LIKE 'Dur%';   -- ANSI-92 (ADO / OLEDB)
```

> ⚠️ **Conséquence concrète** : un `LIKE` écrit avec `*` fonctionne en DAO mais échoue (ou ne renvoie rien) en ADO, et inversement. C'est une cause classique de bugs lorsqu'on déplace du code entre les deux technologies. Vérifiez toujours quel mode s'applique au contexte d'exécution.

---

## Les spécificités propres à Jet

Au-delà des jokers, plusieurs traits distinguent Jet SQL du SQL standard.

**La concaténation se fait avec `&`** (ou `+`), jamais avec l'opérateur standard `||` qui n'existe pas dans Jet. On privilégie `&`, plus robuste face aux valeurs `Null`.

**Les fonctions VBA et Access sont utilisables directement dans le SQL**, grâce au service d'expression. Une requête Jet peut appeler `IIf()`, `Nz()`, `Format()`, `Left()`, `Mid()`, `Date()`, `Now()`, `DateAdd()`, `DateDiff()`, et bien d'autres — chose impossible dans un SGBD classique, qui dispose de ses propres fonctions. On peut même y appeler ses **propres fonctions VBA publiques** (fonctions définies par l'utilisateur), pourvu que la requête s'exécute au sein d'Access. C'est puissant, mais cela **lie la requête à Access** : une telle requête ne tournera pas sur SQL Server.

**`IIf` et `Switch` remplacent `CASE`.** Jet SQL **ne connaît pas** l'expression standard `CASE WHEN … THEN … END`. On exprime la logique conditionnelle avec `IIf(condition, vrai, faux)` ou `Switch(cond1, val1, cond2, val2, …)`.

**Les littéraux ont leur syntaxe propre.** Les dates s'encadrent de dièses `#…#` (section 11.7), spécificité de Jet là où le standard et T-SQL utilisent des apostrophes. Les booléens s'écrivent `True`/`False`, `Yes`/`No`, `On`/`Off`, ou `-1`/`0`.

**`TOP n` (et `TOP n PERCENT`)** limite les résultats, mais avec une particularité signalée à la section 11.2 : Jet **inclut les ex æquo** à la position limite, pouvant renvoyer plus de lignes que demandé. Le mot-clé voisin **`DISTINCTROW`**, propre à Access, élimine les doublons sur la base des lignes entières des tables sous-jacentes — à distinguer de `DISTINCT`.

**`TRANSFORM … PIVOT`** est l'extension Jet pour les **requêtes d'analyse croisée** (*crosstab*), une syntaxe absente du SQL standard.

**Les jointures multiples exigent des parenthèses.** Au-delà de deux tables, Jet impose un **parenthésage explicite** des jointures imbriquées :

```sql
SELECT *
FROM (Clients INNER JOIN Commandes ON Clients.IDClient = Commandes.IDClient)
     INNER JOIN LignesCommande ON Commandes.IDCommande = LignesCommande.IDCommande;
```

Omettre ces parenthèses provoque une erreur — un piège fréquent lorsqu'on écrit le SQL à la main, alors que le concepteur visuel les génère automatiquement.

**La clause `IN 'fichier'`** permet d'interroger une table située dans **une autre base Access**, sans table liée : `SELECT * FROM Clients IN 'C:\Data\autre.accdb';` (à ne pas confondre avec l'opérateur `IN (liste)`).

---

## Les limites de Jet SQL

Face au T-SQL et au SQL moderne, Jet accuse plusieurs absences notables.

**Pas de `FULL OUTER JOIN`.** Jet connaît `INNER`, `LEFT` et `RIGHT JOIN`, mais pas la jointure externe complète. On la simule par une `UNION` d'un `LEFT JOIN` et d'un `RIGHT JOIN`.

**Pas d'expressions de table communes (CTE).** La clause `WITH … AS (…)` n'existe pas. On la remplace par des sous-requêtes ou des requêtes sauvegardées.

**Pas de fonctions de fenêtrage.** Aucune fonction analytique de type `OVER`, `PARTITION BY`, `ROW_NUMBER`, `RANK`. Pour des numérotations ou classements, on recourt à des sous-requêtes corrélées ou à des fonctions de domaine (`DCount`…), souvent moins efficaces.

**Pas d'autres opérateurs ensemblistes ni de `MERGE`.** Jet propose `UNION` (et `UNION ALL`), mais ni `INTERSECT`, ni `EXCEPT`/`MINUS`, ni l'instruction `MERGE` de fusion.

**Pas de procédures stockées, fonctions ou déclencheurs au sens du T-SQL.** Les **requêtes sauvegardées** constituent l'équivalent le plus proche d'une « procédure » (et sont d'ailleurs vues comme telles via `adCmdStoredProc` en ADO), mais sans langage procédural. Les **macros de données** (*data macros*) du format `.accdb` jouent un rôle proche des déclencheurs au niveau des tables, sans être du SQL.

**Des limites de moteur**, enfin, qui ne relèvent pas du SQL mais l'encadrent : la **taille de base plafonnée à 2 Go** (chapitre 18.9) et les **contraintes de concurrence** en environnement multi-utilisateur (chapitre 15.10).

---

## Implications pratiques : la portabilité

Cette connaissance du dialecte a une retombée directe : la **portabilité**. Plus une requête s'appuie sur des spécificités Jet — fonctions VBA dans le SQL, `TRANSFORM/PIVOT`, `IIf`, jokers `*`/`?`, clause `IN 'fichier'` —, moins elle se transposera aisément vers un autre SGBD. À l'inverse, un SQL proche du standard migre plus facilement.

Si une migration vers SQL Server est envisageable (chapitre 23), deux attitudes sont possibles : **éviter** dès le départ les constructions purement Jet lorsqu'on vise la portabilité, ou les **assumer** en sachant précisément ce qu'il faudra réécrire (jointures externes complètes, `CASE`, fonctions de fenêtrage, fonctions VBA…). L'annexe F, dédiée à la correspondance Jet SQL ↔ T-SQL, est l'outil de référence pour cet exercice.

---

## En résumé

Le **Jet SQL** est un dialecte dérivé du SQL standard, doté de **deux modes** (ANSI-89 par défaut, ANSI-92 via OLEDB/ADO) dont la différence la plus piégeuse concerne les **jokers** (`*`/`?` contre `%`/`_`) — source de bugs au passage entre DAO et ADO. Ses **spécificités** incluent la concaténation par `&`, l'usage de **fonctions VBA dans les requêtes** (et de fonctions maison), `IIf`/`Switch` à la place de `CASE`, les dates entre `#`, `TOP` avec ex æquo, `TRANSFORM/PIVOT`, le parenthésage obligatoire des jointures multiples et la clause `IN 'fichier'`. Ses **limites** comprennent l'absence de `FULL OUTER JOIN`, de CTE, de fonctions de fenêtrage, de `CASE`, d'`INTERSECT`/`EXCEPT`/`MERGE`, et de véritables procédures stockées ou déclencheurs SQL — sans oublier les limites de moteur (2 Go, concurrence). Toutes ces particularités conditionnent la **portabilité** du code, déterminante en cas de migration (chapitre 23, annexe F).

Parmi les capacités du dialecte figure justement la possibilité d'envoyer du SQL **directement au serveur** lorsqu'on travaille avec SQL Server — un SQL qui n'est alors plus interprété par Jet, mais par le serveur lui-même, avec son propre dialecte. C'est le mécanisme des requêtes Pass-Through, qu'aborde la section suivante.


⏭️ [11.9. Requêtes Pass-Through vers SQL Server](/11-sql-access-vba/09-requetes-pass-through.md)
