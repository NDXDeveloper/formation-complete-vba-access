🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11.11. Requêtes Union, sous-requêtes et expressions complexes

Pour clore ce chapitre consacré au SQL, nous abordons les constructions les plus élaborées — celles qui permettent d'exprimer ce que les requêtes simples et les fonctions de domaine ne suffisent plus à formuler. La mention des sous-requêtes à la fin de la section précédente nous y conduit naturellement : combiner plusieurs jeux de résultats, imbriquer une requête dans une autre, ou bâtir des expressions conditionnelles complexes.

Ces outils donnent au SQL Jet une réelle expressivité, dans les limites du dialecte rappelées à la section 11.8. Nous verrons aussi une stratégie typiquement Access pour dompter la complexité : l'empilement de requêtes sauvegardées.

---

## Les requêtes UNION

Une requête **`UNION`** combine les résultats de **plusieurs `SELECT`** en un seul jeu de résultats, empilés les uns sous les autres.

```sql
SELECT NomClient, 'Actif' AS Statut FROM Clients
UNION
SELECT NomClient, 'Archivé' AS Statut FROM ClientsArchive
ORDER BY NomClient;
```

Deux variantes existent : **`UNION`** **élimine les doublons**, tandis que **`UNION ALL`** les **conserve** — et s'avère plus rapide, puisqu'il évite le travail de dédoublonnage. On choisit `UNION ALL` dès qu'on sait qu'il n'y a pas de doublons ou qu'on souhaite les garder.

Quelques **règles** encadrent l'opération. Chaque `SELECT` doit comporter le **même nombre de colonnes**, dans le **même ordre**, avec des **types compatibles**. Les **noms de colonnes proviennent du premier `SELECT`** ; un éventuel `ORDER BY`, placé **à la toute fin**, s'applique à l'ensemble et référence ces noms.

Les usages sont nombreux : **fusionner des données** de plusieurs tables (enregistrements actifs et archivés, par exemple), ajouter une **ligne synthétique** à une liste, ou encore **simuler le `FULL OUTER JOIN`** absent de Jet (section 11.8) par l'union d'un `LEFT JOIN` et d'un `RIGHT JOIN`. L'ajout d'une ligne d'en-tête à une liste déroulante en est une application fréquente :

```sql
SELECT IDClient, NomClient FROM Clients
UNION
SELECT 0, '(Tous les clients)'
ORDER BY NomClient;
```

> 💡 Les requêtes `UNION` ne peuvent **pas** être construites dans la grille du concepteur visuel : elles s'écrivent directement en mode SQL.

---

## Les sous-requêtes

Une **sous-requête** est un `SELECT` **imbriqué** dans une autre requête. Jet les prend en charge (avec les limites évoquées en 11.8), à trois emplacements principaux.

Dans la clause **`WHERE`**, avec `IN`, `NOT IN`, `EXISTS`, `NOT EXISTS`, ou un opérateur de comparaison pour une sous-requête scalaire :

```sql
-- Clients ayant passé une grosse commande (IN)
SELECT * FROM Clients
WHERE IDClient IN (SELECT IDClient FROM Commandes WHERE Montant > 1000);

-- Produits plus chers que la moyenne (sous-requête scalaire)
SELECT * FROM Produits
WHERE Prix > (SELECT AVG(Prix) FROM Produits);
```

Dans la **liste `SELECT`**, comme sous-requête scalaire renvoyant une valeur par ligne (souvent **corrélée**, c'est-à-dire référençant l'enregistrement externe) :

```sql
SELECT NomClient,
       (SELECT COUNT(*) FROM Commandes AS O WHERE O.IDClient = C.IDClient) AS NbCommandes
FROM Clients AS C;
```

Dans la clause **`FROM`**, enfin, comme **table dérivée** (à laquelle Jet impose un alias) :

```sql
SELECT T.Region, T.Total
FROM (SELECT Region, SUM(Montant) AS Total FROM Commandes GROUP BY Region) AS T
WHERE T.Total > 10000;
```

### `IN`, `EXISTS` et le piège de `NOT IN`

Pour un simple test d'existence, **`EXISTS` est souvent plus efficace qu'`IN`** : il s'arrête au premier enregistrement trouvé. Surtout, **`NOT IN` recèle un piège** : si la sous-requête renvoie une valeur `Null`, `NOT IN` peut ne renvoyer **aucune ligne**, de façon contre-intuitive. La parade consiste à utiliser **`NOT EXISTS`**, qui se comporte correctement face aux `Null`.

### Sous-requêtes plutôt que fonctions de domaine

C'est l'application directe de la mise en garde de la section 11.10. Là où l'on serait tenté d'appeler un `DCount` ou un `DSum` par ligne, une **sous-requête corrélée** — ou mieux, une **jointure** — fait le même travail bien plus efficacement, en une seule passe du moteur. Pour les besoins répétés, c'est la voie à privilégier.

---

## Les expressions complexes

Le SQL Jet permet des expressions riches dans les clauses `SELECT` et `WHERE`.

La **logique conditionnelle** s'exprime avec `IIf`, `Switch` ou `Choose`, puisque Jet ne connaît pas le `CASE` standard (section 11.8) :

```sql
SELECT NomClient,
       IIf(Solde < 0, 'Débiteur', IIf(Solde = 0, 'À jour', 'Créditeur')) AS Situation
FROM Clients;
```

Les **colonnes calculées** reçoivent un **alias** (`AS`) pour être relues de façon fiable, comme établi en 11.2. Les **agrégats groupés** s'appuient sur `GROUP BY`, filtrés au besoin par `HAVING` :

```sql
SELECT Region, SUM(Montant) AS Total
FROM Commandes
GROUP BY Region
HAVING SUM(Montant) > 10000;
```

Une distinction essentielle oppose ici `WHERE` et `HAVING` : **`WHERE` filtre les lignes *avant* le regroupement**, tandis que **`HAVING` filtre les groupes *après* l'agrégation**. Un critère portant sur une valeur individuelle va dans `WHERE` ; un critère portant sur un agrégat (comme `SUM(...)`) va dans `HAVING`.

---

## Gérer la complexité : empiler les requêtes sauvegardées

Au fil de ces constructions, le SQL peut devenir long et difficile à lire. Or Jet ne dispose pas des **expressions de table communes (CTE)** qui, ailleurs, permettent de nommer des étapes intermédiaires (section 11.8). La réponse idiomatique d'Access est différente, et tout aussi efficace : **sauvegarder les requêtes intermédiaires et les empiler**.

Une requête Access peut en effet utiliser **une autre requête sauvegardée comme source** (`FROM MaRequeteSauvegardee`). On décompose ainsi une logique complexe en **étapes successives**, chacune nommée, lisible et réutilisable — et précompilée, donc performante (section 11.2). Plutôt qu'une unique requête monolithique truffée de sous-requêtes, on construit une chaîne de requêtes claires qui s'appuient les unes sur les autres. Cette technique, qui s'appuie sur la manipulation des requêtes sauvegardées (chapitre 12.1), est la meilleure alliée de la maintenabilité face à la complexité.

---

## Performance et modifiabilité (rappels)

Deux rappels pour refermer. Côté **performance**, les sous-requêtes corrélées et les colonnes calculées par ligne peuvent être lentes ; une **jointure** est souvent préférable, et mérite d'être testée comme alternative (chapitre 18). Côté **modifiabilité**, les requêtes `UNION`, comme celles comportant des agrégats ou des sous-requêtes, ne sont généralement **pas modifiables** (sections 11.2, 9.3 et 10.7) — on les destine à la lecture. Enfin, lorsqu'on construit ces requêtes dynamiquement en VBA, toutes les précautions des sections 11.6 et 11.7 s'appliquent, et le `Debug.Print` du SQL final devient d'autant plus précieux que la requête est complexe.

---

## En résumé

Les **requêtes `UNION`** combinent plusieurs `SELECT` (avec dédoublonnage, ou `UNION ALL` sans), sous réserve d'un même nombre de colonnes aux types compatibles, les noms venant du premier `SELECT`. Les **sous-requêtes** s'imbriquent dans le `WHERE` (`IN`, `EXISTS`, comparaison), la liste `SELECT` (scalaire corrélée) ou le `FROM` (table dérivée) ; on privilégie `EXISTS`/`NOT EXISTS` (gare au piège de `NOT IN` avec les `Null`), et l'on préfère sous-requêtes et jointures aux fonctions de domaine répétées. Les **expressions complexes** reposent sur `IIf`/`Switch`/`Choose` (faute de `CASE`), des colonnes calculées aliasées, et des agrégats `GROUP BY`/`HAVING` — en distinguant bien `WHERE` (avant groupement) de `HAVING` (après). Pour dompter la complexité, l'**empilement de requêtes sauvegardées** remplace avantageusement les CTE absentes de Jet. Performances (préférer les jointures) et modifiabilité (résultats en lecture seule) restent les points de vigilance.

---

## Conclusion du chapitre 11

Ce chapitre a fait le tour du SQL tel qu'on l'emploie depuis VBA dans Access. Nous avons d'abord appris à l'**exécuter** correctement (`RunSQL` vs `Execute`, section 11.1), puis à **lire** (11.2), **modifier les données** (11.3) et **la structure** (11.4). Les deux fils rouges annoncés ont été traités en profondeur : la **sécurité**, via le SQL paramétré et la construction dynamique maîtrisée (11.5 et 11.6), et la **localisation**, via le formatage rigoureux des littéraux (11.7). Nous avons ensuite dressé le portrait du **dialecte Jet SQL**, de ses spécificités et de ses limites (11.8), avant d'explorer les **requêtes Pass-Through** vers le serveur (11.9), les **fonctions de domaine** (11.10) et, enfin, les **constructions avancées** (11.11).

La leçon transversale du chapitre tient en deux principes complémentaires. D'une part, **paramétrer** dès que des valeurs externes entrent dans une requête — c'est la réponse unique à la sécurité comme au formatage. D'autre part, **laisser le moteur travailler** : filtrer, trier et agréger en SQL ensembliste plutôt qu'en code, et préférer les jointures aux traitements ligne à ligne. Maîtriser le SQL, c'est savoir déléguer au moteur ce qu'il fait mieux que VBA.

Plusieurs sections ont renvoyé à la manipulation par code des **requêtes sauvegardées** et des **structures de tables** — paramétrage en DAO, requêtes sauvegardées empilées, création de tables. C'est précisément l'objet du chapitre suivant, qui passe du SQL *écrit* aux objets DAO qui *représentent et manipulent* les requêtes et les tables de la base.


⏭️ [12. QueryDefs et TableDefs](/12-querydefs-tabledefs/README.md)
