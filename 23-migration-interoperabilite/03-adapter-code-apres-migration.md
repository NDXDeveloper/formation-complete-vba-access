🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 23.3. Adapter le code VBA après migration (T-SQL vs Jet SQL)

À l'issue de la migration des objets et des données (sections 23.2), le schéma et les enregistrements vivent désormais sur SQL Server, mais tout le code VBA — formulaires, états, modules — est resté tel quel. La tentation est de croire le travail terminé : on lance l'application, et… elle fonctionne presque. C'est précisément ce « presque » qui rend cette étape délicate. Une partie du code marche sans rien changer, une autre marche mais devient lente, et une dernière partie échoue avec des messages d'erreur déroutants. Cette section explique pourquoi, et comment adapter le code en conséquence.

Il faut distinguer deux niveaux d'adaptation, de nature différente. Le premier concerne le **dialecte SQL** : SQL Server parle le T-SQL, qui diffère du Jet SQL d'Access sur de nombreux points (fonctions, dates, opérateurs). Le second, souvent plus important pour les performances, concerne le **comportement de l'accès aux données** : la manière dont DAO et ADO dialoguent avec un serveur n'est plus la même qu'avec un fichier local.

## Une bonne nouvelle, et un piège : le rôle des tables liées

Dans l'architecture la plus courante après migration (détaillée à la section 23.4), Access conserve son interface et accède aux données du serveur via des **tables liées** par ODBC. Aux yeux de DAO, une table liée ressemble beaucoup à une table locale : on l'ouvre, on la parcourt, on la modifie avec le même code. C'est la bonne nouvelle, et c'est ce qui permet à une grande partie de l'application de continuer à tourner sans réécriture.

Mais il faut comprendre comment une requête atteint réellement le serveur, car cela détermine quand les différences de dialecte importent. Il existe trois voies :

- **Via les tables liées** : Access soumet la requête au moteur ACE local, qui l'interprète, en exécute une partie localement et traduit le reste en appels ODBC vers le serveur. Comme c'est ACE qui interprète, **la syntaxe Jet et les fonctions Access (`#dates#`, `*`, `&`, `Nz`, `IIf`…) restent acceptées**. L'inconvénient est qu'ACE ne pousse pas toujours tout le traitement vers le serveur, d'où des transferts de données inutiles et des lenteurs.
- **Via une requête Pass-Through** (section 11.9) : le texte SQL est envoyé tel quel au serveur, qui l'exécute entièrement. Il doit alors être écrit en **T-SQL pur**. C'est la voie la plus performante pour les traitements ensemblistes.
- **Via une procédure stockée** appelée depuis VBA, écrite et exécutée côté serveur.

Conséquence pratique : **on n'est pas obligé de réécrire immédiatement tout son SQL en T-SQL**. Les requêtes qui transitent par ACE sur des tables liées tolèrent le Jet SQL. C'est seulement le code envoyé directement au serveur — Pass-Through, procédures stockées, T-SQL écrit à la main — qui exige le respect strict du dialecte T-SQL. Reste que s'appuyer sur ACE a un coût en performance ; l'adaptation idéale consiste donc à réécrire les traitements lourds en T-SQL côté serveur, ce qui impose de connaître les différences de dialecte.

## Les différences de dialecte SQL (Jet SQL → T-SQL)

Voici les écarts qui causent le plus d'erreurs dès qu'on écrit du SQL destiné directement au serveur.

### Délimiteurs de dates

Le Jet SQL encadre les dates par des dièses, le T-SQL par des apostrophes, et il est vivement recommandé d'employer un format ISO non ambigu pour éviter toute interprétation dépendant de la langue du serveur (voir section 11.7).

```sql
-- Jet SQL (Access)
WHERE DateCommande >= #2024-01-15#

-- T-SQL (SQL Server)
WHERE DateCommande >= '20240115'
```

### Concaténation de chaînes

Il faut ici être précis pour ne pas se tromper. En VBA, l'opérateur `&` qui sert à *assembler* la chaîne SQL reste inchangé : c'est du code VBA, pas du SQL. Le problème survient uniquement lorsque `&` figure **dans le texte SQL que le serveur va analyser** (par exemple pour concaténer deux colonnes). En T-SQL, `&` est l'opérateur binaire ET ; la concaténation de chaînes s'écrit avec `+` (ou la fonction `CONCAT`).

```sql
-- Jet SQL
SELECT Nom & ' ' & Prenom AS NomComplet

-- T-SQL
SELECT Nom + ' ' + Prenom AS NomComplet
-- ou, plus sûr face aux NULL :
SELECT CONCAT(Nom, ' ', Prenom) AS NomComplet
```

Attention : en T-SQL, `'texte' + NULL` donne `NULL`, alors que `CONCAT` ignore les NULL. C'est un piège fréquent lors de la conversion.

### Caractères génériques de LIKE

Le Jet SQL (en mode par défaut) utilise `*` et `?` ; le T-SQL utilise `%` et `_`.

```sql
-- Jet SQL
WHERE Nom LIKE 'Dup*'      -- commence par Dup
WHERE Code LIKE 'A?C'      -- A, un caractère, C

-- T-SQL
WHERE Nom LIKE 'Dup%'
WHERE Code LIKE 'A_C'
```

### Valeurs booléennes

Access stocke Vrai à **-1**, SQL Server à **1** (type `bit`). Toute comparaison à `-1` ou à `True` dans du T-SQL doit être revue.

```sql
-- Jet SQL
WHERE Actif = -1      -- ou = True

-- T-SQL
WHERE Actif = 1
```

### Les fonctions

C'est la source la plus abondante de différences. De nombreuses fonctions Jet/VBA n'existent pas en T-SQL, ou portent un autre nom, ou prennent leurs arguments dans un autre ordre.

| Fonction Access (Jet) | Équivalent T-SQL | Remarque |
|---|---|---|
| `IIf(c, a, b)` | `CASE WHEN c THEN a ELSE b END` | `IIF` existe aussi en SQL Server récent, mais `CASE` est universel |
| `Nz(x, 0)` | `ISNULL(x, 0)` ou `COALESCE(x, 0)` | `Nz` est propre à Access |
| `Mid(s, i, n)` | `SUBSTRING(s, i, n)` | |
| `InStr(s, sous)` | `CHARINDEX(sous, s)` | **Ordre des arguments inversé** |
| `Len(s)` | `LEN(s)` | `LEN` ignore les espaces finaux |
| `UCase` / `LCase` | `UPPER` / `LOWER` | |
| `Left` / `Right` | `LEFT` / `RIGHT` | Identiques |
| `Trim(s)` | `TRIM(s)` | `TRIM` en versions récentes, sinon `LTRIM(RTRIM(s))` |
| `Date()` | `CAST(GETDATE() AS date)` | |
| `Now()` | `GETDATE()` ou `SYSDATETIME()` | |
| `DateAdd("d", n, d)` | `DATEADD(day, n, d)` | Intervalle **non** entre guillemets |
| `DateDiff("d", d1, d2)` | `DATEDIFF(day, d1, d2)` | Idem pour l'intervalle |
| `Format(x, …)` | `CONVERT(...)` / `FORMAT(...)` | `FORMAT` (récent) est lent ; préférer `CONVERT`/`CAST` |
| `First` / `Last` | *(pas d'équivalent direct)* | Réécrire avec `MIN`/`MAX` ou `TOP` + `ORDER BY` |

La fonction `InStr` mérite une vigilance particulière : non seulement elle devient `CHARINDEX`, mais l'ordre des deux arguments est inversé, ce qui produit un bug silencieux si on transpose mécaniquement.

### Autres différences

Quelques points supplémentaires reviennent régulièrement. Le mot-clé `DISTINCTROW`, propre à Jet, n'existe pas en T-SQL (utiliser `DISTINCT` ou `GROUP BY`). La clause `TOP n` existe dans les deux dialectes, mais le T-SQL préfère la forme parenthésée `TOP (n)`. Enfin, Access ajoute automatiquement des **parenthèses autour des jointures multiples**, syntaxe que le T-SQL n'exige pas ; ces parenthèses doivent être retirées lorsqu'on porte une requête à jointures multiples vers le serveur.

## Les changements de comportement de l'accès aux données

Même avec un SQL correct, la façon dont le code accède aux données doit évoluer. C'est souvent ici que se jouent les vrais problèmes, notamment de performance.

### Plus de Recordset de type Table ni de Seek

Sur une table liée ODBC, le type de Recordset `dbOpenTable` n'est pas disponible : il est réservé aux tables locales Jet. Par conséquent, la méthode `Seek` (vue à la section 9.9), qui repose sur un index de table locale, ne fonctionne plus non plus. Le code qui s'en servait doit être réécrit, idéalement sous forme de requête filtrée côté serveur.

```vba
' Avant : recherche par Seek sur table locale indexée
rs.Index = "PrimaryKey"
rs.Seek "=", lngID

' Après : requête filtrée (le serveur fait le travail)
Set rs = db.OpenRecordset( _
    "SELECT * FROM Clients WHERE ClientID = " & lngID, _
    dbOpenSnapshot)
```

### L'erreur 3622 et l'option dbSeeChanges

C'est l'erreur post-migration la plus fréquente. Lorsqu'on ouvre un Recordset de type Dynaset sur une table SQL Server possédant une colonne `IDENTITY` (l'équivalent du NuméroAuto), DAO exige l'option `dbSeeChanges` ; sans elle, l'ouverture échoue avec l'erreur 3622.

```vba
' Avant (base locale)
Set rs = db.OpenRecordset("Clients", dbOpenDynaset)

' Après (table liée SQL Server avec colonne IDENTITY)
Set rs = db.OpenRecordset("Clients", dbOpenDynaset, dbSeeChanges)
```

### Récupérer l'identifiant auto-généré

En Jet, après un `AddNew`/`Update`, on relisait simplement le champ NuméroAuto. Avec une colonne `IDENTITY` côté serveur, la récupération via `@@IDENTITY` à travers ODBC peut être trompeuse (elle peut renvoyer un identifiant créé par un déclencheur, et non celui de votre insertion). Deux approches fiables existent.

En DAO, après `Update` (avec `dbSeeChanges`), on se repositionne sur le dernier enregistrement modifié pour relire son identifiant :

```vba
rs.AddNew
rs!Nom = "Durand"
rs.Update
rs.Bookmark = rs.LastModified
lngNouvelID = rs!ClientID
```

Côté serveur, la fonction sûre est `SCOPE_IDENTITY()`, que l'on appelle typiquement via une requête Pass-Through :

```sql
INSERT INTO Clients (Nom) VALUES ('Durand');
SELECT SCOPE_IDENTITY() AS NouvelID;
```

### Ne ramener que le strict nécessaire

C'est le piège de performance majeur. Du code qui ouvrait `SELECT * FROM GrosseTable` puis filtrait ou cherchait en VBA était acceptable en local ; il devient catastrophique sur un serveur, car il transfère toute la table à travers le réseau avant de jeter l'essentiel. La règle change radicalement : **le filtrage et le tri doivent être confiés au serveur**, via une clause `WHERE` et `ORDER BY`, et l'on évite d'ouvrir des tables entières. Le même principe s'applique aux formulaires : un formulaire dont le `RecordSource` est une table complète chargera l'ensemble depuis le serveur ; il faut le restreindre par un filtre ou une requête paramétrée (voir sections 18.2 et 18.5).

### Privilégier les traitements ensemblistes côté serveur

Le corollaire du point précédent concerne les modifications en masse. Boucler sur un Recordset pour modifier chaque enregistrement un à un multiplie les allers-retours réseau. Une opération ensembliste unique, exécutée idéalement en Pass-Through, fait le même travail en une seule passe côté serveur.

```vba
' À éviter : modification enregistrement par enregistrement
Do Until rs.EOF
    rs.Edit
    rs!Statut = "Traité"
    rs.Update
    rs.MoveNext
Loop

' Préférer : une seule opération ensembliste
db.Execute "UPDATE Commandes SET Statut = 'Traité' " & _
           "WHERE DateCommande < '20240101'", dbFailOnError
```

Pour la logique métier lourde, on peut aller plus loin et la déporter dans des **procédures stockées** appelées depuis VBA, ce qui concentre le traitement sur le serveur et réduit le code côté client.

### Transactions et verrouillage

Le verrouillage est désormais géré par SQL Server, avec une finesse bien supérieure à celle d'ACE. Par ailleurs, les transactions DAO (`BeginTrans`/`CommitTrans`) sur des tables liées ne correspondent pas toujours à de véritables transactions serveur ; pour un contrôle transactionnel rigoureux côté serveur, on s'appuie sur des requêtes Pass-Through encadrant `BEGIN TRAN`/`COMMIT`, ou sur les transactions ADO (chapitre 14).

## Une démarche d'adaptation méthodique

En pratique, l'adaptation gagne à être systématique plutôt qu'au coup par coup. Il est utile de **parcourir le code à la recherche des motifs à risque** : les dièses de dates, les `&` dans des chaînes SQL destinées au serveur, les `*` et `?` dans des `LIKE`, les comparaisons à `-1`, les fonctions `Nz`, `IIf`, `InStr`, `Format`, les ouvertures de Recordset en `dbOpenTable`, les appels à `Seek`, les `SELECT *` sur de grosses tables, et les boucles de modification enregistrement par enregistrement. Les **fonctions de domaine** (`DLookup`, `DSum`, `DCount`…, section 11.10) continuent de fonctionner sur les tables liées, mais au prix de transferts coûteux : sur des volumes importants, mieux vaut les remplacer par des requêtes ciblées.

Chaque correction doit être **testée** avec des données réalistes, et l'application **profilée** (section 18.1) pour vérifier que la migration a bien amélioré — et non dégradé — les performances. Une migration techniquement réussie qui rendrait l'application plus lente serait un échec ; or c'est exactement ce qui arrive lorsqu'on se contente de relier les tables sans revoir la façon dont le code les sollicite.

## En résumé

Adapter le code après migration suppose de traiter deux fronts. Sur le plan du dialecte, le T-SQL impose ses délimiteurs de dates, son opérateur de concaténation, ses caractères génériques et ses fonctions, dès lors qu'on écrit du SQL destiné directement au serveur. Sur le plan du comportement, l'accès aux données doit être repensé pour exploiter le modèle client-serveur : recordsets filtrés côté serveur, abandon de `Seek` et de `dbOpenTable`, option `dbSeeChanges`, récupération correcte des identifiants, et traitements ensemblistes plutôt que boucles. Ce travail, qu'aucun outil n'automatise, est le véritable cœur d'une migration. Une fois le code adapté, reste à organiser durablement la cohabitation entre le front-end Access et le back-end serveur : c'est l'objet de la section suivante, consacrée à l'architecture hybride.

⏭️ [23.4. Architecture hybride Access + SQL Server (tables liées ODBC)](/23-migration-interoperabilite/04-architecture-hybride-access-sql-server.md)
