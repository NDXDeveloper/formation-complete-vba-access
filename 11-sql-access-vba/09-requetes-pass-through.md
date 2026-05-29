🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11.9. Requêtes Pass-Through vers SQL Server

La section précédente s'est achevée sur une capacité particulière du dialecte : envoyer du SQL **directement au serveur**, sans que Jet l'interprète. C'est exactement ce que réalise une **requête Pass-Through** (« requête directe »). Au lieu d'être traduite et traitée par le moteur ACE, sa requête est transmise telle quelle au serveur — SQL Server, par exemple —, qui l'exécute avec son propre dialecte et ne renvoie que le résultat.

Ce mécanisme est l'allié des architectures Access + serveur : il décharge le moteur ACE, exploite toute la puissance du SGBD distant, et donne accès aux fonctionnalités que Jet ne connaît pas. Il s'inscrit dans la continuité du chapitre 10.9 (connexion ADO aux bases externes), avec lequel il partage le terrain du client-serveur.

---

## Qu'est-ce qu'une requête Pass-Through ?

Une requête Pass-Through est un type de requête Access spécial qui **transmet son SQL directement à une base serveur, sans intervention de Jet/ACE**. Le SQL est écrit dans le dialecte du **serveur** (le T-SQL pour SQL Server), le serveur l'exécute, et seul le résultat revient à Access.

La différence avec une requête normale sur **table liée** est fondamentale. Lorsqu'on interroge une table liée ODBC par une requête Access classique, **Jet est impliqué** : il analyse la requête, en traite parfois une partie localement, traduit ce qu'il peut vers le serveur, et peut rapatrier plus de données que nécessaire pour les traiter côté client. Une Pass-Through, elle, **court-circuite entièrement Jet** : c'est le serveur, et lui seul, qui fait le travail.

---

## Pourquoi utiliser le Pass-Through ?

Quatre motivations principales justifient ce mécanisme.

La **performance** d'abord. Le serveur effectue l'intégralité du travail — filtrage, jointures, agrégations — et ne renvoie que le résultat. Pour une table volumineuse ou une opération complexe, c'est radicalement plus efficace que de faire transiter les données par des tables liées pour les traiter localement.

L'**accès aux fonctionnalités serveur** ensuite. Tout ce que Jet ne sait pas faire (section 11.8) — fonctions de fenêtrage, CTE, `MERGE`, `FULL OUTER JOIN`, fonctions T-SQL spécifiques — devient utilisable, puisque c'est le serveur qui interprète le SQL. Le Pass-Through est la voie pour tirer parti de la pleine puissance du SGBD.

L'**exécution de procédures stockées** enfin, dans le prolongement de ce que permet ADO (section 10.9), ainsi que les **opérations de masse côté serveur** (un gros `UPDATE` ou `DELETE` exécuté entièrement par le serveur).

---

## Une exigence : le dialecte du serveur

Point essentiel et conséquence directe des sections 11.7 et 11.8 : le SQL d'une Pass-Through est du **T-SQL** (pour SQL Server), **pas du Jet SQL**. Toutes les spécificités Jet doivent donc être abandonnées au profit de la syntaxe du serveur. Les dates ne s'encadrent plus de dièses mais d'**apostrophes** (`'2024-12-31'`), les jokers `*`/`?` deviennent `%`/`_`, `IIf` cède la place à `CASE`, et les fonctions VBA n'ont évidemment plus cours.

Écrire une Pass-Through suppose donc de **maîtriser le dialecte cible**. L'annexe F, dédiée à la correspondance Jet SQL ↔ T-SQL, est ici une référence précieuse.

---

## Créer une requête Pass-Through par code

Par code, on crée une Pass-Through via une **`QueryDef` DAO** dont on configure deux propriétés clés : `Connect` (la connexion ODBC au serveur) et `ReturnsRecords` (la requête renvoie-t-elle des lignes ?).

### Lecture (SELECT)

Pour une requête qui renvoie des enregistrements, on positionne `ReturnsRecords = True` et l'on ouvre un recordset sur le résultat :

```vba
Dim qdf As DAO.QueryDef
Dim rs As DAO.Recordset

Set qdf = CurrentDb.CreateQueryDef("")     ' QueryDef temporaire
qdf.Connect = "ODBC;Driver={ODBC Driver 17 for SQL Server};" & _
              "Server=NOM_SERVEUR;Database=MaBase;Trusted_Connection=yes;"
qdf.ReturnsRecords = True
qdf.SQL = "SELECT NomClient, Ville FROM dbo.Clients WHERE Region = 'Nord';"   ' T-SQL

Set rs = qdf.OpenRecordset(dbOpenSnapshot)   ' résultat en lecture seule
Do While Not rs.EOF
    Debug.Print rs!NomClient, rs!Ville
    rs.MoveNext
Loop
rs.Close: Set rs = Nothing
Set qdf = Nothing
```

### Action ou procédure sans résultat

Pour une requête action côté serveur ou une procédure qui ne renvoie pas de lignes, on met `ReturnsRecords = False` et l'on appelle `Execute` :

```vba
qdf.ReturnsRecords = False
qdf.SQL = "UPDATE dbo.Clients SET Actif = 1 WHERE Region = 'Nord';"
qdf.Execute dbFailOnError
```

> ⚠️ **Piège de `ReturnsRecords`** : un réglage incohérent provoque une erreur. Un `SELECT` avec `ReturnsRecords = False`, ou une requête action avec `True`, échoue. La règle : **`True` pour ce qui renvoie des lignes, `False` sinon**.

### Exécuter une procédure stockée

Une procédure stockée s'invoque avec la syntaxe T-SQL `EXEC`, en réglant `ReturnsRecords` selon qu'elle renvoie ou non des lignes :

```vba
qdf.ReturnsRecords = True
qdf.SQL = "EXEC dbo.usp_ClientsParRegion 'Nord';"
Set rs = qdf.OpenRecordset(dbOpenSnapshot)
```

---

## La connexion ODBC

La propriété `Connect` commence toujours par le préfixe **`ODBC;`** suivi d'une chaîne de connexion ODBC vers le serveur. Comme au chapitre 10.9, on peut employer un **DSN** (`ODBC;DSN=MonDSN;…`) ou une chaîne **sans DSN** détaillant le pilote, le serveur et la base. Le nom exact du pilote (`ODBC Driver 17/18 for SQL Server`) dépend de la version installée, à vérifier sur le poste.

Les mêmes prérequis s'appliquent : le **pilote ODBC doit être installé** dans la bonne architecture (chapitre 21.7), et l'on privilégie l'authentification Windows (`Trusted_Connection=yes`) pour ne pas exposer d'identifiants. Si la chaîne est incomplète, Access invite l'utilisateur à fournir les informations manquantes.

---

## Limites et précautions

Le Pass-Through présente plusieurs contraintes à intégrer.

Ses **résultats ne sont pas modifiables**. Un recordset issu d'une Pass-Through est en **lecture seule** : on ne peut pas y modifier les enregistrements, ni lier directement un formulaire modifiable dessus. Pour l'édition, on recourt aux tables liées ou à une Pass-Through d'action.

Le **dialecte serveur est obligatoire** : aucune syntaxe ni fonction Jet n'est admise.

Le **paramétrage est malcommode** : une requête Pass-Through n'accepte pas les paramètres Access (clause `PARAMETERS`). Pour intégrer des valeurs variables, on doit soit **construire le SQL dynamiquement** — avec toute la rigueur de sécurité et de formatage des sections 11.5 et 11.7, en visant cette fois le dialecte du serveur —, soit passer par une **procédure stockée paramétrée**. Dès qu'un vrai paramétrage (notamment des paramètres de sortie) est requis, **ADO (section 10.9) est souvent plus propre** que le Pass-Through.

Enfin, les **erreurs proviennent du serveur** : ce sont des messages SQL Server qu'il faut intercepter et interpréter (gestion d'erreurs du chapitre 13, et collection `Errors` d'ADO en section 10.8 si l'on passe par ADO).

---

## Pass-Through, tables liées ou ADO ?

Trois approches permettent d'atteindre un serveur depuis Access, et elles **coexistent** souvent. Les **tables liées** offrent un accès transparent et **modifiable**, idéal pour lier l'interface, mais Jet traite la requête et peut rapatrier trop de données. Le **Pass-Through** envoie du SQL serveur brut, **rapide** et exploitant les fonctionnalités du SGBD, mais en **lecture seule**. ADO (section 10.9), enfin, est l'approche **programmatique** par excellence, avec une gestion riche des paramètres et des valeurs de sortie.

Un schéma courant combine ces forces : **tables liées** pour la saisie et l'interface, **Pass-Through** pour les états et agrégations lourdes qui seraient lentes via les tables liées. Cette architecture hybride est développée au chapitre 23.4.

---

## En résumé

Une requête **Pass-Through** transmet son SQL **directement au serveur**, sans intervention de Jet, en écrivant dans le **dialecte du serveur** (T-SQL). On l'emploie pour la **performance** (le serveur fait le travail et ne renvoie que le résultat), pour exploiter les **fonctionnalités serveur** absentes de Jet (section 11.8), et pour les **procédures stockées**. Par code, on la crée via une `QueryDef` DAO en réglant `Connect` (`ODBC;…`) et `ReturnsRecords` (`True` pour un `SELECT`, `False` pour une action) — un réglage incohérent provoquant une erreur. Ses limites : résultats **en lecture seule**, dialecte serveur imposé, **paramétrage malcommode** (construire le SQL avec les précautions de 11.5/11.7, ou préférer ADO pour les paramètres). Le Pass-Through cohabite avec les **tables liées** (transparentes, modifiables) et **ADO** (programmatique), un schéma hybride classique réservant le Pass-Through aux traitements lourds (chapitre 23.4).

Nous avons parcouru l'essentiel de l'exécution et de la construction du SQL. Il reste une famille de fonctions très pratiques, à mi-chemin entre le SQL et le code, qui permettent d'obtenir une valeur agrégée ou une recherche ponctuelle sans écrire de requête complète : les fonctions de domaine. C'est l'objet de la section suivante.


⏭️ [11.10. DLookup, DSum, DCount et autres fonctions de domaine](/11-sql-access-vba/10-fonctions-domaine.md)
