🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 23.7. Cohabitation Access avec d'autres SGBD (PostgreSQL, MySQL, MariaDB)

Les sections précédentes ont exploré la migration vers l'écosystème Microsoft : SQL Server pour qui reste sur Access (23.4), SharePoint, Dataverse et Power Apps pour qui le quitte (23.5, 23.6). Reste un dernier cas de figure, fréquent dans les environnements hétérogènes : **rester sur Access tout en dialoguant avec des moteurs de bases de données non-Microsoft**, au premier rang desquels PostgreSQL, MySQL et MariaDB. Il ne s'agit pas ici de migration, mais d'**interopérabilité** : Access demeure en place et travaille sur des données qui vivent dans un autre moteur. C'est le sujet, et la conclusion, de ce chapitre.

## Cohabitation, et non migration

La distinction est importante. Dans les scénarios précédents, on déplaçait des données *hors* d'Access. Ici, on fait **coexister** Access avec un SGBD existant. Plusieurs situations y conduisent : l'organisation exploite déjà un PostgreSQL ou un MySQL — par exemple comme base d'une application web ou d'un progiciel — et souhaite utiliser Access comme interface familière de saisie, de consultation ou de reporting ; ou bien ces moteurs, gratuits et open source, ont été retenus comme back-end (par souci de coût, en alternative à SQL Server), Access jouant alors le rôle de front-end de la même manière qu'à la section 23.4.

La forme la plus simple et la moins risquée de cohabitation est la **consultation en lecture seule** : Access se branche sur une base PostgreSQL ou MySQL pour produire des états et des analyses, sans jamais y écrire, pendant que l'application principale reste maîtresse des données. À l'autre extrémité, Access peut servir de front-end complet en lecture-écriture sur ces moteurs, ce qui suppose plus de précautions, détaillées plus bas.

## Le pont commun : ODBC

Comme pour SQL Server, la passerelle universelle est **ODBC**. Le mécanisme est rigoureusement le même qu'à la section 23.4 : on installe un pilote, on crée des **tables liées** (de préférence sans DSN et par code, via la propriété `Connect` d'un `TableDef`), et l'on manipule ensuite ces tables comme des tables locales. Les **requêtes Pass-Through** (section 11.9) permettent d'envoyer du SQL directement au moteur, et **ADO** offre une alternative pour les connexions à des bases externes (section 10.9). Tout ce qui a été dit en 23.4 sur la liaison, la reliaison au démarrage et la centralisation des chaînes de connexion s'applique ici sans changement ; seul le pilote et le dialecte diffèrent.

## Les pilotes ODBC à installer

Chaque moteur dispose de son propre pilote ODBC, à installer sur chaque poste client.

| SGBD | Pilote ODBC | Remarque |
|---|---|---|
| PostgreSQL | psqlODBC (pilote officiel), en variantes ANSI et Unicode | Préférer la variante **Unicode** pour gérer l'UTF-8 |
| MySQL | MySQL Connector/ODBC (Oracle) | Variantes ANSI et Unicode également |
| MariaDB | MariaDB Connector/ODBC | Le connecteur MySQL/ODBC fonctionne souvent aussi avec MariaDB |

Deux précautions techniques sont déterminantes. D'abord, le **bitness du pilote doit correspondre à celui d'Office**, et non à celui du système d'exploitation : un Office 64 bits exige un pilote ODBC 64 bits, un Office 32 bits un pilote 32 bits. C'est une cause d'échec classique (voir section 21.7). Ensuite, pour éviter les problèmes d'**encodage** (caractères accentués déformés), il faut retenir la variante **Unicode** du pilote et veiller à la cohérence des jeux de caractères et collations (UTF-8 de préférence). Le format exact des chaînes de connexion, qui varie d'un pilote à l'autre, est donné à l'annexe D ; seul le nom du pilote et les paramètres changent par rapport à l'exemple SQL Server vu en 23.4.

```vba
' Exemple de chaîne sans DSN pour PostgreSQL (le motif est le même
' que pour SQL Server, seuls le pilote et les paramètres changent)
tdf.Connect = "ODBC;Driver={PostgreSQL Unicode};" & _
              "Server=srv-pg;Port=5432;Database=gestion;" & _
              "Uid=app_user;Pwd=*****;"
```

## Les différences de dialecte SQL

Chaque moteur possède son propre dialecte SQL, distinct **à la fois** du Jet SQL d'Access et du T-SQL de SQL Server. Comme expliqué en 23.3, cela ne pose pas de problème pour les requêtes qui transitent par le moteur ACE sur des tables liées (ACE traduit la syntaxe Jet, au prix des mêmes réserves de performance). En revanche, dès qu'on écrit du SQL envoyé **directement** au moteur — en Pass-Through ou dans une vue côté serveur — il faut respecter strictement le dialecte cible.

| Aspect | Access (Jet) | PostgreSQL | MySQL / MariaDB |
|---|---|---|---|
| Délimiteur d'identifiant | `[crochets]` | `"guillemets"` | `` `accents graves` `` |
| Limiter le nombre de lignes | `TOP n` | `LIMIT n` | `LIMIT n` |
| Concaténation de chaînes | `&` | `||` | `CONCAT(...)` (ou `||` selon le mode) |
| Auto-incrément | NuméroAuto | `SERIAL` / `IDENTITY` (+ séquences) | `AUTO_INCREMENT` |
| Type booléen | -1 / 0 | type `boolean` (`TRUE`/`FALSE`) | `TINYINT(1)` : 1 / 0 |

Un point mérite une vigilance particulière pour un développeur Access, habitué à l'insensibilité à la casse : la **sensibilité à la casse** de ces moteurs. PostgreSQL replie les identifiants non entre guillemets en minuscules, mais traite les identifiants entre guillemets de façon sensible à la casse, et ses comparaisons de chaînes peuvent l'être aussi selon la collation. C'est une source de surprises (« la table existe pourtant ! ») qu'il faut connaître avant de rédiger des requêtes directes.

## Les pièges spécifiques d'interopérabilité

Au-delà du dialecte, l'interopérabilité avec ces moteurs comporte quelques chausse-trapes propres, qui s'ajoutent aux conditions déjà vues en 23.4.

La règle de la **clé primaire ou de l'index unique** pour qu'une table ou une vue liée soit modifiable reste valable à l'identique : sans identifiant unique connu d'Access, les données sont en lecture seule, et l'on retrouve les symptômes `#Deleted` déjà diagnostiqués.

Les **correspondances de types via ODBC** réservent des surprises spécifiques. Les **booléens** sont gérés différemment selon le moteur (vrai type côté PostgreSQL, `TINYINT(1)` côté MySQL) et peuvent perturber l'affichage dans Access. Les **entiers non signés** de MySQL (`UNSIGNED`), absents du modèle de types d'Access, se mappent mal. Mais le piège le plus célèbre est la **« date zéro » de MySQL** (`0000-00-00`) : cette valeur, historiquement tolérée par MySQL, est invalide pour Access et provoque des erreurs. Il faut donc éviter ces dates nulles dans les données et, si nécessaire, configurer le pilote pour qu'il les convertisse en `NULL`.

Pour les **performances**, la discipline est exactement celle de l'architecture hybride : filtrer côté serveur, s'appuyer sur des vues et des requêtes Pass-Through écrites dans le dialecte du moteur, éviter de rapatrier des tables entières, et maintenir une connexion ouverte pour épargner le coût des reconnexions ODBC (sections 23.4 et 18.8).

## Choisir le moteur — ou s'y adapter

Dans la pratique, le choix du moteur est rarement libre : il est dicté par le système existant avec lequel Access doit cohabiter. À titre de repère, PostgreSQL est le plus riche et le plus conforme aux standards SQL, apprécié pour les besoins relationnels exigeants ; MySQL et MariaDB sont des back-ends très répandus dans le monde web, MariaDB étant le fork communautaire de MySQL, largement compatible avec lui.

Il faut toutefois être lucide sur une différence de fond avec SQL Server. Ces moteurs ne bénéficient pas du même statut « de première classe » auprès d'Access : Microsoft optimise son couple Access/SQL Server et fournit un outil de migration dédié (SSMA, section 23.2), ce qui n'existe pas pour PostgreSQL ou MySQL. L'interopérabilité par ODBC fonctionne bien, mais avec davantage d'aspérités (particularités des pilotes, mappings de types, dates zéro, casse, entiers non signés). Surtout, **migrer vers ces moteurs est plus manuel** : il n'y a pas d'équivalent de SSMA, et l'on procède soit à la main, soit avec des outils tiers. La cohabitation par liaison et reporting est donc très robuste ; une migration complète du back-end vers ces moteurs demande, elle, plus de travail que la voie SQL Server.

## En résumé

Faire cohabiter Access avec PostgreSQL, MySQL ou MariaDB repose sur le même pont ODBC que l'architecture hybride SQL Server : installer le bon pilote (de bitness conforme à Office et en variante Unicode), lier les tables, garantir une clé primaire pour la modifiabilité, gérer les particularités de types — booléens, entiers non signés, dates zéro de MySQL — et écrire en Pass-Through dans le dialecte propre au moteur (identifiants, `LIMIT`, concaténation, casse). C'est une voie d'interopérabilité solide, particulièrement pour la consultation et le reporting, à condition d'accepter qu'elle soit moins outillée que la voie Microsoft.

Avec cette section se referme le chapitre 23. On y aura vu un spectre complet de trajectoires : conserver Access en fiabilisant sa couche de données, que ce soit avec SQL Server (23.4) ou avec un moteur open source (23.7) ; ou quitter tout ou partie d'Access pour le cloud Microsoft, via SharePoint, Dataverse et Power Apps (23.5, 23.6). Le fil conducteur de toutes les solutions « avec Access » est l'interopérabilité par ODBC ; celui des solutions « sans Access » est le passage à la Power Platform. Entre les deux, le bon choix n'est jamais dicté par la mode, mais par une analyse honnête du besoin, du volume, du budget et de l'existant — exactement l'état d'esprit avec lequel ce chapitre s'était ouvert.

⏭️ [24. Bonnes pratiques et ressources](/24-bonnes-pratiques-ressources/README.md)
