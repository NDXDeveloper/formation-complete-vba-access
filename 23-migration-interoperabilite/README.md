🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 23. Migration et interopérabilité

La plupart des applications Access naissent petites. Un utilisateur avancé assemble quelques tables, dessine un formulaire de saisie, ajoute un état d'impression, automatise deux ou trois traitements en VBA — et résout un vrai problème métier en quelques jours. Puis l'application grandit. Le nombre d'enregistrements augmente, des collègues demandent à y accéder, des fonctionnalités s'ajoutent, et ce qui était un outil personnel devient peu à peu une application critique dont dépend tout un service.

C'est précisément à ce moment-là que les limites du moteur ACE commencent à se faire sentir : la barre des 2 Go par fichier qui approche, des ralentissements en accès concurrent, des corruptions de fichier en réseau instable, ou simplement le besoin de partager les données avec d'autres systèmes de l'entreprise. Ce chapitre traite de la question qui se pose alors inévitablement : **comment faire évoluer l'application au-delà de ce qu'Access seul peut offrir ?**

C'est un chapitre à part dans cette formation. Il parle moins de syntaxe VBA que de décisions d'architecture. Les notions abordées ici engagent l'avenir de l'application sur plusieurs années et méritent d'être comprises avant d'écrire la moindre ligne de migration.

## De l'outil personnel à l'application critique

Access est souvent victime de son propre succès. Sa force — permettre de construire vite et sans infrastructure lourde — devient une faiblesse quand l'application franchit certains seuils. Reconnaître ces seuils est la première compétence à acquérir : il s'agit de distinguer un inconfort passager (qui se règle par de l'optimisation, traitée au chapitre 18) d'une limite structurelle (qui appelle un changement d'architecture).

Migrer ou ouvrir une application à d'autres systèmes n'est jamais une fin en soi. C'est une réponse à un besoin concret : fiabiliser la concurrence d'accès, dépasser les limites de volume, sécuriser réellement les données, intégrer l'application au système d'information, ou préparer une transition technologique. Chaque section de ce chapitre part de ce besoin avant de présenter une solution.

## Migration et interopérabilité : deux notions à ne pas confondre

Le titre du chapitre réunit deux démarches voisines mais distinctes, et il est important de les séparer dès le départ.

L'**interopérabilité** consiste à faire travailler Access *avec* d'autres systèmes, sans renoncer à Access. L'application reste en place ; on l'ouvre vers l'extérieur. C'est le cas, par exemple, d'une base Access qui consomme des données d'un serveur PostgreSQL via des tables liées ODBC, tout en conservant ses propres tables locales.

La **migration**, elle, consiste à déplacer tout ou partie de l'application *hors* d'Access. Cette migration peut être partielle — on ne déplace que les données vers un vrai serveur tout en conservant l'interface Access et son code VBA — ou totale — on quitte complètement l'environnement Access au profit d'une autre plateforme.

Cette distinction structure tout le chapitre : certaines sections décrivent comment partir d'Access, d'autres comment l'enrichir en le gardant.

## Trois grandes familles de réponses

Face à une application Access qui atteint ses limites, il existe schématiquement trois trajectoires possibles, et ce chapitre les couvre toutes les trois.

La première consiste à **conserver l'interface Access et son code VBA, en ne migrant que la base de données** vers un moteur plus robuste, typiquement SQL Server. On obtient alors une architecture hybride : un front-end Access (formulaires, états, modules) connecté par tables liées ODBC à un back-end serveur. C'est de loin la trajectoire la plus fréquente, car elle préserve l'investissement considérable que représente le développement VBA tout en réglant les problèmes de volume, de concurrence et de sécurité. C'est aussi celle qui demande le plus de soin dans l'adaptation du code, car le SQL serveur ne se comporte pas exactement comme le SQL d'Access.

La deuxième consiste à **quitter Access pour une autre plateforme**, généralement dans l'écosystème Microsoft 365 : SharePoint Lists, Dataverse, ou Power Apps. Cette voie répond souvent à des contraintes d'accès web, de mobilité, ou de politique d'entreprise (« plus d'application bureautique locale »). Elle est plus radicale, car elle abandonne le modèle VBA et impose de repenser l'application avec d'autres outils. Toute la difficulté est d'en évaluer honnêtement la faisabilité avant de s'y engager.

La troisième relève de l'interopérabilité pure : **faire cohabiter Access avec d'autres SGBD** (PostgreSQL, MySQL, MariaDB), sans migration au sens strict. Access devient alors un client parmi d'autres d'une base partagée, ou un outil de consultation et de reporting branché sur un système existant.

Aucune de ces trajectoires n'est universellement « la bonne ». Le choix dépend du contexte : volumétrie, nombre d'utilisateurs, compétences de l'équipe, budget, stratégie informatique de l'organisation. L'objectif de ce chapitre est de vous donner les clés pour décider en connaissance de cause, puis pour exécuter proprement la trajectoire retenue.

## Ce que vous saurez faire à l'issue de ce chapitre

À la fin du chapitre, vous serez en mesure de reconnaître les signaux qui indiquent qu'une application Access a dépassé son cadre naturel, et d'argumenter une décision de migration auprès d'un décideur. Vous saurez préparer et conduire une migration des données vers SQL Server à l'aide de l'outil dédié, puis adapter le code VBA existant aux spécificités du dialecte T-SQL. Vous saurez concevoir et maintenir une architecture hybride Access + SQL Server via des tables liées ODBC. Vous saurez enfin évaluer la pertinence d'une migration vers SharePoint, Dataverse ou Power Apps, et mettre en place une cohabitation entre Access et d'autres moteurs de bases de données.

## Prérequis

Ce chapitre suppose une bonne maîtrise des chapitres consacrés à l'accès aux données. La manipulation des **tables liées** (section 12.7) et la compréhension des **chaînes de connexion ADO et ODBC** (annexe D) sont particulièrement mobilisées. Une familiarité avec **DAO et ADO** (chapitres 9 et 10) ainsi qu'avec le **SQL dans Access** (chapitre 11), notamment les requêtes Pass-Through (section 11.9), facilitera grandement la lecture. Enfin, la connaissance des **limites du moteur ACE en environnement concurrent** (sections 15.10 et 18.9) éclaire les motivations profondes de toute la démarche de migration.

## Plan du chapitre

Le chapitre suit une progression logique, du « pourquoi » vers le « comment », puis des alternatives à SQL Server.

- **23.1.** [Pourquoi et quand migrer Access vers SQL Server](./01-pourquoi-migrer-sql-server.md) — Les critères de décision : limites atteintes, signaux d'alerte, et ce que la migration apporte réellement.
- **23.2.** [SQL Server Migration Assistant (SSMA) — utilisation et limites](./02-ssma-utilisation.md) — L'outil officiel de migration des données et du schéma, ses possibilités et ses angles morts.
- **23.3.** [Adapter le code VBA après migration (T-SQL vs Jet SQL)](./03-adapter-code-apres-migration.md) — Les différences de dialecte SQL et les ajustements indispensables du code existant.
- **23.4.** [Architecture hybride Access + SQL Server (tables liées ODBC)](./04-architecture-hybride-access-sql-server.md) — Le modèle front-end Access / back-end serveur, le plus répandu en pratique.
- **23.5.** [Migration vers SharePoint Lists et Dataverse](./05-migration-sharepoint-dataverse.md) — Les options de stockage dans l'écosystème Microsoft 365 et leurs contraintes.
- **23.6.** [Migration vers Power Apps — analyse de faisabilité](./06-migration-power-apps.md) — Quand et comment envisager un passage vers une application web/mobile sans VBA.
- **23.7.** [Cohabitation Access avec d'autres SGBD (PostgreSQL, MySQL, MariaDB)](./07-cohabitation-autres-sgbd.md) — L'interopérabilité avec les moteurs open source, hors écosystème Microsoft.

## Un point de vigilance

Une migration n'est pas un clic dans un assistant. C'est un projet, avec ses phases (analyse, préparation, test, bascule, vérification) et ses risques. La tentation est grande de croire qu'un outil automatique réglera tout ; ce chapitre montrera au contraire que la valeur d'une migration réussie tient surtout au travail humain en amont et en aval de l'outil : le nettoyage des données, l'adaptation du code, et la validation rigoureuse du résultat. Abordez ce qui suit avec cette idée en tête.

⏭️ [23.1. Pourquoi et quand migrer Access vers SQL Server](/23-migration-interoperabilite/01-pourquoi-migrer-sql-server.md)
