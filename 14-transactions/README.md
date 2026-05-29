🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 14. Transactions et intégrité des données

Une application de gestion ne manipule presque jamais une donnée isolée. Enregistrer une commande, c'est écrire un en-tête **et** ses lignes ; valider un virement, c'est débiter un compte **et** en créditer un autre ; clôturer un exercice, c'est mettre à jour des dizaines d'enregistrements liés entre eux. Dans tous ces cas, les écritures forment un ensemble indissociable : soit elles aboutissent toutes, soit aucune ne doit subsister. C'est précisément ce que les **transactions** permettent de garantir.

Ce chapitre traite de la façon d'assurer, par le code VBA, que les données d'une base Access restent **cohérentes** lorsque plusieurs écritures s'enchaînent, qu'une erreur survient en cours de traitement, ou que plusieurs utilisateurs travaillent en même temps. L'enjeu n'est pas théorique : une opération interrompue à mi-chemin laisse la base dans un état bancal — une commande sans lignes, un stock décrémenté sans vente associée, un solde faux — souvent difficile à détecter et coûteux à corriger après coup.

## Pourquoi ce chapitre est important

Sans contrôle transactionnel, chaque instruction d'écriture est validée immédiatement et indépendamment des autres. Tant que le traitement se déroule normalement, tout va bien. Le problème surgit dès qu'un imprévu interrompt la séquence : une erreur d'exécution, une contrainte violée sur le dernier enregistrement, une coupure réseau sur une base partagée, ou simplement un utilisateur qui ferme l'application au mauvais moment. Les écritures déjà effectuées subsistent, celles qui restaient à faire sont perdues, et la base se retrouve dans un état qu'aucune règle métier n'autorise.

Ce type d'incohérence est particulièrement insidieux car il ne déclenche aucune alerte : les données paraissent valides au niveau de chaque table, mais elles sont fausses au niveau de l'ensemble. Maîtriser les transactions et les règles d'intégrité, c'est se donner les moyens d'écrire des traitements **fiables par construction**, plutôt que d'espérer que rien ne tournera jamais mal.

## Le problème à résoudre — un exemple

Considérons l'enregistrement d'une commande composée d'un en-tête (table `Commandes`) et de plusieurs lignes (table `LignesCommande`). Le traitement écrit d'abord l'en-tête, récupère son identifiant, puis insère chaque ligne. Si une erreur survient à la troisième ligne — un produit inexistant, une valeur hors limites, une contrainte d'intégrité — les deux premières lignes et l'en-tête sont déjà enregistrés. On obtient alors une commande partielle, incohérente, qui faussera tous les états et calculs ultérieurs.

La réponse consiste à **regrouper l'ensemble des écritures dans une transaction** : tant que tout n'est pas confirmé, rien n'est réellement validé dans la base. Si une erreur interrompt le traitement, on annule l'intégralité du bloc et la base revient exactement à son état initial. C'est ce mécanisme, sa mise en œuvre concrète et ses limites dans Access, que ce chapitre détaille.

## Ce que vous allez apprendre

À l'issue de ce chapitre, vous saurez :

- comprendre ce qu'est une transaction et le principe « tout ou rien » qui la fonde ;
- encadrer un bloc d'écritures par une transaction en DAO comme en ADO ;
- annuler proprement un traitement en cas d'erreur et restaurer l'état antérieur de la base ;
- gérer le cas des transactions imbriquées et en connaître les pièges ;
- détecter et traiter les conflits de mise à jour entre utilisateurs ;
- mettre en place et contrôler des règles d'intégrité référentielle par code ;
- situer les limites du moteur ACE en matière de niveaux d'isolation, notamment face à un serveur lié via ODBC.

## Prérequis

Ce chapitre suppose acquis les notions abordées précédemment, en particulier :

- la manipulation des enregistrements en DAO et en ADO (chapitres 9 et 10) ;
- l'exécution de requêtes action — `INSERT`, `UPDATE`, `DELETE` (chapitre 11) ;
- la gestion structurée des erreurs avec `On Error GoTo` et l'objet `Err` (chapitre 13).

La gestion des erreurs est ici un prérequis essentiel : une transaction sans traitement d'erreur associé n'apporte aucune sécurité, puisque c'est précisément en réaction à une erreur que l'on déclenche une annulation (rollback).

## Transaction et intégrité : deux notions complémentaires

Le titre du chapitre associe deux préoccupations distinctes mais convergentes.

La **transaction** est un mécanisme de contrôle : elle garantit qu'un ensemble d'opérations est traité comme une unité atomique, validée d'un seul tenant ou annulée en totalité. Elle répond à la question « comment enchaîner plusieurs écritures sans risquer un état intermédiaire incohérent ? ».

L'**intégrité des données** désigne l'ensemble des règles qui rendent une donnée valide indépendamment du traitement qui l'a produite : intégrité référentielle entre tables liées, contraintes de domaine, unicité, cohérence métier. Elle répond à la question « comment empêcher qu'une donnée invalide soit enregistrée, quelle que soit la voie d'écriture ? ».

Les deux se renforcent mutuellement. Les transactions protègent la cohérence pendant les traitements ; les règles d'intégrité protègent la cohérence en permanence. Une application robuste s'appuie sur les deux.

## Spécificités d'Access et du moteur ACE

Le moteur de base de données d'Access (ACE, successeur de Jet) prend en charge les transactions, mais avec des caractéristiques propres qu'il faut connaître :

- en **DAO**, les transactions sont gérées au niveau de l'objet `Workspace` (`BeginTrans`, `CommitTrans`, `Rollback`) et s'appliquent globalement à l'espace de travail concerné ;
- en **ADO**, elles sont portées par l'objet `Connection` (`BeginTrans`, `CommitTrans`, `RollbackTrans`) ;
- le moteur autorise un certain niveau de **transactions imbriquées**, dont l'usage demande de la rigueur ;
- contrairement à un SGBD serveur comme SQL Server, ACE n'expose pas de **niveaux d'isolation** configurables ; lorsque l'on travaille avec des **tables liées en ODBC**, le comportement transactionnel et l'isolation sont en grande partie gouvernés par le serveur distant et le pilote, et non par Access.

Ces différences expliquent pourquoi certaines stratégies transactionnelles parfaitement valides côté serveur ne se transposent pas telles quelles dans une application Access purement fichier. Le chapitre les met en perspective, en particulier dans sa dernière section.

## Structure du chapitre

- **14.1.** [Concept de transaction — atomicité et rollback](01-concept-transaction.md) — les fondements : qu'est-ce qu'une transaction, le principe « tout ou rien », validation et annulation.
- **14.2.** [Transactions DAO (BeginTrans, CommitTrans, Rollback)](02-transactions-dao.md) — la mise en œuvre via l'objet `Workspace`.
- **14.3.** [Transactions ADO](03-transactions-ado.md) — la gestion transactionnelle au niveau de la `Connection`.
- **14.4.** [Transactions imbriquées](04-transactions-imbriquees.md) — combiner plusieurs niveaux de transaction et éviter les pièges associés.
- **14.5.** [Gestion des conflits de mise à jour](05-conflits-mise-a-jour.md) — détecter et résoudre les écritures concurrentes.
- **14.6.** [Règles d'intégrité référentielle par code](06-integrite-referentielle.md) — définir et contrôler l'intégrité entre tables par programmation.
- **14.7.** [Niveaux d'isolation (ODBC / serveur lié) — limites du moteur ACE natif](07-niveaux-isolation.md) — ce qu'Access permet, ce qu'il délègue au serveur, et où s'arrêtent ses garanties.

## Positionnement dans la formation

Ce chapitre prolonge naturellement le [chapitre 13 sur la gestion des erreurs](/13-gestion-erreurs/README.md) : transaction et traitement d'erreur fonctionnent de pair, l'annulation d'une transaction étant déclenchée par l'interception d'une erreur.

Il prépare également le [chapitre 15 consacré au multi-utilisateurs et au verrouillage](/15-multi-utilisateurs/README.md), où la question de l'accès concurrent aux données — étroitement liée aux transactions et aux conflits de mise à jour abordés ici — est traitée en profondeur.

⏭️ [14.1. Concept de transaction — atomicité et rollback](/14-transactions/01-concept-transaction.md)
