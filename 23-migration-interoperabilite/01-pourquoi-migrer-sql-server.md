🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 23.1. Pourquoi et quand migrer Access vers SQL Server

Cette section ne décrit pas *comment* migrer — c'est l'objet des sections suivantes — mais *pourquoi* et *à quel moment* le faire. C'est une étape de réflexion, pas de technique. Une migration mal motivée, lancée par principe ou par effet de mode, coûte du temps et de l'argent sans rien résoudre ; à l'inverse, une migration repoussée trop longtemps expose l'organisation à des pertes de données et à des interruptions de service. Savoir décider est donc une compétence en soi, et elle commence par comprendre la nature exacte du problème que SQL Server résout.

## Comprendre la racine du problème : le modèle « fichier partagé »

Pour décider en connaissance de cause, il faut d'abord saisir une différence d'architecture fondamentale entre Access et SQL Server. Cette différence explique à elle seule la plupart des limites que l'on rencontre.

Le moteur de base de données d'Access, appelé ACE (Access Connectivity Engine, successeur de Jet), fonctionne sur un modèle de **fichier partagé**. La base de données est un simple fichier `.accdb` posé sur un disque, éventuellement un partage réseau. Quand plusieurs postes y accèdent, il n'y a pas de programme serveur qui arbitre les demandes : c'est le moteur ACE installé sur *chaque poste client* qui ouvre le fichier et y travaille directement. Lorsqu'une requête s'exécute, le poste client lit les pages d'index et de données dont il a besoin **en les transférant à travers le réseau** jusqu'à sa propre mémoire, puis y effectue le tri, le filtrage et les calculs. Le réseau transporte donc des données brutes, souvent bien plus que ce que la requête finit par afficher.

SQL Server repose sur le modèle inverse, dit **client-serveur**. Un programme serveur tourne en permanence sur une machine dédiée. Le poste client ne lit jamais le fichier de données directement : il envoie une requête au serveur, le serveur l'exécute *sur place* (avec ses propres index, sa mémoire, son processeur) et ne renvoie au client que le résultat — par exemple les 30 lignes qui correspondent au filtre, et non les 500 000 lignes de la table.

Cette opposition — *traitement côté client sur fichier partagé* contre *traitement côté serveur* — est la clé de lecture de toute la section. Presque chaque limite d'Access en découle, et presque chaque avantage de SQL Server aussi.

## Les limites qui motivent une migration

### La limite des 2 Go

Un fichier `.accdb` (comme un ancien `.mdb`) ne peut pas dépasser **2 Go**. Cette limite inclut les données, mais aussi les index, les objets de l'application et l'espace temporaire de travail non encore récupéré par un compactage. En pratique, on commence à ressentir des difficultés bien avant ce plafond. Une stratégie classique consiste à séparer les données dans un fichier back-end distinct du front-end, mais cela ne fait que repousser l'échéance : le back-end reste soumis aux mêmes 2 Go. Une base SQL Server, même en édition gratuite, dépasse largement ce seuil.

### La concurrence d'accès

On lit souvent qu'Access supporte jusqu'à 255 connexions simultanées. Ce chiffre est théorique et trompeur. En conditions réelles, et particulièrement pour des applications où plusieurs utilisateurs écrivent en même temps, le confort se dégrade bien avant : au-delà d'une poignée à une vingtaine d'utilisateurs actifs, les blocages d'enregistrement, les lenteurs et les conflits de mise à jour se multiplient. Le modèle de fichier partagé n'a tout simplement pas été conçu pour arbitrer finement de nombreux accès concurrents en écriture. SQL Server, lui, est bâti pour cela : il gère le verrouillage, les transactions et la concurrence avec un niveau de sophistication sans commune mesure.

### Les performances en réseau

Sur un réseau local rapide et avec peu de données, Access est très performant. Mais à mesure que les tables grossissent ou que le réseau s'étend (sites distants, VPN, Wi-Fi instable), le coût du modèle « tout transférer côté client » devient prohibitif. Une requête qui ramène des centaines de milliers de pages à travers le réseau pour n'en afficher que quelques-unes sature la bande passante et ralentit toute l'application. Le passage à un traitement côté serveur réduit drastiquement le trafic réseau et redonne de la réactivité.

### La fiabilité et le risque de corruption

Parce que le fichier de données est ouvert et modifié simultanément par plusieurs postes, une coupure réseau, un poste qui se fige ou une déconnexion brutale pendant une écriture peut laisser le fichier dans un état incohérent — c'est la redoutée *corruption de base*. Plus il y a d'utilisateurs et plus le réseau est fragile, plus le risque augmente. SQL Server, qui maîtrise lui-même tous les accès au stockage et tient un journal des transactions, est conçu pour récupérer proprement après un incident, sans perte de données validées.

### La sécurité

Depuis l'abandon de la sécurité au niveau utilisateur du format `.mdb`, le format `.accdb` n'offre plus qu'une protection par mot de passe global, qui chiffre le fichier mais ne distingue pas les utilisateurs ni leurs droits (ce sujet est traité au chapitre 20). Concrètement, quiconque a accès au partage réseau a accès au fichier. Pour des données sensibles, des obligations réglementaires ou un besoin d'audit, c'est insuffisant. SQL Server apporte une authentification robuste (intégrée à Windows ou propre au serveur), des permissions granulaires par table, vue ou colonne, le chiffrement, et une traçabilité fine des accès.

### La sauvegarde et la continuité d'activité

Sauvegarder un fichier Access ouvert par des utilisateurs n'est pas fiable : la copie peut être incohérente. En pratique, il faut souvent attendre que tout le monde soit déconnecté. SQL Server permet des sauvegardes *à chaud*, sans interrompre l'activité, ainsi que la restauration à un instant précis (point-in-time recovery) grâce au journal des transactions. Pour une application critique, cette capacité à garantir la continuité de service est souvent, à elle seule, une raison suffisante de migrer.

### Synthèse comparative

| Critère | Access (moteur ACE) | SQL Server |
|---|---|---|
| Taille maximale | 2 Go par fichier | Plusieurs téraoctets (selon l'édition) |
| Modèle d'exécution | Traitement côté client, fichier partagé | Traitement côté serveur (client-serveur) |
| Concurrence en écriture | Quelques utilisateurs confortablement | Des centaines à des milliers |
| Trafic réseau | Élevé (transfert des données brutes) | Réduit (seul le résultat circule) |
| Risque de corruption | Réel en réseau partagé | Très faible (récupération journalisée) |
| Sécurité | Mot de passe global uniquement | Comptes, rôles, permissions fines, chiffrement |
| Sauvegarde à chaud | Non fiable | Native, sans interruption |
| Coût d'entrée | Inclus avec Access | Édition gratuite disponible |

## Ce que SQL Server apporte concrètement

Au-delà de la simple levée des limites, migrer le stockage vers SQL Server ouvre des possibilités nouvelles. Les **transactions** y sont robustes et fiables, garantissant qu'un ensemble d'opérations s'exécute entièrement ou pas du tout, même en cas de panne. Le moteur permet de déporter une partie de la logique côté serveur, via des **procédures stockées**, des **vues** et des **fonctions**, ce qui peut considérablement améliorer les performances et centraliser les règles métier. L'écosystème d'**outils professionnels** (SQL Server Management Studio pour l'administration, l'Agent SQL pour planifier des tâches, les outils de supervision et de diagnostic) facilite l'exploitation au quotidien. Enfin, disposer des données dans un véritable serveur ouvre naturellement la voie à l'**intégration** avec d'autres systèmes, au reporting et à la business intelligence (Power BI, SSRS, etc.).

## La question du coût : les éditions de SQL Server

L'objection la plus fréquente à la migration est le coût supposé de SQL Server. Cette crainte est largement infondée pour la majorité des applications Access, car il existe une **édition gratuite, SQL Server Express**, dont les capacités suffisent à remplacer un back-end Access dans la grande majorité des cas.

| Édition | Coût | Limite de taille par base | Cas d'usage typique |
|---|---|---|---|
| LocalDB | Gratuit | 10 Go | Développement et tests sur un poste |
| Express | Gratuit | 10 Go | Petites et moyennes applications, ressources processeur et mémoire bridées |
| Standard | Payant (licence) | Très élevée | Applications d'entreprise, montée en charge |
| Enterprise | Payant (licence) | Très élevée | Hautes exigences (haute dispo, gros volumes) |
| Azure SQL Database | Abonnement cloud | Selon le palier | Hébergement géré, sans serveur à administrer |

La limite de 10 Go de l'édition Express représente déjà cinq fois la limite d'un fichier Access, tout en apportant le modèle client-serveur, les transactions, la sécurité et les sauvegardes à chaud. Pour beaucoup d'applications qui souffrent surtout de problèmes de concurrence, de fiabilité ou de sécurité plutôt que de volume pur, Express est une solution gratuite et largement suffisante. Le passage à une édition payante ne s'impose que lorsque le volume, le nombre d'utilisateurs ou les besoins de haute disponibilité dépassent ce que l'édition gratuite couvre.

## Quand migrer : les signaux d'alerte

La décision de migrer mûrit généralement lorsque plusieurs des signaux suivants se manifestent : le fichier back-end approche de la limite des 2 Go ou la franchit de plus en plus souvent malgré les compactages ; le nombre d'utilisateurs simultanés augmente et s'accompagne de lenteurs, de blocages ou de conflits de mise à jour récurrents ; des corruptions de base surviennent, surtout en réseau ; l'application manipule des données sensibles ou est soumise à des obligations de sécurité, de confidentialité ou d'audit que le format `.accdb` ne peut satisfaire ; le besoin d'une sauvegarde fiable et d'une continuité de service devient critique parce que l'application est désormais indispensable au métier ; ou encore une croissance prévisible (en volume, en utilisateurs, en exigences) rend la limite proche et justifie d'anticiper plutôt que de subir.

Aucun de ces signaux n'impose à lui seul une migration, mais leur accumulation est un indicateur fiable que l'application a dépassé le cadre naturel d'Access.

## Quand ne pas migrer

Migrer n'est pas toujours la bonne réponse, et le réflexe « SQL Server réglera tout » est trompeur. Avant d'envisager une migration, il faut s'assurer que le problème rencontré n'est pas en réalité un **problème d'optimisation**. Une application lente à cause d'index manquants, de requêtes mal écrites, de recordsets ouverts sans nécessité ou de formulaires chargeant des tables entières restera lente après migration — voire deviendra plus lente si le code n'est pas adapté. Le chapitre 18 traite précisément de ces optimisations, qui suffisent à résoudre une grande partie des inconforts perçus, sans changer d'architecture.

De même, une petite base utilisée par une ou deux personnes, avec un volume modeste et sans exigence de sécurité particulière, ne tire aucun bénéfice d'une migration : le coût en temps et en complexité dépasserait largement le gain. La règle est simple : on migre pour résoudre un problème réel et identifié, jamais par principe.

## Migrer les données, pas forcément toute l'application

Un point essentiel pour bien cadrer la décision : migrer vers SQL Server ne signifie pas, dans la plupart des cas, abandonner Access. La trajectoire la plus courante consiste à ne déplacer que les **données** (le back-end) vers SQL Server, tout en conservant Access comme **front-end** — c'est-à-dire les formulaires, les états et tout le code VBA déjà écrit. Access se connecte alors au serveur par des tables liées, et l'utilisateur ne voit pratiquement pas la différence dans son interface quotidienne.

Cette approche, l'**architecture hybride** détaillée à la section 23.4, préserve l'investissement considérable que représente le développement de l'application, tout en réglant les problèmes de fond liés au moteur de stockage. Il ne faut donc pas confondre « migrer vers SQL Server » et « réécrire l'application » : ce sont deux décisions différentes, et la première n'entraîne pas la seconde.

## Ce que la migration implique réellement

Décider de migrer, c'est s'engager dans un projet, pas déclencher une opération instantanée. Les sections suivantes le montreront en détail, mais il est utile de l'avoir en tête dès maintenant. La migration des données et du schéma s'appuie sur un outil dédié, SSMA, qui automatise une grande partie du travail mais comporte ses propres limites (section 23.2). Surtout, le code VBA existant devra presque toujours être **adapté** : le dialecte SQL de SQL Server (T-SQL) diffère de celui d'Access (Jet SQL) sur de nombreux points — fonctions, gestion des dates, syntaxe — et certaines pratiques de code qui fonctionnaient bien en local peuvent dégrader fortement les performances une fois branchées sur un serveur via ODBC (section 23.3). Une migration réussie tient donc autant à la rigueur du travail humain, en amont (nettoyage des données, préparation) et en aval (adaptation du code, tests), qu'à l'outil employé.

Une fois cette décision prise et bien comprise, la section suivante présente l'outil par lequel commence concrètement toute migration vers SQL Server.

⏭️ [23.2. SQL Server Migration Assistant (SSMA) — utilisation et limites](/23-migration-interoperabilite/02-ssma-utilisation.md)
