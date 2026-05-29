🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15. Multi-utilisateurs et verrouillage

Une application Access est rarement utilisée par une seule personne. Dès qu'une base est partagée sur un réseau et ouverte simultanément par plusieurs collaborateurs, une nouvelle catégorie de problèmes apparaît : deux utilisateurs qui modifient le même enregistrement, des écritures qui se chevauchent, des enregistrements temporairement inaccessibles, des numéros censés être uniques attribués deux fois. Ce chapitre traite de la conception et du comportement d'une application Access en environnement **multi-utilisateurs**, et du mécanisme central qui arbitre les accès concurrents : le **verrouillage**.

Il prolonge directement le [chapitre 14](/14-transactions/README.md), qui s'achevait sur cette limite : le moteur ACE est un moteur fichier, et sa gestion de la concurrence repose sur le verrouillage plutôt que sur des niveaux d'isolation configurables. C'est précisément ce terrain — le partage, les verrous, les sessions et leurs limites — que nous explorons ici.

## Pourquoi ce chapitre est important

Une application qui fonctionne parfaitement pour un utilisateur unique peut se dégrader de façon subtile dès qu'un second s'y connecte. Les symptômes sont caractéristiques et redoutables : modifications silencieusement écrasées, enregistrements bloqués sans explication claire, doublons sur des identifiants supposés uniques, ralentissements brutaux, voire risques de corruption en cas de coupure réseau mal gérée.

Ces problèmes ont un point commun : ils sont **intermittents**. Ils ne surviennent que lorsque deux actions coïncident dans le temps, donc rarement en phase de test, souvent en production sous charge réelle. Ce sont les bugs les plus difficiles à reproduire et à diagnostiquer. La seule parade efficace consiste à **concevoir l'application pour le multi-utilisateur dès le départ**, plutôt qu'à corriger après coup. C'est l'objet de ce chapitre.

## Le problème à résoudre — un exemple

Imaginons une application de saisie de commandes partagée par plusieurs commerciaux. Deux d'entre eux créent une commande au même instant.

Pour numéroter la nouvelle commande, chaque poste calcule le plus grand numéro existant et lui ajoute 1 — soit 1001 dans les deux cas, puisque ni l'un ni l'autre n'a encore enregistré. Les deux commandes sont sauvegardées avec le **même numéro 1001** : l'unicité est rompue. Dans un autre scénario, les deux commerciaux ouvrent la même fiche client pour la corriger ; le second à enregistrer **écrase silencieusement** la modification du premier. Dans un troisième, l'un détient un enregistrement en cours d'édition, et le second se retrouve **bloqué** sans comprendre pourquoi.

Aucun de ces incidents n'est imputable à une erreur de logique « classique » : ils naissent uniquement de la **concurrence**. Les maîtriser suppose une architecture adaptée, une stratégie de verrouillage explicite, et des techniques spécifiques pour les opérations sensibles comme la numérotation.

## Ce que vous allez apprendre

À l'issue de ce chapitre, vous saurez :

- structurer une application en **front-end / back-end** et comprendre pourquoi cette séparation est la fondation du multi-utilisateur ;
- choisir le **mode de partage** approprié pour ouvrir la base ;
- sélectionner et mettre en œuvre une **stratégie de verrouillage** (optimiste ou pessimiste) adaptée au besoin ;
- **détecter et traiter les conflits de verrouillage** lorsqu'un enregistrement est déjà occupé par un autre utilisateur ;
- générer des **numéros séquentiels fiables** en environnement concurrent, sans risque de doublon ;
- contrôler le verrouillage au niveau des formulaires via la propriété `RecordLocks` ;
- utiliser les **TempVars** comme variables globales de session ;
- mettre en place un **suivi des sessions utilisateurs** (fichier de verrouillage et table de sessions) ;
- **lier des tables distantes par code** pour connecter dynamiquement un front-end à son back-end ;
- identifier les **limites du moteur ACE** en environnement concurrent, et reconnaître le moment où une autre architecture s'impose.

## Prérequis

Ce chapitre s'appuie sur des notions développées précédemment :

- les **transactions et les conflits de mise à jour** ([chapitre 14](/14-transactions/README.md)), dont il constitue le prolongement naturel ;
- la manipulation des **recordsets** en DAO et en ADO (chapitres 9 et 10) ;
- les **requêtes action** et les fonctions de domaine en SQL (chapitre 11) ;
- la **gestion des erreurs** ([chapitre 13](/13-gestion-erreurs/README.md)), indispensable pour traiter proprement les situations de verrou ;
- les bases des **formulaires** et de leurs événements (chapitres 6 et 8).

## Vue d'ensemble : partage, verrouillage, session

Le chapitre s'organise autour de trois grandes préoccupations complémentaires.

La première est **l'architecture et le partage** : comment une application est déployée et ouverte par plusieurs utilisateurs à la fois. C'est le rôle de la séparation front-end / back-end, des modes de partage, et de la liaison des tables distantes.

La deuxième est **le verrouillage** : comment l'accès concurrent à une même donnée est arbitré pour éviter les pertes et les corruptions. On y trouve les stratégies optimiste et pessimiste, la détection des conflits de verrou, et le contrôle du verrouillage au niveau des formulaires.

La troisième regroupe **la session et les limites** : le suivi des utilisateurs connectés, les variables globales de session, la numérotation fiable, et enfin les limites intrinsèques du moteur ACE en situation concurrente. Ces sujets déterminent jusqu'où une application Access partagée peut aller, et où commencent les contraintes.

## Spécificités d'Access et du moteur ACE

Il faut garder à l'esprit une réalité structurante : **le moteur ACE est un moteur fichier, et non un système client-serveur**. Le partage repose sur un fichier de base de données accessible simultanément par plusieurs postes, accompagné d'un **fichier de verrouillage** qui recense les sessions ouvertes. Les accès concurrents sont arbitrés par un verrouillage opérant au niveau de la page ou de l'enregistrement.

Ce modèle fonctionne, mais il a des conséquences directes : la concurrence supportée est réelle mais **limitée**, la fiabilité dépend fortement de la qualité du réseau, et le moteur n'offre pas les garanties d'un serveur (pas de niveaux d'isolation configurables, comme vu en [section 14.7](/14-transactions/07-niveaux-isolation.md)). Comprendre ces caractéristiques est la clé pour construire une application partagée robuste — et pour reconnaître le moment où il devient pertinent de migrer vers un serveur (chapitre 23).

## Structure du chapitre

- **15.1.** [Architecture multi-utilisateurs — front-end / back-end](01-architecture-front-end-back-end.md) — la séparation fondatrice entre l'interface et les données.
- **15.2.** [Modes de partage de la base de données](02-modes-partage.md) — ouvertures exclusive et partagée, et leurs implications.
- **15.3.** [Stratégies de verrouillage (Optimistic vs Pessimistic)](03-strategies-verrouillage.md) — les deux approches du verrouillage et leurs compromis.
- **15.4.** [Détection et gestion des conflits de verrouillage](04-detection-conflits-verrouillage.md) — traiter le cas d'un enregistrement déjà verrouillé par un autre.
- **15.5.** [Génération de numéros séquentiels fiables en multi-utilisateur](05-numeros-sequentiels-multi-utilisateurs.md) — éviter les doublons (DMax+1, table de compteurs).
- **15.6.** [Propriété RecordLocks des formulaires](06-recordlocks-formulaires.md) — contrôler le verrouillage au niveau du formulaire.
- **15.7.** [TempVars — variables globales de session](07-tempvars.md) — maintenir un état partagé tout au long de la session.
- **15.8.** [Gestion des sessions utilisateurs (LDB et table de sessions)](08-gestion-sessions-utilisateurs.md) — savoir qui est connecté.
- **15.9.** [Connexion à une base distante via table liée par code](09-tables-liees-par-code.md) — lier et relier les tables back-end par programmation.
- **15.10.** [Limites du moteur ACE en environnement concurrent](10-limites-moteur-ace.md) — jusqu'où Access partagé peut aller.

## Positionnement dans la formation

Ce chapitre fait suite au [chapitre 14](/14-transactions/README.md) : les conflits de mise à jour qui y étaient abordés sous l'angle de l'intégrité y trouvent ici leur cadre complet, celui de l'accès concurrent. Transactions et verrouillage sont les deux faces d'une même exigence de fiabilité en environnement partagé.

Il prépare également le [chapitre 18 sur l'optimisation et la performance](/18-optimisation-performance/README.md) — notamment l'optimisation réseau des bases partagées et les stratégies face à la limite de taille — et le [chapitre 23 sur la migration et l'interopérabilité](/23-migration-interoperabilite/README.md), qui apporte la réponse de fond lorsque les limites du moteur ACE en concurrence sont atteintes : le passage à un serveur de base de données.

⏭️ [15.1. Architecture multi-utilisateurs — front-end / back-end](/15-multi-utilisateurs/01-architecture-front-end-back-end.md)
