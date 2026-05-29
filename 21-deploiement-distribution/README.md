🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 21. Déploiement et distribution

Une application Access n'a de valeur qu'à partir du moment où elle tourne ailleurs que sur le poste de son développeur. Or c'est précisément là que de nombreux projets achoppent : l'application fonctionne parfaitement dans l'environnement de développement, mais se comporte de façon imprévisible — voire refuse de démarrer — une fois copiée sur le poste d'un utilisateur, installée sur un serveur de fichiers ou ouverte avec une version différente d'Access. Le déploiement est ce que l'on appelle parfois le « dernier kilomètre » : techniquement modeste en apparence, il concentre pourtant une part disproportionnée des incidents en production.

Ce chapitre traite de tout ce qui se passe entre le moment où le développement est terminé et celui où l'application est utilisée de façon fiable, par les bonnes personnes, sur des postes que vous ne maîtrisez pas toujours. Il s'agit d'un sujet à part entière, distinct du développement proprement dit, et qui mobilise autant de rigueur.

## Pourquoi ce chapitre est essentiel

Le déploiement est souvent sous-estimé parce qu'il n'apparaît pas dans le cahier des charges fonctionnel. L'utilisateur demande des formulaires, des états et des traitements ; il ne demande jamais « une procédure de reliaison automatique » ou « une gestion des chemins UNC ». Pourtant, ces aspects invisibles conditionnent directement la perception de qualité de l'application.

Une application Access mal déployée génère un coût récurrent considérable : interventions manuelles à chaque mise à jour, fichiers de données corrompus parce que mal partagés, plantages dus à une incompatibilité 32/64 bits, déclarations d'API qui ne fonctionnent plus après un changement de version d'Office. À l'inverse, une stratégie de déploiement maîtrisée permet de livrer des mises à jour sans déranger les utilisateurs, de protéger le code, et de réduire drastiquement le support.

Maîtriser le déploiement, c'est aussi passer du statut de développeur à celui de fournisseur d'une solution logicielle durable.

## Le déploiement d'une application Access : un contexte particulier

Contrairement à d'autres environnements, Access impose un modèle de distribution spécifique qu'il faut comprendre avant d'aborder les techniques.

Le premier principe structurant est la séparation **front-end / back-end**. En production, les tables de données sont isolées dans un fichier dédié (le back-end), stocké à un emplacement partagé, tandis que l'interface — formulaires, états, requêtes et code VBA — est regroupée dans un fichier distinct (le front-end), dont chaque utilisateur possède sa propre copie locale. Cette architecture, déjà évoquée au chapitre 15 sous l'angle du multi-utilisateur, prend ici une dimension nouvelle : elle conditionne entièrement la façon dont on installe l'application, dont on la met à jour, et dont on relie les deux fichiers entre eux.

Le second principe est l'existence de **plusieurs formats et modes d'exécution**. Le front-end peut être distribué sous forme de fichier `.accdb` modifiable, de fichier `.accde` compilé et verrouillé, et il peut être exécuté soit avec une installation complète d'Access, soit avec l'**Access Runtime**, une version gratuite et allégée destinée aux postes ne disposant pas de licence Office complète. Le choix entre ces options dépend du niveau de protection souhaité, du contrôle que l'on veut conserver sur le code, et du parc de licences disponible chez le client.

Enfin, le déploiement Access se heurte à des contraintes d'environnement que le développeur ne peut pas ignorer : diversité des **versions d'Access** présentes sur les postes, coexistence d'installations **32 bits et 64 bits**, fiabilité des **chemins réseau**, et politiques de sécurité des entreprises. Autant de variables qui n'existent pas, ou peu, pendant la phase de développement, et qui se révèlent uniquement au moment de la distribution.

## Ce que couvre ce chapitre

Le chapitre suit une progression logique, depuis la préparation de l'application jusqu'à son installation automatisée sur les postes cibles.

Les premières sections s'intéressent à la **préparation du livrable** : le packaging d'une application sous forme de fichier compilé ACCDE et la mise à disposition via le runtime, puis l'usage de l'**Access Runtime** lui-même pour déployer sans licence complète. Vous y verrez comment transformer un projet de développement en un produit fini, protégé et prêt à être distribué.

Vient ensuite la question du **cycle de vie en production**, c'est-à-dire la capacité à faire évoluer l'application après sa première installation. Cela passe par les stratégies de mise à jour automatique du front-end, qui évitent de devoir réinstaller manuellement chaque poste, et par la **reliaison automatique des tables back-end**, indispensable lorsque l'emplacement des données change ou lorsqu'un nouveau poste est déployé.

Une partie est consacrée aux **contraintes d'infrastructure** : le déploiement en réseau local et la gestion correcte des chemins UNC, sujet à la fois simple et source d'innombrables dysfonctionnements lorsqu'il est négligé.

Enfin, le chapitre traite des **enjeux de compatibilité**, parmi les plus délicats du déploiement Access : la compatibilité entre les différentes versions du logiciel (2016, 2019, 2021 et Microsoft 365) et la distinction fondamentale entre architectures **32 bits et 64 bits**, qui impacte directement les déclarations d'API Windows abordées au chapitre 22. Il se conclut sur l'**installation automatisée**, où l'on industrialise l'ensemble du processus au moyen de scripts et de paramètres de ligne de commande.

## Objectifs d'apprentissage

À l'issue de ce chapitre, vous serez en mesure de :

- Préparer un livrable professionnel en choisissant le format de distribution adapté (`.accdb`, `.accde`) et en protégeant votre code VBA.
- Déployer une application sur des postes dépourvus de licence Access complète grâce au runtime.
- Mettre en place une stratégie de mise à jour du front-end qui ne nécessite aucune intervention manuelle poste par poste.
- Concevoir un mécanisme de reliaison automatique des tables back-end robuste face aux changements d'emplacement.
- Gérer correctement les chemins réseau et anticiper les problèmes liés aux chemins UNC.
- Diagnostiquer et prévenir les incompatibilités entre versions d'Access et entre architectures 32 et 64 bits.
- Automatiser l'installation de l'application au moyen de scripts et de paramètres de démarrage.

## Prérequis

Ce chapitre suppose que votre application est fonctionnellement terminée et qu'elle adopte une architecture séparant les données de l'interface. Une bonne compréhension du modèle front-end / back-end et des tables liées (chapitres 12 et 15) est recommandée, de même qu'une familiarité avec les options de démarrage de l'application (chapitre 17). Les notions de sécurité abordées au chapitre 20, en particulier la compilation en ACCDE, sont directement réutilisées ici.

## Environnement et versions

Les techniques présentées s'appliquent aux versions modernes d'Access reposant sur le moteur ACE et le format `.accdb` (Access 2007 et ultérieures). Sauf mention contraire, les exemples sont compatibles avec Access 2016, 2019, 2021 et Microsoft 365. Les particularités propres à une version donnée, ou à une architecture 32 ou 64 bits, sont signalées explicitement dans les sections concernées.

⏭️ [21.1. Packaging d'une application Access (ACCDE, runtime)](/21-deploiement-distribution/01-packaging-accde-runtime.md)
