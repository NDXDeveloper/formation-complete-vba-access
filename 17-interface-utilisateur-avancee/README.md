🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 17. Interface utilisateur avancée

Les chapitres précédents ont permis de construire des applications Access qui *fonctionnent* : des formulaires liés aux données (chapitre 6), des états mis en page (chapitre 7), des contrôles réagissant aux événements de l'utilisateur (chapitre 8), le tout structuré selon des principes solides (chapitre 16). Mais une application qui fonctionne n'est pas encore une application qu'on a *plaisir* à utiliser, ni une application qui inspire confiance. Entre les deux se trouve une couche de finition que ce chapitre a pour objet de combler : tout ce qui transforme une base de données fonctionnelle mais brute en un logiciel professionnel, cohérent et agréable.

## Pourquoi soigner l'interface d'une application Access

Access traîne une réputation d'outil « technique », et l'interface par défaut n'aide pas à la corriger : volet de navigation visible sur le côté, ruban Office complet exposant des centaines de commandes sans rapport avec l'application, fenêtres qui s'ouvrent les unes sur les autres sans logique apparente. Pour un développeur, c'est l'environnement de travail ; pour l'utilisateur final, c'est une source de confusion et un signal que l'outil n'est pas « fini ». La qualité perçue d'une application se joue largement sur ces détails d'interface, bien avant que l'utilisateur n'évalue la pertinence de la logique métier qui se cache derrière.

Cette finition prend une importance toute particulière au moment du déploiement. Une application distribuée via **Access Runtime** (chapitre 21) ne donne accès ni à l'éditeur, ni au volet de navigation, ni au ruban de conception : l'utilisateur ne voit *que* l'interface que vous avez explicitement construite. Tout ce qui n'a pas été soigné — un ruban resté générique, une fenêtre de chargement absente, des raccourcis clavier hasardeux — devient alors visible et figé. Maîtriser l'interface avancée, c'est s'assurer que l'application livrée ressemble à un produit conçu pour ses utilisateurs, et non à une base de données qu'on leur aurait simplement transmise.

## Au programme de ce chapitre

Les techniques abordées se répartissent en quelques grandes intentions. Plusieurs concernent la **personnalisation de l'enveloppe applicative** — ce cadre Office que l'utilisateur perçoit avant même d'ouvrir un formulaire : remplacer le ruban Office par un ruban dédié piloté par votre code, redéfinir les menus contextuels du clic droit, et verrouiller l'environnement au démarrage pour que l'application se présente comme un logiciel autonome plutôt que comme une session Access.

D'autres organisent la **structure et la navigation** : offrir un point d'entrée clair d'où l'utilisateur atteint chaque fonction de l'application sans jamais se perdre dans une accumulation de fenêtres.

Une troisième famille soigne le **retour visuel et l'expérience au lancement** : l'écran d'accueil qui s'affiche pendant l'initialisation, et la barre de progression qui rassure l'utilisateur durant un traitement long en lui montrant que l'application travaille au lieu de paraître figée.

Vient ensuite l'**ergonomie de la saisie**, souvent négligée et pourtant décisive pour les utilisateurs intensifs : maîtriser l'ordre de tabulation, le déplacement du focus et la navigation entièrement au clavier, afin qu'une saisie répétitive devienne fluide et rapide.

Enfin, l'**apparence et la cohérence visuelle** : mettre en forme les données dynamiquement par code selon leur contenu, et appliquer à l'ensemble de l'application un thème homogène, pour que tous les écrans partagent la même identité plutôt que d'accumuler des choix esthétiques disparates.

## Prérequis

Ce chapitre se situe dans le bloc avancé de la formation et suppose une bonne aisance avec les formulaires et les contrôles. La maîtrise du **modèle objet Form** et des modes d'affichage (chapitre 6), du **cycle de vie et des événements** des formulaires et des contrôles (chapitre 8) ainsi que de l'objet **DoCmd** (chapitre 5) est nécessaire pour suivre les exemples sans difficulté. Plusieurs sections gagnent par ailleurs à être abordées après le chapitre 16 : la personnalisation du ruban et certaines briques d'interface réutilisables s'appuient naturellement sur les **modules de classe** et les bonnes pratiques de structuration vues en programmation orientée objet.

## Plan du chapitre

- 17.1. [Ruban personnalisé (Ribbon) avec XML et callbacks VBA](01-ruban-personnalise-xml.md) — remplacer le ruban Office par un ruban sur mesure, défini en XML et piloté par des fonctions de rappel (*callbacks*) VBA.
- 17.2. [Menus contextuels personnalisés (CommandBars)](02-menus-contextuels.md) — redéfinir les menus du clic droit pour n'exposer que les actions pertinentes au contexte.
- 17.3. [Formulaire de navigation (Navigation Form) contrôlé par VBA](03-formulaire-navigation.md) — bâtir un point d'entrée central et le piloter par code pour structurer l'accès aux fonctions de l'application.
- 17.4. [Splash screen et écran de chargement](04-splash-screen.md) — afficher un écran d'accueil pendant l'initialisation et masquer la phase de démarrage à l'utilisateur.
- 17.5. [Barre de progression personnalisée dans un formulaire](05-barre-progression.md) — donner un retour visuel sur l'avancement d'un traitement long pour éviter l'impression d'application figée.
- 17.6. [Gestion du focus et navigation clavier avancée](06-focus-navigation-clavier.md) — contrôler l'ordre de tabulation et le déplacement du focus pour une saisie rapide et entièrement au clavier.
- 17.7. [Options de démarrage et verrouillage de l'interface](07-options-demarrage.md) — configurer le lancement de l'application et masquer l'environnement Access pour la présenter comme un logiciel autonome.
- 17.8. [Mise en forme conditionnelle par code](08-mise-en-forme-conditionnelle.md) — adapter dynamiquement l'apparence des contrôles en fonction des données affichées.
- 17.9. [Thèmes visuels et personnalisation cohérente de l'application](09-themes-visuels.md) — appliquer une identité visuelle homogène à l'ensemble des écrans de l'application.

## À l'issue de ce chapitre

Vous serez en mesure de transformer une application Access fonctionnelle en un produit fini : un logiciel doté de son propre ruban et de ses propres menus, présentant un écran d'accueil au lancement, guidant l'utilisateur par une navigation claire, signalant ses traitements longs par un retour visuel, optimisé pour une saisie au clavier, verrouillé pour masquer son origine Access, et habillé d'une apparence cohérente d'un écran à l'autre. Ces compétences préparent directement le **déploiement** (chapitre 21), où cette interface soignée devient la seule chose que verront vos utilisateurs.

## Un mot sur le dosage

L'interface est un moyen, jamais une fin. Un ruban superbe et une animation de chargement léchée ne rachètent pas une logique métier défaillante ou des données incohérentes ; à l'inverse, une application au cœur solide mérite une interface à la hauteur. Le bon réflexe consiste à investir dans la finition à proportion de la durée de vie de l'application et du nombre de ses utilisateurs : un outil personnel et éphémère se contente du strict nécessaire, tandis qu'une application destinée à être déployée et utilisée quotidiennement justifie pleinement le soin décrit dans ce chapitre. Chaque technique présentée ici est à mobiliser quand elle sert réellement l'utilisateur — et à laisser de côté quand elle ne ferait qu'ajouter de la complexité pour le plaisir de l'effet.

---

> 🔗 **Situé dans le parcours :** ce chapitre prolonge les chapitres 6 (Formulaires), 7 (États) et 8 (Contrôles et événements) en y ajoutant la couche de finition, et prépare le chapitre 21 (Déploiement et distribution), où l'application habillée est packagée pour ses utilisateurs.

⏭️ [17.1. Ruban personnalisé (Ribbon) avec XML et callbacks VBA](/17-interface-utilisateur-avancee/01-ruban-personnalise-xml.md)
