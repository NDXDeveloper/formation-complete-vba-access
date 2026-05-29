🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13. Gestion des erreurs

## Introduction

Aucune application n'est à l'abri d'une erreur d'exécution. Un fichier déplacé, un enregistrement supprimé par un autre utilisateur, une clé primaire en double, une conversion de type impossible, une connexion réseau interrompue : autant de situations que le code ne peut pas toujours empêcher, mais qu'il doit impérativement savoir affronter. La différence entre une application Access amateur et une application professionnelle ne tient pas tant à l'absence d'erreurs qu'à la **manière dont elles sont gérées** lorsqu'elles surviennent.

Par défaut, lorsqu'une erreur non interceptée se produit en VBA, l'exécution s'arrête brutalement et Access affiche une boîte de dialogue technique, illisible pour l'utilisateur final, comportant les boutons *Fin* et *Débogage*. Dans le meilleur des cas, l'utilisateur est désorienté ; dans le pire, il clique sur *Débogage* et se retrouve face au code source de l'application, voire modifie ou interrompt un traitement en cours. Et lorsque l'application est déployée avec Access Runtime, ce comportement par défaut provoque tout simplement la fermeture sans avertissement de l'application — perdant au passage les données non enregistrées.

Ce chapitre présente les mécanismes, les patterns et les bonnes pratiques qui permettent de reprendre le contrôle : anticiper les erreurs, les intercepter proprement, informer l'utilisateur dans un langage compréhensible, journaliser l'incident pour le diagnostic, et — chaque fois que c'est possible — permettre à l'application de poursuivre son fonctionnement.

## Pourquoi la gestion des erreurs est particulière dans Access

VBA partage avec ses cousins (Excel, Word) un même modèle de gestion d'erreurs reposant sur les instructions `On Error` et l'objet `Err`. Mais l'environnement Access introduit des sources d'erreurs et des mécanismes d'interception qui lui sont propres, et qu'il faut maîtriser conjointement.

Dans une application Access, les erreurs proviennent en réalité de **trois couches distinctes** :

1. **Le moteur d'exécution VBA** lui-même — division par zéro, indice de tableau hors limites, variable objet non définie, type incompatible. Ce sont les erreurs « classiques » que l'on retrouve dans tout langage VBA.

2. **Le moteur de base de données** (Jet/ACE via DAO, ou un fournisseur OLE DB via ADO) — violation de clé primaire, violation d'intégrité référentielle, verrou impossible à poser, requête SQL malformée, échec d'un appel ODBC vers un serveur distant. Ces erreurs remontent souvent sous la forme d'un simple numéro générique côté VBA, alors que le détail réel est stocké dans une **collection d'erreurs** spécifique (`DBEngine.Errors` pour DAO, `Connection.Errors` pour ADO).

3. **L'interface d'Access** elle-même — erreurs déclenchées par la liaison des contrôles, la validation des champs, les actions de l'utilisateur dans un formulaire. Celles-ci ne passent pas par le `On Error` habituel mais sont notifiées via des **événements dédiés**, principalement `Form_Error` et `Report_Error`, accompagnés du code d'erreur Access dans le paramètre `DataErr`.

Gérer correctement les erreurs dans Access, c'est donc savoir traiter ces trois couches, comprendre laquelle est concernée selon le contexte, et choisir le bon point d'interception.

## Ce que vous allez apprendre

À l'issue de ce chapitre, vous serez en mesure de :

- reconnaître les erreurs les plus fréquemment rencontrées en VBA Access et en identifier la cause ;
- mettre en place une structure de gestion d'erreurs robuste et reproductible avec `On Error GoTo` ;
- utiliser `On Error Resume Next` à bon escient, sans masquer involontairement de vrais problèmes ;
- exploiter pleinement l'objet `Err` (numéro, description, source) et les instructions de reprise (`Resume`, `Resume Next`, `Resume <étiquette>`) ;
- intercepter spécifiquement les erreurs DAO et ADO via leurs collections respectives ;
- journaliser les incidents dans une table de log pour faciliter le diagnostic en production ;
- centraliser le traitement des erreurs dans un module réutilisable plutôt que de dupliquer la logique partout ;
- déclencher vos propres erreurs métier avec `Err.Raise` ;
- comprendre comment une erreur se propage entre les procédures appelantes et appelées, et concevoir une hiérarchie de gestion cohérente.

## Le modèle de gestion d'erreurs de VBA en bref

Contrairement à des langages plus récents, VBA ne dispose **pas** de bloc structuré de type `try / catch / finally`. Le langage repose sur un mécanisme plus ancien, dit « par déroutement » : on indique en début de procédure ce que VBA doit faire en cas d'erreur, à l'aide d'une instruction `On Error`.

Trois grandes options structurent ce mécanisme :

- `On Error GoTo <étiquette>` — en cas d'erreur, l'exécution saute immédiatement vers une étiquette de la procédure où se trouve le code de traitement. C'est la forme de référence, celle qui permet une gestion explicite et contrôlée.
- `On Error Resume Next` — en cas d'erreur, VBA ignore l'instruction fautive et passe à la suivante. Puissant mais dangereux s'il est mal employé, car il peut faire passer silencieusement des bugs réels.
- `On Error GoTo 0` — désactive toute gestion d'erreur en cours et rétablit le comportement par défaut (affichage de la boîte de dialogue système).

Tout au long du chapitre, ces instructions seront combinées à l'objet `Err`, qui renseigne sur l'erreur survenue, et aux instructions `Resume`, qui déterminent comment reprendre l'exécution une fois l'erreur traitée. L'ensemble forme un système cohérent dont chaque section détaille un aspect.

## Vue d'ensemble du chapitre

Le chapitre progresse du diagnostic vers la mise en œuvre, puis vers l'architecture :

- **[13.1. Erreurs courantes en VBA Access — catalogue et causes](/13-gestion-erreurs/01-erreurs-courantes-access.md)** dresse un inventaire raisonné des erreurs que vous rencontrerez le plus souvent, avec leurs causes typiques, afin d'apprendre à les reconnaître rapidement.

- **[13.2. On Error GoTo — structure robuste de gestion d'erreurs](/13-gestion-erreurs/02-on-error-goto.md)** établit le squelette standard d'une procédure correctement protégée, fondement de toutes les techniques suivantes.

- **[13.3. On Error Resume Next — utilisation contrôlée](/13-gestion-erreurs/03-on-error-resume-next.md)** explique quand et comment recourir à cette instruction sans tomber dans le piège des erreurs masquées.

- **[13.4. Objet Err — numéro, description, source](/13-gestion-erreurs/04-objet-err.md)** détaille l'objet central qui décrit l'erreur survenue et permet d'adapter le traitement à son numéro.

- **[13.5. Erreurs DAO et ADO — interception spécifique](/13-gestion-erreurs/05-erreurs-dao-ado.md)** aborde les particularités des erreurs issues du moteur de base de données et de leurs collections dédiées.

- **[13.6. Journalisation des erreurs dans une table Access](/13-gestion-erreurs/06-journalisation-erreurs-table.md)** montre comment conserver une trace exploitable des incidents, indispensable pour le support en production.

- **[13.7. Module de gestion centralisée des erreurs](/13-gestion-erreurs/07-module-centralise-erreurs.md)** propose de factoriser toute la logique de traitement dans un point unique, réutilisable dans l'ensemble de l'application.

- **[13.8. Création d'erreurs personnalisées avec Err.Raise](/13-gestion-erreurs/08-erreurs-personnalisees-raise.md)** explique comment générer vos propres erreurs métier pour signaler des situations spécifiques à votre application.

- **[13.9. Hiérarchie de propagation des erreurs entre procédures](/13-gestion-erreurs/09-propagation-erreurs.md)** traite de la manière dont une erreur remonte la pile des appels et de la façon de répartir les responsabilités entre procédures.

## Prérequis

Pour tirer pleinement profit de ce chapitre, il est recommandé d'être à l'aise avec les notions abordées précédemment, en particulier les **procédures `Sub` et fonctions `Function`** (chapitre 3), les **structures de contrôle** et la **portée des variables**. Une familiarité avec les objets **DAO** (chapitre 9) et **ADO** (chapitre 10) est utile pour la section 13.5, mais n'est pas indispensable à la compréhension des principes généraux.

## Approche du chapitre

La gestion des erreurs n'est pas une fonctionnalité que l'on « ajoute à la fin » d'un projet : c'est une discipline qui se conçoit dès l'écriture de chaque procédure. Les techniques présentées ici ne sont pas des artifices défensifs, mais des éléments structurants de la qualité logicielle. Une application bien protégée est une application dans laquelle l'erreur devient un événement prévu et traité, plutôt qu'un accident subi.

Les sections suivantes adoptent une progression volontairement incrémentale : on part de la simple reconnaissance des erreurs, on construit ensuite une structure de protection fiable, puis on industrialise cette structure jusqu'à une véritable architecture de gestion d'erreurs applicable à l'échelle d'un projet entier.

---


⏭️ [13.1. Erreurs courantes en VBA Access — catalogue et causes](/13-gestion-erreurs/01-erreurs-courantes-access.md)
