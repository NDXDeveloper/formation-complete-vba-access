🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 16. Programmation orientée objet en VBA Access

La plupart du code VBA écrit pour Access est **procédural** : des modules remplis de procédures et de fonctions, souvent rattachées aux événements des formulaires. Cette approche convient parfaitement aux petites applications. Mais VBA permet aussi une véritable **programmation orientée objet** (POO), au moyen des **modules de classe**, et c'est cette dimension — longtemps sous-exploitée dans le monde Access — que ce chapitre explore. La POO n'est pas une fin en soi : c'est un ensemble d'outils qui, employés à bon escient, rendent une application plus structurée, plus robuste et plus facile à faire évoluer.

## Pourquoi ce chapitre est important

À mesure qu'une application grandit, le code procédural montre ses limites. Les règles métier se dispersent dans les événements des formulaires, les mêmes requêtes SQL se retrouvent dupliquées en plusieurs endroits, et le code d'accès aux données se mêle aux gestionnaires de clic des boutons. Conséquence : ajouter une fonctionnalité oblige à toucher de nombreux formulaires, une modification du schéma impose de traquer le SQL partout, et tester quoi que ce soit suppose de cliquer dans l'interface.

La POO apporte des réponses à ces difficultés : l'**encapsulation** (regrouper données et comportements, masquer les détails internes), l'**abstraction** (cacher la mécanique d'accès aux données derrière une interface claire), la **séparation des responsabilités** (distinguer interface, logique métier et accès aux données), la **réutilisabilité** et une bien meilleure **testabilité**. Ces bénéfices prennent toute leur valeur dans les applications volumineuses et durables — et facilitent considérablement une éventuelle migration, comme on le verra.

## Le problème à résoudre — un exemple

Considérons une application de gestion de clients écrite de façon procédurale. La validation d'un client est codée dans l'événement `BeforeUpdate` d'un formulaire ; la même requête de recherche de clients est réécrite dans trois formulaires différents ; l'enregistrement passe par du code DAO inséré directement dans un `Click` de bouton. Le jour où la règle de validation change, il faut la corriger à plusieurs endroits ; le jour où l'on migre vers un serveur, il faut réécrire du SQL disséminé dans toute l'application.

Une conception orientée objet réorganise tout cela : une classe `Client` encapsule un client et ses règles ; un `ClientRepository` centralise l'accès aux données ; la logique métier vit dans une couche dédiée ; et les formulaires se contentent d'**orchestrer** ces objets. Une règle ne se corrige plus qu'à un seul endroit, et changer le moteur de données n'affecte que le repository.

## Ce que vous allez apprendre

À l'issue de ce chapitre, vous saurez :

- créer et instancier des **modules de classe** ;
- exposer un état interne de façon contrôlée via les **propriétés** (`Property Get`, `Let`, `Set`) ;
- doter une classe de **méthodes** et appliquer l'**encapsulation** ;
- concevoir des classes **événementielles** avec `RaiseEvent` et `WithEvents` ;
- construire des **collections personnalisées** et exploiter le polymorphisme par **interfaces** (`Implements`) ;
- abstraire l'accès aux données avec le **patron Repository** ;
- mettre en œuvre les patrons **Singleton, Factory et Observer** en VBA ;
- structurer une application en **couches** (DAL, BLL, UI).

## Prérequis

Ce chapitre suppose une bonne maîtrise du VBA procédural :

- les **procédures et fonctions** ([section 3.3](/03-rappels-fondamentaux/03-procedures-fonctions.md)) et la **portée des variables** ([section 3.4](/03-rappels-fondamentaux/04-portee-variables.md)) ;
- les **collections** et le `Dictionary` ([section 3.7](/03-rappels-fondamentaux/07-collections-dictionary.md)) ;
- l'accès aux données en **DAO et ADO** (chapitres 9 et 10), que les patrons d'accès viendront encapsuler ;
- la **gestion des erreurs** ([chapitre 13](/13-gestion-erreurs/README.md)) ;
- les **formulaires** et leurs événements (chapitres 6 et 8), que la POO viendra alléger.

## Vue d'ensemble : ce que la POO apporte

Le chapitre s'articule autour de quelques notions fondamentales. L'**encapsulation** consiste à réunir dans un objet les données et les comportements qui les concernent, en n'exposant à l'extérieur qu'une interface maîtrisée. L'**abstraction** permet de manipuler un concept (un client, un dépôt de données) sans se soucier de son implémentation. Le **polymorphisme**, obtenu en VBA par les interfaces, autorise plusieurs classes à être traitées de façon uniforme. La **composition** assemble des objets entre eux, et la **conception événementielle** laisse des objets réagir aux événements d'autres objets. Ces briques se combinent ensuite dans les patrons et l'architecture en couches présentés en fin de chapitre.

## Spécificités — et limites — de la POO en VBA

Il est essentiel de cadrer ce que VBA permet et ce qu'il ne permet pas, car ces limites façonnent les techniques employées.

VBA offre des **modules de classe** dotés de propriétés, de méthodes et d'événements, l'encapsulation par des membres privés, le polymorphisme par interfaces (`Implements`), et des événements personnalisés (`RaiseEvent` / `WithEvents`).

En revanche, VBA **ne propose pas d'héritage d'implémentation** : on ne peut pas faire dériver une classe d'une autre pour en hériter le code. On privilégie donc la **composition** et les **interfaces** plutôt que l'héritage. De même, VBA ne connaît pas de **constructeur paramétré** — l'événement `Class_Initialize` ne reçoit aucun argument —, ce qui conduit à initialiser les objets après création ou à recourir à des **fonctions fabriques**. Enfin, il n'y a ni surcharge de méthodes, ni véritables membres statiques de classe. Ces contraintes ne sont pas des obstacles, mais elles orientent vers un style propre à VBA, sur lequel s'appuient les sections suivantes.

## Une approche pragmatique

La POO est un **outil, pas un dogme**. Pour un petit utilitaire ou un traitement ponctuel, le procédural reste souvent le choix le plus simple et le plus lisible. La POO gagne sa place dans les applications **vastes et durables**, là où la structure et la séparation des responsabilités sont rentables sur le long terme. L'écueil inverse — sur-concevoir une petite application en multipliant les classes inutiles — est tout aussi réel. Ce chapitre vise donc à donner les moyens de choisir, pas à imposer la POO partout.

## Structure du chapitre

- **16.1.** [Modules de classe — création et instanciation](01-modules-classe.md) — la brique de base : définir une classe et créer des objets.
- **16.2.** [Propriétés (Property Get, Let, Set)](02-proprietes-get-let-set.md) — exposer un état interne de manière contrôlée.
- **16.3.** [Méthodes et encapsulation](03-methodes-encapsulation.md) — doter les objets de comportements et masquer leurs détails.
- **16.4.** [Événements personnalisés dans une classe (RaiseEvent / WithEvents)](04-evenements-personnalises-classe.md) — concevoir des objets qui émettent et écoutent des événements.
- **16.5.** [Collections personnalisées et interfaces (Implements)](05-collections-interfaces.md) — regrouper des objets et obtenir le polymorphisme.
- **16.6.** [Patron Repository — abstraction de l'accès aux données](06-patron-repository.md) — isoler la mécanique d'accès aux données derrière une interface.
- **16.7.** [Patrons Singleton, Factory et Observer en VBA](07-singleton-factory-observer.md) — adapter les patrons de conception classiques à VBA.
- **16.8.** [Architecture en couches (DAL, BLL, UI) appliquée à Access](08-architecture-en-couches.md) — la synthèse : organiser l'application en couches distinctes.

## Positionnement dans la formation

Ce chapitre s'appuie sur les **fondamentaux procéduraux** du [chapitre 3](/03-rappels-fondamentaux/README.md) et sur les technologies d'accès aux données des chapitres 9 et 10, qu'il vient encapsuler et structurer.

Il éclaire plusieurs chapitres ultérieurs. La testabilité qu'apporte la POO prépare les **tests unitaires** ([section 19.4](/19-debogage-tests/04-tests-unitaires-framework.md)). L'abstraction de l'accès aux données par le patron Repository facilite grandement la **migration vers un serveur** ([chapitre 23](/23-migration-interoperabilite/README.md)), puisque changer de moteur de données ne touche alors qu'une couche. Enfin, les principes d'**architecture applicative** et de **refactoring** développés ici trouvent leur prolongement dans le [chapitre 24 sur les bonnes pratiques](/24-bonnes-pratiques-ressources/README.md), notamment aux sections [24.2](/24-bonnes-pratiques-ressources/02-architecture-applicative.md) et [24.5](/24-bonnes-pratiques-ressources/05-revue-code-refactoring.md).

⏭️ [16.1. Modules de classe — création et instanciation](/16-poo-vba-access/01-modules-classe.md)
