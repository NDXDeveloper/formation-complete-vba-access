🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 4. Modèle objet Access

Les chapitres précédents ont préparé le terrain : on sait désormais où VBA s'inscrit, comment se servir de l'éditeur, et l'on a révisé les fondamentaux du langage. Reste à découvrir ce que VBA **manipule** réellement dans Access — son **modèle objet**. C'est l'objet de ce chapitre, sans doute l'un des plus structurants de toute la formation, car tout ce qui suit — `DoCmd`, formulaires, états, accès aux données — s'appuie sur lui.

## Deux modèles objet en un

La section 1.2 l'a annoncé : ce qui distingue VBA d'une application Office à l'autre, ce n'est pas le langage mais le **modèle objet**. Sous Access, ce modèle présente une particularité : il en réunit en réalité **deux**, complémentaires.

- D'un côté, le modèle de l'**application Access** : l'objet `Application`, les formulaires et les états, les collections `All…`, les objets utilitaires `Screen` et `SysCmd`, ainsi que l'objet `DoCmd` (traité à part, au chapitre 5). C'est le versant **interface et pilotage**.
- De l'autre, le modèle **DAO** (*Data Access Objects*) : `DBEngine`, `Workspace`, `Database`. C'est le versant **moteur de données**.

Ce chapitre présente les deux, et surtout la **façon dont ils se rejoignent** — par exemple via `CurrentProject` pour l'application et `CurrentDb` pour les données.

## Pourquoi ce chapitre est essentiel

Maîtriser cette hiérarchie, c'est savoir **atteindre n'importe quel objet** depuis le code : un formulaire ouvert, la liste de toutes les tables, la base de données courante. C'est aussi comprendre des distinctions qui, mal saisies, deviennent des sources d'erreurs — au premier rang desquelles le choix entre `CurrentDb()` et `DBEngine(0)(0)` (section 4.6).

Le chapitre 3 a par ailleurs montré que ce modèle est constellé de **collections** (`Forms`, `AllTables`, `Fields`…) : tout le savoir-faire du `For Each` (section 3.2) y trouve directement son application.

## Objectifs du chapitre

À l'issue de ce chapitre, vous serez en mesure de :

- vous représenter la **hiérarchie** globale des objets Access et y situer chaque grande brique ;
- exploiter l'objet **`Application`** et ce qu'il expose ;
- distinguer **`CurrentProject`** et **`CurrentData`**, et utiliser les collections **`All…`** pour énumérer les objets ;
- comprendre le modèle **DAO** (`DBEngine`, `Workspace`, `Database`) ;
- saisir la différence cruciale entre **`CurrentDb()`** et **`DBEngine(0)(0)`**, et savoir lequel employer ;
- tirer parti des objets utilitaires **`Screen`** et **`SysCmd`**, souvent oubliés.

## Prérequis

Ce chapitre suppose les chapitres 1 à 3 assimilés, en particulier la notion de modèle objet (section 1.2) et la manipulation des collections (sections 3.2 et 3.7). Côté langage, la maîtrise des objets — déclaration, mot-clé `Set`, `For Each` — est indispensable.

## Au programme de ce chapitre

Le chapitre se compose des sept sections suivantes.

- **4.1 — [Vue d'ensemble de la hiérarchie des objets Access](/04-modele-objet-access/01-hierarchie-objets-access.md)**
  La carte générale du modèle objet, pour situer l'ensemble avant d'entrer dans le détail.

- **4.2 — [L'objet Application](/04-modele-objet-access/02-objet-application.md)**
  La racine du modèle Access : ce qu'elle représente et ce qu'elle donne à manipuler.

- **4.3 — [CurrentProject et CurrentData](/04-modele-objet-access/03-currentproject-currentdata.md)**
  Les deux conteneurs de l'application courante : ses objets côté projet et côté données.

- **4.4 — [Collections AllForms, AllReports, AllTables, AllQueries...](/04-modele-objet-access/04-collections-allobjects.md)**
  Énumérer tous les objets d'un type, qu'ils soient ouverts ou non.

- **4.5 — [DBEngine, Workspace et Database (DAO)](/04-modele-objet-access/05-dbengine-workspace-database.md)**
  Le sommet du modèle DAO et la chaîne qui mène jusqu'à la base de données.

- **4.6 — [CurrentDb() vs DBEngine(0)(0) — différences et cas d'usage](/04-modele-objet-access/06-currentdb-vs-dbengine.md)**
  Une distinction discrète mais déterminante pour un accès aux données fiable.

- **4.7 — [Screen et SysCmd — objets utilitaires souvent oubliés](/04-modele-objet-access/07-screen-syscmd.md)**
  Deux objets pratiques pour connaître l'objet actif et dialoguer avec Access.

## Comment aborder ce chapitre

Commencez par la **vue d'ensemble (4.1)** : elle donne la carte qui rend lisibles toutes les sections suivantes. Le reste du chapitre détaille ensuite chaque branche, du versant application (4.2 à 4.4) au versant données (4.5 et 4.6), avant les objets utilitaires (4.7).

La section **4.6**, en particulier, mérite une lecture attentive : la distinction entre `CurrentDb()` et `DBEngine(0)(0)` est un grand classique des bugs Access. Une fois ce chapitre acquis, vous disposerez de la **carte mentale** indispensable pour aborder sereinement `DoCmd` (chapitre 5), les formulaires (chapitre 6), les états (chapitre 7) et l'accès aux données (chapitres 9 et 10).

---


⏭️ [4.1. Vue d'ensemble de la hiérarchie des objets Access](/04-modele-objet-access/01-hierarchie-objets-access.md)
