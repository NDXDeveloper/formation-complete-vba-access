🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 3. Rappels fondamentaux de programmation VBA

Les deux premiers chapitres ont posé le décor — la place de VBA dans Access, puis l'environnement de développement. Avant d'entrer dans les sujets propres à Access (modèle objet, `DoCmd`, formulaires, accès aux données…), ce chapitre fait une halte sur les **fondamentaux du langage VBA** lui-même : variables, structures de contrôle, procédures, portée, tableaux, chaînes, dates et collections.

## La nature de ce chapitre : un condensé de référence

Une précision s'impose d'emblée. Comme l'ont montré les sections 1.2 et 2.3, le langage VBA est **identique** dans toutes les applications Office ; seul le **modèle objet** change de l'une à l'autre. Les fondamentaux présentés ici ne sont donc **pas spécifiques à Access** : ce sont ceux de VBA en général. Ce chapitre est, à dessein, un **condensé de référence** — un rappel des notions essentielles, et non un cours complet construit à partir de zéro.

> 📘 **Pour un apprentissage approfondi** de ces fondamentaux, avec explications détaillées et progression pédagogique complète, reportez-vous à la **[Formation VBA pour Excel](https://github.com/NDXDeveloper/formation-complete-vba-excel)** associée. Le présent chapitre vise à **rafraîchir** ces bases et à garantir que vous disposez du socle nécessaire pour suivre la suite, en signalant au passage les points qui prennent une saveur particulière dans le contexte Access.

## Objectifs du chapitre

À l'issue de ce chapitre, vous aurez :

- **révisé** les briques fondamentales du langage VBA (variables et types, structures de contrôle, procédures et fonctions) ;
- **consolidé** des notions souvent mal maîtrisées, comme la portée et la durée de vie des variables, ou la distinction entre tableaux statiques et dynamiques ;
- **rafraîchi** la manipulation des chaînes de caractères et des dates, ainsi que l'usage des collections et du `Dictionary` ;
- repéré quelques **spécificités utiles** dans le contexte Access, qui seront développées dans les chapitres dédiés.

## Prérequis

Ce chapitre suppose que l'environnement est prêt (chapitres 1 et 2). Côté langage, une **première exposition à VBA** facilite grandement la lecture : l'objectif est de réviser, pas de découvrir. Si vous débutez complètement en programmation VBA, le plus efficace est de commencer par la **[Formation VBA pour Excel](https://github.com/NDXDeveloper/formation-complete-vba-excel)**, puis de revenir ici pour aborder Access.

## Au programme de ce chapitre

Le chapitre se compose des sept sections suivantes.

- **3.1 — [Variables, types de données et constantes](/03-rappels-fondamentaux/01-variables-types-constantes.md)**
  Déclarer et typer des variables, choisir le bon type de données, définir des constantes.

- **3.2 — [Structures de contrôle (If, Select Case, boucles)](/03-rappels-fondamentaux/02-structures-controle.md)**
  Conditionner et répéter l'exécution : tests, choix multiples, boucles.

- **3.3 — [Procédures Sub et fonctions Function](/03-rappels-fondamentaux/03-procedures-fonctions.md)**
  Organiser le code en procédures et fonctions, leur passer des arguments, renvoyer des valeurs.

- **3.4 — [Portée des variables et durée de vie](/03-rappels-fondamentaux/04-portee-variables.md)**
  Comprendre où une variable est visible et combien de temps elle conserve sa valeur.

- **3.5 — [Tableaux statiques et dynamiques](/03-rappels-fondamentaux/05-tableaux.md)**
  Manipuler des ensembles de valeurs indexées, à taille fixe ou redimensionnable.

- **3.6 — [Gestion des chaînes de caractères et des dates](/03-rappels-fondamentaux/06-chaines-dates.md)**
  Les fonctions de manipulation de texte et de dates, avec leurs subtilités.

- **3.7 — [Collections et Dictionary (Scripting.Dictionary)](/03-rappels-fondamentaux/07-collections-dictionary.md)**
  Stocker et retrouver des données par clé, au-delà des simples tableaux.

## Comment aborder ce chapitre

L'usage de ce chapitre dépend de votre niveau :

- si vous **maîtrisez déjà VBA**, parcourez-le en diagonale ou utilisez-le comme **mémo** auquel revenir au besoin ;
- si vos bases sont **un peu rouillées**, une lecture suivie vous remettra rapidement à niveau avant les chapitres spécifiques à Access ;
- si vous **débutez en VBA**, suivez d'abord la formation Excel mentionnée plus haut, puis revenez ici.

Les chapitres suivants supposent ces fondamentaux acquis. Une fois ce socle en place, vous pourrez vous concentrer sur ce qui fait la **spécificité d'Access** — son modèle objet et ses mécanismes propres — sans buter sur la syntaxe du langage.

---


⏭️ [3.1. Variables, types de données et constantes](/03-rappels-fondamentaux/01-variables-types-constantes.md)
