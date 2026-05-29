🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 2. Interface et environnement de développement

Le chapitre précédent a posé les fondations conceptuelles : la place de VBA dans Access, ses différences avec Excel, son rapport aux macros, la structure d'une application et les formats de fichiers. Vous avez aussi, en section 1.4, configuré les principales options de développement. Il est temps d'entrer concrètement dans l'**atelier** où s'écrit le code : l'éditeur VBA et son environnement.

L'éditeur VBA — le **VBE** (*Visual Basic Editor*) — est un environnement à part entière, distinct de la fenêtre Access, et partagé par toutes les applications Office. C'est là que l'on écrit, organise, exécute et met au point le code. On y passe l'essentiel de son temps de développement : il est donc payant de bien le connaître.

Une bonne maîtrise de cet environnement — ses fenêtres, l'organisation des projets, les types de modules, les mécanismes de liaison aux bibliothèques — fait gagner un temps considérable et permet d'éviter quelques pièges classiques : références manquantes, mauvais choix de module, ou confusion entre liaison précoce et tardive. Ce chapitre vous installe confortablement dans l'atelier avant que les chapitres suivants ne vous fassent vraiment programmer.

## Objectifs du chapitre

À l'issue de ce chapitre, vous serez en mesure de :

- ouvrir l'éditeur VBA depuis Access et vous repérer dans son interface ;
- comprendre l'organisation d'un projet VBA grâce à l'explorateur de projets ;
- distinguer les différents types de modules et savoir lequel employer selon le besoin ;
- exploiter les fenêtres essentielles, en particulier Propriétés et Exécution immédiate ;
- activer et gérer les références aux bibliothèques COM (DAO, ADO, Scripting…) ;
- choisir en connaissance de cause entre liaison précoce (*early binding*) et liaison tardive (*late binding*) ;
- convertir des macros Access existantes en code VBA.

## Prérequis

Ce chapitre suppose la lecture du chapitre 1, en particulier la compréhension de la structure d'une application Access (section 1.5) et la configuration des options de développement (section 1.4). Les notions fondamentales du langage VBA (variables, procédures, fonctions) ne sont pas réexpliquées ici ; elles sont rappelées au chapitre 3 et développées dans la formation VBA pour Excel associée.

## Au programme de ce chapitre

Le chapitre se compose des sept sections suivantes.

- **2.1 — [Accès à l'éditeur VBA (Alt+F11) depuis Access](/02-interface-environnement/01-acces-editeur-vba.md)**
  Ouvrir le VBE par toutes les voies possibles et faire le tour de son interface générale.

- **2.2 — [Explorateur de projets : modules, formulaires, états](/02-interface-environnement/02-explorateur-projets.md)**
  Comprendre comment le code est organisé en projet, et naviguer entre modules, formulaires et états.

- **2.3 — [Types de modules : Standard, Classe, Form, Report](/02-interface-environnement/03-types-modules.md)**
  Distinguer les natures de modules et leurs rôles respectifs, pour placer le bon code au bon endroit.

- **2.4 — [Fenêtre Propriétés et fenêtre Exécution immédiate](/02-interface-environnement/04-fenetres-proprietes-execution.md)**
  Maîtriser deux fenêtres incontournables, l'une pour inspecter les propriétés, l'autre pour tester et déboguer.

- **2.5 — [Références et bibliothèques COM à activer (DAO, ADO, Scripting…)](/02-interface-environnement/05-references-bibliotheques-com.md)**
  Savoir quelles bibliothèques activer, comment et pourquoi, pour étendre les capacités de VBA.

- **2.6 — [Liaisons précoces (Early Binding) vs liaisons tardives (Late Binding)](/02-interface-environnement/06-early-vs-late-binding.md)**
  Comprendre les deux façons de référencer un objet externe, avec leurs avantages et leurs compromis.

- **2.7 — [Conversion de macros Access en code VBA](/02-interface-environnement/07-conversion-macros-vba.md)**
  Mettre en pratique le pont évoqué au chapitre 1 : transformer des macros en code éditable.

## Comment aborder ce chapitre

Les sections suivent une progression naturelle : on entre d'abord dans l'éditeur (2.1), on découvre l'organisation du code (2.2 et 2.3), on apprend à se servir des fenêtres clés (2.4), puis on aborde la connexion aux bibliothèques externes (2.5 et 2.6) et, enfin, la migration depuis les macros (2.7).

Si vous débutez avec le VBE, lisez ces sections dans l'ordre : elles construisent une familiarité progressive avec l'outil. Si vous le connaissez déjà, les sections 2.5 et 2.6 — références et liaison — méritent une attention particulière, car ce sont des sources fréquentes d'erreurs, y compris chez les développeurs expérimentés. L'objectif du chapitre n'est pas encore de produire des traitements complets, mais d'être **à l'aise dans l'environnement** : un préalable qui rendra tout le reste de la formation plus fluide.

---


⏭️ [2.1. Accès à l'éditeur VBA (Alt+F11) depuis Access](/02-interface-environnement/01-acces-editeur-vba.md)
