🔝 Retour au [Sommaire](/SOMMAIRE.md)

# Annexes

Les annexes regroupent la **documentation de référence** de cette formation : fiches techniques, tableaux de correspondance, listes normalisées et modèles de code. Contrairement aux chapitres, elles ne sont pas conçues pour être lues de façon linéaire, mais pour être **consultées ponctuellement**, au moment précis où une information est nécessaire pendant le développement.

Là où les 24 chapitres expliquent les concepts, les illustrent et les mettent en perspective, les annexes répondent à une logique différente : un accès immédiat à la donnée exacte. Un numéro d'erreur à identifier, la syntaxe d'une fonction Jet SQL à vérifier, une chaîne de connexion à recopier, un préfixe de nommage à appliquer, un raccourci clavier à mémoriser. Ce sont les pages que l'on garde ouvertes dans un second onglet pendant que l'on écrit du code.

## Pourquoi des annexes séparées

Plusieurs informations reviennent en permanence dans la pratique de VBA pour Access, mais n'ont leur place dans aucun chapitre en particulier — ou alors elles seraient dispersées sur plusieurs d'entre eux. Les rassembler dans des annexes dédiées présente trois avantages :

- **Centralisation** : une seule adresse pour chaque type d'information de référence, au lieu d'une recherche à travers les chapitres.
- **Densité** : un format tabulaire ou en liste, sans le texte explicatif des chapitres, pour aller à l'essentiel.
- **Pérennité** : ces pages restent utiles longtemps après la première lecture de la formation, comme aide-mémoire de travail.

## Comment utiliser cette section

Les annexes se consultent par recherche ciblée plutôt que par lecture intégrale. En pratique :

- Repérez l'annexe correspondant à votre besoin grâce à la vue d'ensemble ci-dessous.
- Utilisez la recherche de votre navigateur ou éditeur (`Ctrl+F`) à l'intérieur d'une annexe pour retrouver rapidement un élément précis (un numéro d'erreur, un nom de fonction, un préfixe…).
- Gardez à l'esprit que les annexes **complètent** les chapitres : elles fournissent la donnée brute, tandis que les chapitres en expliquent le contexte et le bon usage. En cas de doute sur l'emploi d'un élément, le chapitre concerné reste la référence.

## Vue d'ensemble des annexes

Les onze annexes peuvent se regrouper selon leur fonction.

### Références d'objets et de propriétés

- **A. Référence des propriétés et méthodes DoCmd** — Récapitulatif des principales actions de l'objet `DoCmd`, de leurs arguments et de leur usage courant. Le complément direct du chapitre 5.
- **B. Propriétés essentielles des formulaires et états** — Liste des propriétés les plus utilisées des objets `Form` et `Report` (source de données, filtres, affichage, événements…), avec leur signification.

### Diagnostic et gestion des erreurs

- **C. Codes d'erreur courants Access / DAO / ADO** — Table de correspondance entre les numéros d'erreur, leur description, leurs causes probables et les pistes de résolution. À garder sous la main lors du débogage (chapitre 13).

### Connexions et SQL

- **D. Chaînes de connexion ADO/ODBC pour les SGBD courants** — Modèles de chaînes de connexion prêts à adapter pour Access, SQL Server, Oracle, PostgreSQL, MySQL et autres bases (chapitres 10 et 23).
- **E. Référence Jet SQL — syntaxe et fonctions** — Aide-mémoire de la syntaxe du dialecte SQL propre à Access (Jet/ACE) et de ses fonctions intégrées (chapitre 11).
- **F. Correspondance Jet SQL ↔ T-SQL (SQL Server)** — Tableau de correspondance entre les deux dialectes, indispensable lors d'une migration ou d'une cohabitation avec SQL Server (chapitre 23).

### Conventions et productivité

- **G. Préfixes de nommage Leszynski/Reddick pour Access** — Liste normalisée des préfixes recommandés pour nommer objets, contrôles et variables de façon cohérente (chapitre 24).
- **I. Raccourcis clavier de l'éditeur VBA** — Tableau des raccourcis de l'éditeur (VBE) pour gagner en rapidité au quotidien (chapitre 2).
- **K. Modèles de code réutilisables (snippets)** — Bibliothèque de blocs de code prêts à l'emploi, illustrant les patterns récurrents abordés tout au long de la formation.

### Compatibilité système

- **H. Déclarations API Windows compatibles 32/64 bits** — Déclarations `Declare PtrSafe` prêtes à utiliser, compatibles avec les deux architectures (chapitres 21 et 22).

### Vocabulaire

- **J. Glossaire des termes Access et base de données** — Définitions des principaux termes employés dans la formation, utile pour lever toute ambiguïté de vocabulaire.

---

> 💡 **Bon réflexe** : avant de chercher une information sur le web, vérifiez si une annexe y répond déjà. Elles sont alignées sur les conventions et les choix techniques retenus dans cette formation, ce qui évite les incohérences.

⏭️ [A. Référence des propriétés et méthodes DoCmd](/annexes/a-reference-docmd.md)
