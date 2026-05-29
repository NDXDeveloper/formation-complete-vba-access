🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 6. Formulaires (Forms)

Avec ce chapitre, nous abordons le cœur de l'expérience utilisateur d'une application Access : les **formulaires**. Ce sont eux qui constituent l'**interface** — l'endroit où l'utilisateur consulte, saisit et manipule les données. Après avoir posé le modèle objet (chapitre 4) et l'automatisation par `DoCmd` (chapitre 5), il est temps de plonger dans l'objet qui structure le dialogue avec l'utilisateur.

## Le formulaire, couche d'interface de l'application

La section 1.5 a présenté les formulaires comme la **couche de présentation** d'une application, et le chapitre 4 a montré comment les atteindre via la collection `Forms`. Ce chapitre approfondit l'objet **`Form`** lui-même : ses propriétés et méthodes, ses modes d'affichage, ses sous-formulaires, sa source de données, ainsi que plusieurs patterns d'usage — boîtes de dialogue, formulaires indépendants, communication entre écrans.

On y réutilisera abondamment des notions déjà rencontrées : l'ouverture par `DoCmd.OpenForm`, le passage de paramètres par `OpenArgs` et le filtrage par `Me.Filter` (chapitre 5) y trouvent ici leur traitement complet.

## Périmètre : le formulaire, pas (encore) ses contrôles

Une précision importante sur ce qui est couvert. Ce chapitre traite du **formulaire en tant qu'objet** — le contenant. Les **contrôles** qu'il abrite (zones de texte, listes déroulantes, cases à cocher…) et le **cycle des événements** (ouverture, validation, clic, perte de focus…) font l'objet du **chapitre 8**. Les deux chapitres sont complémentaires : l'un porte sur le **contenant**, l'autre sur le **contenu et son comportement**.

Pour mémoire, les **états** — pendants des formulaires pour la restitution et l'impression — sont, eux, traités au **chapitre 7**.

## Objectifs du chapitre

À l'issue de ce chapitre, vous serez en mesure de :

- exploiter le **modèle objet `Form`** : ses propriétés et méthodes essentielles ;
- accéder aux **formulaires ouverts** via `Forms!NomFormulaire` ;
- choisir et comprendre les **modes d'affichage** (simple, continu, feuille de données) ;
- mettre en place des **sous-formulaires** et leur liaison parent/enfant ;
- piloter le **`RecordSource`** et les **filtres** par code ;
- passer des paramètres à l'ouverture grâce à **`OpenArgs`** ;
- créer des **formulaires modaux** et des **boîtes de dialogue** personnalisées ;
- concevoir des **formulaires indépendants** (unbound) pour les saisies complexes ;
- organiser la **communication entre formulaires** sans couplage trop fort.

## Prérequis

Ce chapitre suppose les chapitres 1 à 5 assimilés, en particulier le modèle objet (chapitre 4, notamment la collection `Forms` et la distinction ouverts/existants des sections 4.1 et 4.4) et `DoCmd` (chapitre 5, surtout `OpenForm` en section 5.2). La maîtrise des objets, des collections et des événements de base est indispensable ; le **cycle complet des événements** sera, lui, détaillé au chapitre 8.

## Au programme de ce chapitre

Le chapitre se compose des neuf sections suivantes.

- **6.1 — [Modèle objet Form — propriétés et méthodes essentielles](/06-formulaires/01-modele-objet-form.md)**
  Les propriétés et méthodes incontournables de l'objet `Form`.

- **6.2 — [Accès aux formulaires ouverts via Forms!NomFormulaire](/06-formulaires/02-acces-formulaires-ouverts.md)**
  Atteindre un formulaire ouvert et ses contrôles depuis le code.

- **6.3 — [Formulaires en mode simple, continu et feuille de données](/06-formulaires/03-modes-affichage-formulaire.md)**
  Les trois modes d'affichage et leurs usages respectifs.

- **6.4 — [Sous-formulaires — liaison parent/enfant et propriété LinkMasterFields](/06-formulaires/04-sous-formulaires.md)**
  Imbriquer des formulaires et synchroniser leurs données.

- **6.5 — [RecordSource dynamique et filtres par code](/06-formulaires/05-recordsource-filtres-dynamiques.md)**
  Changer la source de données et filtrer un formulaire par programmation.

- **6.6 — [Propriété OpenArgs — passage de paramètres à l'ouverture](/06-formulaires/06-openargs.md)**
  Transmettre des informations à un formulaire au moment de l'ouvrir.

- **6.7 — [Formulaires modaux et boîtes de dialogue personnalisées](/06-formulaires/07-formulaires-modaux-dialogues.md)**
  Créer des dialogues sur mesure qui suspendent l'exécution du code.

- **6.8 — [Formulaires indépendants (unbound forms) pour saisies complexes](/06-formulaires/08-formulaires-independants.md)**
  Concevoir des écrans non liés à une source de données unique.

- **6.9 — [Communication entre formulaires (sans couplage fort)](/06-formulaires/09-communication-formulaires.md)**
  Faire dialoguer les formulaires tout en préservant leur indépendance.

## Comment aborder ce chapitre

La progression va du **fondamental** vers l'**architectural**. Les sections 6.1 et 6.2 posent l'objet `Form` et la façon d'y accéder ; 6.3 et 6.4 en décrivent la **structure** (modes d'affichage, sous-formulaires) ; 6.5 et 6.6 en exploitent le **dynamisme** (source de données, filtres, paramètres) ; enfin, 6.7 à 6.9 présentent des **patterns** plus avancés (dialogues, formulaires indépendants, communication).

Une lecture suivie est recommandée pour les débutants. Gardez à l'esprit que ce chapitre et le **chapitre 8** se complètent : dès que vous aurez besoin de réagir aux actions de l'utilisateur sur les contrôles, c'est là qu'il faudra vous reporter. Les propriétés essentielles des formulaires sont par ailleurs recensées en **annexe B**.

---


⏭️ [6.1. Modèle objet Form — propriétés et méthodes essentielles](/06-formulaires/01-modele-objet-form.md)
