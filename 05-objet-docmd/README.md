🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 5. L'objet DoCmd

Le chapitre précédent a cartographié le modèle objet d'Access et a, à plusieurs reprises, renvoyé un objet à ce chapitre-ci : **`DoCmd`**. C'est l'objet de l'**action** par excellence — celui par lequel on *fait faire* des choses à Access : ouvrir un formulaire, lancer une requête, imprimer un état, importer des données. Suffisamment central pour mériter son propre chapitre, il est l'un des objets les plus utilisés au quotidien.

## DoCmd, le levier d'automatisation d'Access

La section 1.2 l'avait souligné : `DoCmd` est un objet **propre à Access**, sans équivalent direct dans Excel. Il regroupe sous forme de **méthodes** les **actions** que l'on retrouve aussi dans les macros — c'est d'ailleurs pourquoi la conversion d'une macro en VBA (section 2.7) produit massivement des appels à `DoCmd`. Maîtriser cet objet, c'est donc disposer du principal levier d'automatisation de l'application.

On y trouve, organisées en grandes familles, les actions courantes : **ouvrir et fermer** des objets, **naviguer** entre enregistrements, **exécuter** des requêtes, **importer et exporter** des données, **imprimer**, ainsi que diverses commandes utilitaires.

## DoCmd n'est pas toujours le meilleur outil

Une nuance, toutefois, traverse ce chapitre. `DoCmd` est central et commode, mais il n'est pas systématiquement le **meilleur** choix. Pour certaines opérations, des **méthodes objet directes**, plus modernes, se révèlent plus claires ou plus puissantes — modifier la propriété `Filter` d'un formulaire plutôt que recourir à `DoCmd.ApplyFilter`, par exemple. Ce chapitre enseigne donc `DoCmd` en profondeur, tout en indiquant — à la **section 5.8** — quand lui préférer une alternative.

## Objectifs du chapitre

À l'issue de ce chapitre, vous serez en mesure de :

- comprendre le **rôle central** de `DoCmd` dans l'automatisation d'Access ;
- **ouvrir et fermer** formulaires, états, requêtes et tables ;
- **naviguer** et vous déplacer entre les enregistrements ;
- **exécuter** des requêtes avec `RunSQL` et `OpenQuery` ;
- **importer et exporter** des données (`TransferDatabase`, `TransferSpreadsheet`, `TransferText`) ;
- **imprimer** et **exporter** des objets (`PrintOut`, `OutputTo`) ;
- employer les **méthodes utilitaires** (`SetWarnings`, `Hourglass`, `Beep`, `Quit`) ;
- savoir quand préférer les **alternatives modernes** à `DoCmd`.

## Prérequis

Ce chapitre suppose les chapitres 1 à 4 assimilés, en particulier la nature de `DoCmd` (section 1.2), sa parenté avec les macros (sections 1.3 et 2.7) et le modèle objet Access (chapitre 4). La maîtrise des arguments de procédure — notamment les **arguments nommés** (section 3.3) — est ici très utile, car les méthodes de `DoCmd` en comportent souvent beaucoup.

## Au programme de ce chapitre

Le chapitre se compose des huit sections suivantes.

- **5.1 — [Rôle central de DoCmd dans l'automatisation Access](/05-objet-docmd/01-role-docmd.md)**
  Ce qu'est `DoCmd`, sa logique, et sa place parmi les outils d'automatisation.

- **5.2 — [Ouverture et fermeture d'objets (OpenForm, OpenReport, OpenQuery, OpenTable)](/05-objet-docmd/02-ouverture-fermeture-objets.md)**
  Les méthodes les plus utilisées : ouvrir et fermer les objets de l'application.

- **5.3 — [Navigation et déplacement entre enregistrements](/05-objet-docmd/03-navigation-enregistrements.md)**
  Se déplacer dans les enregistrements d'un formulaire ou d'une feuille de données.

- **5.4 — [Exécution de requêtes (RunSQL, OpenQuery)](/05-objet-docmd/04-execution-requetes.md)**
  Lancer des requêtes action et ouvrir des requêtes par code.

- **5.5 — [Import et export de données (TransferDatabase, TransferSpreadsheet, TransferText)](/05-objet-docmd/05-import-export-donnees.md)**
  Échanger des données avec d'autres bases, Excel et des fichiers texte.

- **5.6 — [Impression et aperçu (PrintOut, OutputTo)](/05-objet-docmd/06-impression-apercu.md)**
  Imprimer, prévisualiser et exporter des objets vers différents formats.

- **5.7 — [Autres méthodes utiles (SetWarnings, Hourglass, Beep, Quit)](/05-objet-docmd/07-autres-methodes-utiles.md)**
  Les commandes utilitaires du quotidien et leurs précautions d'emploi.

- **5.8 — [Alternatives modernes à DoCmd (méthodes objet directes)](/05-objet-docmd/08-alternatives-modernes-docmd.md)**
  Prendre du recul : quand une méthode objet directe vaut mieux que `DoCmd`.

## Comment aborder ce chapitre

Commencez par la **section 5.1**, qui pose la logique d'ensemble de `DoCmd`. Les sections 5.2 à 5.7 détaillent ensuite chaque **famille d'actions** ; elles se prêtent aussi à une consultation **au fil des besoins**, comme une référence. La section 5.8, enfin, prend de la hauteur et mérite une lecture attentive : elle vous évitera de recourir à `DoCmd` là où une approche plus directe serait préférable.

Pour une **référence exhaustive** des méthodes et arguments de `DoCmd`, reportez-vous à l'**annexe A**, consultable en permanence.

---


⏭️ [5.1. Rôle central de DoCmd dans l'automatisation Access](/05-objet-docmd/01-role-docmd.md)
