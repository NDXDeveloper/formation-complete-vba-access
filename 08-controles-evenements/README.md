🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8. Contrôles et événements

Les **contrôles** sont les éléments d'interface d'un formulaire — zones de texte, listes déroulantes, cases à cocher, boutons, onglets. Les **événements** sont ce qui rend une application vivante : du code se déclenche en réaction aux actions de l'utilisateur (clic, saisie, changement de valeur) et aux moments clés du système (ouverture, mise à jour, perte de focus). Ensemble, ils forment le cœur de la **programmation événementielle**, le modèle sur lequel repose toute application Access.

Ce chapitre approfondit ces deux notions : les contrôles et leurs propriétés, puis le modèle d'événements qui orchestre l'interactivité.

---

## Deux thèmes complémentaires

Le chapitre se lit en deux temps :

- **Les contrôles (8.1 à 8.6)** — les briques de l'interface : leur hiérarchie, les types les plus courants (texte, listes, booléens, onglets, ActiveX) et la liaison aux données.
- **Les événements (8.7 à 8.12)** — le moteur de l'interactivité : le cycle de vie d'un formulaire, les événements de formulaire et de contrôle, la validation des saisies et l'annulation d'actions.

| | Contrôles (8.1–8.6) | Événements (8.7–8.12) |
|---|---|---|
| Nature | Éléments visuels et leurs propriétés | Déclencheurs de code |
| Question | *Avec quoi* l'utilisateur interagit ? | *Quand* le code s'exécute-t-il ? |
| Enjeu principal | Choix du bon contrôle et liaison aux données | Maîtrise de **l'ordre** des événements |

---

## La programmation événementielle au cœur d'Access

Une application Access ne s'exécute pas de façon linéaire : elle **réagit**. Comprendre **quand** chaque événement se déclenche — et dans **quel ordre** — est la compétence déterminante de ce chapitre. La plupart des comportements imprévus (validation qui ne se déclenche pas, valeur lue trop tôt, message affiché en double) trouvent leur origine dans une mauvaise compréhension de la séquence des événements.

C'est pourquoi le **cycle de vie d'un formulaire** (section 8.7) est un passage obligé, sur lequel s'appuient toutes les sections suivantes consacrées aux événements.

---

## Contenu du chapitre

- **8.1. Hiérarchie des contrôles dans Access** : comment Access organise les contrôles et comment y accéder par code.
- **8.2. Contrôles texte, liste déroulante et liste (TextBox, ComboBox, ListBox)** : les contrôles de saisie et de sélection les plus courants.
- **8.3. Contrôles booléens (CheckBox, OptionGroup, ToggleButton)** : les choix oui/non et les choix exclusifs.
- **8.4. Contrôle onglet (TabControl) — navigation multi-pages** : organiser une interface dense en plusieurs pages.
- **8.5. Contrôles ActiveX dans Access** : intégrer des composants externes, avec leurs précautions.
- **8.6. Propriété ControlSource — contrôles liés et indépendants** : la liaison d'un contrôle à un champ, distinction fondamentale entre contrôles liés et indépendants.
- **8.7. Cycle de vie complet d'un formulaire (ordre des événements)** : la séquence des événements, fondation de tout le reste.
- **8.8. Événements de formulaire (Current, BeforeUpdate, AfterUpdate, Dirty)** : réagir au niveau de l'enregistrement.
- **8.9. Événements de contrôle (Change, GotFocus, LostFocus, Click, DblClick)** : réagir au niveau du contrôle.
- **8.10. Validation des données et messages d'erreur personnalisés** : contrôler les saisies et guider l'utilisateur.
- **8.11. Manipulation dynamique des contrôles (création, visibilité, activation)** : agir sur les contrôles par code.
- **8.12. Annulation d'événements avec Cancel — patterns courants** : empêcher une action (saisie, mise à jour, fermeture).

---

## Prérequis et chapitres liés

Pour aborder ce chapitre dans les meilleures conditions, il est utile de connaître :

- le **chapitre 3 (Rappels fondamentaux VBA)**, car les gestionnaires d'événements sont des procédures ;
- le **chapitre 4 (Modèle objet Access)** pour la hiérarchie des objets, dont les contrôles font partie ;
- le **chapitre 6 (Formulaires)**, puisque les contrôles vivent sur les formulaires et que ce chapitre en prolonge naturellement l'étude.

Ce chapitre éclaire en retour le **chapitre 7 (États)**, qui réutilise plusieurs de ces concepts (événements, manipulation de contrôles), et prépare les chapitres consacrés à l'**interface utilisateur avancée** (chapitre 17).

---

## Fil conducteur

La progression va du **statique** au **dynamique** : d'abord les contrôles et leur liaison aux données (ce qui compose l'interface), puis les événements et la validation (ce qui la fait réagir), enfin la manipulation par code et l'annulation d'actions (ce qui la rend pleinement pilotable). Aborder les sections dans l'ordre est recommandé, le cycle de vie des événements (8.7) servant de clé de voûte à toute la seconde moitié du chapitre.

⏭️ [8.1. Hiérarchie des contrôles dans Access](/08-controles-evenements/01-hierarchie-controles.md)
