🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7. États (Reports)

Les **états** (reports) sont les objets dédiés à la **restitution** des données dans Access : impression, aperçu, export PDF, diffusion par e-mail. Là où les formulaires servent à saisir et à consulter de façon interactive, les états produisent une sortie **mise en forme et en lecture seule**, destinée à être imprimée ou transmise.

Ce chapitre traite du pilotage des états par VBA : adapter dynamiquement leur source de données, calculer des totaux, mettre en forme conditionnellement, générer des états par code, les exporter et automatiser leur envoi.

---

## États et formulaires : une parenté, des différences

Les états partagent avec les formulaires de nombreux concepts (modèle objet, sections, contrôles, propriété `RecordSource`), au point que beaucoup de techniques du chapitre 6 se transposent directement. Mais leur finalité et leur fonctionnement diffèrent sur des points essentiels :

| Aspect | Formulaires (chapitre 6) | États (ce chapitre) |
|---|---|---|
| Finalité | Interaction, saisie, consultation | Présentation, impression, diffusion |
| Accès aux données | Lecture / écriture | Lecture seule |
| Sortie | Écran | Aperçu, papier, PDF, RTF, HTML, e-mail |
| Événements clés | `Current`, `BeforeUpdate`, `AfterUpdate`, `Dirty` | `Open`, `NoData`, `Format`, `Print`, `Close` |
| Traitement des données | Un enregistrement courant | Par **passes de mise en forme**, section par section |

Cette dernière ligne est fondamentale : un état ne traite pas « un enregistrement courant » comme un formulaire, mais parcourt ses données en **mettant en forme chaque section** (parfois en plusieurs passes). C'est ce modèle qui explique le rôle particulier des événements `Format` et `Print`, abordés dès les sections 7.2 et 7.3.

---

## Ce que VBA apporte aux états

La conception visuelle d'un état couvre déjà beaucoup de besoins. VBA intervient lorsqu'il faut du **dynamisme** ou de l'**automatisation** :

- **Sources et filtres dynamiques** : produire un même état pour des périmètres variables (un client, une période, un service) sans multiplier les états.
- **Calculs et cumuls** : totaux par groupe, totaux généraux, numérotation, cumuls progressifs au fil des enregistrements.
- **Mise en forme conditionnelle par code** : faire varier l'apparence selon les valeurs (alertes, seuils, alternance de lignes).
- **Génération programmatique** : créer ou modifier des états et leurs contrôles à la volée.
- **Export multi-format** : produire des fichiers PDF, RTF ou HTML.
- **Diffusion automatisée** : paramétrer un état puis l'envoyer par e-mail, en lot si nécessaire.

---

## Contenu du chapitre

- **7.1. Modèle objet Report — sections et propriétés** : la hiérarchie d'un état et les propriétés essentielles à connaître pour le manipuler par code.
- **7.2. Événements d'un état (Open, Activate, NoData, Format, Print, Close)** : le cycle de vie d'un état et l'usage précis de chaque événement.
- **7.3. Sections (en-tête, détail, pied de page, en-têtes/pieds de groupe)** : le rôle de chaque section et son comportement lors de la mise en forme.
- **7.4. Manipulation dynamique du RecordSource et des filtres** : adapter la source d'un état au moment de l'ouverture, dans le prolongement des techniques du chapitre 6.5.
- **7.5. Calculs et totaux dans les états** : agrégats, cumuls et totaux par niveau de regroupement.
- **7.6. Sous-états et liaison parent/enfant** : imbriquer un état dans un autre et synchroniser les données.
- **7.7. Génération d'états par code (création, modification de contrôles)** : produire et modifier des états programmatiquement.
- **7.8. Export d'états (PDF, RTF, HTML) par DoCmd.OutputTo** : générer des fichiers dans les principaux formats.
- **7.9. Gestion de l'événement NoData — éviter les états vides** : détecter l'absence de données et empêcher l'impression d'un état vide.
- **7.10. États paramétrés et envoi par e-mail automatisé** : combiner paramétrage, export et envoi pour automatiser la diffusion.

---

## Prérequis et chapitres liés

Pour tirer le meilleur parti de ce chapitre, il est utile de maîtriser au préalable :

- le **chapitre 5 (DoCmd)**, notamment `OpenReport`, `OutputTo` et l'aperçu/impression (sections 5.2 et 5.6) ;
- le **chapitre 6 (Formulaires)**, dont de nombreux concepts (modèle objet, `RecordSource`, sections, contrôles) se transposent aux états ;
- le **chapitre 8 (Contrôles et événements)** pour la logique événementielle et la manipulation des contrôles ;
- les **chapitres 9 (DAO)** et **11 (SQL)** pour comprendre et construire les sources de données des états.

---

## Fil conducteur

Le chapitre suit une progression naturelle : d'abord **comprendre** la structure d'un état et son fonctionnement (7.1 à 7.3), puis **agir** sur ses données et son rendu (7.4 et 7.5), ensuite **composer et générer** des états plus riches (7.6 et 7.7), enfin **produire et diffuser** les résultats (7.8 à 7.10). Cette logique « comprendre → agir → composer → diffuser » permet d'aborder les sections dans l'ordre, chacune s'appuyant sur les précédentes.

⏭️ [7.1. Modèle objet Report — sections et propriétés](/07-etats-reports/01-modele-objet-report.md)
