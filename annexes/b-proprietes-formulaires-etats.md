🔝 Retour au [Sommaire](/SOMMAIRE.md)

# B. Propriétés essentielles des formulaires et états

Cette annexe rassemble les propriétés les plus utilisées des objets `Form` et `Report`. Elle complète les **chapitres 6 et 7** : ceux-ci expliquent le modèle objet et les usages, tandis que cette annexe sert d'aide-mémoire pour retrouver une propriété, ses valeurs admises et son rôle.

Il s'agit d'une **sélection des propriétés essentielles**, et non de la liste exhaustive. La feuille de propriétés (`F4` en mode Création) et l'Explorateur d'objets (`F2`) listent l'intégralité des propriétés disponibles.

## Comment accéder aux propriétés

Trois voies principales :

- **En conception** : feuille de propriétés (`F4`), organisée par onglets (Format, Données, Événement, Autres).
- **Depuis le module de l'objet** : le mot-clé `Me` désigne le formulaire ou l'état courant.
  ```vba
  Me.Caption = "Liste des clients"
  Me.Filter = "Ville = 'Rouen'"
  Me.FilterOn = True
  ```
- **Depuis l'extérieur**, sur un objet **ouvert**, via les collections `Forms` et `Reports` :
  ```vba
  Forms!frmClients.Caption = "Clients actifs"
  Reports![rptFactures].Filter = "Annee = 2026"
  ```

Pour une propriété personnalisée ou rarement exposée, on passe par la collection `Properties` de l'objet.

## Convention de lecture

- Les valeurs sont données sous forme de **constantes** quand elles existent ; les valeurs numériques sont indiquées à titre informatif.
- La mention **(lecture seule)** signale une propriété non modifiable par code.
- ⚠️ Certaines propriétés ne s'appliquent que lorsque l'objet est **(ré)ouvert** : les modifier sur un objet déjà affiché reste sans effet immédiat (voir *Points de vigilance*).

---

## Propriétés communes aux formulaires et aux états

Ces propriétés existent à la fois sur `Form` et sur `Report` et concernent l'identité et la source de données.

| Propriété | Valeurs / type | Rôle |
|---|---|---|
| `Name` | Chaîne | Nom de l'objet (lecture seule à l'exécution) |
| `Caption` | Chaîne | Texte affiché dans la barre de titre / l'onglet |
| `RecordSource` | Chaîne | Table, requête ou instruction SQL alimentant l'objet |
| `Filter` | Chaîne | Condition de filtrage (clause SQL **sans** le mot `WHERE`) |
| `FilterOn` | Booléen | Active (`True`) ou non le filtre défini par `Filter` |
| `OrderBy` | Chaîne | Tri (liste de champs **sans** `ORDER BY`, ex. `"Nom ASC, Prenom"`) |
| `OrderByOn` | Booléen | Active ou non le tri défini par `OrderBy` |
| `Visible` | Booléen | Visibilité de l'objet |
| `Tag` | Chaîne | Étiquette libre, exploitable par le développeur |
| `HasModule` | Booléen | Indique si l'objet possède un module de code associé |
| `Section(index)` | Objet `Section` | Accès aux sections (voir plus bas) |

> Définir `RecordSource` à l'exécution provoque un rechargement complet des données (équivalent d'un `Requery`). À utiliser avec discernement dans les traitements fréquents (chapitre 18.5).

---

## Propriétés spécifiques aux formulaires

### Données et permissions

| Propriété | Valeurs / type | Rôle |
|---|---|---|
| `RecordsetType` | 0 Feuille réponse dynamique · 1 (maj incohérentes) · 2 Capture instantanée | Type du jeu d'enregistrements lié |
| `Recordset` | Objet Recordset | Jeu d'enregistrements sous-jacent (DAO/ADO) ; assignable pour lier le formulaire (chapitre 6.5) |
| `RecordsetClone` | Recordset DAO | Clone pour parcourir/manipuler les enregistrements sans déplacer l'affichage (lecture seule — chapitre 9.12) |
| `AllowEdits` | Booléen | Autorise la modification des enregistrements |
| `AllowAdditions` | Booléen | Autorise l'ajout d'enregistrements |
| `AllowDeletions` | Booléen | Autorise la suppression d'enregistrements |
| `DataEntry` | Booléen | À l'ouverture, n'affiche qu'un nouvel enregistrement vierge |
| `RecordLocks` | 0 Aucun · 1 Tous les enregistrements · 2 Enregistrement modifié | Stratégie de verrouillage en multi-utilisateur (chapitre 15.6) |
| `Dirty` | Booléen | `True` si l'enregistrement courant est modifié et non enregistré ; affecter `False` force l'enregistrement |
| `NewRecord` | Booléen | `True` si l'on est positionné sur un nouvel enregistrement (lecture seule) |
| `CurrentRecord` | Long | Numéro de l'enregistrement courant (lecture seule) |

### Affichage et fenêtre

| Propriété | Valeurs / constantes | Rôle |
|---|---|---|
| `DefaultView` | 0 Formulaire simple · 1 Formulaires continus · 2 Feuille de données · 5 Double affichage (split) | Mode d'affichage par défaut (chapitre 6.3) |
| `BorderStyle` | 0 Aucune · 1 Fine · 2 Ajustable · 3 Boîte de dialogue | Style de bordure de la fenêtre |
| `Modal` | Booléen | Fenêtre modale (capture le focus dans l'application) |
| `PopUp` | Booléen | Fenêtre indépendante, au-dessus des autres |
| `ScrollBars` | 0 Aucune · 1 Horizontale · 2 Verticale · 3 Les deux | Barres de défilement |
| `RecordSelectors` | Booléen | Affiche la barre de sélection d'enregistrement à gauche |
| `NavigationButtons` | Booléen | Affiche les boutons de navigation en bas |
| `DividingLines` | Booléen | Lignes de séparation entre sections/enregistrements |
| `AutoCenter` | Booléen | Centre le formulaire à l'ouverture |
| `AutoResize` | Booléen | Ajuste la taille de la fenêtre au formulaire |
| `ControlBox` | Booléen | Affiche le menu système (icône, fermeture) |
| `MinMaxButtons` | 0 Aucun · 1 Réduction · 2 Agrandissement · 3 Les deux | Boutons réduire/agrandir |
| `CloseButton` | Booléen | Active le bouton de fermeture de la fenêtre |
| `Width` | Long (twips) | Largeur du formulaire (1 cm ≈ 567 twips) |

### Comportement et événements

| Propriété | Valeurs / type | Rôle |
|---|---|---|
| `OpenArgs` | Variant | Argument transmis par `DoCmd.OpenForm` (lecture seule — chapitre 6.6) |
| `Cycle` | 0 Tous les enregistrements · 1 Enregistrement courant · 2 Page courante | Comportement de la touche Tab en fin de formulaire |
| `KeyPreview` | Booléen | Le formulaire intercepte les événements clavier avant ses contrôles |
| `TimerInterval` | Long (millisecondes) | Période de déclenchement de l'événement `Timer` ; `0` le désactive |
| `Painting` | Booléen | Active/désactive le rafraîchissement de l'affichage (optimisation) |

**Propriétés d'événement** (chaînes désignant la procédure ou `[Event Procedure]`) les plus utilisées : `OnOpen`, `OnLoad`, `OnCurrent`, `BeforeUpdate`, `AfterUpdate`, `BeforeInsert`, `AfterInsert`, `OnDirty`, `OnUndo`, `OnClose`, `OnUnload`, `OnActivate`, `OnDeactivate`, `OnTimer`, `OnError`. Le détail de ces événements et de leur ordre figure aux chapitres 8.7 et 8.8.

---

## Propriétés spécifiques aux états

### Données et regroupement

| Propriété | Valeurs / type | Rôle |
|---|---|---|
| `RecordSource`, `Filter`, `FilterOn`, `OrderBy`, `OrderByOn` | — | Voir *Propriétés communes* (chapitre 7.4) |
| `HasData` | Integer | `0` si l'état n'a aucune donnée ; utile pour détecter un état vide (lecture seule — préférer l'événement `NoData`, chapitre 7.9) |
| `GroupLevel(n)` | Objet | Niveau de regroupement/tri ; expose `ControlSource`, `SortOrder`, `GroupOn`, `GroupInterval`, `KeepTogether` (chapitres 7.3 et 7.5) |
| `DateGrouping` | 0 Système américain · 1 Paramètres locaux | Mode de regroupement des dates |

### Mise en page et pagination

| Propriété | Valeurs / type | Rôle |
|---|---|---|
| `Page` | Long | Numéro de la page courante (disponible pendant la mise en forme / l'impression) |
| `Pages` | Long | Nombre total de pages (disponible une fois la mise en page calculée) |
| `Caption` | Chaîne | Titre de la fenêtre d'aperçu |
| `Width` | Long (twips) | Largeur de la zone d'impression |

### Formatage dynamique (dans l'événement Format/Print)

Ces propriétés se manipulent dans les événements de section pour des mises en page avancées (étiquettes, sauts conditionnels, multi-colonnes).

| Propriété | Valeurs / type | Rôle |
|---|---|---|
| `FormatCount` | Long | Nombre de déclenchements de l'événement `Format` pour la section (lecture seule) |
| `PrintCount` | Long | Nombre de déclenchements de l'événement `Print` pour la section (lecture seule) |
| `MoveLayout` | Booléen | Indique si Access avance à la position d'impression suivante |
| `NextRecord` | Booléen | Indique si Access passe à l'enregistrement suivant |
| `PrintSection` | Booléen | Indique si la section est effectivement imprimée |

---

## Propriétés des sections

Une **section** (objet `Section`) est commune aux formulaires et aux états. On y accède par index (constantes `AcSection`) ou par nom.

### Constantes de section — `AcSection`

| Constante | Valeur | Section |
|---|---|---|
| `acDetail` | 0 | Détail |
| `acHeader` | 1 | En-tête de formulaire/état |
| `acFooter` | 2 | Pied de formulaire/état |
| `acPageHeader` | 3 | En-tête de page |
| `acPageFooter` | 4 | Pied de page |
| `acGroupLevel1Header` | 5 | En-tête de groupe (niveau 1) |
| `acGroupLevel1Footer` | 6 | Pied de groupe (niveau 1) |
| `acGroupLevel2Header` | 7 | En-tête de groupe (niveau 2) |
| `acGroupLevel2Footer` | 8 | Pied de groupe (niveau 2) |

### Propriétés d'une section

| Propriété | Valeurs / type | Rôle |
|---|---|---|
| `Visible` | Booléen | Affiche ou masque la section |
| `Height` | Long (twips) | Hauteur de la section |
| `CanGrow` | Booléen | La section s'agrandit pour afficher tout le contenu |
| `CanShrink` | Booléen | La section se réduit si le contenu est plus court |
| `KeepTogether` | Booléen | Évite de scinder la section sur deux pages (états) |
| `ForceNewPage` | 0 Aucun · 1 Avant section · 2 Après section · 3 Avant et après | Saut de page autour de la section (états) |
| `NewRowOrCol` | 0 Aucun · 1 Avant · 2 Après · 3 Avant et après | Saut de colonne (états multi-colonnes) |
| `RepeatSection` | Booléen | Répète l'en-tête de groupe en haut de chaque page (états) |
| `DisplayWhen` | 0 Toujours · 1 Impression seule · 2 Écran seul | Conditions d'affichage de la section |
| `BackColor` | Long | Couleur de fond (utile pour l'alternance de lignes, chapitre 17.8) |

Exemple d'accès :

```vba
Me.Section(acDetail).CanGrow = True
Me.Section(acHeader).Visible = False
```

---

## Points de vigilance

> ⚠️ **Propriétés « de conception » modifiables seulement à l'ouverture.** Des propriétés comme `DefaultView`, `BorderStyle`, `ScrollBars`, `RecordSelectors`, `NavigationButtons`, `ControlBox`, `MinMaxButtons`, `Cycle` peuvent être affectées par code, mais le changement ne s'applique qu'au **prochain affichage** de l'objet — pas sur une fenêtre déjà ouverte. À l'inverse, `Caption`, `RecordSource`, `Filter`/`FilterOn`, `OrderBy`/`OrderByOn`, `AllowEdits`/`AllowAdditions`/`AllowDeletions`, `Visible`, `TimerInterval` et les propriétés de section s'appliquent immédiatement.

> ⚠️ **`Filter` et `OrderBy` n'incluent jamais leurs mots-clés SQL.** On écrit `Me.Filter = "Ville = 'Rouen'"` (pas `WHERE`) et `Me.OrderBy = "Nom ASC"` (pas `ORDER BY`). Un filtre défini reste inactif tant que `FilterOn` n'est pas à `True` (idem `OrderByOn`).

> ⚠️ **`Me` n'est valable que dans le module de l'objet.** Depuis l'extérieur, on utilise `Forms!Nom` ou `Reports!Nom`, et l'objet doit être **ouvert** sous peine d'erreur d'exécution.

> ⚠️ **`Page` et `Pages` (états) ne sont fiables que pendant la mise en forme.** Les lire en dehors des événements de section/impression ou de l'aperçu peut renvoyer une valeur non significative.

> ⚠️ **`RecordsetClone` vs `Recordset`.** `RecordsetClone` sert à parcourir et localiser des enregistrements **sans** déplacer la position affichée par le formulaire ; pour synchroniser l'affichage, on recopie ensuite le signet (`Me.Bookmark = rs.Bookmark`). Détails au chapitre 9.12.

---

> 💡 **À retenir.** La grande majorité de ces propriétés sont communes aux formulaires et aux états dès qu'il s'agit de la **source de données** (`RecordSource`, `Filter`, `OrderBy`). Les différences portent surtout sur l'**interaction** (propre aux formulaires : navigation, saisie, fenêtre) et la **pagination/regroupement** (propre aux états). Pour toute propriété absente de cette sélection, la feuille de propriétés (`F4`) et l'Explorateur d'objets (`F2`) restent les références complètes.

⏭️ [C. Codes d'erreur courants Access / DAO / ADO](/annexes/c-codes-erreur-access-dao-ado.md)
