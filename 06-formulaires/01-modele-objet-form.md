🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 6.1. Modèle objet Form — propriétés et méthodes essentielles

Ce chapitre commence par l'essentiel : l'objet **`Form`** et ses **propriétés et méthodes** les plus utilisées. Cette section est la boîte à outils sur laquelle s'appuieront toutes les suivantes — filtrage (6.5), sous-formulaires (6.4), dialogues (6.7)… Rappelons d'abord ce qu'est, du point de vue du code, un formulaire.

## Le formulaire est un objet

Comme l'a montré la section 2.3, un formulaire est une **instance de son module de classe** (le « code derrière le formulaire »). À ce titre, il possède des **propriétés**, des **méthodes** et des **événements** (ces derniers traités au chapitre 8) — et l'on peut même lui ajouter ses propres membres (voir plus bas).

Dans le code d'un formulaire, le mot-clé **`Me`** désigne **cette instance**, c'est-à-dire le formulaire lui-même. C'est la façon la plus directe et la plus sûre de référencer ses propriétés, ses contrôles et ses méthodes.

## Référencer le formulaire et ses contrôles

```vba
' Depuis le code du formulaire, Me désigne le formulaire lui-même
Me.Caption = "Fiche client"
Me.RecordSource = "qryClientsActifs"
Me.AllowAdditions = False

' Référencer un contrôle du formulaire
Me.txtNom = "Dupont"
Me.txtNom.Enabled = False
```

Depuis **l'extérieur** du formulaire, on y accède par la collection `Forms` (`Forms!frmClients`), traitée à la section suivante (6.2). L'accès aux **contrôles** et la notation `!` relèvent du chapitre 8 ; retenez ici que `Me.txtNom` et `Me!txtNom` désignent le même contrôle.

## Les propriétés essentielles

Le tableau suivant récapitule les propriétés les plus employées de l'objet `Form`, avec un renvoi vers la section qui les approfondit.

| Propriété | Rôle | Détail |
|---|---|---|
| `RecordSource` | source de données du formulaire | section 6.5 |
| `Recordset` / `RecordsetClone` | jeu d'enregistrements sous-jacent | section 9.12 |
| `Filter` / `FilterOn` | filtre et son activation | sections 5.8 et 6.5 |
| `OrderBy` / `OrderByOn` | tri et son activation | — |
| `AllowEdits` / `AllowAdditions` / `AllowDeletions` | autorisations de données | — |
| `DataEntry` | ouverture en saisie seule (nouveaux enreg.) | — |
| `Dirty` | l'enregistrement courant est-il modifié ? | sections 5.8 et 8.8 |
| `NewRecord` | est-on sur un nouvel enregistrement ? | — |
| `Caption` | titre du formulaire | — |
| `Visible` | visible / masqué | sections 5.2 et 6.7 |
| `Modal` / `PopUp` | comportement modal / fenêtre indépendante | section 6.7 |
| `DefaultView` | mode d'affichage par défaut | section 6.3 |
| `Name` | nom du formulaire | — |
| `ActiveControl` | contrôle ayant le focus | chapitre 8 |
| `Controls` | collection des contrôles | chapitre 8 |
| `Parent` | formulaire parent (cas d'un sous-formulaire) | section 6.4 |
| `OpenArgs` | paramètre transmis à l'ouverture | section 6.6 |

## Les méthodes essentielles

Côté méthodes, quelques-unes reviennent constamment.

```vba
Me.Requery              ' réexécute la source de données
Me.Refresh              ' relit les enregistrements courants
Me.Recalc               ' recalcule les contrôles calculés
Me.Repaint              ' force le rafraîchissement de l'affichage
Me.Undo                 ' annule les modifications de l'enregistrement courant
Me.txtNom.SetFocus      ' donne le focus à un contrôle
```

À cette liste s'ajoutent `SetFocus` (donner le focus au formulaire) et `GoToPage` (atteindre une page, pour les formulaires multi-pages).

## Requery, Refresh, Recalc, Repaint : ne pas confondre

Ces quatre méthodes se ressemblent mais ont des effets bien distincts — une source de confusion fréquente.

| Méthode | Effet | Coût |
|---|---|---|
| `Requery` | **réexécute** la source : prend en compte les **ajouts/suppressions** et **revient au 1er enregistrement** | élevé |
| `Refresh` | **relit** les enregistrements **courants** (modifications faites par d'autres), **conserve la position** ; n'ajoute ni ne supprime | modéré |
| `Recalc` | recalcule les **contrôles calculés** du formulaire | faible |
| `Repaint` | redessine l'**écran** (effet purement visuel) | faible |

En résumé : `Requery` pour rafraîchir l'**ensemble** des données (au prix d'un retour au début), `Refresh` pour voir les **modifications** des enregistrements déjà chargés, `Recalc` pour les **calculs**, `Repaint` pour l'**affichage**. L'optimisation de ces rafraîchissements est abordée au chapitre 18.

## Ajouter ses propres membres au formulaire

Puisque le module d'un formulaire est une **classe** (section 2.3), on peut lui ajouter ses **propres propriétés et méthodes** `Public`. Le formulaire devient alors un objet plus riche, qui expose une interface propre — par exemple une propriété renvoyant l'identifiant du client affiché :

```vba
' Dans le module du formulaire : un membre public personnalisé
Public Property Get ClientID() As Long
    ClientID = Nz(Me.txtClientID, 0)
End Property
```

Cette possibilité est précieuse pour la **communication entre formulaires** (section 6.9) et, plus largement, pour la programmation orientée objet (chapitre 16).

## À retenir

- Un formulaire est une **instance de son module de classe** ; dans son code, **`Me`** le désigne et permet d'accéder à ses propriétés, contrôles et méthodes.
- Propriétés essentielles : **données** (`RecordSource`, `Filter`/`FilterOn`, `AllowEdits`…, `Dirty`), **état/apparence** (`Caption`, `Visible`, `Modal`, `DefaultView`) et **contexte** (`OpenArgs`, `Parent`).
- Méthodes essentielles : `Requery`, `Refresh`, `Recalc`, `Repaint`, `Undo`, `SetFocus`.
- **Ne pas confondre** : `Requery` (réexécute, voit ajouts/suppressions), `Refresh` (relit les enreg. courants), `Recalc` (calculs), `Repaint` (affichage).
- Le module de formulaire étant une **classe**, on peut y ajouter des **membres publics** (utile pour la communication entre formulaires, section 6.9, et la POO, chapitre 16).

---


⏭️ [6.2. Accès aux formulaires ouverts via Forms!NomFormulaire](/06-formulaires/02-acces-formulaires-ouverts.md)
