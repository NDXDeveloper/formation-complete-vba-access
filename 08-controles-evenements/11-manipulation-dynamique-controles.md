🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8.11. Manipulation dynamique des contrôles (création, visibilité, activation)

Les contrôles ne sont pas figés : on peut, par code et à l'exécution, modifier leur **visibilité**, leur **activation**, leur **apparence** et leur **position** — voire en **créer** ou en **supprimer**. Ces manipulations rendent l'interface adaptative : afficher un champ selon un choix, verrouiller des contrôles selon un statut, mettre en évidence une donnée.

Cette section distingue les manipulations **de propriétés à l'exécution** (le travail quotidien) des manipulations **de structure** (création/suppression), qui obéissent aux mêmes contraintes que la génération d'états (chapitre 7.7).

---

## 8.11.1. Deux catégories : propriétés à l'exécution et structure

| Catégorie | Exemples | Contrainte |
|---|---|---|
| **Propriétés à l'exécution** | `Visible`, `Enabled`, `Locked`, couleurs, `Caption`, position | Réalisable **à tout moment**, dans les événements |
| **Structure** | `CreateControl`, `DeleteControl` | Exige le **mode Création** ; **impossible en ACCDE** |

La quasi-totalité des besoins relève de la première catégorie. La seconde, rare et contraignante, est à éviter quand une alternative existe (section 8.11.10).

---

## 8.11.2. Modifier les propriétés à l'exécution

Les propriétés d'un contrôle se règlent par code dans n'importe quel événement (`Open`, `Load`, `Current`, `AfterUpdate`, `Click`…) :

```vba
Me.txtMontant.Visible = False
Me.cmdSupprimer.Enabled = False
Me.lblTitre.Caption = "Fiche client"
Me.txtNom.ForeColor = vbRed
Me.txtNom.FontBold = True
```

La position et la taille s'expriment **en twips** (1 cm ≈ 567 twips) :

```vba
Me.txtNote.Width = 4000
Me.txtNote.Top = 1200
```

Pour une logique **par enregistrement**, ces réglages se placent dans l'événement `Current` (chapitres 8.7 et 8.8).

---

## 8.11.3. Enabled ou Locked : la distinction

Deux propriétés contrôlent l'accès à un contrôle, avec des effets **différents** souvent confondus :

- **`Enabled = False`** : le contrôle est **grisé**, ne peut **pas recevoir le focus** ni être saisi.
- **`Locked = True`** : le contrôle reste **accessible et lisible** (on peut y placer le curseur, copier son contenu), mais **non modifiable**.

Leurs combinaisons couvrent les besoins courants :

| `Enabled` | `Locked` | Comportement |
|---|---|---|
| `True` | `False` | Normal, modifiable |
| `True` | `True` | Lisible et sélectionnable, **non modifiable** (lecture seule) |
| `False` | `False` | **Grisé**, ni focus ni saisie |
| `False` | `True` | Affiché **normalement** (non grisé), mais ni focus ni saisie |

**Recommandation** : pour une lecture seule où l'utilisateur doit pouvoir lire/copier la valeur, employer `Locked = True` (en laissant `Enabled = True`). Pour désactiver complètement un contrôle, `Enabled = False`. La dernière combinaison (`Enabled = False`, `Locked = True`) sert à présenter un contrôle d'aspect normal mais inaccessible.

---

## 8.11.4. Visibilité et activation conditionnelles

Le pattern le plus fréquent : afficher ou activer un contrôle **selon une valeur**, dans l'événement adéquat.

```vba
' Afficher le motif uniquement si le statut est « Annulé »
Private Sub grpStatut_AfterUpdate()
    Me.txtMotif.Visible = (Me.grpStatut.Value = 3)
End Sub
```

```vba
' Verrouiller le montant pour un enregistrement validé (par enregistrement)
Private Sub Form_Current()
    Me.txtMontant.Locked = (Me.Statut.Value = "Validé")
End Sub
```

L'expression booléenne affectée directement à `Visible`/`Locked`/`Enabled` rend le code concis et lisible.

---

## 8.11.5. Agir sur plusieurs contrôles (parcours et recursion)

Pour agir sur un **ensemble** de contrôles, on parcourt la collection `Controls` en filtrant par `ControlType`, et l'on **recurse** dans les conteneurs (chapitre 8.1) :

```vba
Sub VerrouillerSaisie(ByVal conteneur As Object, ByVal verrou As Boolean)
    Dim ctl As Control
    For Each ctl In conteneur.Controls
        Select Case ctl.ControlType
            Case acTextBox, acComboBox, acCheckBox, acListBox
                ctl.Locked = verrou
            Case acTabCtl, acPage, acOptionGroup
                VerrouillerSaisie ctl, verrou   ' descendre dans les conteneurs
        End Select
    Next ctl
End Sub
```

Appel depuis le module du formulaire :

```vba
VerrouillerSaisie Me, True    ' verrouiller tous les champs de saisie
```

---

## 8.11.6. La propriété Tag — marquer un sous-ensemble de contrôles

La propriété **`Tag`** est une chaîne libre, vide par défaut, que le développeur peut renseigner pour **marquer** des contrôles et les traiter ensemble. C'est une alternative élégante au filtrage par type lorsqu'on veut agir sur un sous-ensemble arbitraire :

```vba
' Verrouiller uniquement les contrôles marqués "saisie" dans leur Tag
Dim ctl As Control
For Each ctl In Me.Controls
    If ctl.Tag = "saisie" Then ctl.Locked = True
Next ctl
```

On marque par exemple « obligatoire » les champs requis, ou « admin » les contrôles réservés, pour les manipuler en bloc selon le contexte.

---

## 8.11.7. Déplacer le focus avant de masquer ou désactiver

Un contrôle **invisible** ou **désactivé** ne peut **pas** recevoir le focus : tenter d'y déplacer le focus déclenche une erreur. Surtout, masquer ou désactiver le contrôle qui **détient actuellement** le focus provoque une erreur. Il faut donc **déplacer le focus ailleurs au préalable** :

```vba
Me.txtAutre.SetFocus            ' déplacer le focus d'abord
Me.txtMontant.Visible = False   ' puis masquer le contrôle
```

---

## 8.11.8. Éviter le scintillement (propriété Painting)

De nombreuses modifications successives de contrôles peuvent provoquer un **scintillement** désagréable. La propriété **`Painting`** du formulaire permet de **suspendre** l'affichage le temps des changements, puis de le rétablir :

```vba
Me.Painting = False
' ... nombreuses modifications de contrôles ...
Me.Painting = True
```

`Me.Repaint` force un rafraîchissement immédiat lorsque nécessaire. Ces aspects de performance d'affichage sont abordés au chapitre 18.

---

## 8.11.9. Manipulation structurelle : CreateControl et DeleteControl

Créer ou supprimer des contrôles par code repose sur `CreateControl` et `DeleteControl`, pendants de `CreateReportControl`/`DeleteReportControl` (chapitre 7.7), avec **les mêmes contraintes** :

```vba
DoCmd.OpenForm "F_Test", acDesign        ' mode Création obligatoire
CreateControl "F_Test", acTextBox, acDetail, , "Nom", 200, 200, 3000, 300
DoCmd.Close acForm, "F_Test", acSaveYes
```

`CreateControl` prend le **nom** du formulaire, le **type** de contrôle, la **section**, l'éventuel **parent** (groupe, page), le **champ lié** et la position en twips. `DeleteControl` supprime un contrôle par son nom.

> **Deux contraintes majeures** (identiques au chapitre 7.7) : ces opérations exigent le formulaire ouvert en **mode Création**, et sont **impossibles dans une base compilée en ACCDE** (chapitre 20.3).

---

## 8.11.10. Concevoir plutôt que créer : la voie recommandée

Créer des contrôles à l'exécution est **rare, verbeux et inopérant en ACCDE**. La voie pragmatique consiste presque toujours à **concevoir** tous les contrôles nécessaires en mode Création — même ceux destinés à rester cachés — puis à jouer sur leur **`Visible`** et leur **`Enabled`/`Locked`** selon le contexte.

Cette approche est **bien plus simple**, fonctionne en ACCDE, et couvre la quasi-totalité des besoins d'interface adaptative. On réserve `CreateControl` aux cas exceptionnels (générateurs de formulaires, interfaces définies par métadonnées).

---

## 8.11.11. Pièges et bonnes pratiques

- **Distinguer `Enabled` et `Locked`** : grisé/inaccessible d'un côté, lisible mais non modifiable de l'autre.
- **Déplacer le focus avant de masquer/désactiver** le contrôle qui le détient, sous peine d'erreur.
- **Ne pas faire `SetFocus`** vers un contrôle invisible ou désactivé.
- **Suspendre l'affichage** (`Painting = False`) pour de nombreuses modifications, afin d'éviter le scintillement (chapitre 18).
- **Préférer concevoir et masquer** plutôt que créer des contrôles à l'exécution : plus simple et compatible ACCDE.
- **`CreateControl`/`DeleteControl` exigent le mode Création** et sont **impossibles en ACCDE** (chapitres 7.7 et 20.3).
- **Utiliser la propriété `Tag`** pour agir sur un sous-ensemble arbitraire de contrôles.
- **Positions et tailles en twips** (1 cm ≈ 567 twips).
- **Pour l'apparence conditionnelle récurrente**, envisager la mise en forme conditionnelle (chapitre 17.8) plutôt que du code répétitif.

---

## 8.11.12. Récapitulatif

- La manipulation dynamique se divise en deux catégories : les **propriétés à l'exécution** (`Visible`, `Enabled`, `Locked`, couleurs, position) — réalisables partout — et la **structure** (`CreateControl`/`DeleteControl`), contrainte au mode Création et **impossible en ACCDE**.
- La distinction **`Enabled`** (grisé, inaccessible) / **`Locked`** (lisible mais non modifiable) est essentielle ; leurs combinaisons couvrent lecture seule et désactivation.
- La visibilité et l'activation **conditionnelles** se règlent dans l'événement adéquat (souvent `Current` pour le « par enregistrement »), par affectation directe d'une expression booléenne.
- On agit sur **plusieurs contrôles** via la collection `Controls`, le filtrage par `ControlType` et la **recursion** (chapitre 8.1), ou via la propriété **`Tag`** pour un sous-ensemble marqué.
- Toujours **déplacer le focus** avant de masquer/désactiver le contrôle actif, et **suspendre l'affichage** (`Painting`) pour limiter le scintillement (chapitre 18).
- La voie recommandée est de **concevoir** les contrôles (même cachés) et de jouer sur leurs propriétés, plutôt que de les **créer** par code — solution simple, compatible ACCDE, suffisante dans la quasi-totalité des cas.

⏭️ [8.12. Annulation d'événements avec Cancel — patterns courants](/08-controles-evenements/12-annulation-evenements-cancel.md)
