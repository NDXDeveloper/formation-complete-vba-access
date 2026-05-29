🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8.4. Contrôle onglet (TabControl) — navigation multi-pages

Lorsqu'un formulaire devient dense, le **contrôle onglet** permet de répartir ses contrôles sur plusieurs **pages** dont une seule est visible à la fois — par exemple « Général », « Adresse », « Historique ». C'est un **conteneur** (au sens du chapitre 8.1) qui héberge des pages, lesquelles hébergent à leur tour leurs propres contrôles.

Cette section décrit la structure du contrôle onglet, sa navigation par code, son événement `Change`, ainsi qu'une limite importante : l'impossibilité d'annuler un changement de page.

---

## 8.4.1. Rôle et structure du contrôle onglet

Le contrôle onglet organise une interface en pages thématiques. Sa structure est **imbriquée sur deux niveaux** :

```
Contrôle onglet (TabControl)
 ├── Page « Général »
 │    ├── txtNom
 │    └── txtPrenom
 ├── Page « Adresse »
 │    ├── txtRue
 │    └── txtVille
 └── Page « Historique »
      └── ctlSousForm
```

Le contrôle onglet contient une collection de **pages** ; chaque page est elle-même un conteneur possédant sa propre collection `Controls`. C'est cette imbrication qui explique pourquoi les contrôles d'une page ne sont pas enfants directs du formulaire (chapitre 8.1).

---

## 8.4.2. La collection Pages

Le contrôle onglet expose une collection **`Pages`** :

```vba
Debug.Print Me.TabCtl0.Pages.Count          ' nombre de pages
Debug.Print Me.TabCtl0.Pages(0).Caption     ' libellé de la première page
Debug.Print Me.TabCtl0.Pages("pgAdresse").Caption   ' par le nom
```

Chaque élément de cette collection est un objet **`Page`**, avec ses propres propriétés (section 8.4.4) et ses propres contrôles.

---

## 8.4.3. Propriétés du contrôle onglet

| Propriété | Rôle |
|---|---|
| `Value` | Indice (à partir de **0**) de la page **actuellement affichée** — lecture/écriture |
| `Pages` | Collection des pages |
| `Style` | Apparence : `0` Onglets, `1` Boutons, `2` Aucun |
| `MultiRow` | Affiche les onglets sur plusieurs rangées |
| `TabFixedHeight`, `TabFixedWidth` | Dimensions fixes des onglets |

La propriété **`Value`** est centrale : elle indique la page visible et permet d'en changer par code.

---

## 8.4.4. Propriétés d'une page

| Propriété | Rôle |
|---|---|
| `Caption` | Libellé affiché sur l'onglet |
| `Visible` | Affiche ou masque la page (et son onglet) |
| `PageIndex` | Position de la page (à partir de 0) |
| `Controls` | Collection des contrôles de la page |

```vba
Me.TabCtl0.Pages("pgGeneral").Caption = "Informations générales"
```

---

## 8.4.5. Naviguer entre les pages par code (Value)

Pour afficher une page par programmation, on affecte son indice à la propriété `Value` du contrôle onglet (indice **à partir de 0**) :

```vba
Me.TabCtl0.Value = 1   ' affiche la deuxième page
```

C'est le mécanisme de base de toute navigation pilotée par code (boutons d'assistant, redirection après une action, etc.).

---

## 8.4.6. L'événement Change

L'événement **`Change`** du contrôle onglet se déclenche à **chaque changement de page**, qu'il provienne d'un clic utilisateur ou d'une affectation par code :

```vba
Private Sub TabCtl0_Change()
    Select Case Me.TabCtl0.Value
        Case 0: Debug.Print "Page Général"
        Case 1: Debug.Print "Page Adresse"
        Case 2: Debug.Print "Page Historique"
    End Select
End Sub
```

C'est l'endroit privilégié pour réagir à l'arrivée sur une page : charger des données à la demande (section 8.4.11), rafraîchir un affichage, etc.

---

## 8.4.7. Afficher ou masquer des pages dynamiquement

La propriété `Visible` d'une page permet d'en masquer l'onglet selon le contexte — par exemple réserver une page d'administration à certains profils :

```vba
Me.TabCtl0.Pages("pgAdmin").Visible = False
```

> **Précaution** : ne pas masquer la page **actuellement affichée**. Basculer d'abord sur une autre page (`Me.TabCtl0.Value = 0`), puis masquer celle voulue.

---

## 8.4.8. Accéder aux contrôles d'une page

L'**accès par nom** est transparent à l'imbrication : un contrôle situé sur une page se référence directement, sans mention de la page (chapitre 8.1) :

```vba
Me.txtVille.Value = "Paris"   ' fonctionne même si txtVille est sur une page d'onglet
```

En revanche, la propriété `Parent` d'un tel contrôle renvoie la **page**, et l'**énumération** de tous les contrôles exige une recursion dans le contrôle onglet et ses pages (section 8.1.8). Les contrôles d'une page précise se parcourent via la collection `Controls` de cette page :

```vba
Dim ctl As Control
For Each ctl In Me.TabCtl0.Pages("pgAdresse").Controls
    Debug.Print ctl.Name
Next ctl
```

---

## 8.4.9. Navigation de type assistant (boutons Précédent/Suivant)

Le contrôle onglet sert souvent de base à un **assistant** multi-étapes. Des boutons « Précédent » et « Suivant » font alors progresser la propriété `Value` :

```vba
Private Sub cmdSuivant_Click()
    If Me.TabCtl0.Value < Me.TabCtl0.Pages.Count - 1 Then
        Me.TabCtl0.Value = Me.TabCtl0.Value + 1
    End If
End Sub

Private Sub cmdPrecedent_Click()
    If Me.TabCtl0.Value > 0 Then
        Me.TabCtl0.Value = Me.TabCtl0.Value - 1
    End If
End Sub
```

Pour un véritable assistant, on règle souvent `Style = 2` (Aucun) afin de **masquer les onglets** et de forcer la navigation par les boutons, ce qui permet de contrôler le flux (section suivante).

---

## 8.4.10. Validation au changement de page : une limite à connaître

Point **important** : l'événement `Change` du contrôle onglet **n'est pas annulable** (il ne reçoit pas d'argument `Cancel`). Il se déclenche **après** le changement de page. On ne peut donc **pas** empêcher l'utilisateur de quitter une page en cliquant sur un autre onglet.

Pour imposer une validation avant de passer à l'étape suivante, deux approches :

- **Navigation par boutons** (assistant) : valider dans le clic du bouton « Suivant », et ne changer `Value` qu'en cas de succès. Combiné à `Style = 2` (onglets masqués), c'est la méthode la plus fiable.
- **Validation au niveau de l'enregistrement** : s'appuyer sur l'événement `BeforeUpdate` du formulaire (chapitres 8.8 et 8.10), qui, lui, est annulable et garantit la cohérence avant l'enregistrement.

```vba
Private Sub cmdSuivant_Click()
    If Not DonneesPageValides() Then Exit Sub   ' validation maîtrisée
    Me.TabCtl0.Value = Me.TabCtl0.Value + 1
End Sub
```

---

## 8.4.11. Chargement différé des pages (performance)

Tous les contrôles d'un onglet — y compris ceux des pages **non affichées** — sont chargés à l'ouverture du formulaire. Un sous-formulaire placé sur une page peu consultée pèse donc sur le temps d'ouverture, même si l'utilisateur ne l'ouvre jamais.

Pour l'éviter, on **diffère le chargement** : on laisse vide la propriété `SourceObject` du sous-formulaire (chapitre 6.4), et on ne l'affecte que lorsque la page est affichée pour la première fois, dans l'événement `Change` :

```vba
Private Sub TabCtl0_Change()
    If Me.TabCtl0.Value = 2 Then                        ' page Historique
        If Len(Me.ctlSousForm.SourceObject) = 0 Then
            Me.ctlSousForm.SourceObject = "F_Historique"  ' chargé à la demande
        End If
    End If
End Sub
```

Les techniques d'optimisation des formulaires sont approfondies au chapitre 18.

---

## 8.4.12. Pièges et bonnes pratiques

- **`Value` est indexée à partir de 0** : la première page a l'indice 0.
- **L'événement `Change` n'est pas annulable** : pour valider avant de quitter une page, recourir à la navigation par boutons ou à `BeforeUpdate` (sections 8.4.10, 8.8, 8.10).
- **Ne pas masquer la page active** : basculer d'abord sur une autre page.
- **Accès par nom transparent**, mais `Parent` renvoie la page et l'énumération exige une recursion (chapitre 8.1).
- **Différer le chargement** des sous-formulaires des pages peu consultées via `SourceObject` (chapitres 6.4 et 18).
- **`Style = 2` (Aucun)** pour les assistants, afin de maîtriser le flux par les boutons.
- **Distinguer la collection `Pages` et la propriété `Value`** : l'une regroupe les pages, l'autre désigne la page affichée.

---

## 8.4.13. Récapitulatif

- Le contrôle onglet est un **conteneur** organisant l'interface en **pages** ; sa structure est imbriquée sur deux niveaux (onglet → pages → contrôles).
- La collection **`Pages`** donne accès aux pages ; la propriété **`Value`** (indexée à partir de **0**) désigne et change la page affichée.
- L'événement **`Change`** réagit à tout changement de page, mais **n'est pas annulable** : pour valider avant de quitter une page, employer la navigation par **boutons** ou l'événement **`BeforeUpdate`** (chapitres 8.8 et 8.10).
- Les pages se masquent dynamiquement via `Visible` (sans masquer la page active) ; les contrôles d'une page s'accèdent par **nom** de façon transparente, mais leur `Parent` est la **page** (chapitre 8.1).
- Pour les performances, **différer le chargement** des sous-formulaires des pages peu consultées (`SourceObject` vide puis affecté dans `Change` — chapitres 6.4 et 18).
- Pour un **assistant**, masquer les onglets (`Style = 2`) et piloter la progression par des boutons agissant sur `Value`.

⏭️ [8.5. Contrôles ActiveX dans Access](/08-controles-evenements/05-controles-activex.md)
