🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8.8. Événements de formulaire (Current, BeforeUpdate, AfterUpdate, Dirty)

Quatre événements de niveau **formulaire** gouvernent la logique liée aux enregistrements : `Current` (réagir à l'enregistrement courant), `Dirty` (détecter une modification), `BeforeUpdate` (valider avant l'enregistrement) et `AfterUpdate` (réagir après). Ils se distinguent des événements de **contrôle** (chapitre 8.9), qui concernent un élément isolé. Maîtriser ces quatre événements, c'est tenir l'essentiel de la logique de saisie d'un formulaire.

Cette section approfondit chacun d'eux, dans le prolongement du cycle de vie présenté au chapitre 8.7.

---

## 8.8.1. Événements de formulaire ou de contrôle ?

| Niveau | Concerne… | Événements |
|---|---|---|
| **Formulaire** | L'**enregistrement** dans son ensemble | `Current`, `BeforeUpdate`, `AfterUpdate`, `Dirty` |
| **Contrôle** | Un **contrôle** isolé | `Change`, `Enter`, `Exit`, `BeforeUpdate`/`AfterUpdate` du contrôle… (chapitre 8.9) |

Les événements de formulaire raisonnent au niveau de l'enregistrement courant ; ils sont le bon endroit pour la logique métier qui porte sur l'ensemble d'une ligne (validation inter-champs, réaction au changement d'enregistrement).

---

## 8.8.2. Current — réagir à l'enregistrement courant

`Form_Current` se déclenche **chaque fois** qu'un enregistrement devient courant : à l'ouverture, à chaque navigation, après un `Requery`, après ajout ou suppression. Il ne concerne **pas** l'édition, mais la question « *sur quel enregistrement suis-je maintenant ?* ».

C'est l'emplacement de la logique d'interface **par enregistrement** : activer/désactiver des contrôles, ajuster un affichage, verrouiller un champ selon les données de la ligne.

```vba
Private Sub Form_Current()
    ' Verrouiller le montant pour un enregistrement validé
    Me.txtMontant.Locked = (Me.Statut.Value = "Validé")
End Sub
```

### La propriété NewRecord

`Current` se déclenche aussi lorsqu'on arrive sur l'**enregistrement vierge** de saisie. La propriété **`NewRecord`** permet de le détecter et d'adapter le comportement :

```vba
Private Sub Form_Current()
    If Me.NewRecord Then
        Me.txtNumero.Value = "(nouveau)"
    Else
        Me.txtMontant.Locked = (Me.Statut.Value = "Validé")
    End If
End Sub
```

> Important : dans `Current`, l'enregistrement n'est **pas** encore modifié. Ce n'est donc **pas** l'endroit pour valider une saisie (cela revient au `BeforeUpdate`, section 8.8.4).

---

## 8.8.3. Dirty — détecter la première modification

`Form_Dirty` se déclenche à la **première** modification de l'enregistrement courant — la première frappe après que la ligne est devenue courante. Il ne se déclenche **qu'une fois** par « salissure » : les modifications suivantes ne le redéclenchent pas (pour réagir à chaque frappe, voir l'événement `Change` au chapitre 8.9).

Son usage typique est la **détection** : activer un bouton « Enregistrer », positionner un indicateur, signaler des modifications en attente.

```vba
Private Sub Form_Dirty(Cancel As Integer)
    Me.cmdEnregistrer.Enabled = True   ' activer dès la 1re modification
End Sub
```

### La propriété Dirty

Au-delà de l'événement, la **propriété** `Dirty` est très employée :

- **lire** `Me.Dirty` indique s'il existe des modifications non enregistrées (`True`/`False`) ;
- **affecter** `Me.Dirty = False` **force l'enregistrement** de la ligne courante ;
- `Me.Undo` **annule** les modifications en cours.

```vba
' Forcer l'enregistrement s'il y a des modifications
If Me.Dirty Then Me.Dirty = False

' Annuler les modifications en cours
Me.Undo
```

C'est cette propriété que l'on reproduit manuellement sur un formulaire indépendant (chapitre 6.8).

---

## 8.8.4. BeforeUpdate — valider avant l'enregistrement

`Form_BeforeUpdate(Cancel)` se déclenche **juste avant** que l'enregistrement ne soit écrit dans la table — par navigation, fermeture, ou enregistrement explicite (`Me.Dirty = False`). Il est **annulable** : c'est **l'emplacement privilégié de la validation** de niveau enregistrement.

Affecter `Cancel = True` **empêche l'écriture** ; l'enregistrement reste alors « sale » et l'utilisateur demeure dessus pour corriger.

```vba
Private Sub Form_BeforeUpdate(Cancel As Integer)
    ' Validation inter-champs : date de fin postérieure à la date de début
    If Not IsNull(Me.txtDateFin.Value) And Not IsNull(Me.txtDateDebut.Value) Then
        If Me.txtDateFin.Value < Me.txtDateDebut.Value Then
            MsgBox "La date de fin doit suivre la date de début.", vbExclamation
            Cancel = True
            Me.txtDateFin.SetFocus
        End If
    End If
End Sub
```

Pourquoi ici plutôt que dans les événements de contrôle ? Parce que le `BeforeUpdate` du formulaire se déclenche **systématiquement** avant toute sauvegarde, alors qu'un événement de contrôle peut être **contourné** si l'utilisateur ne visite jamais ce contrôle. La validation et l'annulation sont approfondies aux chapitres 8.10 et 8.12.

---

## 8.8.5. AfterUpdate — réagir après l'enregistrement

`Form_AfterUpdate` se déclenche **après** l'enregistrement réussi de la ligne. Il n'est **pas** annulable (l'écriture a déjà eu lieu). On y place les traitements **postérieurs** à la sauvegarde : journalisation, rafraîchissement d'un formulaire lié, remise à zéro d'indicateurs.

```vba
Private Sub Form_AfterUpdate()
    Me.cmdEnregistrer.Enabled = False    ' désactiver après enregistrement
    ' Journaliser, rafraîchir un formulaire dépendant, etc.
End Sub
```

---

## 8.8.6. Les deux niveaux de BeforeUpdate / AfterUpdate

Distinction essentielle (annoncée au chapitre 8.7) : ces deux noms d'événements existent **aux deux niveaux** :

| Niveau | Se déclenche quand… | Portée |
|---|---|---|
| **Contrôle** | Un contrôle modifié **perd le focus** | Validation de la valeur du contrôle dans le tampon de l'enregistrement |
| **Formulaire** | L'**enregistrement entier** est sauvegardé | Validation et écriture de la ligne dans la table |

Lorsque les deux interviennent, l'ordre est : `BeforeUpdate`/`AfterUpdate` **du contrôle** (en quittant chaque contrôle modifié), puis `BeforeUpdate`/`AfterUpdate` **du formulaire** (à la sauvegarde du record). Les événements de contrôle sont détaillés au chapitre 8.9.

---

## 8.8.7. Les événements connexes (Undo, Delete, Error)

Au-delà des quatre événements principaux, le formulaire en expose d'autres, utiles pour la gestion des données :

- **`Undo(Cancel)`** : se déclenche lorsque l'utilisateur annule ses modifications (touche Échap) ; annulable.
- **`Delete(Cancel)`** : se déclenche **avant** la suppression de chaque enregistrement ; annulable, pour bloquer une suppression interdite.
- **`BeforeDelConfirm(Cancel, Response)`** et **`AfterDelConfirm(Status)`** : encadrent la boîte de confirmation de suppression.
- **`Error(DataErr, Response)`** : intercepte les erreurs du **moteur** survenant sur le formulaire (champ requis, violation de règle, doublon), pour afficher un message personnalisé.

```vba
Private Sub Form_Error(DataErr As Integer, Response As Integer)
    If DataErr = 3022 Then               ' doublon sur un index unique
        MsgBox "Cette valeur existe déjà.", vbExclamation
        Response = acDataErrContinue     ' supprime le message standard
    Else
        Response = acDataErrDisplay
    End If
End Sub
```

L'événement `Error` est complémentaire de la gestion d'erreurs `On Error` (chapitre 13) : il capte les erreurs **de données** que `On Error` ne voit pas.

---

## 8.8.8. Patterns courants

- **`Current`** : ajuster l'interface à la ligne courante, gérer le cas `NewRecord`.
- **`Dirty`** : activer un bouton « Enregistrer » ou positionner un indicateur de modification.
- **`BeforeUpdate`** : valider l'enregistrement (le **principal** lieu de validation), annuler si invalide.
- **`AfterUpdate`** : journaliser et rafraîchir après une sauvegarde réussie.

---

## 8.8.9. Pièges et bonnes pratiques

- **Valider dans le `BeforeUpdate` du formulaire** : il se déclenche toujours avant la sauvegarde, contrairement aux événements de contrôle, contournables (chapitres 8.10 et 8.12).
- **Logique par enregistrement dans `Current`**, pas dans `Load` : `Load` ne se déclenche qu'une fois (chapitre 8.7).
- **Ne pas valider dans `Current`** : l'enregistrement n'y est pas encore modifié.
- **`Dirty` se déclenche une seule fois** par modification, pas à chaque frappe : pour le suivi frappe par frappe, utiliser `Change` (chapitre 8.9).
- **`Cancel = True` dans `BeforeUpdate` laisse l'enregistrement « sale »** : l'utilisateur doit corriger ou annuler (`Me.Undo`).
- **`Me.Dirty = False` force l'enregistrement** et peut donc déclencher `BeforeUpdate` : en tenir compte.
- **Distinguer les deux niveaux** de `BeforeUpdate`/`AfterUpdate` (contrôle vs formulaire).
- **Traiter `NewRecord`** dans `Current` pour différencier création et modification.
- **Utiliser l'événement `Error`** pour personnaliser les messages d'erreurs de données (doublons, champs requis — chapitre 13).

---

## 8.8.10. Récapitulatif

- Les événements de **formulaire** raisonnent au niveau de l'**enregistrement**, à la différence des événements de **contrôle** (chapitre 8.9).
- **`Current`** se déclenche à chaque changement d'enregistrement : c'est le lieu de la logique d'interface **par enregistrement** ; la propriété **`NewRecord`** y distingue création et modification.
- **`Dirty`** signale la **première** modification ; la **propriété `Dirty`** indique l'état non enregistré, `Me.Dirty = False` force l'enregistrement et `Me.Undo` l'annule.
- **`BeforeUpdate`** (annulable) est le lieu **privilégié de la validation** avant écriture ; `Cancel = True` bloque l'enregistrement.
- **`AfterUpdate`** réagit **après** une sauvegarde réussie (journalisation, rafraîchissement).
- `BeforeUpdate`/`AfterUpdate` existent à **deux niveaux** (contrôle puis formulaire) ; des événements connexes (`Undo`, `Delete`, `Error`) complètent la gestion des données, l'événement **`Error`** captant les erreurs de données (chapitre 13).

⏭️ [8.9. Événements de contrôle (Change, GotFocus, LostFocus, Click, DblClick)](/08-controles-evenements/09-evenements-controle.md)
