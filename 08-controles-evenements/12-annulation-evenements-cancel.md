🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8.12. Annulation d'événements avec Cancel — patterns courants

De nombreux événements exposent un argument **`Cancel`** qui permet au code de **bloquer l'action** sur le point de se produire : empêcher un enregistrement invalide, refuser une fermeture, interdire une suppression, retenir le focus sur un champ. Ce mécanisme, rencontré à plusieurs reprises dans les chapitres 7 et 8, est ici rassemblé et systématisé.

Cette section clôt le chapitre en présentant le fonctionnement de `Cancel`, la liste des événements annulables, les patterns d'usage les plus courants, et — point crucial — ce que `Cancel` **ne peut pas** annuler.

---

## 8.12.1. Le mécanisme Cancel

Certaines procédures d'événement reçoivent un argument **`Cancel As Integer`**. Affecter `Cancel = True` (équivalent à `-1`, ou toute valeur non nulle) **annule l'action par défaut** qui a déclenché l'événement :

```vba
Private Sub Form_BeforeUpdate(Cancel As Integer)
    Cancel = True   ' l'enregistrement n'aura pas lieu
End Sub
```

`Cancel` est un paramètre passé **par référence** : c'est en le modifiant que le code communique sa décision à Access. Laissé à `False` (`0`), l'action se déroule normalement.

L'**action** annulée dépend de l'événement : l'ouverture, la sauvegarde, la fermeture, la suppression, la sortie d'un contrôle, la mise en forme d'une section…

---

## 8.12.2. Quels événements sont annulables ?

Tous les événements ne sont pas annulables. Voici les principaux qui le sont, et ce que `Cancel` interrompt dans chaque cas :

| Événement | `Cancel = True` annule… |
|---|---|
| **`Open`** (formulaire/état) | l'ouverture |
| **`BeforeUpdate`** (formulaire) | l'enregistrement de la ligne |
| **`BeforeInsert`** (formulaire) | la création d'un nouvel enregistrement |
| **`Delete`** (formulaire) | la suppression d'un enregistrement |
| **`Unload`** (formulaire) | la fermeture |
| **`Undo`** (formulaire) | l'annulation des modifications |
| **`BeforeUpdate`** (contrôle) | la validation de la valeur du contrôle |
| **`Exit`** (contrôle) | la sortie du contrôle (le focus reste) |
| **`DblClick`** (contrôle) | l'action par défaut du double-clic |
| **`NoData`** (état) | l'affichage/impression de l'état vide |
| **`Format`** / **`Print`** (état) | la mise en forme / l'impression de la section |

Les sections suivantes détaillent les patterns associés aux plus utilisés.

---

## 8.12.3. Pattern : valider avant l'enregistrement (Form BeforeUpdate)

C'est l'usage le plus important de `Cancel` (chapitres 8.8 et 8.10) : bloquer l'écriture d'un enregistrement invalide. La ligne reste « sale » et l'utilisateur demeure dessus pour corriger.

```vba
Private Sub Form_BeforeUpdate(Cancel As Integer)
    If Len(Nz(Me.txtNom.Value, "")) = 0 Then
        MsgBox "Le nom est obligatoire.", vbExclamation
        Cancel = True
        Me.txtNom.SetFocus
    End If
End Sub
```

---

## 8.12.4. Pattern : valider un champ en gardant le focus (Control Exit / BeforeUpdate)

Au niveau d'un contrôle, `Cancel` **retient le focus** sur le champ : la valeur n'est pas validée, l'utilisateur ne peut pas quitter le contrôle tant qu'il n'a pas corrigé.

```vba
Private Sub txtEmail_BeforeUpdate(Cancel As Integer)
    If Len(Nz(Me.txtEmail.Value, "")) > 0 And InStr(Me.txtEmail.Value, "@") = 0 Then
        MsgBox "Adresse e-mail invalide.", vbExclamation
        Cancel = True            ' la valeur n'est pas validée
    End If
End Sub
```

> **Attention au piégeage** : un `Cancel = True` dans `Exit` ou le `BeforeUpdate` d'un contrôle empêche l'utilisateur de quitter le champ. Mal géré (par exemple sur une condition qui ne peut être satisfaite), il **bloque** l'utilisateur dans le contrôle. Toujours laisser une issue, et préférer souvent la validation au niveau du **formulaire** (chapitre 8.10), moins risquée.

---

## 8.12.5. Pattern : confirmer avant la fermeture (Form Unload)

L'événement `Unload` étant annulable (chapitre 8.7), il permet de demander confirmation avant de fermer, et d'annuler la fermeture si l'utilisateur se ravise :

```vba
Private Sub Form_Unload(Cancel As Integer)
    If Me.Dirty Then
        If MsgBox("Des modifications ne sont pas enregistrées. Fermer quand même ?", _
                  vbYesNo + vbQuestion) = vbNo Then
            Cancel = True        ' la fermeture est annulée
        End If
    End If
End Sub
```

---

## 8.12.6. Pattern : empêcher une suppression (Form Delete)

L'événement `Delete` se déclenche **avant** la suppression de chaque enregistrement ; `Cancel = True` la bloque — par exemple pour protéger un enregistrement selon son statut :

```vba
Private Sub Form_Delete(Cancel As Integer)
    If Me.Statut.Value = "Validé" Then
        MsgBox "Un enregistrement validé ne peut pas être supprimé.", vbExclamation
        Cancel = True
    End If
End Sub
```

---

## 8.12.7. Pattern : annuler l'ouverture (Open) — et l'erreur 2501

Annuler dans l'événement `Open` empêche l'ouverture du formulaire ou de l'état (chapitres 7.2, 7.9 et 8.7) :

```vba
Private Sub Form_Open(Cancel As Integer)
    If Not CurrentProject.AllForms("F_Criteres").IsLoaded Then
        MsgBox "Ouvrez d'abord le formulaire de critères.", vbExclamation
        Cancel = True
    End If
End Sub
```

**Conséquence cruciale** : lorsque l'objet a été ouvert par `DoCmd.OpenForm`/`DoCmd.OpenReport`, l'annulation **remonte** au code appelant sous la forme de l'**erreur 2501**. Sans gestion, l'utilisateur voit ce message technique. Le code appelant doit donc l'intercepter (comme pour les états vides, chapitre 7.9) :

```vba
Sub OuvrirFormulaire()
    On Error GoTo Gestion
    DoCmd.OpenForm "F_X"
    Exit Sub

Gestion:
    If Err.Number = 2501 Then
        ' ouverture annulée volontairement : ce n'est pas une vraie erreur
    Else
        MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical
    End If
End Sub
```

La gestion structurée des erreurs est traitée au chapitre 13.

---

## 8.12.8. Ce que Cancel ne peut pas annuler

Savoir ce qui **n'est pas** annulable est aussi important :

- **`Current`** (formulaire) n'est **pas** annulable : on ne peut pas bloquer l'arrivée sur un enregistrement via `Current`. Pour empêcher de quitter une ligne invalide, on bloque la **sauvegarde** dans `BeforeUpdate`.
- **`Change`** (contrôle onglet) n'est **pas** annulable : impossible de bloquer un changement de page par cet événement (chapitre 8.4) ; recourir à la navigation par boutons.
- **`AfterUpdate`**, **`Load`**, **`Close`** ne sont **pas** annulables : l'action a déjà eu lieu (after) ou est définitive.

En résumé : on **prévient** une action via son événement *Before*/*Exit*/*Open*/*Unload* (annulable), jamais via son événement *After* ou via `Current`/`Change`.

---

## 8.12.9. Cancel et les autres mécanismes (Response, KeyAscii, Undo)

`Cancel` n'est pas le seul moyen d'infléchir un comportement ; il ne faut pas le confondre avec des mécanismes voisins :

| Mécanisme | Où | Effet |
|---|---|---|
| **`Cancel = True`** | Événements *Before*, `Exit`, `Open`, `Unload`, `Delete`… | Annule l'action par défaut |
| **`Response`** | `Error`, `NotInList` | Indique à Access **comment** réagir (chapitres 8.8 et 8.2) |
| **`KeyAscii = 0`** | `KeyPress` | Rejette un caractère saisi (chapitre 8.9) |
| **`Me.Undo`** | À tout moment | Annule les modifications **en cours** d'un enregistrement (chapitre 8.8) |

Chacun répond à un besoin distinct : bloquer une action (`Cancel`), orienter une réponse (`Response`), filtrer une frappe (`KeyAscii`), ou défaire des saisies (`Undo`).

---

## 8.12.10. Pièges et bonnes pratiques

- **Toujours informer l'utilisateur** lorsqu'on annule une action : un message clair explique *pourquoi* (chapitre 8.10).
- **Gérer l'erreur 2501** dans le code appelant dès qu'on annule un `Open` (chapitres 7.9 et 13).
- **Éviter de piéger l'utilisateur** avec un `Cancel` dans `Exit`/`BeforeUpdate` de contrôle : laisser toujours une issue, préférer la validation au formulaire (chapitre 8.10).
- **Ne pas tenter d'annuler un événement non annulable** (`Current`, `Change` d'onglet, `AfterUpdate`) : utiliser l'événement *Before* approprié.
- **`Cancel = True` modifie l'état** : la ligne reste « sale » (`BeforeUpdate`), le formulaire reste ouvert (`Unload`), le focus reste sur le contrôle (`Exit`) — en tenir compte.
- **Distinguer `Cancel`, `Response`, `KeyAscii` et `Undo`** : quatre mécanismes pour quatre besoins.

---

## 8.12.11. Récapitulatif

- L'argument **`Cancel`** de certains événements permet, en l'affectant à `True`, d'**annuler l'action par défaut** (ouverture, enregistrement, fermeture, suppression, sortie d'un contrôle…).
- Les événements annulables incluent `Open`, `BeforeUpdate` (formulaire et contrôle), `Delete`, `Unload`, `Undo`, `Exit`, `DblClick`, ainsi que `NoData`, `Format`, `Print` côté états.
- Les **patterns courants** : valider avant enregistrement (`BeforeUpdate` du formulaire), valider un champ en gardant le focus (`Exit`/`BeforeUpdate` de contrôle, avec prudence), confirmer avant fermeture (`Unload`), empêcher une suppression (`Delete`), annuler une ouverture (`Open`).
- Annuler un `Open` ouvert par `DoCmd` provoque l'**erreur 2501** côté appelant, à **intercepter** (chapitres 7.9 et 13).
- Certains événements ne sont **pas** annulables (`Current`, `Change` d'onglet, `AfterUpdate`, `Load`, `Close`) : on prévient une action via son événement *Before*, jamais *After*.
- `Cancel` se distingue des mécanismes voisins **`Response`**, **`KeyAscii`** et **`Me.Undo`**, chacun répondant à un besoin précis.

---

> ✅ **Fin du chapitre 8.** Le modèle des contrôles et des événements est désormais complet : la hiérarchie et les types de contrôles (8.1–8.5), leur liaison aux données (8.6), puis le cycle de vie et les événements (8.7–8.9), la validation (8.10), la manipulation dynamique (8.11) et l'annulation d'actions (8.12). Cette progression « du statique au dynamique » donne les clés de la programmation événementielle dans Access.

⏭️ [9. DAO – Data Access Objects](/09-dao-data-access-objects/README.md)
