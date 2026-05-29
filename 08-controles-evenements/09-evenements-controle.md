🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8.9. Événements de contrôle (Change, GotFocus, LostFocus, Click, DblClick)

Là où les événements de formulaire raisonnent au niveau de l'enregistrement (chapitre 8.8), les événements de **contrôle** concernent un élément isolé : la prise et la perte de focus, chaque frappe, les clics, les touches. Ils permettent une interaction fine — recherche au fil de la saisie, validation d'un champ, raccourcis clavier, actions au double-clic.

Cette section organise ces événements par familles, en détaillant les cinq plus courants (`Change`, `GotFocus`, `LostFocus`, `Click`, `DblClick`) et en clarifiant des distinctions souvent mal comprises.

---

## 8.9.1. Les familles d'événements de contrôle

| Famille | Événements | Rôle |
|---|---|---|
| **Focus** | `Enter`, `GotFocus`, `Exit`, `LostFocus` | Prise et perte du focus |
| **Modification** | `Change`, `BeforeUpdate`, `AfterUpdate` (contrôle) | Édition et validation de la valeur |
| **Souris** | `Click`, `DblClick`, `MouseDown`, `MouseMove`, `MouseUp` | Interaction à la souris |
| **Clavier** | `KeyDown`, `KeyPress`, `KeyUp` | Interaction au clavier |

---

## 8.9.2. Les événements de focus : Enter, GotFocus, Exit, LostFocus

Quatre événements gèrent le déplacement du focus :

- **`Enter`** : juste avant que le contrôle ne reçoive le focus ;
- **`GotFocus`** : lorsque le contrôle reçoit le focus ;
- **`Exit(Cancel)`** : juste avant que le contrôle ne perde le focus — **annulable** ;
- **`LostFocus`** : lorsque le contrôle perd le focus.

L'ordre est `Enter → GotFocus` à l'arrivée, `Exit → LostFocus` au départ.

### Enter/Exit ou GotFocus/LostFocus : la distinction

La différence est subtile mais importante :

- **`Enter`** et **`Exit`** ne se déclenchent que lors d'un déplacement de focus **entre contrôles d'un même formulaire**.
- **`GotFocus`** et **`LostFocus`** se déclenchent **chaque fois** que le contrôle obtient ou perd réellement le focus — y compris lorsqu'on **bascule entre fenêtres** ou applications.

Ainsi, basculer vers une autre application puis revenir déclenche `LostFocus`/`GotFocus`, mais **pas** `Exit`/`Enter`. Pour une logique d'entrée/sortie de champ, on privilégie donc `Enter`/`Exit`.

```vba
' Sélectionner tout le texte à l'entrée, pour faciliter la réécriture
Private Sub txtRecherche_Enter()
    Me.txtRecherche.SelStart = 0
    Me.txtRecherche.SelLength = Len(Nz(Me.txtRecherche.Value, ""))
End Sub
```

`Exit` étant annulable, il permet une validation **avant de quitter** un contrôle (avec la réserve de la section 8.9.10).

---

## 8.9.3. Change — réagir à chaque modification

`Change` se déclenche **à chaque modification** du contenu d'un contrôle : pour une zone de texte, à **chaque frappe** (caractère ajouté ou supprimé). C'est l'événement de l'interaction « en temps réel ».

Dans `Change`, on lit la **saisie en cours** via la propriété **`Text`** (et non `Value`, qui ne sera mise à jour qu'à la validation — chapitre 8.2). Le cas d'usage emblématique est la **recherche au fil de la frappe** :

```vba
Private Sub txtRecherche_Change()
    Dim strSaisie As String
    strSaisie = Me.txtRecherche.Text          ' saisie en cours
    Me.Filter = "Nom LIKE '*" & Replace(strSaisie, "'", "''") & "*'"
    Me.FilterOn = (Len(strSaisie) > 0)
End Sub
```

> **`Change` ≠ `AfterUpdate`** : `Change` se déclenche à **chaque frappe** (valeur non validée), `AfterUpdate` **une fois** à la sortie du contrôle (valeur validée). De plus, affecter une valeur **par code** (`Me.ctl.Value = …`) ne déclenche **pas** `Change`.

---

## 8.9.4. BeforeUpdate et AfterUpdate au niveau contrôle

Ces deux noms, déjà rencontrés au niveau formulaire (chapitre 8.8), existent aussi au niveau **contrôle** :

- **`BeforeUpdate(Cancel)`** du contrôle : juste **avant** que la valeur modifiée ne soit validée dans le tampon de l'enregistrement — **annulable** ;
- **`AfterUpdate`** du contrôle : **après** cette validation.

Ils ne se déclenchent **que si** le contrôle a été **modifié**, au moment où il perd le focus.

```vba
Private Sub txtCodePostal_BeforeUpdate(Cancel As Integer)
    If Len(Nz(Me.txtCodePostal.Value, "")) <> 5 Then
        MsgBox "Le code postal doit comporter 5 chiffres.", vbExclamation
        Cancel = True
    End If
End Sub
```

> **Quel niveau pour valider ?** Le `BeforeUpdate` **du contrôle** offre un retour immédiat sur un champ, mais peut être **contourné** si l'utilisateur ne visite jamais ce contrôle. Pour une validation **fiable et systématique**, le `BeforeUpdate` **du formulaire** (chapitre 8.8) reste la référence. On combine souvent les deux : feedback immédiat au contrôle, validation autoritaire au formulaire (chapitre 8.10).

---

## 8.9.5. Click et DblClick

- **`Click`** se déclenche au clic. C'est **l'**événement des **boutons de commande** (`cmdX_Click`), mais il se déclenche aussi sur d'autres contrôles (sélection d'un élément, par exemple).
- **`DblClick(Cancel)`** se déclenche au double-clic ; il est **annulable**.

Pour un bouton, `Click` concentre l'action :

```vba
Private Sub cmdEnregistrer_Click()
    If Me.Dirty Then Me.Dirty = False
End Sub
```

Sur une **liste**, le double-clic sert classiquement à « ouvrir » l'élément sélectionné :

```vba
Private Sub lstClients_DblClick(Cancel As Integer)
    If Not IsNull(Me.lstClients.Value) Then
        DoCmd.OpenForm "F_Client", , , "ClientID=" & Me.lstClients.Value
    End If
End Sub
```

---

## 8.9.6. Les événements clavier (KeyDown, KeyPress, KeyUp)

Trois événements captent le clavier :

- **`KeyDown(KeyCode, Shift)`** et **`KeyUp(KeyCode, Shift)`** : touches **physiques**, y compris non-caractères (touches de fonction, flèches). `KeyCode` donne la touche, `Shift` l'état des touches Maj/Ctrl/Alt.
- **`KeyPress(KeyAscii)`** : touches de **caractère** uniquement. `KeyAscii` est le code ASCII du caractère ; l'affecter à `0` **annule** la frappe.

`KeyPress` permet ainsi de **restreindre** la saisie :

```vba
Private Sub txtCode_KeyPress(KeyAscii As Integer)
    ' N'autoriser que les chiffres (et la touche Retour arrière, code 8)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> 8 Then
        KeyAscii = 0    ' caractère rejeté
    End If
End Sub
```

> La propriété **`KeyPreview`** du formulaire, à `True`, fait recevoir les événements clavier **par le formulaire avant les contrôles** — pratique pour gérer des raccourcis à l'échelle du formulaire.

---

## 8.9.7. Les événements souris (MouseDown, MouseMove, MouseUp)

Pour des interactions plus fines, les événements souris fournissent le bouton pressé, l'état des touches et les coordonnées :

```vba
Private Sub txtNom_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
    If Button = acRightButton Then
        ' déclencher un menu contextuel personnalisé (chapitre 17.2)
    End If
End Sub
```

Ils servent notamment à détecter le **clic droit** (`Button = acRightButton`) pour les menus contextuels, ou à gérer des comportements de survol.

---

## 8.9.8. L'ordre des événements dans un contrôle

En consolidant (chapitre 8.7) :

```
Entrée dans le contrôle :        Enter → GotFocus
Frappe d'un caractère :          KeyDown → KeyPress → Change → KeyUp
Sortie d'un contrôle modifié :   BeforeUpdate → AfterUpdate → Exit → LostFocus
```

Cet enchaînement explique, par exemple, pourquoi une validation placée dans `AfterUpdate` s'exécute avant la perte effective du focus (`LostFocus`).

---

## 8.9.9. Patterns courants

- **`Change`** : recherche/filtre au fil de la frappe, compteur de caractères, activation conditionnelle.
- **`Enter`** : sélectionner tout le texte pour faciliter la réécriture.
- **`Exit`** (annulable) : validation d'un champ avant de le quitter.
- **`Click`** : action d'un bouton.
- **`DblClick`** : ouvrir le détail de l'élément sélectionné d'une liste.
- **`KeyPress`** : restreindre la saisie (chiffres seuls, par exemple).

---

## 8.9.10. Pièges et bonnes pratiques

- **`Change` ≠ `AfterUpdate`** : l'un à chaque frappe (non validé), l'autre une fois à la sortie (validé) ; dans `Change`, lire **`Text`**, pas `Value`.
- **`Enter`/`Exit` ne réagissent qu'aux déplacements intra-formulaire** ; `GotFocus`/`LostFocus` se déclenchent aussi au changement de fenêtre.
- **`Exit` annulable peut piéger l'utilisateur** : un `Cancel = True` mal géré l'empêche de quitter le champ ; à utiliser avec discernement.
- **La validation de contrôle est contournable** : pour une validation fiable, préférer le `BeforeUpdate` du **formulaire** (chapitres 8.8 et 8.10).
- **`Click` se déclenche sur de nombreux contrôles**, pas seulement les boutons.
- **`KeyPress` ne voit que les caractères** ; pour les touches de fonction ou les flèches, utiliser `KeyDown`. Affecter `KeyAscii = 0` rejette un caractère.
- **Une affectation par code ne déclenche pas `Change`/`AfterUpdate`** : en tenir compte si l'on attend une réaction.
- **`KeyPreview`** permet de centraliser la gestion clavier au niveau du formulaire.

---

## 8.9.11. Récapitulatif

- Les événements de **contrôle** se répartissent en quatre familles : **focus** (`Enter`/`GotFocus`/`Exit`/`LostFocus`), **modification** (`Change`, `BeforeUpdate`/`AfterUpdate` du contrôle), **souris** (`Click`, `DblClick`, `MouseDown`…) et **clavier** (`KeyDown`, `KeyPress`, `KeyUp`).
- `Enter`/`Exit` ne réagissent qu'aux déplacements **entre contrôles d'un même formulaire** ; `GotFocus`/`LostFocus` réagissent à tout gain/perte de focus.
- **`Change`** se déclenche à **chaque frappe** (lire `Text`, chapitre 8.2), à la différence d'`AfterUpdate` (une fois, valeur validée).
- Les `BeforeUpdate`/`AfterUpdate` de **contrôle** valident un **champ** (le `BeforeUpdate` étant annulable), mais sont **contournables** : la validation fiable se place au **formulaire** (chapitres 8.8 et 8.10).
- **`Click`** anime les boutons ; **`DblClick`** sert à ouvrir l'élément d'une liste ; les événements **clavier** permettent de restreindre la saisie (`KeyAscii = 0`) et, via `KeyPreview`, une gestion centralisée.
- L'**ordre** des événements d'un contrôle (entrée, frappe, sortie) s'inscrit dans le cycle de vie global du chapitre 8.7.

⏭️ [8.10. Validation des données et messages d'erreur personnalisés](/08-controles-evenements/10-validation-donnees.md)
