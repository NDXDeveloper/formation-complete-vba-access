🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8.10. Validation des données et messages d'erreur personnalisés

La **validation** garantit la qualité et l'intégrité des données saisies. Elle ne se limite pas à bloquer les erreurs : bien menée, elle **guide** l'utilisateur par des messages clairs indiquant *quoi* corriger et *comment*. Cette section consolide les approches de validation évoquées tout au long du chapitre — `BeforeUpdate` de contrôle et de formulaire (chapitres 8.8 et 8.9), restriction à la frappe (chapitre 8.9) — et les organise en une stratégie cohérente, complétée par la gestion des messages d'erreur.

---

## 8.10.1. Les niveaux de validation (défense en profondeur)

La validation s'organise sur plusieurs niveaux, du plus proche de l'interface au plus proche des données :

| Niveau | Mécanisme | Caractéristique |
|---|---|---|
| **Saisie** | Restriction à la frappe (`KeyPress`), masque de saisie | Empêche l'entrée invalide |
| **Contrôle** | `BeforeUpdate` du contrôle (VBA) | Retour immédiat sur un champ |
| **Formulaire** | `BeforeUpdate` du formulaire (VBA) | Validation **autoritaire** avant enregistrement |
| **Table** | `ValidationRule`, `Required`, index uniques | **Dernier rempart**, imposé par le moteur |

Le principe est la **défense en profondeur** : l'interface valide pour le confort et la pédagogie, la **table** garantit l'intégrité même si l'interface est contournée (import, modification directe).

---

## 8.10.2. Validation déclarative sans code (InputMask, ValidationRule)

Avant d'écrire du code, on exploite les validations **déclaratives**, qui ne demandent aucune programmation :

- **`InputMask`** (masque de saisie) : impose un format pendant la frappe (par exemple `00000` pour cinq chiffres).
- **`ValidationRule`** + **`ValidationText`** : une règle simple et le message affiché en cas de violation, applicables à un contrôle ou à un champ.
- **`Required`**, **`Format`**, **`DecimalPlaces`** : obligation de saisie, mise en forme.

```
' Propriétés d'un champ ou d'un contrôle :
'   ValidationRule : >=0
'   ValidationText : La valeur doit être positive.
```

On réserve la validation par code aux règles que le déclaratif ne peut exprimer : contrôles **inter-champs**, règles **métier** complexes, messages **contextuels**.

---

## 8.10.3. Validation par code : où la placer ?

Deux événements portent la validation par code (chapitres 8.8 et 8.9) :

- le **`BeforeUpdate` du contrôle** : retour **immédiat** sur un champ, mais **contournable** si l'utilisateur ne visite pas ce contrôle ;
- le **`BeforeUpdate` du formulaire** : validation **systématique** avant tout enregistrement — la place **fiable**.

La bonne pratique combine les deux : feedback immédiat au niveau du contrôle, **validation autoritaire** au niveau du formulaire.

---

## 8.10.4. Validation autoritaire : le BeforeUpdate du formulaire

C'est le **lieu de référence** de la validation : il se déclenche avant chaque écriture, quel que soit le chemin emprunté par l'utilisateur. On y vérifie champs obligatoires, cohérence inter-champs et règles métier, et l'on **annule** l'enregistrement si nécessaire :

```vba
Private Sub Form_BeforeUpdate(Cancel As Integer)
    If Not FormulaireValide() Then Cancel = True
End Sub
```

La logique de contrôle est isolée dans une **fonction dédiée**, plus lisible et réutilisable (section 8.10.7).

---

## 8.10.5. Validation d'un champ : le BeforeUpdate du contrôle

Pour un retour immédiat à la sortie d'un champ, on valide dans le `BeforeUpdate` du contrôle ; `Cancel = True` **conserve le focus** sur le contrôle :

```vba
Private Sub txtCodePostal_BeforeUpdate(Cancel As Integer)
    If Len(Nz(Me.txtCodePostal.Value, "")) <> 5 Then
        MsgBox "Le code postal doit comporter 5 chiffres.", vbExclamation
        Cancel = True
    End If
End Sub
```

> Cette validation reste **complémentaire** de celle du formulaire : un champ jamais visité ne déclenche pas son `BeforeUpdate`. Ne jamais s'y fier seul pour une règle critique.

---

## 8.10.6. Prévenir plutôt que rejeter : restriction à la saisie

Mieux vaut souvent **empêcher** une saisie invalide que la rejeter après coup. La restriction à la frappe via `KeyPress` (chapitre 8.9) y pourvoit — par exemple n'autoriser que des chiffres :

```vba
Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> 8 Then
        KeyAscii = 0    ' caractère rejeté
    End If
End Sub
```

Associée à un masque de saisie, cette prévention réduit le besoin de messages d'erreur.

---

## 8.10.7. Collecter et présenter plusieurs erreurs ensemble

Plutôt que d'interrompre l'utilisateur dès la **première** erreur par une succession de boîtes de message, il est préférable de **rassembler** tous les problèmes et de les présenter **en une fois**. La fonction de validation accumule les anomalies dans une chaîne :

```vba
Private Function FormulaireValide() As Boolean
    Dim strErreurs As String

    If Len(Nz(Me.txtNom.Value, "")) = 0 Then
        strErreurs = strErreurs & "- Le nom est obligatoire." & vbCrLf
    End If

    If Not IsNull(Me.txtEmail.Value) And InStr(Nz(Me.txtEmail.Value, ""), "@") = 0 Then
        strErreurs = strErreurs & "- L'adresse e-mail est invalide." & vbCrLf
    End If

    If Not IsNull(Me.txtFin.Value) And Not IsNull(Me.txtDebut.Value) Then
        If Me.txtFin.Value < Me.txtDebut.Value Then
            strErreurs = strErreurs & "- La date de fin précède la date de début." & vbCrLf
        End If
    End If

    If Len(strErreurs) > 0 Then
        MsgBox "Veuillez corriger les points suivants :" & vbCrLf & vbCrLf & strErreurs, _
               vbExclamation, "Validation"
        FormulaireValide = False
    Else
        FormulaireValide = True
    End If
End Function
```

Cette approche offre une bien meilleure expérience qu'une cascade de messages successifs.

---

## 8.10.8. Messages d'erreur personnalisés : bonnes pratiques

Un bon message d'erreur :

- **précise ce qui ne va pas** et, si possible, **comment le corriger** (« Le code postal doit comporter 5 chiffres » plutôt que « Saisie invalide ») ;
- **positionne le focus** sur le champ fautif (`Me.txtChamp.SetFocus`) pour faciliter la correction ;
- emploie l'**icône adaptée** (`vbExclamation` pour un avertissement, `vbCritical` pour une erreur grave) ;
- **remplace** les messages cryptiques du moteur par un libellé compréhensible (section 8.10.10).

```vba
If Len(Nz(Me.txtNom.Value, "")) = 0 Then
    MsgBox "Le nom du client est obligatoire.", vbExclamation, "Champ requis"
    Me.txtNom.SetFocus
    Cancel = True
End If
```

---

## 8.10.9. Le retour visuel (surlignage des champs)

Le message peut être complété par un **retour visuel** signalant les champs concernés, par exemple en modifiant leur couleur de fond :

```vba
Me.txtNom.BackColor = RGB(255, 230, 230)   ' fond rosé pour un champ en erreur
```

Pour un rendu cohérent et automatique, on s'appuie souvent sur la **mise en forme conditionnelle** (chapitre 17.8), qui colore un contrôle selon une règle sans code répétitif.

---

## 8.10.10. Intercepter les erreurs du moteur (événement Error)

Même validés en amont, certains contrôles ne se déclenchent qu'au niveau du **moteur** : doublon sur un index unique, champ requis défini sur la table, violation de `ValidationRule`. Le moteur affiche alors un message technique. L'événement **`Error` du formulaire** (chapitre 8.8) permet de le **remplacer** par un message clair :

```vba
Private Sub Form_Error(DataErr As Integer, Response As Integer)
    Select Case DataErr
        Case 3022     ' doublon sur un index unique
            MsgBox "Cette valeur existe déjà ; elle doit être unique.", vbExclamation
            Response = acDataErrContinue
        Case 3314     ' champ obligatoire non renseigné
            MsgBox "Un champ obligatoire n'est pas renseigné.", vbExclamation
            Response = acDataErrContinue
        Case Else
            Response = acDataErrDisplay   ' message standard pour les autres cas
    End Select
End Sub
```

`Response = acDataErrContinue` supprime le message standard (on l'a remplacé) ; `acDataErrDisplay` le laisse s'afficher. C'est le complément indispensable à la validation par code lorsque les contraintes de table entrent en jeu.

---

## 8.10.11. La stratégie complète

En synthèse, une validation robuste combine les niveaux :

1. **Prévenir** à la saisie (`InputMask`, `KeyPress`) ce qui peut l'être ;
2. **Signaler immédiatement** au niveau du contrôle (`BeforeUpdate` de contrôle) ;
3. **Valider de façon autoritaire** au niveau du formulaire (`BeforeUpdate` de formulaire), avec collecte des erreurs ;
4. **Garantir l'intégrité** par les contraintes de **table** (chapitres 12.5 et 14.6), dernier rempart ;
5. **Traduire** les erreurs du moteur en messages clairs via l'événement **`Error`**.

C'est la défense en profondeur de la section 8.10.1, appliquée concrètement.

---

## 8.10.12. Pièges et bonnes pratiques

- **Ne pas se fier aux seuls événements de contrôle** : contournables ; la validation fiable est au **formulaire** (chapitres 8.8 et 8.9).
- **Ne pas se fier à la seule validation d'interface** : un import ou une saisie directe la contourne ; doubler par des **contraintes de table** (chapitres 12.5 et 14.6).
- **Privilégier la validation déclarative** (`InputMask`, `ValidationRule`) quand elle suffit.
- **Collecter les erreurs** et les présenter en une fois, plutôt qu'une cascade de `MsgBox`.
- **Rédiger des messages précis** et **positionner le focus** sur le champ fautif.
- **Protéger contre `Null`** dans les tests (`Nz`, `IsNull`).
- **`Cancel = True` sur `Exit`/`BeforeUpdate` de contrôle peut piéger l'utilisateur** : à manier avec discernement (chapitre 8.9).
- **Ne pas valider dans `Current`** : l'enregistrement n'y est pas encore modifié (chapitre 8.8).
- **Traduire les erreurs du moteur** via l'événement `Error` pour éviter les messages techniques.

---

## 8.10.13. Récapitulatif

- La validation s'organise en **défense en profondeur** : saisie, contrôle, formulaire et **table** — chaque niveau renforçant les autres.
- Les validations **déclaratives** (`InputMask`, `ValidationRule`/`ValidationText`, `Required`) couvrent les cas simples sans code ; le code prend le relais pour les règles inter-champs et métier.
- La validation **autoritaire** se place dans le **`BeforeUpdate` du formulaire** (systématique, annulable) ; le `BeforeUpdate` de **contrôle** offre un retour immédiat mais **contournable**.
- Mieux vaut **prévenir** à la frappe (`KeyPress`, masque) que rejeter, et **collecter** toutes les erreurs pour les présenter en une fois.
- Un bon **message** est précis, positionne le focus et emploie l'icône adaptée ; un **retour visuel** (surlignage, chapitre 17.8) le complète.
- L'événement **`Error`** (chapitre 8.8) traduit les erreurs du **moteur** (doublon, champ requis, règle de validation) en messages clairs ; les contraintes de **table** (chapitres 12.5 et 14.6) restent le dernier rempart de l'intégrité.

⏭️ [8.11. Manipulation dynamique des contrôles (création, visibilité, activation)](/08-controles-evenements/11-manipulation-dynamique-controles.md)
