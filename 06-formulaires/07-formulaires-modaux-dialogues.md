🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 6.7. Formulaires modaux et boîtes de dialogue personnalisées

Les fonctions natives `MsgBox` et `InputBox` suffisent pour un message simple ou une saisie textuelle unique. Dès que l'on a besoin de plusieurs champs, d'une liste déroulante, de validations ou d'une mise en forme soignée, il faut concevoir une **boîte de dialogue personnalisée** : un formulaire ordinaire que l'on transforme en fenêtre modale et que l'on apprend à interroger une fois refermé.

Cette section explique les propriétés qui définissent un comportement modal, le mécanisme de suspension du code propre au mode dialogue, et surtout le patron incontournable qui permet à une boîte de dialogue de **renvoyer une valeur** au code appelant.

---

## 6.7.1. Modal ou non modal : le concept

Un formulaire **modal** monopolise l'interaction : tant qu'il est ouvert, l'utilisateur ne peut pas accéder aux autres objets de l'application. C'est le comportement attendu d'une boîte de dialogue qui exige une réponse avant de poursuivre.

Un formulaire **non modal** (le comportement par défaut) laisse l'utilisateur naviguer librement entre plusieurs fenêtres ouvertes.

La modalité combine deux notions distinctes : le **verrouillage du focus** (l'utilisateur ne peut pas cliquer ailleurs) et la **suspension du code** (la procédure appelante s'interrompt). Ces deux comportements ne sont pas pilotés par les mêmes leviers, comme le détaillent les sections suivantes.

---

## 6.7.2. Les propriétés qui définissent une boîte de dialogue

Trois propriétés du formulaire, réglables en mode création, donnent l'apparence et le comportement d'une boîte de dialogue :

| Propriété | Valeur dialogue | Effet |
|---|---|---|
| `Modal` | `True` | L'utilisateur doit fermer le formulaire avant d'accéder aux autres objets (verrouille le focus) |
| `PopUp` | `True` | Le formulaire flotte au-dessus des autres fenêtres, hors de la zone d'affichage normale |
| `BorderStyle` | `Dialog` (3) | Bordure fine, barre de titre avec seul bouton de fermeture, fenêtre non redimensionnable |

La combinaison **`Modal = True` + `PopUp = True` + `BorderStyle = Dialog`** produit l'aspect classique d'une boîte de dialogue.

> **Distinction importante** : régler ces propriétés en mode création verrouille bien le focus, mais **ne suspend pas** le code appelant. Pour obtenir la suspension — indispensable au patron de renvoi de valeur — il faut ouvrir le formulaire avec l'argument `acDialog`, présenté ci-après.

---

## 6.7.3. Ouvrir un formulaire en mode dialogue (acDialog)

Le sixième argument de `DoCmd.OpenForm`, `WindowMode`, accepte la valeur `acDialog` :

```vba
DoCmd.OpenForm "F_Dialogue", , , , , acDialog
```

L'ouverture en `acDialog` produit deux effets :

1. Elle force le formulaire en mode **modal** et **pop-up** pour cette ouverture (quelles que soient ses propriétés en création).
2. Surtout, elle **suspend l'exécution** de la procédure appelante : le code s'arrête sur la ligne `OpenForm` et ne reprend que lorsque le formulaire est **fermé** ou **masqué** (`Visible = False`).

C'est cette suspension qui rend possible le dialogue interactif : le code appelant attend la réponse de l'utilisateur avant de continuer.

```vba
Debug.Print "Avant"
DoCmd.OpenForm "F_Dialogue", , , , , acDialog
' --- le code est suspendu ici tant que F_Dialogue est ouvert et visible ---
Debug.Print "Après"   ' s'exécute seulement après fermeture ou masquage du dialogue
```

---

## 6.7.4. Le patron « renvoyer une valeur » : masquer plutôt que fermer

Puisque le code reprend aussi bien à la **fermeture** qu'au **masquage**, on exploite cette différence pour transmettre un résultat :

- Le bouton **OK** ne ferme pas le formulaire ; il le **masque** (`Me.Visible = False`). Le code appelant reprend alors que le formulaire est toujours chargé en mémoire — ses contrôles restent lisibles.
- Le bouton **Annuler** (ou la croix de fermeture) **ferme** réellement le formulaire. Le code appelant reprend, mais le formulaire n'existe plus.

Code à placer dans le formulaire de dialogue :

```vba
' --- Bouton OK : masquer (surtout pas fermer) ---
Private Sub cmdOK_Click()
    Me.Visible = False
End Sub

' --- Bouton Annuler : fermer réellement ---
Private Sub cmdAnnuler_Click()
    DoCmd.Close acForm, Me.Name
End Sub
```

Le code appelant peut alors lire les valeurs du formulaire **masqué** avant de le fermer lui-même :

```vba
DoCmd.OpenForm "F_Dialogue", , , , , acDialog
' Le code reprend ici (formulaire masqué ou fermé)

If CurrentProject.AllForms("F_Dialogue").IsLoaded Then
    ' OK a été cliqué → le formulaire est masqué, ses contrôles sont lisibles
    MsgBox "Valeur saisie : " & Forms("F_Dialogue").txtValeur.Value
    DoCmd.Close acForm, "F_Dialogue"   ' fermeture explicite indispensable
Else
    ' Annulé → le formulaire a été fermé
    MsgBox "Saisie annulée."
End If
```

---

## 6.7.5. Détecter OK ou Annuler

La distinction OK / Annuler repose, dans le patron ci-dessus, sur l'**état de chargement** du formulaire au retour : masqué (donc encore chargé) pour OK, fermé (donc déchargé) pour Annuler.

La propriété `IsLoaded` de la collection `AllForms` (chapitre 4.4) constitue le test idiomatique :

```vba
If CurrentProject.AllForms("F_Dialogue").IsLoaded Then
    ' OK
Else
    ' Annulé
End If
```

> La croix de fermeture de la barre de titre déclenche elle aussi la fermeture du formulaire : elle est donc naturellement assimilée à un « Annuler », sans code supplémentaire.

D'autres approches existent pour transmettre le résultat : affecter une **`TempVar`** dans le dialogue puis la lire au retour (chapitre 15.7), ou exposer une **propriété publique** dans le module de classe du dialogue (chapitres 6.9 et 16.2). Le test sur `IsLoaded` reste cependant le plus simple pour les cas courants.

---

## 6.7.6. Exemple complet : une fonction de saisie réutilisable

L'ensemble du patron se synthétise élégamment dans une **fonction** qui ouvre le dialogue, en lit le résultat et le renvoie. L'invite est transmise via `OpenArgs` (chapitre 6.6) :

```vba
Public Function DemanderValeur(ByVal strInvite As String) As Variant
    ' Ouvre F_Dialogue en mode modal et attend la réponse de l'utilisateur.
    ' Renvoie la valeur saisie, ou Null si l'utilisateur annule.

    DoCmd.OpenForm "F_Dialogue", , , , , acDialog, strInvite

    If CurrentProject.AllForms("F_Dialogue").IsLoaded Then
        DemanderValeur = Forms("F_Dialogue").txtValeur.Value
        DoCmd.Close acForm, "F_Dialogue"
    Else
        DemanderValeur = Null
    End If
End Function
```

Appel depuis n'importe où dans l'application :

```vba
Dim varReponse As Variant

varReponse = DemanderValeur("Saisissez le code du dossier :")

If IsNull(varReponse) Then
    MsgBox "Opération annulée."
Else
    MsgBox "Vous avez saisi : " & varReponse
End If
```

Le formulaire de dialogue affiche l'invite reçue dans son événement `Open` :

```vba
Private Sub Form_Open(Cancel As Integer)
    If Not IsNull(Me.OpenArgs) Then
        Me.lblInvite.Caption = Me.OpenArgs
    End If
End Sub
```

---

## 6.7.7. Formulaires modaux sans patron de dialogue

Toutes les fenêtres modales ne servent pas à renvoyer une valeur. Un formulaire de **paramètres**, d'**à propos** ou un **assistant** peut simplement exiger d'être refermé avant de poursuivre, sans interrogation de ses contrôles au retour.

Dans ce cas, il suffit de régler `Modal = True` (et éventuellement `PopUp = True`) en mode création, ou d'ouvrir le formulaire normalement après l'avoir conçu comme modal. L'utilisateur reste cantonné à cette fenêtre, mais le code appelant n'a pas besoin d'être suspendu : on ne recherche pas de résultat.

---

## 6.7.8. Boîte de dialogue personnalisée ou MsgBox / InputBox ?

| Besoin | Solution adaptée |
|---|---|
| Message + boutons standards (OK/Annuler, Oui/Non) | `MsgBox` |
| Une unique saisie textuelle simple | `InputBox` |
| Plusieurs champs, listes déroulantes, validations, valeurs par défaut, mise en forme | **Boîte de dialogue personnalisée** |

`MsgBox` renvoie directement un code de bouton (`vbYes`, `vbNo`, `vbCancel`…) et reste imbattable pour une confirmation. `InputBox` ne gère qu'une saisie texte, sans contrôle de validité. Dès que l'entrée est **structurée**, le formulaire personnalisé s'impose : il offre une expérience cohérente et toute la puissance des contrôles Access.

---

## 6.7.9. Pièges et bonnes pratiques

- **Pour renvoyer une valeur, masquer et non fermer** : le bouton OK doit exécuter `Me.Visible = False`. Un `DoCmd.Close` détruirait le formulaire et rendrait ses contrôles illisibles.
- **Toujours fermer le formulaire masqué après lecture** : une fois les valeurs récupérées, exécuter `DoCmd.Close`. Sinon le formulaire reste chargé (invisible) en mémoire et persiste dans `AllForms`.
- **`acDialog` est requis pour suspendre le code** : régler la seule propriété `Modal` verrouille le focus mais ne met pas la procédure en pause.
- **Ne pas confondre `Modal` et `PopUp`** : `Modal` verrouille le focus ; `PopUp` gère le flottement au premier plan. Une boîte de dialogue les combine généralement.
- **Concevoir le dialogue pour la touche Échap** : la croix de fermeture équivaut à « Annuler » ; veiller à ce que ce chemin soit géré proprement (aucune écriture de données non confirmée).
- **Valider avant de masquer** : effectuer les contrôles de saisie dans le bouton OK et n'exécuter `Me.Visible = False` qu'une fois la validation réussie, sans quoi le code appelant lira des valeurs invalides.

---

## 6.7.10. Récapitulatif

- Une boîte de dialogue personnalisée est un formulaire réglé avec `Modal = True`, `PopUp = True` et `BorderStyle = Dialog`.
- L'ouverture avec `WindowMode = acDialog` force la modalité **et suspend le code** appelant jusqu'à la fermeture ou au masquage du formulaire.
- Le patron de renvoi de valeur exploite cette double reprise : le bouton **OK masque** le formulaire (`Visible = False`), le bouton **Annuler le ferme**.
- Au retour, `CurrentProject.AllForms("…").IsLoaded` (chapitre 4.4) distingue OK (encore chargé) d'Annuler (fermé). Le code lit alors les contrôles du formulaire masqué, puis le **ferme explicitement**.
- Une **fonction réutilisable** combinée à `OpenArgs` (chapitre 6.6) encapsule proprement tout le mécanisme.
- Pour une simple confirmation, préférer `MsgBox` ; pour une entrée structurée, la boîte de dialogue personnalisée s'impose.

⏭️ [6.8. Formulaires indépendants (unbound forms) pour saisies complexes](/06-formulaires/08-formulaires-independants.md)
