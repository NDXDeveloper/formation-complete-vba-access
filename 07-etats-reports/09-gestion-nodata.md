🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7.9. Gestion de l'événement NoData — éviter les états vides

Un état lancé sur une sélection qui ne renvoie aucun enregistrement s'imprime malgré tout : en-têtes seuls, zone de détail vide, et souvent des `#Erreur` dans les calculs. Le résultat est confus et peu professionnel. Access fournit un événement dédié, **`NoData`**, pour traiter ce cas. Mais son usage cache un piège : annuler un état vide provoque l'**erreur 2501** dans le code appelant, qu'il faut impérativement gérer.

Cette section réunit et approfondit les éléments évoqués au fil du chapitre (l'événement `NoData` en 7.2, la propriété `HasData` en 7.1, l'erreur 2501 en 7.8) et présente les stratégies pour éviter proprement les états vides.

---

## 7.9.1. Le problème des états vides

Sans traitement, un état dépourvu de données :

- imprime ses en-têtes (état, page, groupe) mais une zone de **détail vide** ;
- affiche fréquemment des **`#Erreur`** dans les contrôles calculés des en-têtes/pieds ;
- produit un artefact **trompeur**, qui laisse croire à un dysfonctionnement.

L'objectif est donc soit d'**empêcher** l'ouverture d'un tel état, soit d'afficher à la place un **message clair**.

---

## 7.9.2. L'événement NoData

`Report_NoData` se déclenche lorsque la source de l'état ne renvoie **aucun enregistrement**. Il survient après `Open` (et `Activate`), **avant** toute mise en forme (chapitre 7.2) :

```vba
Private Sub Report_NoData(Cancel As Integer)
    ' Déclenché uniquement si l'état n'a aucune donnée
End Sub
```

C'est le point d'entrée naturel pour réagir à l'absence de données.

---

## 7.9.3. Annuler un état vide (et l'erreur 2501)

La réaction la plus courante consiste à **annuler** l'état en affectant `Cancel = True` :

```vba
Private Sub Report_NoData(Cancel As Integer)
    Cancel = True
End Sub
```

L'état ne s'affiche alors pas. **Mais** lorsqu'il a été ouvert par `DoCmd.OpenReport` (ou exporté par `DoCmd.OutputTo`, chapitre 7.8), cette annulation **remonte** au code appelant sous la forme de l'**erreur 2501** (« L'action OuvrirÉtat a été annulée »). Sans gestion, l'utilisateur voit ce message d'erreur technique au lieu d'une information compréhensible.

> **Règle à retenir** : annuler dans `NoData` impose **toujours** d'intercepter l'erreur 2501 du côté qui lance l'état.

---

## 7.9.4. Gérer l'erreur 2501 côté appelant

Le code qui ouvre l'état doit donc traiter spécifiquement l'erreur 2501 :

```vba
Sub OuvrirEtatFactures(ByVal lngID As Long)
    On Error GoTo Gestion

    DoCmd.OpenReport "E_Factures", acViewPreview, , "ClientID=" & lngID
    Exit Sub

Gestion:
    If Err.Number = 2501 Then
        MsgBox "Aucune donnée à afficher pour cette sélection.", _
               vbInformation, "État"
    Else
        MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical
    End If
End Sub
```

Cette répartition des rôles est propre : l'état se contente d'**annuler** (sans message), et l'appelant décide du **message** adapté au contexte. L'état reste ainsi réutilisable depuis plusieurs endroits, chacun affichant son propre message. La gestion structurée des erreurs est détaillée au chapitre 13.

> **Éviter le double message** : si l'on affiche déjà un `MsgBox` dans `NoData`, ne pas en afficher un second à l'interception du 2501 — choisir **un seul** des deux emplacements.

---

## 7.9.5. L'approche proactive : tester avant d'ouvrir

La stratégie souvent la plus propre consiste à **ne pas ouvrir** un état voué à être vide. On teste la présence de données **avant** l'ouverture, à l'aide de `DCount` (chapitre 11.10), et l'on n'ouvre l'état que s'il y a quelque chose à afficher :

```vba
Sub OuvrirEtatFactures(ByVal lngID As Long)
    If DCount("*", "T_Factures", "ClientID=" & lngID) = 0 Then
        MsgBox "Aucune facture pour ce client.", vbInformation, "État"
        Exit Sub
    End If

    DoCmd.OpenReport "E_Factures", acViewPreview, , "ClientID=" & lngID
End Sub
```

Cette approche évite entièrement le mécanisme `NoData`/erreur 2501 et offre un message immédiat. Elle est idéale lorsqu'un formulaire lance l'état avec des critères connus.

> **Cohérence du critère** : le filtre de `DCount` doit correspondre exactement à celui de l'état. Si l'état repose sur une requête complexe, compter sur **la même requête** plutôt que sur la table brute, sous peine de divergence entre le test et le contenu réel.

---

## 7.9.6. La propriété HasData

À la différence de `NoData` (un **événement** qui se déclenche une fois si l'état est vide), **`HasData`** est une **propriété** que l'on peut interroger à tout moment dans le code de l'état (chapitre 7.1) :

```vba
If Me.HasData Then
    ' l'état contient des enregistrements
End If
```

`HasData` est particulièrement utile à l'**intérieur** d'un état pour piloter l'affichage de certains éléments, et indispensable pour les **sous-états** (section suivante).

---

## 7.9.7. Afficher un message plutôt qu'annuler

Dans certains contextes, on préfère produire l'**enveloppe** de l'état avec un message « Aucune donnée » plutôt que de l'annuler — par exemple quand un destinataire attend systématiquement un document. On place alors une étiquette masquée par défaut, que `NoData` rend visible (sans annuler) :

```vba
Private Sub Report_NoData(Cancel As Integer)
    ' Ne pas annuler : afficher un message dans l'état
    Me.lblAucuneDonnee.Visible = True
End Sub
```

Le contrôle `lblAucuneDonnee` (par exemple « Aucune donnée pour les critères sélectionnés ») est invisible en temps normal. Avec cette approche, il peut être judicieux de **masquer aussi** les contrôles calculés susceptibles d'afficher `#Erreur` en l'absence de données.

---

## 7.9.8. Cas des sous-états

Rappel important du chapitre 7.6 : pour un **sous-état** vide, il ne faut **pas** annuler dans son `NoData`, car cela provoque une erreur lors du rendu du conteneur dans l'état parent. On gère l'absence de données autrement :

- côté **parent**, en testant `HasData` du sous-état pour masquer son conteneur ;
- ou en laissant le sous-état afficher un message « Aucune donnée ».

```vba
' Dans l'état parent : masquer le sous-état vide
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    Me.ctlSousEtat.Visible = Me.ctlSousEtat.Report.HasData
End Sub
```

---

## 7.9.9. Choisir la bonne stratégie

| Situation | Stratégie recommandée |
|---|---|
| Un formulaire lance l'état avec des critères connus | **Test proactif** `DCount` avant ouverture (7.9.5) |
| L'état est lancé de multiples façons, ou le critère est complexe | **Annuler dans `NoData`** + gérer l'erreur 2501 (7.9.3–7.9.4) |
| Un document est attendu même vide | **Afficher un message** sans annuler (7.9.7) |
| Sous-état | **Ne jamais annuler** ; masquer via `HasData` côté parent (7.9.8) |

En pratique, le **test proactif** est le plus simple et le plus confortable quand il est applicable ; l'**annulation avec gestion du 2501** est le filet de sécurité universel.

---

## 7.9.10. Pièges et bonnes pratiques

- **Toujours intercepter l'erreur 2501** dès qu'on annule dans `NoData` (ou `Open`) : sinon l'utilisateur voit un message technique.
- **Un seul message** : afficher l'information *soit* dans `NoData`, *soit* à l'interception du 2501, jamais les deux.
- **Privilégier le test proactif** (`DCount`) lorsque les critères sont connus à l'avance.
- **Aligner le critère de test** sur le filtre réel de l'état (même requête, même `WHERE` — chapitre 7.4).
- **Ne jamais annuler le `NoData` d'un sous-état** : masquer le conteneur via `HasData` côté parent (chapitre 7.6).
- **Masquer les contrôles calculés** lorsqu'on affiche un message plutôt que d'annuler, pour éviter les `#Erreur`.
- **Distinguer `NoData` (événement) et `HasData` (propriété)** : le premier réagit à l'absence, le second se teste à tout moment.

---

## 7.9.11. Récapitulatif

- Un état sans données s'imprime vide et affiche des `#Erreur` ; l'événement **`NoData`** permet d'y remédier (déclenché après `Open`, si la source est vide — chapitre 7.2).
- Annuler avec `Cancel = True` empêche l'affichage, **mais** provoque l'**erreur 2501** côté appelant, qui doit être **interceptée** (chapitres 7.8 et 13).
- La répartition propre : l'état **annule**, l'appelant **gère le 2501** et choisit le message — pour un état réutilisable, sans double message.
- L'**approche proactive** (`DCount` avant ouverture — chapitre 11.10) évite tout le mécanisme et reste la plus simple quand les critères sont connus.
- La propriété **`HasData`** (chapitre 7.1) se teste à tout moment ; elle est indispensable pour gérer les **sous-états** vides, dont le `NoData` ne doit **jamais** être annulé (chapitre 7.6).
- Trois stratégies selon le besoin : **test proactif**, **annulation + gestion du 2501**, ou **message dans l'état** sans annuler.

⏭️ [7.10. États paramétrés et envoi par email automatisé](/07-etats-reports/10-etats-parametres-envoi-email.md)
