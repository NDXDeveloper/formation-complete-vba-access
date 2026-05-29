🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7.6. Sous-états et liaison parent/enfant

Un **sous-état** est un état imbriqué dans un autre. Il permet de présenter des données **liées** — typiquement une relation un-à-plusieurs comme une facture (état principal) et ses lignes (sous-état) — ou de juxtaposer des blocs indépendants sur un même rapport. Le mécanisme est l'exact pendant des sous-formulaires (chapitre 6.4) : un même type de contrôle conteneur, et une liaison automatique par champs maître/enfant.

Cette section décrit le contrôle de sous-état, sa liaison parent/enfant, son pilotage par code, ainsi que les limites propres aux états (en-têtes de page non rendus, performances, gestion de l'absence de données).

---

## 7.6.1. Qu'est-ce qu'un sous-état ?

Un sous-état repose sur deux éléments :

- un **contrôle conteneur** placé dans l'état principal (le contrôle « sous-formulaire/sous-état », identique pour les deux usages) ;
- un **objet source** (`SourceObject`) — un état, un formulaire, une table ou une requête — affiché à l'intérieur de ce contrôle.

Usages typiques :

- relation **un-à-plusieurs** (en-tête + lignes) ;
- affichage de **données liées** à chaque enregistrement principal ;
- **juxtaposition** de synthèses indépendantes sur un seul état ;
- imbrication sur **plusieurs niveaux**.

---

## 7.6.2. Le contrôle sous-état et ses propriétés clés

Le contrôle conteneur expose plusieurs propriétés essentielles :

| Propriété | Rôle |
|---|---|
| `SourceObject` | Objet affiché (`"Report.E_Lignes"` pour un état) |
| `LinkMasterFields` | Champ(s) de l'état **principal** servant de lien |
| `LinkChildFields` | Champ(s) du **sous-état** servant de lien |
| `CanGrow` / `CanShrink` | Laisse la zone du sous-état s'adapter à son contenu |
| `Report` | Donne accès à l'objet `Report` imbriqué |
| `Visible` | Affiche ou masque le sous-état |

> **Activer `CanGrow`** sur le contrôle est presque toujours indispensable : sans cela, le contenu du sous-état est **tronqué** à la hauteur fixe du conteneur.

---

## 7.6.3. La liaison automatique : LinkMasterFields et LinkChildFields

Les propriétés `LinkMasterFields` (côté principal) et `LinkChildFields` (côté sous-état) créent la **synchronisation parent/enfant** : pour chaque enregistrement de l'état principal, Access filtre automatiquement le sous-état sur la valeur de liaison.

```vba
' Lier le sous-état des lignes à la facture courante
Me.ctlSousEtat.LinkMasterFields = "FactureID"
Me.ctlSousEtat.LinkChildFields = "FactureID"
```

Placé dans la **section Détail**, le sous-état est ainsi ré-interrogé à chaque enregistrement principal, n'affichant que les lignes correspondantes. Pour plusieurs champs de liaison, on les sépare par des points-virgules. C'est le strict équivalent, pour les états, de la liaison des sous-formulaires (chapitre 6.4).

> Un sous-état peut aussi être **indépendant** : en laissant `LinkMasterFields`/`LinkChildFields` vides, on affiche un bloc autonome (une synthèse globale, par exemple).

---

## 7.6.4. Accéder au sous-état par code

Un sous-état n'apparaît **pas** dans la collection `Reports` : il n'est pas « ouvert » de façon indépendante. On y accède via la propriété **`Report`** du contrôle conteneur :

```vba
' Objet Report imbriqué
Dim rptEnfant As Report
Set rptEnfant = Me.ctlSousEtat.Report

' Tester la présence de données, lire un contrôle du sous-état
If rptEnfant.HasData Then
    Debug.Print rptEnfant.txtTotalLignes.Value
End If
```

La syntaxe est donc `Me.NomControle.Report` puis l'accès habituel aux contrôles et propriétés de l'état imbriqué.

---

## 7.6.5. Remonter un total du sous-état vers l'état principal

On peut afficher, dans l'état principal, une valeur calculée par le sous-état (son total, par exemple). L'expression référence le contrôle via la propriété `Report` du conteneur :

```
' Source d'un contrôle de l'état principal
=[ctlSousEtat].[Report]![txtTotalLignes]
```

Le contrôle référencé doit se trouver dans une section **calculée une fois** du sous-état (son pied d'état). L'expression n'est valide que si le sous-état **contient des données** ; sinon elle renvoie une erreur (voir 7.6.8).

> **Total général à travers plusieurs sous-états** : additionner les totaux de toutes les **instances** d'un sous-état placé dans la section Détail est malaisé. Il est généralement plus simple de calculer ce total directement sur les données sous-jacentes — via une fonction de domaine (`DSum`) ou une colonne de requête — dans l'état principal (chapitre 7.5).

---

## 7.6.6. Changer le sous-état dynamiquement (SourceObject)

La propriété `SourceObject` peut être modifiée par code pour **changer l'objet affiché** au moment de l'ouverture :

```vba
Private Sub Report_Open(Cancel As Integer)
    ' Choisir le sous-état selon un critère
    If Forms!F_Criteres!optDetaille.Value Then
        Me.ctlSousEtat.SourceObject = "Report.E_LignesDetaillees"
    Else
        Me.ctlSousEtat.SourceObject = "Report.E_LignesResumees"
    End If
End Sub
```

Affecter `SourceObject = ""` **vide** le conteneur, ce qui revient à ne rien afficher.

---

## 7.6.7. Affichage conditionnel d'un sous-état

Pour masquer un sous-état dépourvu de données, on agit sur la visibilité du contrôle dans l'événement `Format` de la section qui le contient :

```vba
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    ' Masquer le sous-état s'il n'a aucune ligne pour l'enregistrement courant
    Me.ctlSousEtat.Visible = Me.ctlSousEtat.Report.HasData
End Sub
```

> La propriété `HasData` du sous-état reflète son état pour la valeur de liaison courante. Comme tout traitement dans `Format`, ce code n'opère qu'en aperçu et à l'impression (chapitre 7.2).

---

## 7.6.8. Sous-états et absence de données (NoData)

Un sous-état sans enregistrement correspondant déclenche **son propre** événement `NoData`. Attention : y placer `Cancel = True` (comme on le ferait pour un état principal vide, chapitre 7.9) provoque une **erreur** lors du rendu du conteneur dans l'état parent.

Bonne pratique pour un sous-état : **ne pas annuler** dans `NoData`, mais gérer l'affichage vide autrement — en testant `HasData` côté parent (section 7.6.7) pour masquer le conteneur, ou en affichant un libellé « Aucune donnée ». La gestion de `NoData` est détaillée au chapitre 7.9.

---

## 7.6.9. Limites et performances

Deux particularités des sous-états méritent l'attention :

- **Les en-têtes et pieds de *page* d'un sous-état ne sont pas rendus** lorsqu'il est imbriqué : seules ses sections en-tête/pied d'état, de groupe et Détail s'affichent. Tout élément à placer en haut/bas de page doit donc figurer dans l'état **principal**.
- **Performances** : un sous-état placé dans la section Détail est **ré-interrogé pour chaque enregistrement** de l'état principal. Avec de nombreux enregistrements, cela peut fortement ralentir la production. Lorsque c'est possible, un **état unique avec regroupements** (chapitre 7.5) est plus performant qu'un état à sous-états. Les techniques d'optimisation sont traitées au chapitre 18.

---

## 7.6.10. Pièges et bonnes pratiques

- **Activer `CanGrow`** sur le contrôle conteneur pour éviter de tronquer le contenu du sous-état.
- **Respecter le format de `SourceObject`** : `"Report.NomEtat"` pour un état.
- **Accéder au sous-état via `.Report`** : il n'est pas dans la collection `Reports`.
- **Vérifier `HasData`** avant de lire un contrôle du sous-état ou de remonter un total, sous peine d'erreur.
- **Ne jamais annuler le `NoData` d'un sous-état** : gérer l'affichage vide côté parent (chapitre 7.9).
- **Placer les éléments de page dans l'état principal** : les en-têtes/pieds de page des sous-états ne s'impriment pas.
- **Préférer les regroupements aux sous-états** quand la relation s'y prête, pour de meilleures performances (chapitres 7.5 et 18).
- **Calculer les totaux globaux sur les données** (`DSum`, requête) plutôt qu'en agrégeant des instances de sous-état.

---

## 7.6.11. Récapitulatif

- Un sous-état associe un **contrôle conteneur** et un **objet source** (`SourceObject`), sur le modèle des sous-formulaires (chapitre 6.4).
- La liaison parent/enfant repose sur **`LinkMasterFields`** et **`LinkChildFields`**, qui filtrent automatiquement le sous-état pour chaque enregistrement principal ; un sous-état peut aussi être **indépendant**.
- On accède au sous-état via **`Me.NomControle.Report`** (il n'est pas dans la collection `Reports`), et l'on **remonte un total** par l'expression `=[ctl].[Report]![txtTotal]`, à condition que le sous-état ait des données.
- `SourceObject` se change par code pour afficher dynamiquement un autre sous-état ; la visibilité se pilote dans `Format` selon `HasData`.
- Pour un sous-état vide, **ne pas annuler** son `NoData` ; gérer l'affichage côté parent (chapitre 7.9).
- Limites clés : **en-têtes/pieds de page non rendus** et **ré-interrogation par enregistrement** — d'où l'intérêt de préférer les **regroupements** quand c'est possible (chapitres 7.5 et 18).

⏭️ [7.7. Génération d'états par code (création, modification de contrôles)](/07-etats-reports/07-generation-etats-par-code.md)
