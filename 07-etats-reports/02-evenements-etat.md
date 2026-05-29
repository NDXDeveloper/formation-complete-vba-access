🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7.2. Événements d'un état (Open, Activate, NoData, Format, Print, Close)

Un état réagit à des **événements** déclenchés tout au long de son cycle de vie : à son ouverture, lorsqu'il découvre qu'il n'a aucune donnée, au moment de mettre en forme chaque section, puis à sa fermeture. Savoir **quel événement se produit à quel moment** est la clé pour piloter correctement un état : un même traitement placé dans `Open`, dans `Format` ou dans `Print` ne produira pas du tout le même résultat.

Cette section présente les deux familles d'événements d'un état, leur ordre de déclenchement, et l'usage précis de chacun — en particulier la distinction essentielle entre `Format` et `Print`.

---

## 7.2.1. Deux familles d'événements : niveau état et niveau section

Les événements d'un état se répartissent en deux catégories :

- **Les événements de l'état** (niveau rapport), qui rythment sa vie globale : `Open`, `Activate`, `NoData`, `Page`, `Close`, `Deactivate`, `Error`.
- **Les événements de section**, déclenchés *pour chaque section* au fil de la mise en forme : `Format`, `Print`, `Retreat`.

Cette distinction reflète le modèle de traitement d'un état décrit au chapitre 7 : un état parcourt ses données en **mettant en forme ses sections les unes après les autres**, ce qui justifie l'existence d'événements spécifiques au niveau section.

---

## 7.2.2. L'ordre des événements

À l'ouverture d'un état contenant des données, les événements se succèdent globalement ainsi :

```
Open  →  Activate  →  [ Format / Print par section, Page par page ]  →  Close  →  Deactivate
```

Si l'état ne contient **aucun** enregistrement, l'événement `NoData` s'intercale après `Open` (et `Activate`), avant toute mise en forme :

```
Open  →  Activate  →  NoData  →  Close  →  Deactivate
```

Au sein de la phase de mise en forme, chaque section déclenche `Format` (avant la mise en page), puis `Print` (avant l'impression effective), et éventuellement `Retreat` lorsque Access doit revenir en arrière.

---

## 7.2.3. Open — préparer ou annuler l'ouverture

`Report_Open` est le **premier** événement, déclenché **avant** l'exécution de la requête source. C'est l'endroit idéal pour :

- définir dynamiquement la source de données (`RecordSource`, `Filter`, `OrderBy` — voir section 7.4) ;
- demander des paramètres à l'utilisateur ;
- adapter le titre de l'aperçu (`Caption`) ;
- **annuler** l'ouverture si une condition n'est pas remplie, via l'argument `Cancel`.

```vba
Private Sub Report_Open(Cancel As Integer)
    Me.Caption = "Ventes au " & Format(Date, "dd/mm/yyyy")

    ' Annuler l'ouverture si un formulaire de critères n'est pas ouvert
    If Not CurrentProject.AllForms("F_Criteres").IsLoaded Then
        MsgBox "Veuillez d'abord ouvrir le formulaire de critères.", vbExclamation
        Cancel = True
    End If
End Sub
```

> Affecter `Cancel = True` dans `Open` interrompt l'ouverture. Si l'état a été lancé par `DoCmd.OpenReport`, l'annulation provoque l'erreur **2501** dans le code appelant, qu'il convient d'intercepter (chapitre 13).

---

## 7.2.4. Activate, Deactivate et Close

- **`Activate`** se déclenche lorsque la fenêtre de l'état devient active (par exemple en aperçu). Utile pour afficher un ruban ou un menu contextuel propre à l'état.
- **`Deactivate`** se déclenche lorsqu'une autre fenêtre prend le focus.
- **`Close`** se déclenche à la fermeture de l'état : c'est l'endroit du **nettoyage** (libération de recordsets ouverts, réinitialisation de variables — chapitre 9.11).

```vba
Private Sub Report_Close()
    ' Libérer d'éventuelles ressources ouvertes à l'ouverture de l'état
    Set m_rsDonnees = Nothing
End Sub
```

---

## 7.2.5. NoData — l'absence de données

`Report_NoData` se déclenche lorsque la source de l'état ne renvoie **aucun enregistrement**. Sans traitement, un état vide s'imprime malgré tout (en-têtes seuls, valeurs `#Erreur` dans les calculs). On y remédie en annulant l'état :

```vba
Private Sub Report_NoData(Cancel As Integer)
    MsgBox "Aucune donnée à imprimer pour ces critères.", vbInformation, "État"
    Cancel = True
End Sub
```

Là encore, `Cancel = True` génère l'erreur 2501 côté appelant. Le traitement complet de l'événement `NoData` — y compris la gestion de cette erreur — fait l'objet de la section 7.9.

---

## 7.2.6. Format et Print : le cœur de la mise en forme

Ces deux événements, déclenchés **pour chaque section**, sont les plus puissants et les plus subtils.

### Format

`Format` se déclenche **avant** qu'Access ne mette la section en page, c'est-à-dire **avant** que la disposition finale ne soit connue. À ce stade, on peut encore :

- modifier l'apparence d'un contrôle (mise en forme conditionnelle) ;
- afficher ou masquer une section ;
- calculer des valeurs qui influencent la disposition ;
- piloter le flux de mise en forme (section 7.2.8).

```vba
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    ' Mettre en évidence les montants négatifs
    If Nz(Me.txtMontant.Value, 0) < 0 Then
        Me.txtMontant.ForeColor = vbRed
    Else
        Me.txtMontant.ForeColor = vbBlack
    End If
End Sub
```

> Chaque section possède sa propre procédure d'événement, nommée d'après la section : `Detail_Format`, `ReportHeader_Format`, `PageHeaderSection_Format`, `GroupHeader0_Format`, etc.

### Print

`Print` se déclenche **après** la mise en forme, juste **avant l'impression physique** de la section. La disposition est alors figée. On y place les traitements liés à l'impression **effective** : cumuls fondés sur ce qui est réellement imprimé, ou dessin direct (méthodes `Line`, `Circle`, `Print`).

```vba
Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)
    ' Traitement lié à l'impression réelle de la section
End Sub
```

### Retreat

`Retreat` se déclenche lorsqu'Access doit **revenir en arrière** sur des sections déjà formatées (par exemple pour respecter `KeepTogether`). Si l'on a effectué des cumuls dans `Format`, c'est dans `Retreat` qu'il faut éventuellement les **annuler** pour rétablir la cohérence.

---

## 7.2.7. Format ou Print : lequel choisir ?

| Critère | `Format` | `Print` |
|---|---|---|
| Moment | Avant la mise en page | Après la mise en page, avant impression |
| Disposition | Pas encore figée | Figée |
| Usages typiques | Mise en forme conditionnelle, masquage, calculs influençant la disposition | Cumuls liés à l'impression réelle, dessin (`Line`, `Circle`) |
| Peut se produire sans l'autre | Une section peut être **formatée sans être imprimée** | Une section peut être **réimprimée** plusieurs fois |

**Règle pratique** : tout ce qui touche à l'**apparence** ou à la **structure** se place dans `Format` ; tout ce qui doit correspondre exactement à ce qui est **physiquement imprimé** se place dans `Print`. Attention au piège classique : une section peut être formatée puis abandonnée (`Retreat`), si bien qu'un cumul réalisé dans `Format` peut compter une donnée qui ne sera pas imprimée.

---

## 7.2.8. Piloter le flux : MoveLayout, NextRecord, PrintSection

Dans l'événement `Format`, trois propriétés booléennes de l'état contrôlent la progression de la mise en forme :

- **`MoveLayout`** : Access avance-t-il à l'emplacement d'impression suivant ?
- **`NextRecord`** : passe-t-il à l'enregistrement suivant ?
- **`PrintSection`** : la section est-elle imprimée ?

Par défaut, les trois valent `True` (comportement normal). Leurs combinaisons permettent des effets avancés, notamment pour les étiquettes de publipostage :

| `MoveLayout` | `NextRecord` | `PrintSection` | Effet |
|---|---|---|---|
| True | True | True | Comportement normal (défaut) |
| True | False | True | Réimprime la même donnée à l'emplacement suivant |
| True | True | False | Avance en laissant un emplacement vide |
| False | True | True | Imprime par superposition |

```vba
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    ' Exemple : laisser un emplacement vide sans avancer l'enregistrement
    Me.MoveLayout = True
    Me.NextRecord = False
    Me.PrintSection = False
End Sub
```

---

## 7.2.9. FormatCount, PrintCount et les passes de mise en forme

Access peut mettre en forme une même section **plusieurs fois** (par exemple lorsqu'il évalue si elle tient sur la page). Deux compteurs renseignent sur ces passes :

- **`FormatCount`** (argument de `Format`) : nombre de fois que `Format` s'est déclenché pour l'occurrence courante de la section.
- **`PrintCount`** (argument de `Print`) : nombre de fois que `Print` s'est déclenché.

Pour n'exécuter un traitement **qu'une seule fois** par occurrence, on teste le compteur :

```vba
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    If FormatCount = 1 Then
        ' Code exécuté une seule fois, même si la section est reformatée
    End If
End Sub
```

Cette précaution évite, par exemple, de compter deux fois une valeur dans un cumul.

---

## 7.2.10. Page et Error

- **`Page`** se déclenche après la mise en forme d'une page, avant son impression : idéal pour dessiner des éléments à l'échelle de la page (bordures, filigranes) via les méthodes graphiques de l'état.
- **`Error`** (`Report_Error(DataErr, Response)`) intercepte les erreurs du moteur survenant pendant la production de l'état. La gestion des erreurs est traitée au chapitre 13.

---

## 7.2.11. Aperçu/Impression vs mode État (Report View)

Distinction **cruciale** : les événements `Format`, `Print` et `Retreat` ne se déclenchent **qu'en mode Aperçu avant impression et à l'impression**. En **mode État** (Report View, introduit pour la consultation interactive à l'écran), il n'y a ni pagination ni impression : Access déclenche à la place des événements proches de ceux des formulaires (`Load`, `Current`…) et **non** `Format`/`Print`.

Conséquence : une mise en forme conditionnelle écrite dans `Detail_Format` **ne s'applique pas** en mode État. Pour un rendu cohérent dans les deux modes, on recourt à la **mise en forme conditionnelle** intégrée ou par code (chapitre 17.8), ou l'on impose l'ouverture en aperçu (`acViewPreview`).

---

## 7.2.12. Pièges et bonnes pratiques

- **Choisir l'événement selon le moment** : préparer la source dans `Open`, mettre en forme dans `Format`, agir sur l'impression réelle dans `Print`, nettoyer dans `Close`.
- **Tester `FormatCount = 1`** pour les traitements à n'exécuter qu'une fois : une section peut être formatée plusieurs fois.
- **Compenser dans `Retreat`** les cumuls effectués dans `Format`, sous peine d'incohérences quand Access revient en arrière.
- **Intercepter l'erreur 2501** côté appelant lorsqu'on annule dans `Open` ou `NoData` (chapitres 7.9 et 13).
- **Se souvenir des deux modes** : `Format`/`Print` n'existent qu'en aperçu/impression ; pour le mode État, prévoir une mise en forme conditionnelle adaptée (chapitre 17.8).
- **Nommer correctement les procédures de section** : `Detail_Format`, `GroupHeader0_Format`, etc., d'après le nom de la section.

---

## 7.2.13. Récapitulatif

- Les événements d'un état se répartissent entre **niveau état** (`Open`, `Activate`, `NoData`, `Page`, `Close`, `Deactivate`, `Error`) et **niveau section** (`Format`, `Print`, `Retreat`).
- L'ordre global est `Open → Activate → (Format/Print, Page) → Close → Deactivate`, avec `NoData` intercalé si l'état est vide.
- `Open` prépare ou annule l'ouverture (source, titre, `Cancel`) ; `Close` nettoie les ressources.
- `NoData` gère l'absence de données (`Cancel = True`), traité en détail au chapitre 7.9.
- **`Format`** agit avant la mise en page (apparence, masquage, flux) ; **`Print`** agit avant l'impression réelle (cumuls effectifs, dessin). Une section peut être formatée sans être imprimée, ou réimprimée plusieurs fois.
- `MoveLayout`, `NextRecord` et `PrintSection` pilotent le flux dans `Format` ; `FormatCount`/`PrintCount` renseignent sur les passes.
- En **mode État (Report View)**, `Format`/`Print` ne se déclenchent pas : prévoir une mise en forme conditionnelle adaptée (chapitre 17.8).

⏭️ [7.3. Sections : En-tête, Détail, Pied de page, En-têtes/pieds de groupe](/07-etats-reports/03-sections-etat.md)
