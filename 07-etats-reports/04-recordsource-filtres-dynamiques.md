🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7.4. Manipulation dynamique du RecordSource et des filtres

Un même état sert rarement un seul cas : on veut l'imprimer pour **un** client, pour **une** période, pour **un** service. Plutôt que de multiplier les états, on adapte dynamiquement leur source de données par code. Les mécanismes sont, pour l'essentiel, ceux déjà présentés pour les formulaires au chapitre 6.5 (propriétés `RecordSource`, `Filter`, `OrderBy`, et les pièges de construction du SQL). Cette section met l'accent sur ce qui est **spécifique aux états** : l'argument `WhereCondition` de `DoCmd.OpenReport`, le moment où régler la source, et la primauté des regroupements sur le tri.

Rappelons qu'un état est en **lecture seule** : on adapte la source uniquement pour ce qu'il **affiche**.

---

## 7.4.1. Trois leviers pour adapter les données d'un état

| Levier | Action | Usage typique |
|---|---|---|
| Argument `WhereCondition` de `OpenReport` | Filtre l'état à l'ouverture, sans modifier sa conception | **La voie idiomatique** : filtrer par contexte |
| Propriété `RecordSource` | Redéfinit la requête source (colonnes, jointures, agrégats) | Changer la nature des données |
| Propriétés `Filter` / `FilterOn` | Filtrent la source par code, dans `Open` | Cas particuliers |

Pour la grande majorité des besoins (« imprimer cet état pour tel périmètre »), l'argument `WhereCondition` est la solution la plus simple et la plus propre.

---

## 7.4.2. WhereCondition : la voie idiomatique pour filtrer un état

La méthode `DoCmd.OpenReport` accepte un argument `WhereCondition` en **quatrième** position :

```vba
DoCmd.OpenReport ReportName, View, FilterName, WhereCondition, WindowMode, OpenArgs
```

`WhereCondition` reçoit une clause `WHERE` **sans le mot-clé `WHERE`**, appliquée à la source existante de l'état. L'état conserve donc un `RecordSource` général, et c'est l'appelant qui décide du périmètre :

```vba
' Imprimer les factures d'un client précis
DoCmd.OpenReport "E_Factures", acViewPreview, , "ClientID = " & Me.cboClient.Value
```

Pour une plage de dates (format ISO non ambigu — voir 7.4.9) :

```vba
Dim strWhere As String
strWhere = "DateFacture BETWEEN #" & Format(Me.txtDebut, "yyyy\-mm\-dd") & _
           "# AND #" & Format(Me.txtFin, "yyyy\-mm\-dd") & "#"

DoCmd.OpenReport "E_Factures", acViewPreview, , strWhere
```

> En interne, `WhereCondition` règle le `Filter` et le `FilterOn` de l'état pour cette ouverture. C'est pourquoi le filtrage opère sans toucher à la conception de l'état.

**Attention au mode d'affichage** : `acViewPreview` ouvre l'aperçu, tandis que `acViewNormal` **envoie directement l'état à l'imprimante**. Pour un affichage à l'écran, toujours préciser `acViewPreview`. Les modes d'affichage et `OpenReport` sont détaillés au chapitre 5.2.

---

## 7.4.3. Définir le RecordSource dynamiquement (dans Open)

Lorsqu'il ne s'agit pas seulement de filtrer, mais de changer la **requête elle-même** (colonnes différentes, jointures, agrégation), on redéfinit la propriété `RecordSource`. Le bon endroit est l'événement **`Open`** (chapitre 7.2), qui se déclenche **avant** l'exécution de la requête source :

```vba
Private Sub Report_Open(Cancel As Integer)
    Me.RecordSource = "SELECT * FROM T_Factures WHERE Annee = " & Year(Date)
End Sub
```

La construction de cette chaîne SQL obéit exactement aux mêmes règles — et aux mêmes pièges — que pour les formulaires : format des dates, gestion des apostrophes, séparateur décimal. Ces points sont traités au chapitre 6.5 et approfondis aux chapitres 11.5 (paramétrage) et 11.7 (localisation).

> Régler `RecordSource` **après** la phase d'ouverture (par exemple dans un événement de section) n'a pas d'effet utile : la source doit être fixée avant que l'état ne récupère ses données.

---

## 7.4.4. Les propriétés Filter et FilterOn

Comme un formulaire, un état dispose de `Filter` et `FilterOn`, à régler dans `Open` :

```vba
Private Sub Report_Open(Cancel As Integer)
    Me.Filter = "Statut = 'Payée'"
    Me.FilterOn = True
End Sub
```

Cette approche reste cependant **moins courante** pour un état que l'argument `WhereCondition`, qui réalise le même filtrage en une seule instruction au moment de l'ouverture, sans code dans l'état. On réserve `Filter`/`FilterOn` aux cas où la condition dépend d'un calcul interne à l'état.

---

## 7.4.5. Le tri : OrderBy et la primauté des regroupements

Point **spécifique aux états** : le tri d'un état est avant tout gouverné par sa définition de **regroupement et de tri** (volet « Regrouper, trier et total » / collection `GroupLevel`, chapitre 7.5). En conséquence, affecter `Me.OrderBy` peut **rester sans effet** si des niveaux de regroupement imposent déjà un ordre.

Pour trier dynamiquement un état de façon fiable, deux voies :

- modifier l'ordre directement dans la requête du `RecordSource` (clause `ORDER BY`) ;
- agir sur la collection `GroupLevel` (chapitre 7.5).

C'est une différence notable avec les formulaires, où `OrderBy`/`OrderByOn` suffisent généralement.

---

## 7.4.6. WhereCondition, RecordSource ou Filter : que choisir ?

| Besoin | Solution recommandée |
|---|---|
| Restreindre l'état à un périmètre (client, période, service) | **`WhereCondition`** de `OpenReport` |
| Changer les colonnes, les jointures ou agréger différemment | Propriété `RecordSource` (dans `Open`) |
| Filtrer selon une condition calculée dans l'état | `Filter` / `FilterOn` (dans `Open`) |
| Imposer un tri | Requête `ORDER BY` ou `GroupLevel` (pas `OrderBy` si regroupements) |

**Règle pratique** : commencer par `WhereCondition` ; ne redéfinir le `RecordSource` que lorsque la *structure* de la requête doit changer.

---

## 7.4.7. Construire un filtre multi-critères depuis un formulaire

Le scénario le plus fréquent : un formulaire de critères, doté de champs facultatifs, lance l'état avec un `WhereCondition` assemblé à la volée. Le pattern reprend la construction incrémentale du chapitre 6.5, mais alimente cette fois l'argument `WhereCondition` :

```vba
Private Sub cmdImprimer_Click()
    Dim strWhere As String

    If Not IsNull(Me.cboClient) Then
        strWhere = strWhere & " AND ClientID = " & Me.cboClient
    End If

    If Not IsNull(Me.txtDebut) Then
        strWhere = strWhere & " AND DateFacture >= #" & _
                   Format(Me.txtDebut, "yyyy\-mm\-dd") & "#"
    End If

    If Not IsNull(Me.cboStatut) Then
        strWhere = strWhere & " AND Statut = '" & _
                   Replace(Me.cboStatut, "'", "''") & "'"
    End If

    ' Retirer le " AND " initial
    If Len(strWhere) > 0 Then strWhere = Mid(strWhere, 6)

    DoCmd.OpenReport "E_Factures", acViewPreview, , strWhere
End Sub
```

Si aucun critère n'est renseigné, `strWhere` reste vide et l'état s'imprime sans filtre — comportement généralement souhaité.

---

## 7.4.8. Référencer un formulaire de critères : code plutôt que requête

Il est tentant de baser la requête d'un état directement sur un formulaire de critères, par exemple :

```sql
SELECT * FROM T_Factures WHERE ClientID = Forms!F_Criteres!cboClient
```

Cette approche fonctionne, mais crée une **dépendance fragile** : si le formulaire `F_Criteres` n'est pas ouvert au moment de l'impression, l'état ne s'ouvre pas (le paramètre est introuvable). De plus, gérer des critères facultatifs dans une requête fixe est malaisé.

Construire le `WhereCondition` **par code** (section 7.4.7) est préférable : l'état reste indépendant du formulaire de critères, et l'on gère aisément les critères absents. C'est l'application aux états du principe de découplage vu au chapitre 6.9.

---

## 7.4.9. Pièges et bonnes pratiques

- **Préférer `WhereCondition` pour filtrer** : c'est l'idiome propre aux états ; il évite tout code dans l'état lui-même.
- **Toujours préciser `acViewPreview`** pour un affichage écran : `acViewNormal` lance l'impression directe.
- **Régler `RecordSource` dans `Open`**, jamais plus tard : la source doit être fixée avant la récupération des données.
- **Format des dates en ISO** (`yyyy-mm-dd`), **apostrophes doublées**, **point décimal** : les mêmes règles qu'au chapitre 6.5 s'appliquent (détails en 11.7).
- **Paramétrer pour prévenir l'injection SQL** dès que des saisies utilisateur entrent dans la condition (chapitre 11.5).
- **Le tri vient des regroupements** : `OrderBy` est souvent ignoré ; agir via la requête ou `GroupLevel` (chapitre 7.5).
- **Anticiper l'absence de résultat** : un filtre trop restrictif déclenche l'événement `NoData` (chapitre 7.9) et, si l'on y annule l'état, l'erreur 2501 côté appelant (chapitre 13).
- **Découpler de tout formulaire de critères** en construisant la condition par code plutôt qu'en la référençant dans une requête fixe.

---

## 7.4.10. Récapitulatif

- Un état s'adapte par trois leviers : l'argument **`WhereCondition`** de `DoCmd.OpenReport` (le plus idiomatique), la propriété **`RecordSource`** (changer la requête), et **`Filter`/`FilterOn`** (cas particuliers).
- `WhereCondition` reçoit une clause `WHERE` sans mot-clé et filtre l'état à l'ouverture, sans toucher à sa conception ; il règle en interne `Filter`/`FilterOn`.
- Le `RecordSource` se redéfinit dans l'événement **`Open`**, en respectant les règles de construction SQL du chapitre 6.5 (dates, apostrophes, décimales).
- Le **tri** d'un état dépend de ses **regroupements** : `OrderBy` est souvent sans effet ; on agit via la requête ou la collection `GroupLevel` (chapitre 7.5).
- Le pattern de référence est la **construction d'un `WhereCondition` multi-critères** depuis un formulaire, qui découple l'état du formulaire de critères (chapitre 6.9).
- Penser à `acViewPreview` pour l'écran, au paramétrage SQL (chapitre 11.5) et à la gestion de `NoData` / erreur 2501 (chapitres 7.9 et 13).

⏭️ [7.5. Calculs et totaux dans les états](/07-etats-reports/05-calculs-totaux.md)
