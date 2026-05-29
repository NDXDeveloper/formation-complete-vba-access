🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 6.5. RecordSource dynamique et filtres par code

La source de données d'un formulaire n'est jamais figée. Access expose plusieurs propriétés qui permettent, depuis le code VBA, de redéfinir **ce que** le formulaire affiche (`RecordSource`), **quels enregistrements** restent visibles (`Filter`, `FilterOn`) et **dans quel ordre** ils apparaissent (`OrderBy`, `OrderByOn`). Maîtriser ces leviers est indispensable pour construire des formulaires de recherche, des écrans multi-critères ou des tableaux de bord adaptatifs.

Cette section détaille chacun de ces mécanismes, leurs différences, et les pièges classiques (formatage des dates, délimiteurs de chaînes, performance).

---

## 6.5.1. La propriété RecordSource

`RecordSource` définit l'origine des données d'un formulaire **lié** (bound form). Elle accepte trois formes :

| Valeur acceptée | Exemple | Usage |
|---|---|---|
| Nom de table | `T_Clients` | Affichage direct d'une table |
| Nom de requête sauvegardée | `R_ClientsActifs` | Réutilisation d'une requête de l'application |
| Instruction SQL complète | `SELECT * FROM T_Clients WHERE Actif = True` | Source calculée dynamiquement |

La troisième forme est la plus puissante : elle permet de composer la source à la volée, sans avoir à créer une requête sauvegardée pour chaque cas.

```vba
' Affecter une table
Me.RecordSource = "T_Clients"

' Affecter une requête existante
Me.RecordSource = "R_ClientsActifs"

' Affecter une instruction SQL
Me.RecordSource = "SELECT * FROM T_Clients WHERE Ville = 'Paris'"
```

> **Point clé** : modifier `RecordSource` provoque automatiquement une nouvelle exécution de la requête. Il n'est pas nécessaire d'appeler `Me.Requery` après une réaffectation.

---

## 6.5.2. RecordSource dynamique : construire la source au moment de l'exécution

L'intérêt majeur d'un `RecordSource` modifiable par code est d'adapter la source aux choix de l'utilisateur. Le scénario typique est un formulaire de recherche où l'on assemble une clause `WHERE` selon les critères saisis.

```vba
Private Sub cmdRechercher_Click()
    Dim strSQL As String

    strSQL = "SELECT NumCommande, DateCommande, Montant " & _
             "FROM T_Commandes " & _
             "WHERE ClientID = " & Me.cboClient.Value & _
             " ORDER BY DateCommande DESC"

    Me.RecordSource = strSQL
End Sub
```

Cette approche **réduit le volume de données à la source** : seuls les enregistrements correspondant au critère sont rapatriés. C'est particulièrement avantageux lorsque le back-end est distant (table liée) et que la table d'origine contient un grand nombre de lignes.

### Réaffectation du RecordSource ou Requery ?

Deux opérations différentes méritent d'être distinguées :

- **Réaffecter `RecordSource`** : on change la *définition* de la requête (nouveaux critères, nouvelles colonnes, nouveau tri intégré). Indispensable quand la clause `WHERE` ou `SELECT` évolue.
- **`Me.Requery`** : on ré-exécute la requête **déjà définie**. Utile pour rafraîchir la liste après un ajout/suppression effectué ailleurs, sans changer la source.

```vba
' Rafraîchir la source actuelle (récupère ajouts/suppressions)
Me.Requery
```

> **À ne pas confondre avec `Me.Refresh`** : `Refresh` ne fait que relire les enregistrements **présents** dans le jeu courant (pour refléter des modifications de valeurs) ; il n'intègre **ni** les nouveaux enregistrements **ni** les suppressions. `Requery` reconstruit le jeu complet.

---

## 6.5.3. Les propriétés Filter et FilterOn

À la différence de `RecordSource`, le couple `Filter` / `FilterOn` agit **sans modifier la source** : il restreint l'affichage du jeu d'enregistrements déjà défini.

- **`Filter`** : reçoit une expression équivalente à une clause `WHERE`, **sans le mot-clé `WHERE`**.
- **`FilterOn`** : booléen qui active (`True`) ou désactive (`False`) le filtre.

```vba
' Définir puis activer un filtre
Me.Filter = "Ville = 'Paris' AND Actif = True"
Me.FilterOn = True

' Désactiver le filtre (réaffiche tous les enregistrements)
Me.FilterOn = False
```

Quelques comportements à connaître :

- Définir `Filter` alors que `FilterOn` vaut déjà `True` applique le nouveau filtre **immédiatement**.
- Définir `Filter` avec `FilterOn = False` mémorise l'expression sans l'appliquer.
- Le filtre s'applique aussi bien aux formulaires simples (il restreint la navigation aux enregistrements correspondants) qu'aux formulaires continus ou en feuille de données.

### Alternative : DoCmd.ApplyFilter

`DoCmd.ApplyFilter` réalise en une seule instruction ce que font `Filter` et `FilterOn` réunis :

```vba
' Équivalent à : Me.Filter = "..." : Me.FilterOn = True
DoCmd.ApplyFilter , "Ville = 'Paris'"

' Lever tous les filtres
DoCmd.ShowAllRecords
```

L'approche orientée objet (`Me.Filter` / `Me.FilterOn`) reste préférable car elle est plus lisible, plus facile à déboguer et indépendante de l'objet `DoCmd`.

---

## 6.5.4. Tri dynamique : OrderBy et OrderByOn

Sur le même modèle que le filtre, le tri d'affichage se pilote par `OrderBy` (la clause `ORDER BY` sans le mot-clé) et `OrderByOn`.

```vba
Me.OrderBy = "Nom ASC, Prenom ASC"
Me.OrderByOn = True
```

Ce mécanisme est idéal pour offrir un tri interactif (par exemple au clic sur l'en-tête d'une colonne en formulaire continu) sans toucher au `RecordSource`.

```vba
Private Sub lblNom_Click()
    Me.OrderBy = "Nom ASC"
    Me.OrderByOn = True
End Sub
```

---

## 6.5.5. Filter ou RecordSource : que choisir ?

Les deux approches restreignent les enregistrements visibles, mais selon des logiques différentes :

| Critère | Modifier `RecordSource` | Utiliser `Filter` |
|---|---|---|
| Niveau d'action | Redéfinit la requête source | Filtre le jeu déjà défini |
| Volume rapatrié | Réduit à la source (idéal si back-end distant) | La source reste large |
| Colonnes affichées | Modifiables (`SELECT`) | Inchangées |
| Réactivité | Reconstruit le jeu | Bascule rapide active/inactif |
| Conservation du filtre d'origine | Écrase la requête | Se superpose proprement |

**Règle pratique :**

- Pour une **forte réduction** du volume (ex. passer de 500 000 à 50 lignes), surtout sur un back-end partagé : redéfinir le `RecordSource` avec une clause `WHERE` ciblée.
- Pour un **filtrage interactif léger** d'un jeu déjà chargé (ex. affiner, puis revenir à la liste complète) : `Filter` / `FilterOn`.

> **Pour mémoire** : il existe aussi `ServerFilter`, qui applique un filtre au niveau du serveur. Cette propriété concerne principalement les projets connectés à SQL Server et a peu d'intérêt avec le moteur natif ACE/Jet, où la maîtrise du `RecordSource` joue le même rôle.

---

## 6.5.6. Pattern : construction incrémentale d'une clause WHERE multi-critères

Un formulaire de recherche réaliste combine plusieurs critères facultatifs. Le pattern consiste à **accumuler** les conditions, puis à les injecter dans le `RecordSource` (ou dans `Filter`).

```vba
Private Sub cmdAppliquer_Click()
    Dim strWhere As String
    Dim strSQL   As String

    ' Construction conditionnelle des critères
    If Not IsNull(Me.txtNom) Then
        strWhere = strWhere & " AND Nom LIKE '*" & _
                   Replace(Me.txtNom, "'", "''") & "*'"
    End If

    If Not IsNull(Me.cboVille) Then
        strWhere = strWhere & " AND Ville = '" & _
                   Replace(Me.cboVille, "'", "''") & "'"
    End If

    If Not IsNull(Me.txtDateDebut) Then
        strWhere = strWhere & " AND DateInscription >= #" & _
                   Format(Me.txtDateDebut, "yyyy\-mm\-dd") & "#"
    End If

    ' Suppression du " AND " initial superflu
    If Len(strWhere) > 0 Then
        strWhere = " WHERE " & Mid(strWhere, 6)
    End If

    strSQL = "SELECT * FROM T_Clients" & strWhere & " ORDER BY Nom"
    Me.RecordSource = strSQL
End Sub
```

Le même principe s'applique avec `Filter`, en omettant simplement le mot-clé `WHERE` :

```vba
If Len(strWhere) > 0 Then
    Me.Filter = Mid(strWhere, 6)   ' on retire " AND "
    Me.FilterOn = True
Else
    Me.FilterOn = False             ' aucun critère : tout afficher
End If
```

---

## 6.5.7. Pièges classiques et bonnes pratiques

### Formatage des dates

Le moteur Access interprète les littéraux de date entre dièses (`#…#`) au **format américain** (`mm/dd/yyyy`). Une date construite avec le format régional français produit des résultats faux ou des erreurs.

```vba
' DANGEREUX : dépend des paramètres régionaux
strSQL = "... WHERE D >= #" & Me.txtDate & "#"

' FIABLE : format ISO non ambigu, toujours correctement interprété
strSQL = "... WHERE D >= #" & Format(Me.txtDate, "yyyy\-mm\-dd") & "#"
```

Le format ISO `yyyy-mm-dd` est recommandé car il lève toute ambiguïté entre jour et mois.

### Délimiteurs et apostrophes dans les chaînes

Une valeur texte contenant une apostrophe (par exemple `O'Brien`) casse une chaîne SQL délimitée par des apostrophes. Deux solutions :

```vba
' Solution 1 : doubler les apostrophes
strSQL = "... WHERE Nom = '" & Replace(Me.txtNom, "'", "''") & "'"

' Solution 2 : délimiter avec des guillemets via Chr(34)
strSQL = "... WHERE Nom = " & Chr(34) & Me.txtNom & Chr(34)
```

Doubler les apostrophes (solution 1) est la pratique la plus robuste et la plus portable.

### Séparateur décimal des nombres

Dans une expression SQL, le séparateur décimal doit toujours être le **point**, quels que soient les paramètres régionaux. La concaténation directe d'un nombre saisi en français (virgule) génère une erreur. Ce point est traité en détail au chapitre 11.7 (*Localisation et formatage dans le SQL dynamique*).

### Gérer l'absence de résultats

Une recherche peut ne renvoyer aucun enregistrement. Plutôt que de laisser un formulaire vide sans explication, on peut tester le nombre de lignes via le clone du jeu :

```vba
Me.RecordSource = strSQL
Me.Requery

If Me.RecordsetClone.RecordCount = 0 Then
    MsgBox "Aucun enregistrement ne correspond aux critères.", _
           vbInformation, "Recherche"
End If
```

Le `RecordsetClone` est étudié spécifiquement au chapitre 9.12.

### Sécurité : injection SQL

Toute construction de SQL par concaténation de valeurs saisies par l'utilisateur expose à l'**injection SQL**. Les techniques de protection (requêtes paramétrées, fonctions d'échappement) font l'objet du chapitre 11.5 (*SQL paramétré — prévention des injections SQL*). Pour des formulaires sensibles ou des bases partagées, privilégier ces approches plutôt que la simple concaténation.

---

## 6.5.8. Récapitulatif

- `RecordSource` définit la source d'un formulaire lié : table, requête sauvegardée ou instruction SQL. Le modifier par code permet d'adapter dynamiquement les données affichées et **réduit le volume rapatrié** lorsqu'on intègre une clause `WHERE`.
- Réaffecter `RecordSource` déclenche automatiquement une nouvelle exécution ; inutile d'ajouter un `Requery`.
- `Filter` / `FilterOn` restreignent l'affichage **sans modifier la source** ; idéaux pour un filtrage interactif léger.
- `OrderBy` / `OrderByOn` pilotent le tri d'affichage sur le même principe que le filtre.
- `Requery` reconstruit le jeu complet (ajouts/suppressions inclus) ; `Refresh` se limite à relire les valeurs des enregistrements présents.
- Les pièges majeurs concernent le **formatage des dates** (format ISO conseillé), la **gestion des apostrophes** (les doubler), le **séparateur décimal** (toujours le point) et la **sécurité** (injection SQL).

⏭️ [6.6. Propriété OpenArgs — passage de paramètres à l'ouverture](/06-formulaires/06-openargs.md)
