🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11.6. Construction dynamique de SQL — bonnes pratiques et pièges

La section précédente a posé une distinction essentielle : les **valeurs** se paramètrent, mais les **identifiants** et la **structure** d'une requête ne le peuvent pas. Or de nombreuses situations réelles exigent justement de faire varier cette structure à l'exécution — c'est l'objet de la construction dynamique de SQL. Loin d'être un pis-aller, elle est parfois indispensable ; encore faut-il la pratiquer avec méthode et sans rouvrir les failles que le paramétrage a refermées.

Cette section expose les cas où la construction dynamique est légitime, les bonnes pratiques pour l'assembler proprement, et les pièges récurrents. Le formatage précis des valeurs littérales (dates, nombres, apostrophes) y est volontairement laissé de côté : il fait l'objet, à lui seul, de la section 11.7.

---

## Quand la construction dynamique est-elle légitime ?

Plusieurs besoins ne peuvent être couverts par le seul paramétrage et justifient de bâtir la requête à la volée.

Les **identifiants variables** d'abord : un nom de table, une colonne de tri, le sens d'un `ORDER BY` choisis à l'exécution. Comme établi en 11.5, un paramètre ne peut pas les remplacer.

Les **critères optionnels** ensuite, cas le plus fréquent : un formulaire de recherche multicritères où l'utilisateur ne renseigne qu'une partie des filtres. C'est la *structure même* du `WHERE` qui change selon les champs remplis — on inclut une condition, ou pas.

Les **listes `IN` de longueur variable** : filtrer sur un ensemble de valeurs dont le nombre n'est pas connu à l'avance se prête mal au paramétrage classique.

Enfin, les besoins plus avancés — **liste de colonnes dynamique**, **jointures conditionnelles**, colonnes de type analyse croisée — relèvent eux aussi de la construction dynamique. Celle-ci est donc une compétence à part entière, et non une mauvaise habitude à proscrire.

---

## L'idéal hybride : structure dynamique, valeurs paramétrées

Il serait erroné de retenir de la section 11.5 qu'il ne faut « jamais concaténer ». La bonne lecture est plus nuancée : **on construit la structure dynamiquement, mais on paramètre les valeurs**. Ces deux approches ne s'opposent pas — elles se combinent.

Concrètement, on décide à l'exécution *quelles* conditions inclure (structure dynamique), tandis que les *valeurs* de ces conditions transitent par des paramètres (sécurité et formatage assurés). C'est le meilleur des deux mondes : la souplesse de la construction dynamique et la robustesse du paramétrage. Nous verrons plus loin comment maintenir structure et paramètres synchronisés.

---

## Bonnes pratiques de construction

### 1. Toujours tracer le SQL final

C'est l'habitude la plus rentable de toute la construction dynamique : **afficher la chaîne assemblée** dans la fenêtre Exécution immédiate avant de l'exécuter.

```vba
Debug.Print strSQL
```

Cette simple ligne permet d'**inspecter** la requête réellement produite, d'y repérer un espace manquant ou un délimiteur oublié, et même de **copier-coller** le résultat dans le concepteur de requêtes pour le tester directement. Face à une requête dynamique qui échoue, c'est presque toujours le premier réflexe à avoir (les techniques de débogage sont détaillées au chapitre 19.2).

### 2. Gérer les espaces entre fragments

Le bug numéro un de la concaténation est l'**espace manquant** entre deux fragments. Coller `"FROM Clients"` et `"WHERE ..."` produit `"FROM ClientsWHERE ..."` — une requête invalide.

```vba
' ❌ Espaces manquants
strSQL = "SELECT * FROM Clients" & "WHERE Actif = True"   ' "ClientsWHERE"

' ✅ Espaces présents (en tête ou en queue de fragment, de façon cohérente)
strSQL = "SELECT * FROM Clients" & " WHERE Actif = True"
```

La règle : intégrer systématiquement un **espace de séparation** au début (ou à la fin) de chaque fragment, et s'y tenir.

### 3. Construire un `WHERE` conditionnel proprement

Pour assembler un `WHERE` à partir de critères optionnels, l'astuce du **`WHERE 1=1`** simplifie considérablement le code. En partant d'une condition toujours vraie et neutre, chaque critère ultérieur s'ajoute uniformément avec `AND`, sans avoir à gérer le cas particulier de la première condition :

```vba
Dim strSQL As String
strSQL = "SELECT * FROM Clients WHERE 1=1"      ' base neutre

If Len(Me.txtVille & "") > 0 Then
    strSQL = strSQL & " AND Ville = '" & Me.txtVille & "'"   ' formatage : voir 11.7
End If
If IsNumeric(Me.txtMontantMin & "") Then
    strSQL = strSQL & " AND Solde >= " & Me.txtMontantMin
End If

strSQL = strSQL & ";"
Debug.Print strSQL
```

Sans le `1=1`, il faudrait déterminer si chaque condition est la première (pour écrire `WHERE`) ou non (pour écrire `AND`) — une gymnastique source d'erreurs. Une alternative plus structurée consiste à **collecter les conditions dans un tableau** puis à les assembler avec `Join(conditions, " AND ")`, en ne préfixant `WHERE` que si le tableau n'est pas vide. Les deux approches sont valables ; le `1=1` a pour lui la simplicité.

### 4. Construire une liste `IN` de longueur variable

Pour filtrer sur un ensemble de valeurs, on construit le contenu de la clause `IN` par concaténation, idéalement avec `Join` :

```vba
Dim ids() As String
ids = Split("12,27,43", ",")          ' valeurs (numériques ici)
Dim strSQL As String
strSQL = "SELECT * FROM Commandes WHERE IDClient IN (" & Join(ids, ", ") & ");"
```

Pour des valeurs **numériques**, l'insertion directe convient. Pour des valeurs **texte**, chacune doit être encadrée d'apostrophes et correctement échappée — un formatage que la section 11.7 détaille. Attention également à ne pas produire une liste vide (`IN ()`), qui provoquerait une erreur : on vérifie qu'au moins une valeur est présente.

### 5. Valider les identifiants par liste blanche

Le rappel est crucial, car la construction dynamique réintroduit le risque d'injection sur les identifiants. **Jamais** on n'insère une saisie utilisateur brute en position de nom de colonne, de table ou de tri. On la **valide contre une liste fixe** d'identifiants autorisés, selon le patron établi en 11.5 :

```vba
Dim colTri As String
Select Case Me.cboTri.Value
    Case "Nom":   colTri = "NomClient"
    Case "Ville": colTri = "Ville"
    Case Else:    colTri = "IDClient"      ' valeur sûre par défaut
End Select
strSQL = strSQL & " ORDER BY " & colTri    ' colTri est maîtrisé
```

### 6. Centraliser le formatage et l'échappement

Plutôt que de répéter partout le formatage des valeurs, on l'**encapsule dans des fonctions utilitaires** — une fonction pour les dates, une pour les chaînes, etc. Ce réflexe garantit la cohérence et concentre la logique délicate en un seul endroit. Ces fonctions de formatage sont précisément le sujet de la section 11.7.

---

## Garder structure et paramètres synchronisés

Voici la concrétisation de l'idéal hybride. Lorsqu'on combine un `WHERE` dynamique avec des paramètres, il faut **ajouter chaque emplacement `?` et son paramètre ensemble**, dans le même ordre. Les paramètres ADO étant positionnels (section 10.6), cette synchronisation suffit à les maintenir alignés :

```vba
Dim cmd As ADODB.Command
Set cmd = New ADODB.Command
cmd.ActiveConnection = CurrentProject.Connection
cmd.CommandType = adCmdText

Dim strSQL As String
strSQL = "SELECT * FROM Clients WHERE 1=1"

If Len(Me.txtVille & "") > 0 Then
    strSQL = strSQL & " AND Ville = ?"                      ' on ajoute le ?
    cmd.Parameters.Append cmd.CreateParameter( _
        "pVille", adVarWChar, adParamInput, 50, Me.txtVille) ' ET son paramètre
End If
If IsNumeric(Me.txtMontantMin & "") Then
    strSQL = strSQL & " AND Solde >= ?"
    cmd.Parameters.Append cmd.CreateParameter( _
        "pSolde", adCurrency, adParamInput, , CCur(Me.txtMontantMin))
End If

cmd.CommandText = strSQL & ";"
Dim rs As ADODB.Recordset
Set rs = cmd.Execute
```

La structure varie selon les critères saisis, mais aucune valeur n'est concaténée : chaque `?` ajouté est immédiatement accompagné de son paramètre. On obtient une requête à la fois **flexible** (structure dynamique) et **sûre** (valeurs paramétrées) — la réconciliation parfaite des sections 11.5 et 11.6.

---

## Pièges courants (récapitulatif)

La construction dynamique concentre un certain nombre de pièges, dont voici la synthèse. L'**espace manquant** entre fragments reste le plus fréquent. Viennent ensuite l'**oubli des délimiteurs** autour des valeurs (apostrophes pour le texte, dièses pour les dates — section 11.7) et les **séparateurs en trop** (un `AND` ou une virgule traînant en fin de chaîne, qu'évite l'approche `1=1` ou `Join`). L'**insertion de saisie brute** rouvre la faille d'injection, pour les valeurs (paramétrer) comme pour les identifiants (liste blanche). Les noms comportant **espaces ou mots réservés** exigent des **crochets** `[ ]`. L'**absence de `Debug.Print`** prive du moyen le plus simple de diagnostiquer. Enfin, les **monolithes illisibles** — d'interminables concaténations en une seule expression — nuisent à la maintenance : mieux vaut construire la requête par fragments clairs et la tracer.

---

## En résumé

La construction dynamique de SQL est **légitime et parfois incontournable** : identifiants variables, critères optionnels, listes `IN` de longueur variable, structures conditionnelles. La clé est de **ne pas l'opposer au paramétrage** mais de les **combiner** — structure dynamique, valeurs paramétrées —, en ajoutant chaque `?` et son paramètre de façon synchronisée. Les bonnes pratiques tiennent en quelques réflexes : **tracer** systématiquement le SQL final (`Debug.Print`), **gérer les espaces** entre fragments, assembler un `WHERE` conditionnel proprement (astuce `WHERE 1=1` ou `Join`), construire les listes `IN` avec soin, **valider les identifiants par liste blanche**, et **centraliser le formatage** dans des fonctions dédiées. Les pièges — espaces, délimiteurs, séparateurs traînants, saisie brute, mots réservés — se neutralisent par la méthode et la traçabilité.

Reste le sujet délibérément écarté jusqu'ici, mais omniprésent dès qu'on insère une valeur littérale dans une chaîne SQL : son **formatage**. Une date, un nombre, une apostrophe mal formatés produisent des erreurs subtiles, parfois dépendantes des paramètres régionaux du poste. C'est l'objet, déterminant, de la section suivante.


⏭️ [11.7. Localisation et formatage dans le SQL dynamique (dates US/ISO, séparateur décimal)](/11-sql-access-vba/07-localisation-formatage-sql.md)
