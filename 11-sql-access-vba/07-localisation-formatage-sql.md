🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11.7. Localisation et formatage dans le SQL dynamique (dates US/ISO, séparateur décimal)

Nous abordons le second fil rouge annoncé dans l'introduction du chapitre, et l'un des pièges les plus insidieux d'Access : le **formatage des valeurs littérales** dans le SQL dynamique. Le sujet a été délibérément écarté des sections précédentes pour lui consacrer un traitement complet, tant il est source d'erreurs — et tant ces erreurs sont sournoises.

Ce qui rend ce piège redoutable, c'est qu'il se manifeste rarement chez le développeur et frappe chez l'utilisateur. Un code qui fonctionne parfaitement sur le poste qui l'a écrit peut **échouer, ou pire, produire silencieusement des résultats faux** sur une machine aux paramètres régionaux différents. C'est le syndrome du *« ça marche sur mon poste »* dans toute sa splendeur.

---

## Le problème : le SQL n'attend pas le format local

La racine du problème tient en une phrase : **le moteur Jet attend les valeurs littérales dans son propre format, indépendant des paramètres régionaux du poste**. Or, dès qu'on concatène une date ou un nombre dans une chaîne SQL via les mécanismes habituels de VBA (`CStr`, l'opérateur `&`), c'est le format **local** qui s'applique — celui de l'utilisateur, pas celui qu'attend Jet.

Deux exemples illustrent la gravité du décalage. Une date française `03/04/2024` (3 avril) insérée telle quelle sera interprétée par Jet comme `03/04/2024` au format américain, soit le **4 mars** — une date valide mais fausse, sans la moindre erreur pour alerter. Un nombre décimal français `1234,56` (avec une virgule) deviendra, dans le SQL, `1234,56` où la virgule est comprise comme un **séparateur d'arguments** — provoquant cette fois une erreur de syntaxe.

Le pire cas est le premier : un **résultat faux mais plausible**, qui ne déclenche aucune alarme et peut fausser des données ou des calculs pendant des mois avant d'être détecté.

---

## La solution définitive : paramétrer

Avant d'entrer dans les règles de formatage, rappelons la leçon des sections 10.5, 10.6 et 11.5 : **le paramétrage élimine purement et simplement ce problème**.

Lorsqu'on passe une valeur par un paramètre, on transmet une **valeur typée native** — une `Date`, un `Double`, une `Currency` — et non du texte à interpréter. Le moteur reçoit la valeur dans son type exact et n'a aucun formatage à décoder. Ni dièse, ni point décimal, ni apostrophe à gérer : le problème **disparaît**.

```vba
' Aucun formatage requis : la date et le nombre transitent typés
cmd.CommandText = "SELECT * FROM Commandes WHERE DateCommande = ? AND Montant >= ?;"
cmd.Parameters.Append cmd.CreateParameter("pDate", adDate, adParamInput, , Me.txtDate)
cmd.Parameters.Append cmd.CreateParameter("pMontant", adCurrency, adParamInput, , CCur(Me.txtMontant))
```

> 💡 **Le message à retenir** : si vous paramétrez vos valeurs, **vous n'avez pas besoin du reste de cette section**. Les règles de formatage qui suivent ne concernent que les cas où la concaténation d'une valeur littérale est inévitable — situation qui, pour des *valeurs*, devrait rester rare.

---

## Quand on doit insérer une valeur : les trois règles

Il arrive néanmoins qu'on doive insérer une valeur en clair — par exemple dans un contexte où le paramétrage n'est pas disponible. Trois règles, par type, doivent alors être appliquées sans exception.

### Les dates : dièses et format US/ISO

Une date littérale dans le SQL Jet doit être **encadrée de dièses** `#` et exprimée dans un format **non ambigu**. Le format **ISO `aaaa-mm-jj`** est le plus sûr, car il ne prête à aucune confusion entre jour et mois (le format américain `mm/jj/aaaa` fonctionne aussi, mais reste ambigu à l'œil).

La clé est de **formater explicitement** la date avec `Format`, sans jamais se reposer sur la conversion par défaut :

```vba
Public Function SQLDate(ByVal d As Date) As String
    SQLDate = Format$(d, "\#yyyy\-mm\-dd\#")          ' produit #2024-12-31#
End Function

' Avec composante horaire si le champ en comporte une
Public Function SQLDateHeure(ByVal d As Date) As String
    SQLDateHeure = Format$(d, "\#yyyy\-mm\-dd hh\:nn\:ss\#")
End Function
```

Un détail technique mérite explication : les **barres obliques inverses** (`\`) échappent les caractères littéraux. Sans elles, `Format` traiterait `/` comme un séparateur de date et `:` comme un séparateur d'heure — tous deux **remplacés selon les paramètres régionaux** ! En les échappant (`\-`, `\:`) et en encadrant de `\#`, on garantit un résultat **strictement indépendant de la locale**. (Notez aussi l'usage de `nn` pour les minutes, et non `mm`, qui désigne les mois en VBA.)

### Les nombres : point décimal, jamais de virgule

Un nombre décimal dans le SQL doit utiliser le **point** comme séparateur décimal, **jamais la virgule**, et **sans séparateur de milliers**. Le piège, en locale française, est que `CStr` et l'opérateur `&` produisent une **virgule**.

La fonction `Str` constitue ici la parade : elle utilise **toujours le point**, indépendamment de la locale, et n'insère aucun séparateur de milliers. Elle ajoute en revanche un espace de tête pour les nombres positifs, qu'on supprime par `Trim` :

```vba
Public Function SQLNombre(ByVal n As Variant) As String
    SQLNombre = Trim$(Str$(n))                        ' "1234.56", point garanti
End Function
```

À l'inverse, on **proscrit** `"... = " & monDouble`, qui produira `... = 1234,56` en France — une erreur de syntaxe SQL.

### Le texte : apostrophes et doublement

Une valeur texte doit être **encadrée d'apostrophes**, et ses apostrophes internes **doublées** pour ne pas rompre la chaîne (le cas `O'Connor` vu en 11.5). La fonction utilitaire renvoie directement la valeur prête à concaténer, délimiteurs compris :

```vba
Public Function SQLTexte(ByVal s As String) As String
    SQLTexte = "'" & Replace(s, "'", "''") & "'"      ' 'O''Connor'
End Function
```

---

## Avant / après

Le contraste résume l'enjeu de cette section. Voici une même requête, construite naïvement puis correctement :

```vba
' ❌ NAÏF — dépendant de la locale, potentiellement faux ou en erreur
strSQL = "SELECT * FROM Commandes WHERE DateCommande = #" & Me.txtDate & "# " & _
         "AND Montant >= " & Me.txtMontant & ";"
'   → date mal interprétée, virgule décimale = erreur, selon le poste

' ✅ FORMATAGE EXPLICITE — indépendant de la locale
strSQL = "SELECT * FROM Commandes WHERE DateCommande = " & SQLDate(Me.txtDate) & " " & _
         "AND Montant >= " & SQLNombre(Me.txtMontant) & ";"

' ✅✅ PARAMÉTRÉ — aucun formatage, et la sécurité en prime (solution préférée)
cmd.CommandText = "SELECT * FROM Commandes WHERE DateCommande = ? AND Montant >= ?;"
cmd.Parameters.Append cmd.CreateParameter("pDate", adDate, adParamInput, , Me.txtDate)
cmd.Parameters.Append cmd.CreateParameter("pMontant", adCurrency, adParamInput, , CCur(Me.txtMontant))
```

La hiérarchie est claire : **paramétrer** en priorité ; **formater explicitement** via les fonctions utilitaires lorsqu'on doit concaténer ; et **jamais** se fier à la conversion par défaut.

---

## Cas particuliers : Null et dialecte cible

Deux situations méritent une attention spécifique.

Les **valeurs `Null`** ne se comparent pas avec `=`. Une condition `WHERE Champ = Null` ne renvoie jamais rien : il faut écrire `WHERE Champ IS NULL`. Lorsqu'une valeur peut être absente, la construction du critère doit donc **bifurquer** entre `IS NULL` et une comparaison normale. Pour *insérer* une valeur nulle, on écrit le mot-clé `Null` sans apostrophes : `INSERT INTO … VALUES (…, Null, …)`.

Le **dialecte cible** change la donne dès qu'on ne s'adresse plus au moteur ACE. Les dièses `#` sont une **spécificité de Jet** : SQL Server, par exemple, attend ses dates entre **apostrophes** (`'2024-12-31'`), sans dièse. Aussi, lorsqu'on construit du SQL destiné à une **requête pass-through** (section 11.9) ou à une **base externe** (chapitre 10.9), le formatage des littéraux doit suivre les règles du SGBD visé, et non celles de Jet. Les correspondances de syntaxe entre Jet SQL et T-SQL figurent à l'annexe F.

---

## La boîte à outils centralisée

Les trois fonctions présentées ci-dessus — `SQLDate`, `SQLNombre`, `SQLTexte` — constituent la **boîte à outils de formatage** évoquée à la section 11.6 comme bonne pratique de centralisation. Réunies dans un module utilitaire, elles concentrent en un seul endroit toute la logique délicate du formatage, garantissent la cohérence à travers le projet, et rendent le code dynamique nettement plus lisible et plus sûr. Dès qu'une concaténation de valeur est inévitable, **on passe par ces fonctions**, jamais par une conversion ad hoc.

---

## En résumé

Le SQL Jet attend ses valeurs littérales dans un format **indépendant des paramètres régionaux** — ce que les conversions VBA par défaut (`CStr`, `&`) ne respectent pas, d'où des bugs sournois et dépendants du poste, jusqu'au résultat faux mais silencieux. La **solution définitive est de paramétrer** : les valeurs typées dispensent de tout formatage. Lorsque la concaténation d'une valeur est inévitable, trois règles s'imposent : **dates** entre dièses au format ISO `\#aaaa-mm-jj\#` via `Format` (en échappant les séparateurs) ; **nombres** avec un point décimal via `Str`, jamais de virgule ; **texte** entre apostrophes avec doublement des apostrophes internes. On centralise ces règles dans des fonctions utilitaires (`SQLDate`, `SQLNombre`, `SQLTexte`), on traite les `Null` avec `IS NULL`, et l'on adapte le formatage au dialecte visé dès qu'on quitte le moteur ACE.

Nous avons jusqu'ici manié le SQL d'Access en supposant connue sa syntaxe. Mais ce dialecte — le Jet SQL — possède ses propres particularités, ses fonctions spécifiques et ses limites, qui le distinguent du SQL standard et des autres moteurs. La section suivante en dresse le portrait.


⏭️ [11.8. Dialecte SQL Access (Jet SQL) — spécificités et limites](/11-sql-access-vba/08-dialecte-jet-sql.md)
