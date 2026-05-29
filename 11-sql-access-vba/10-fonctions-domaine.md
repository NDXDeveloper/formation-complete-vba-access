🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11.10. DLookup, DSum, DCount et autres fonctions de domaine

Comme l'annonçait la section précédente, il existe une famille de fonctions à mi-chemin entre le SQL et le code : les **fonctions de domaine**. Elles permettent d'obtenir une valeur unique — une recherche, un comptage, une somme — sur une table ou une requête, sans avoir à écrire de requête complète ni à ouvrir un recordset. C'est leur grande commodité : une seule ligne suffit là où il faudrait sinon plusieurs.

Mais cette commodité a un revers que cette section ne manquera pas de souligner : ces fonctions sont **lentes**, et leur critère obéit aux mêmes règles — et aux mêmes dangers — que tout SQL dynamique. Bien employées, elles rendent de grands services ; mal employées, elles plombent les performances ou rouvrent des failles.

---

## Que sont les fonctions de domaine ?

Les fonctions de domaine (le préfixe « D » signifie *Domain*) effectuent un calcul ou une recherche sur un **domaine** — une table ou une requête — et renvoient une **valeur unique**. Leur particularité appréciable est d'être utilisables **partout** : dans le code VBA, bien sûr, mais aussi dans les expressions de requête, les sources de contrôle de formulaires et d'états, les valeurs par défaut, les règles de validation. Un zone de texte peut ainsi afficher `=DLookup(...)` sans la moindre ligne de code.

---

## La liste et la signature commune

Toutes les fonctions de domaine partagent la même structure à trois arguments :

```
Fonction(Expression, Domaine, [Critère])
```

L'**`Expression`** est le champ (ou l'expression) sur lequel opérer, fourni sous forme de **chaîne**. Le **`Domaine`** est le nom de la table ou de la requête, également en chaîne. Le **`Critère`**, facultatif, est un fragment de clause `WHERE` (sans le mot `WHERE`) qui filtre le domaine.

| Fonction | Rôle |
|---|---|
| `DLookup` | Renvoie la valeur d'un champ (premier enregistrement correspondant) |
| `DCount` | Compte les enregistrements |
| `DSum` | Somme d'un champ |
| `DAvg` | Moyenne d'un champ |
| `DMax` | Valeur maximale |
| `DMin` | Valeur minimale |
| `DFirst` | Première valeur (ordre arbitraire) |
| `DLast` | Dernière valeur (ordre arbitraire) |
| `DStDev` / `DStDevP` | Écart-type (échantillon / population) |
| `DVar` / `DVarP` | Variance (échantillon / population) |

Quelques exemples illustrent leur concision :

```vba
maVille  = DLookup("Ville", "Clients", "IDClient = 42")      ' valeur d'un champ
nb       = DCount("*", "Commandes", "IDClient = 42")         ' comptage
total    = DSum("Montant", "Commandes", "Annee = 2024")      ' somme
dernier  = DMax("DateCommande", "Commandes")                 ' max sur tout le domaine
nbTotal  = DCount("*", "Clients")                            ' sans critère
```

Notez que `DCount("*", …)` compte **toutes les lignes**, tandis que `DCount("[Champ]", …)` ne compte que les enregistrements où le champ n'est pas `Null`.

---

## Le critère : une chaîne SQL comme les autres

Voici le point d'intégration majeur avec le reste du chapitre. Le **critère est un fragment de SQL construit sous forme de chaîne** — il est donc soumis **exactement** aux mêmes règles que tout SQL dynamique (sections 11.5 à 11.7).

Concrètement, les valeurs insérées dans le critère doivent être correctement **formatées** : apostrophes pour le texte, dièses et format ISO pour les dates, point décimal pour les nombres. On réutilise donc naturellement les fonctions utilitaires de la section 11.7 :

```vba
' Date et texte correctement formatés dans le critère
total = DSum("Montant", "Commandes", "DateCommande >= " & SQLDate(dateDebut))
ville = DLookup("Ville", "Clients", "Nom = " & SQLTexte(Me.txtNom))
```

Et surtout, le critère est **vulnérable à l'injection** si l'on y insère une saisie utilisateur brute, comme nous l'avons signalé en 11.5. La même discipline s'impose : **formater les valeurs, valider les identifiants**. Quant aux jokers d'un éventuel `LIKE` dans le critère, ils suivent les règles vues en 11.8 (généralement `*`/`?`, le critère étant évalué par le service d'expression d'Access).

---

## Gérer l'absence de résultat : `Null` (et `0` pour `DCount`)

Comportement essentiel à connaître : lorsqu'**aucun enregistrement ne correspond** au critère, la plupart des fonctions de domaine renvoient **`Null`** — `DLookup`, `DSum`, `DMax`, `DMin`, `DAvg`, `DFirst`, `DLast`. Seule `DCount` fait exception et renvoie **`0`** (elle compte, et zéro est un comptage valide).

Comme ces fonctions retournent un `Variant`, affecter un résultat `Null` à une variable non-Variant déclenche une erreur. On neutralise donc systématiquement le `Null` avec `Nz` :

```vba
Dim total As Currency
total = Nz(DSum("Montant", "Commandes", "Annee = 2024"), 0)   ' 0 si aucune commande
```

Ce réflexe — envelopper de `Nz` toute fonction de domaine susceptible de ne rien trouver — évite une catégorie entière d'erreurs.

---

## `DLookup` renvoie le premier enregistrement trouvé

`DLookup` mérite une mise en garde : si le critère correspond à **plusieurs** enregistrements, elle en renvoie **un seul, arbitrairement** — sans garantie sur lequel, car elle ne trie pas. Elle est donc adaptée aux recherches sur une valeur **unique** (typiquement par clé primaire).

Si le besoin est d'obtenir « le plus récent » ou « le plus grand », `DLookup` n'est pas l'outil : aucun argument de tri n'existe. On utilise alors `DMax`/`DMin` sur le champ pertinent, ou un recordset ouvert avec un `ORDER BY` explicite (chapitres 9 et 10).

---

## Le piège des performances

C'est le revers de la commodité, et il est sérieux. Les fonctions de domaine sont **lentes** : chaque appel ouvre et parcourt le domaine. L'impact, négligeable pour un appel isolé, devient **désastreux** dans trois situations.

Dans une **boucle** d'abord : appeler `DLookup` ou `DSum` à chaque itération multiplie les parcours du domaine et effondre les performances.

```vba
' ❌ ANTI-PATRON : une fonction de domaine par tour de boucle
Do While Not rs.EOF
    montant = DSum("Montant", "Commandes", "IDClient = " & rs!IDClient)  ' très lent
    rs.MoveNext
Loop
```

Dans une **requête**, comme colonne calculée évaluée **ligne par ligne** : un `DLookup` placé dans un `SELECT` s'exécute une fois par enregistrement renvoyé — catastrophique sur un grand résultat. Dans la **source de contrôle d'un formulaire continu**, enfin, où la fonction est réévaluée pour chaque enregistrement visible.

La parade est systématique : pour des besoins **répétés ou par ligne**, on remplace les fonctions de domaine par une **jointure**, une **sous-requête** (section 11.11) ou un **recordset** ouvert une seule fois. Une jointure qui calcule toutes les sommes en une passe est incomparablement plus rapide qu'un `DSum` appelé par enregistrement. Ces considérations rejoignent le chapitre 18 sur la performance.

---

## Fonctions de domaine ou SQL ?

La décision se résume à quelques cas. Pour une **valeur unique ponctuelle** — un comptage, une recherche par clé, une somme isolée —, surtout dans un **contexte non-code** (source de contrôle, valeur par défaut, règle de validation), la fonction de domaine est concise et appropriée. Pour des **agrégats groupés** par catégorie, elle est inadaptée : une fonction de domaine ne renvoie qu'une valeur, là où un `GROUP BY` produit un résultat par groupe — il faut alors une vraie requête. Et pour tout besoin **répété ou ligne par ligne**, on privilégie la **jointure, la sous-requête ou le recordset**, bien plus performants.

En somme : **commodité pour l'isolé, requête pour le groupé, jointure pour le répété**.

---

## En résumé

Les **fonctions de domaine** (`DLookup`, `DCount`, `DSum`, `DAvg`, `DMax`, `DMin`, `DFirst`, `DLast`, `DStDev`/`DStDevP`, `DVar`/`DVarP`) renvoient une **valeur unique** sur une table ou une requête, selon la signature commune `Fonction(Expression, Domaine, [Critère])`, et s'utilisent aussi bien en code que dans les expressions et sources de contrôle. Leur **critère est du SQL dynamique** : il obéit aux règles de formatage (section 11.7) et de sécurité (section 11.5). On gère l'**absence de résultat** par `Nz` (`Null`, sauf `DCount` qui renvoie `0`), et l'on garde à l'esprit que `DLookup` renvoie un enregistrement **arbitraire** en cas de correspondances multiples. Leur grand piège est la **performance** : commodes pour une valeur isolée, elles sont à proscrire dans les boucles, les colonnes calculées par ligne et les formulaires continus, où une **jointure, une sous-requête ou un recordset** s'imposent.

Cette mention des sous-requêtes nous amène à la dernière section du chapitre, consacrée aux constructions SQL les plus élaborées : les requêtes d'union, les sous-requêtes et les expressions complexes — les outils pour exprimer en SQL ce que les fonctions de domaine et les requêtes simples ne suffisent plus à formuler.


⏭️ [11.11. Requêtes Union, sous-requêtes et expressions complexes](/11-sql-access-vba/11-union-sous-requetes.md)
