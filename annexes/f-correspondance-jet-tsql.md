🔝 Retour au [Sommaire](/SOMMAIRE.md)

# F. Correspondance Jet SQL ↔ T-SQL (SQL Server)

Cette annexe met en regard le dialecte SQL d'Access (Jet/ACE) et le **Transact-SQL** de SQL Server. Elle prolonge l'**annexe E** (référence Jet SQL) et soutient le **chapitre 23.3** (adapter le code VBA après migration). Elle sert aussi de référence pour rédiger des **requêtes Pass-Through** exécutées côté serveur (chapitre 11.9).

Le sens d'usage est généralement **Jet → T-SQL** (migration d'Access vers SQL Server), mais la correspondance se lit dans les deux sens.

## Différences de syntaxe fondamentales

| Élément | Jet SQL (Access) | T-SQL (SQL Server) | Remarque |
|---|---|---|---|
| Chaîne de caractères | `"texte"` ou `'texte'` | `'texte'` seulement | En T-SQL, `"` désigne un **identifiant** (si `QUOTED_IDENTIFIER ON`) |
| Concaténation | `&` | `+` ou `CONCAT(...)` | `+` propage le `NULL` ; `CONCAT` l'ignore |
| Littéral date | `#2026-05-29#` | `'2026-05-29'` ou `'20260529'` | Pas de `#` en T-SQL ; format ISO conseillé |
| Identifiant avec espace | `[Nom du champ]` | `[Nom du champ]` | Crochets compatibles des deux côtés |
| Test de plage | `BETWEEN … AND …` | `BETWEEN … AND …` | Identique |
| Limiter les lignes | `TOP n` | `TOP (n)` | Parenthèses requises en T-SQL moderne ; `OFFSET … FETCH` pour la pagination |
| Booléen | `True`/`False` (`-1`/`0`) | `1`/`0` (type `bit`) | ⚠️ Access stocke `True = -1`, SQL Server `bit = 1` |
| Division entière | `\` | `/` (sur entiers) | `/` est entière si les deux opérandes sont entiers |
| Puissance | `^` | `POWER(x, y)` | — |
| Modulo | `Mod` | `%` | — |
| Commentaires | non fiables | `--` et `/* */` | T-SQL gère les deux |
| Paramètres | clause `PARAMETERS`, `[invite]` | `@variable` déclarée | Réécriture nécessaire |

---

## Jointures et structure des requêtes

| Construction | Jet | T-SQL |
|---|---|---|
| Jointures multiples | **Parenthèses obligatoires** | Parenthèses facultatives |
| Jointure externe complète | **Non supportée** (`UNION` de `LEFT`+`RIGHT`) | `FULL OUTER JOIN` natif |
| Élimination de doublons | `DISTINCT` ou `DISTINCTROW` | `DISTINCT` seul (`DISTINCTROW` n'existe pas) |
| Suppression | `DELETE * FROM …` ou `DELETE FROM …` | `DELETE FROM …` (pas de `*`) |
| Création de table par requête | `SELECT … INTO nouvelle FROM …` | `SELECT … INTO nouvelle FROM …` (table permanente) |
| Base externe | `… FROM t IN 'C:\base.accdb'` | Serveur lié ou `OPENROWSET` / `OPENQUERY` |

### Analyse croisée

```sql
-- Jet : TRANSFORM / PIVOT
TRANSFORM Sum(Montant)
SELECT Vendeur FROM Ventes
GROUP BY Vendeur
PIVOT Format(DateVente, "mmm");

-- T-SQL : opérateur PIVOT (ou agrégation conditionnelle)
SELECT Vendeur, [janv], [févr], [mars]
FROM (SELECT Vendeur, Montant, FORMAT(DateVente,'MMM') AS Mois FROM Ventes) src
PIVOT (SUM(Montant) FOR Mois IN ([janv],[févr],[mars])) p;
```

---

## Types de données

| Type Jet (DDL) | Type Access (interface) | Type T-SQL | Remarque |
|---|---|---|---|
| `BIT` / `YESNO` | Oui/Non | `bit` | `True (-1)` → `1` |
| `BYTE` | Octet | `tinyint` | — |
| `SHORT` | Entier | `smallint` | — |
| `LONG` / `INTEGER` | Entier long | `int` | — |
| `SINGLE` | Réel simple | `real` | — |
| `DOUBLE` | Réel double | `float` | — |
| `CURRENCY` | Monétaire | `money` ou `decimal(19,4)` | `decimal` souvent préféré |
| `DATETIME` | Date/Heure | `datetime2` (ou `date`, `time`) | `datetime2` recommandé |
| `TEXT(n)` | Texte court | `nvarchar(n)` | Access est **Unicode** → `nvarchar` |
| `MEMO` / `LONGTEXT` | Texte long | `nvarchar(max)` | — |
| `COUNTER` / `AUTOINCREMENT` | NuméroAuto | `int IDENTITY(1,1)` | — |
| `GUID` | Identif. réplication | `uniqueidentifier` | — |
| `DECIMAL(p,s)` | Décimal | `decimal(p,s)` / `numeric(p,s)` | — |
| `LONGBINARY` / `OLEOBJECT` | Objet OLE | `varbinary(max)` | `image` est obsolète |

---

## Fonctions

Correspondance des fonctions intégrées les plus courantes.

### Texte

| Jet / VBA | T-SQL | Remarque |
|---|---|---|
| `Left(s, n)` | `LEFT(s, n)` | — |
| `Right(s, n)` | `RIGHT(s, n)` | — |
| `Mid(s, début, n)` | `SUBSTRING(s, début, n)` | Position 1-based des deux côtés |
| `Len(s)` | `LEN(s)` | `LEN` ignore les espaces de fin ; `DATALENGTH` pour les octets |
| `Trim(s)` | `TRIM(s)` | Sinon `LTRIM(RTRIM(s))` sur versions anciennes |
| `UCase` / `LCase` | `UPPER` / `LOWER` | — |
| `Replace(s, a, b)` | `REPLACE(s, a, b)` | — |
| `InStr(début, s, cherche)` | `CHARINDEX(cherche, s, début)` | ⚠️ **Ordre des arguments inversé** |
| `String(n, car)` | `REPLICATE(car, n)` | Ordre inversé |
| `Space(n)` | `SPACE(n)` | — |
| `Asc` / `Chr` | `ASCII` / `CHAR` | `UNICODE` / `NCHAR` pour l'Unicode |
| `Format(v, modèle)` | `FORMAT(v, modèle)` ou `CONVERT` | `FORMAT` (2012+) est lent ; comportement différent |

### Numériques

| Jet / VBA | T-SQL | Remarque |
|---|---|---|
| `Abs(n)` | `ABS(n)` | — |
| `Int(n)` | `FLOOR(n)` | Arrondi vers le bas (négatifs compris) |
| `Fix(n)` | `ROUND(n, 0, 1)` (troncature) | ⚠️ `Fix` tronque vers zéro ≠ `FLOOR` pour les négatifs |
| `Round(n, d)` | `ROUND(n, d)` | ⚠️ `Round` VBA = arrondi **bancaire** ; `ROUND` T-SQL = arrondi **arithmétique** |
| `Sgn(n)` | `SIGN(n)` | — |
| `Sqr(n)` | `SQRT(n)` | — |
| `n ^ p` | `POWER(n, p)` | — |
| `Rnd()` | `RAND()` | — |

### Date et heure

| Jet / VBA | T-SQL | Remarque |
|---|---|---|
| `Date()` | `CAST(GETDATE() AS date)` | Date sans heure |
| `Now()` | `GETDATE()` (ou `SYSDATETIME()`) | — |
| `Time()` | `CAST(GETDATE() AS time)` | — |
| `DateAdd("d", n, d)` | `DATEADD(day, n, d)` | Codes d'intervalle différents (voir ci-dessous) |
| `DateDiff("d", d1, d2)` | `DATEDIFF(day, d1, d2)` | — |
| `DatePart("m", d)` | `DATEPART(month, d)` | — |
| `Year` / `Month` / `Day` | `YEAR` / `MONTH` / `DAY` | — |
| `Hour` / `Minute` / `Second` | `DATEPART(hour/minute/second, …)` | — |
| `DateSerial(a, m, j)` | `DATEFROMPARTS(a, m, j)` | (2012+) |
| `Weekday(d)` | `DATEPART(weekday, d)` | ⚠️ Dépend du réglage `SET DATEFIRST` |

**Codes d'intervalle** (`DateAdd`/`DateDiff` ↔ `DATEADD`/`DATEDIFF`) :

| Jet | T-SQL | Unité |
|---|---|---|
| `"yyyy"` | `year` | Année |
| `"q"` | `quarter` | Trimestre |
| `"m"` | `month` | Mois |
| `"d"` | `day` | Jour |
| `"ww"` | `week` | Semaine |
| `"h"` | `hour` | Heure |
| `"n"` | `minute` | ⚠️ Minute = `"n"` en Jet (`"m"` = mois) |
| `"s"` | `second` | Seconde |

### Conversion, Null et logique

| Jet / VBA | T-SQL | Remarque |
|---|---|---|
| `Nz(x, val)` | `ISNULL(x, val)` ou `COALESCE(x, val)` | Remplace un `NULL` |
| `IsNull(x)` | `x IS NULL` | ⚠️ **Piège majeur** : voir ci-dessous |
| `IIf(c, a, b)` | `IIF(c, a, b)` ou `CASE WHEN c THEN a ELSE b END` | `IIF` : 2012+ |
| `Switch(c1,v1,c2,v2…)` | `CASE WHEN c1 THEN v1 WHEN c2 THEN v2 … END` | — |
| `Choose(i, v1, v2…)` | `CHOOSE(i, v1, v2…)` | (2012+) |
| `CInt` / `CLng` / `CDbl` | `CAST(x AS int / bigint / float)` | `TRY_CAST` pour une conversion sûre |
| `CStr` / `CDate` | `CAST(x AS nvarchar / datetime2)` | `CONVERT` pour un format précis |
| `Val(s)` | `TRY_CAST(s AS float)` | — |
| `IsNumeric` / `IsDate` | `ISNUMERIC` / `ISDATE` ou `TRY_CAST` | `ISNUMERIC` est permissif |

> ⚠️ **`IsNull` (Jet) ≠ `ISNULL` (T-SQL).** En Jet, `IsNull(x)` **teste** si une valeur est nulle et renvoie un booléen. En T-SQL, `ISNULL(x, val)` **remplace** une valeur nulle par `val`. Ce sont deux fonctions différentes au nom identique. La traduction correcte est :  
> - test Jet `IsNull(x)` → T-SQL `x IS NULL`  
> - remplacement Jet `Nz(x, val)` → T-SQL `ISNULL(x, val)`

### Agrégats

| Jet / VBA | T-SQL | Remarque |
|---|---|---|
| `Count` / `Sum` / `Avg` / `Min` / `Max` | identiques | — |
| `StDev` / `StDevP` | `STDEV` / `STDEVP` | — |
| `Var` / `VarP` | `VAR` / `VARP` | — |
| `First` / `Last` | *(aucun équivalent)* | Utiliser `TOP (1)` + `ORDER BY`, ou fonctions de fenêtrage |

---

## Ce qui exige une réécriture

Certains éléments d'Access n'ont pas d'équivalent T-SQL direct et imposent une refonte :

- **Fonctions de domaine** (`DLookup`, `DSum`, `DCount`, `DMax`…) : réécrire en **sous-requêtes** ou `SELECT` scalaires. Elles n'existent pas en T-SQL.
- **`TRANSFORM` / `PIVOT`** (analyse croisée) : utiliser l'opérateur `PIVOT` de T-SQL ou une agrégation conditionnelle (`SUM(CASE WHEN … THEN … END)`).
- **`First` / `Last`** : remplacer par `TOP (1) … ORDER BY` ou `ROW_NUMBER()`.
- **Fonctions VBA personnalisées** appelées dans une requête : les recréer en **fonctions/procédures T-SQL**, ou effectuer le calcul côté client.
- **Requêtes paramétrées** : la clause `PARAMETERS` devient des variables `@param` déclarées (procédures stockées, requêtes paramétrées).
- **Base externe (`IN '…'`)** : remplacer par un **serveur lié** ou `OPENQUERY`/`OPENROWSET`.

---

## Points de vigilance

> ⚠️ **`ISNULL` : test ou remplacement ?** Le piège n°1 de la migration (voir encadré ci-dessus). `Nz` → `ISNULL`/`COALESCE` ; `IsNull` → `IS NULL`.

> ⚠️ **Arrondi.** `Round` (VBA) arrondit au pair le plus proche (arrondi bancaire) ; `ROUND` (T-SQL) arrondit à l'unité supérieure en cas de 0,5. Les résultats peuvent différer sur les valeurs `.5`.

> ⚠️ **Booléens.** `True` valant `-1` côté Access et `1` côté SQL Server, adapter les comparaisons : `WHERE Actif = True` (Access) devient `WHERE Actif = 1` (T-SQL).

> ⚠️ **`InStr` vs `CHARINDEX`.** L'ordre des arguments est inversé (`InStr(chaîne, cherche)` ↔ `CHARINDEX(cherche, chaîne)`). Erreur silencieuse fréquente.

> ⚠️ **`Int` vs `Fix`.** `Int` arrondit vers le bas (= `FLOOR`) ; `Fix` tronque vers zéro. Pour les nombres négatifs, ces deux fonctions diffèrent, et seul `Int` correspond à `FLOOR`.

> ⚠️ **Unicode.** Le type Texte d'Access est Unicode : le faire correspondre à `nvarchar`, pas `varchar`, pour préserver les caractères accentués et spéciaux.

> ⚠️ **Pass-Through.** Une requête Pass-Through s'exécute **entièrement** côté serveur : elle doit être écrite en T-SQL pur, sans fonction VBA ni joker Jet (chapitre 11.9).

---

> 💡 **À retenir.** La majorité des fonctions ont un équivalent quasi direct (`Left`/`LEFT`, `Year`/`YEAR`, `Sum`/`SUM`). Les vrais points de friction sont peu nombreux mais sournois : **`ISNULL` qui change de sens**, l'**ordre inversé de `CHARINDEX`**, l'**arrondi bancaire**, les **booléens `-1`/`1`**, et tout ce qui est **propre à Access** (fonctions de domaine, `TRANSFORM`, `First`/`Last`) qui doit être réécrit en sous-requêtes ou objets T-SQL. La démarche complète de migration est détaillée aux chapitres 23.1 à 23.4.

⏭️ [G. Préfixes de nommage Leszynski/Reddick pour Access](/annexes/g-prefixes-leszynski-reddick.md)
