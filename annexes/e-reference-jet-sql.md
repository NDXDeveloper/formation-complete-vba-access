🔝 Retour au [Sommaire](/SOMMAIRE.md)

# E. Référence Jet SQL — syntaxe et fonctions

Cette annexe récapitule la syntaxe du dialecte SQL d'Access (moteur Jet, devenu ACE) et ses fonctions intégrées. Elle complète le **chapitre 11**, et en particulier le **11.8** (spécificités et limites du dialecte Jet). Pour la correspondance avec T-SQL lors d'une migration, voir l'**annexe F**.

## Les deux modes SQL d'Access

Point fondamental : Access reconnaît **deux dialectes SQL**, et le comportement de certaines requêtes en dépend.

| Mode | Activé par | Caractères génériques | DDL étendu |
|---|---|---|---|
| **ANSI-89 (Jet traditionnel)** | Exécution via **DAO**, interface Access par défaut | `*` `?` `#` | Non |
| **ANSI-92 (SQL-92)** | Exécution via **ADO / OLE DB**, ou option de base activée | `%` `_` | Oui (CHECK, CREATE VIEW, CREATE PROCEDURE…) |

> ⚠️ **Conséquence directe.** Une même requête `LIKE` ne renvoie pas le même résultat selon qu'elle est lancée par DAO (`LIKE "Dur*"`) ou par ADO (`LIKE "Dur%"`). Toujours adapter les jokers au mode utilisé (chapitre 11.6).

---

## Délimiteurs, littéraux et identifiants

| Élément | Syntaxe Jet | Exemple |
|---|---|---|
| Chaîne de caractères | guillemets `"` ou apostrophes `'` | `"Rouen"` ou `'Rouen'` |
| Concaténation | `&` (sûr avec Null) ou `+` | `Nom & " " & Prenom` |
| Date / heure | encadrée de `#` | `#2026-05-29#` |
| Booléen | `True`/`False`, `Yes`/`No`, `-1`/`0` | `Actif = True` |
| Valeur nulle | `Null` | `IS NULL` |
| Identifiant avec espace/caractère spécial | crochets `[ ]` | `[Nom du client]` |

> ⚠️ **Format des dates.** Les littéraux date sont toujours encadrés de `#`. Pour rester indépendant de la locale, utiliser le **format ISO** `#aaaa-mm-jj#` (ou `#aaaa-mm-jj hh:nn:ss#`). Avec des `/`, Jet attend le format **américain** `#mm/jj/aaaa#`, source d'erreurs en environnement français (chapitre 11.7).

---

## Opérateurs

- **Comparaison** : `=`, `<>`, `<`, `<=`, `>`, `>=`
- **Logiques** : `AND`, `OR`, `NOT`, `XOR`, `EQV`, `IMP`
- **Plage** : `BETWEEN valeur1 AND valeur2`
- **Liste** : `IN (valeur1, valeur2, …)`
- **Motif** : `LIKE`
- **Nullité** : `IS NULL`, `IS NOT NULL`
- **Arithmétiques** : `+`, `-`, `*`, `/`, `\` (division entière), `Mod`, `^` (puissance)

---

## Caractères génériques (LIKE)

| Rôle | Mode ANSI-89 (DAO) | Mode ANSI-92 (ADO) |
|---|---|---|
| Toute suite de caractères | `*` | `%` |
| Un caractère quelconque | `?` | `_` |
| Un chiffre quelconque | `#` | *(non disponible)* |
| Un caractère d'une liste | `[a-m]` | `[a-m]` |
| Un caractère hors liste | `[!a-m]` | `[!a-m]` |

```sql
-- Mode Jet (DAO)
SELECT * FROM Clients WHERE Nom LIKE "Dur*";
-- Mode ANSI-92 (ADO)
SELECT * FROM Clients WHERE Nom LIKE "Dur%";
```

---

## Instruction SELECT

```sql
SELECT [DISTINCT | DISTINCTROW] [TOP n [PERCENT]] liste_champs
FROM sources
[WHERE condition]
[GROUP BY champs]
[HAVING condition_agrégat]
[ORDER BY champs [ASC | DESC]];
```

- **`DISTINCT`** élimine les doublons sur les colonnes sélectionnées ; **`DISTINCTROW`** (propre à Access) élimine les doublons sur l'ensemble des colonnes des tables sous-jacentes — utile dans les jointures.
- **`TOP n`** / **`TOP n PERCENT`** limite le nombre de lignes (en cas d'ex æquo, `TOP` peut renvoyer plus de `n` lignes).
- **Alias** : `SELECT Prix * Quantite AS Total`.

### Jointures

```sql
SELECT c.Nom, cmd.DateCommande
FROM Clients AS c
INNER JOIN Commandes AS cmd ON c.ID = cmd.ClientID;
```

- Types pris en charge : `INNER JOIN`, `LEFT JOIN`, `RIGHT JOIN`.
- ⚠️ **Pas de `FULL OUTER JOIN`** en Jet : le simuler par l'`UNION` d'un `LEFT JOIN` et d'un `RIGHT JOIN`.
- ⚠️ **Jointures multiples** : Jet **exige des parenthèses** explicites pour enchaîner plusieurs jointures.

```sql
FROM (Clients AS c
INNER JOIN Commandes AS cmd ON c.ID = cmd.ClientID)
INNER JOIN Lignes AS l ON cmd.ID = l.CommandeID;
```

---

## Requêtes action

```sql
-- Insertion de valeurs
INSERT INTO Clients (Nom, Ville) VALUES ("Durand", "Rouen");

-- Insertion à partir d'une requête
INSERT INTO Archive (Nom, Ville)
SELECT Nom, Ville FROM Clients WHERE Actif = False;

-- Mise à jour (jointure autorisée en Jet)
UPDATE Clients SET Remise = 0.1 WHERE CA > 10000;

-- Suppression
DELETE FROM Clients WHERE Actif = False;     -- DELETE * FROM … est aussi accepté

-- Création de table à partir d'une requête (make-table)
SELECT Nom, Ville INTO ClientsRouen
FROM Clients WHERE Ville = "Rouen";
```

---

## DDL — définition des données

```sql
CREATE TABLE Clients (
    ID            COUNTER       PRIMARY KEY,
    Nom           TEXT(50)      NOT NULL,
    Email         TEXT(100),
    DateInscr     DATETIME,
    Solde         CURRENCY,
    Actif         BIT
);

ALTER TABLE Clients ADD COLUMN Telephone TEXT(20);
ALTER TABLE Clients ALTER COLUMN Nom TEXT(100);
ALTER TABLE Clients DROP COLUMN Telephone;
ALTER TABLE Commandes ADD CONSTRAINT fkClient
    FOREIGN KEY (ClientID) REFERENCES Clients (ID);

CREATE UNIQUE INDEX idxEmail ON Clients (Email);
DROP INDEX idxEmail ON Clients;

DROP TABLE ClientsRouen;
```

### Types de données en DDL

| Type Jet (DDL) | Synonymes courants | Type Access (interface) | Type DAO/VBA |
|---|---|---|---|
| `BIT` | `LOGICAL`, `YESNO` | Oui/Non | Boolean |
| `BYTE` | `INTEGER1` | Octet | Byte |
| `SHORT` | `SMALLINT`, `INTEGER2` | Entier | Integer |
| `LONG` | `INTEGER`, `INT`, `INTEGER4` | Entier long | Long |
| `SINGLE` | `REAL`, `FLOAT4` | Réel simple | Single |
| `DOUBLE` | `FLOAT`, `FLOAT8`, `NUMBER` | Réel double | Double |
| `CURRENCY` | `MONEY` | Monétaire | Currency |
| `DATETIME` | `DATE`, `TIME` | Date/Heure | Date |
| `TEXT(n)` | `CHAR`, `VARCHAR`, `STRING` | Texte court | String |
| `MEMO` | `LONGTEXT`, `NOTE` | Texte long | String |
| `COUNTER` | `AUTOINCREMENT` | NuméroAuto | Long |
| `GUID` | `UNIQUEIDENTIFIER` | Identif. réplication | — |
| `DECIMAL(p,s)` | `NUMERIC` | Décimal | — |
| `LONGBINARY` | `OLEOBJECT`, `IMAGE`, `GENERAL` | Objet OLE | — |

> Contraintes possibles : `PRIMARY KEY`, `UNIQUE`, `NOT NULL`, `FOREIGN KEY … REFERENCES …`. Les contraintes `CHECK` ne sont disponibles qu'en **mode ANSI-92** (donc via ADO/OLE DB).

---

## Requêtes spéciales Access

### Analyse croisée (TRANSFORM / PIVOT)

```sql
TRANSFORM Sum(Montant) AS TotalVentes
SELECT Vendeur
FROM Ventes
GROUP BY Vendeur
PIVOT Format(DateVente, "mmm");
```

### Union

```sql
SELECT Nom FROM Clients
UNION             -- UNION ALL pour conserver les doublons
SELECT Nom FROM Prospects;
```

### Requête sur une base externe (clause IN)

```sql
SELECT * FROM Clients IN 'C:\Data\Externe.accdb';
```

### Requête paramétrée (clause PARAMETERS)

```sql
PARAMETERS [Date début] DateTime, [Date fin] DateTime;
SELECT * FROM Commandes
WHERE DateCommande BETWEEN [Date début] AND [Date fin];
```

---

## Fonctions intégrées utilisables en SQL

Ces fonctions, issues du moteur d'expression Access/VBA, sont utilisables dans les requêtes Jet.

### Texte

| Fonction | Rôle |
|---|---|
| `Left(ch, n)`, `Right(ch, n)`, `Mid(ch, début, n)` | Extraction de sous-chaîne |
| `Len(ch)` | Longueur |
| `Trim`, `LTrim`, `RTrim` | Suppression des espaces |
| `UCase`, `LCase` | Majuscules / minuscules |
| `Replace(ch, cherche, remplace)` | Remplacement |
| `InStr([début,] ch, cherche)` | Position d'une sous-chaîne |
| `Format(valeur, modèle)` | Mise en forme (texte, nombre, date) |
| `Asc(ch)`, `Chr(code)` | Caractère ↔ code |

### Numériques

| Fonction | Rôle |
|---|---|
| `Abs(n)` | Valeur absolue |
| `Int(n)`, `Fix(n)` | Partie entière (gestion des négatifs différente) |
| `Round(n [, décimales])` | Arrondi |
| `Sgn(n)` | Signe (-1, 0, 1) |
| `Sqr(n)` | Racine carrée |
| `Val(ch)` | Conversion texte → nombre |

### Date et heure

| Fonction | Rôle |
|---|---|
| `Date()`, `Now()`, `Time()` | Date du jour / date-heure / heure |
| `DateAdd(intervalle, n, date)` | Ajout d'un intervalle |
| `DateDiff(intervalle, d1, d2)` | Écart entre deux dates |
| `DatePart(intervalle, date)` | Extraction d'une composante |
| `DateSerial(a, m, j)`, `TimeSerial(h, m, s)` | Construction d'une date/heure |
| `Year`, `Month`, `Day`, `Hour`, `Minute`, `Second` | Composantes |
| `Weekday(date)`, `WeekdayName`, `MonthName` | Jour/nom du jour, nom du mois |

> Codes d'intervalle (`DateAdd`/`DateDiff`/`DatePart`) : `"yyyy"` année, `"q"` trimestre, `"m"` mois, `"ww"` semaine, `"d"` jour, `"h"` heure, `"n"` minute, `"s"` seconde.

### Conversion de type

`CBool`, `CByte`, `CCur`, `CDate`, `CDbl`, `CInt`, `CLng`, `CSng`, `CStr`, `CVar`, `CDec` — conversion explicite vers le type indiqué.

### Logique et gestion du Null

| Fonction | Rôle |
|---|---|
| `IIf(condition, siVrai, siFaux)` | Condition en ligne |
| `Choose(index, val1, val2, …)` | Sélection par position |
| `Switch(cond1, val1, cond2, val2, …)` | Première condition vraie |
| `Nz(valeur [, siNull])` | Remplace un `Null` (par `0` ou `""` par défaut) |
| `IsNull`, `IsNumeric`, `IsDate` | Tests de nature |

### Fonctions d'agrégation (SQL)

`Count`, `Sum`, `Avg`, `Min`, `Max`, `First`, `Last`, `StDev`, `StDevP`, `Var`, `VarP`.

> `First` et `Last` sont propres à Jet (premier/dernier enregistrement rencontré, non déterministe sans tri).

### Fonctions de domaine

`DLookup`, `DSum`, `DCount`, `DAvg`, `DMin`, `DMax`, `DFirst`, `DLast`, `DStDev`, `DVar` — agrégation sur un domaine sans requête explicite (détail au chapitre 11.10).

---

## Points de vigilance

> ⚠️ **Jokers selon le mode.** `*`/`?`/`#` en mode Jet (DAO) ; `%`/`_` en mode ANSI-92 (ADO). C'est la première cause de filtres `LIKE` « qui ne renvoient rien ».

> ⚠️ **Pas de `FULL OUTER JOIN`** et **parenthèses obligatoires** pour les jointures multiples.

> ⚠️ **Null contaminant.** Tout calcul ou concaténation avec un `Null` renvoie `Null`. Protéger avec `Nz()` et préférer `&` à `+` pour concaténer des chaînes.

> ⚠️ **Fonctions VBA et requêtes Pass-Through.** Les fonctions ci-dessus ne sont évaluées **que par Access**. Une requête Pass-Through, exécutée côté SQL Server, ne peut pas les utiliser : il faut alors les équivalents T-SQL (annexe F et chapitres 11.9, 23.3).

> ⚠️ **Commentaires.** Le SQL Jet **ne gère pas de manière fiable** les commentaires (`--`, `/* */`) dans les requêtes sauvegardées. Documenter le SQL côté VBA plutôt que dans la requête.

---

> 💡 **À retenir.** Le dialecte Jet diverge du SQL standard sur quelques points structurants : **deux modes** (jokers `*`/`?` ou `%`/`_`), **délimiteurs `#` pour les dates**, **concaténation par `&`**, **absence de `FULL OUTER JOIN`** et **parenthèses imposées** sur les jointures multiples. La syntaxe complète et les fonctions sont documentées dans l'aide du moteur ACE ; pour passer à T-SQL, l'annexe F fournit la table de correspondance.

⏭️ [F. Correspondance Jet SQL ↔ T-SQL (SQL Server)](/annexes/f-correspondance-jet-tsql.md)
