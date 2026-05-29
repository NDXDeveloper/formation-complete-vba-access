🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11.4. Requêtes DDL : CREATE TABLE, ALTER TABLE, DROP TABLE

Les requêtes action de la section précédente agissaient sur les **données** — elles ajoutent, modifient ou suppriment des enregistrements. Le SQL permet aussi d'agir sur la **structure** de la base elle-même : créer des tables, ajouter ou modifier des colonnes, définir des index et des contraintes, supprimer des objets. C'est le rôle du **DDL** (*Data Definition Language*, langage de définition de données), par opposition au DML (*Data Manipulation Language*) qui manipule le contenu.

Trois instructions structurent ce langage : **`CREATE`** (créer), **`ALTER`** (modifier) et **`DROP`** (supprimer). Comme les requêtes action, elles s'exécutent via `CurrentDb.Execute` ou `DoCmd.RunSQL` (section 11.1). Une question de fond traverse cependant toute la section : le DDL-SQL n'est pas le seul moyen de modifier la structure par code — DAO propose une approche objet (chapitre 12). Nous verrons quand préférer l'un à l'autre.

---

## CREATE TABLE : créer une table

L'instruction `CREATE TABLE` définit une nouvelle table, ses colonnes, leurs types et leurs contraintes :

```sql
CREATE TABLE Clients (
    IDClient    AUTOINCREMENT PRIMARY KEY,
    NomClient   TEXT(50) NOT NULL,
    Ville       TEXT(50),
    DateInscription DATETIME,
    Actif       YESNO,
    Solde       CURRENCY
);
```

Exécutée depuis VBA :

```vba
CurrentDb.Execute "CREATE TABLE Clients (" & _
    "IDClient AUTOINCREMENT PRIMARY KEY, " & _
    "NomClient TEXT(50) NOT NULL, " & _
    "Ville TEXT(50), Actif YESNO);", dbFailOnError
```

### Les types de données Jet SQL

Les noms de types employés en DDL Jet diffèrent de ceux affichés dans l'interface d'Access. Le tableau suivant établit la correspondance.

| Type Access | Mot-clé Jet SQL | Remarque |
|---|---|---|
| Texte court | `TEXT(n)` | `n` = longueur (synonymes `CHAR`, `VARCHAR`) |
| Texte long (Mémo) | `MEMO` (`LONGTEXT`) | |
| Octet | `BYTE` | |
| Entier (16 bits) | `SHORT` (`SMALLINT`) | l'« Entier » d'Access |
| Entier long (32 bits) | `LONG` (`INTEGER`, `INT`) | l'« Entier long » d'Access |
| Réel simple | `SINGLE` | |
| Réel double | `DOUBLE` | |
| Monétaire | `CURRENCY` | |
| Date/Heure | `DATETIME` | |
| Oui/Non | `YESNO` (`BIT`, `LOGICAL`) | |
| NuméroAuto | `AUTOINCREMENT` (`COUNTER`) | germe/pas : `AUTOINCREMENT(1,1)` |
| Identifiant réplication | `GUID` | |

> ⚠️ **Piège des noms d'entiers** : en Jet SQL, le mot-clé `INTEGER` correspond à l'**Entier long** (32 bits) d'Access, et non à l'« Entier » (16 bits). Pour ce dernier, il faut écrire `SHORT`. Cette inversion par rapport à l'intuition est source de confusion : préférez les mots-clés explicites **`SHORT`** (16 bits) et **`LONG`** (32 bits) pour lever toute ambiguïté.

### Clés et contraintes

`CREATE TABLE` accepte plusieurs contraintes. **`PRIMARY KEY`** définit la clé primaire, **`NOT NULL`** rend une colonne obligatoire, **`UNIQUE`** interdit les doublons, et **`DEFAULT`** fixe une valeur par défaut. On peut nommer les contraintes avec la clause **`CONSTRAINT`** et définir des **clés étrangères** :

```sql
CREATE TABLE Commandes (
    IDCommande AUTOINCREMENT PRIMARY KEY,
    IDClient   LONG NOT NULL,
    Montant    CURRENCY,
    CONSTRAINT fk_client FOREIGN KEY (IDClient) REFERENCES Clients (IDClient)
);
```

(La clause `DEFAULT` et certaines contraintes avancées peuvent nécessiter le mode **ANSI-92** ou l'exécution via le fournisseur ACE — voir plus bas.)

---

## ALTER TABLE : modifier une table existante

`ALTER TABLE` fait évoluer une table déjà créée. On y **ajoute**, **modifie** ou **supprime** des colonnes :

```sql
ALTER TABLE Clients ADD COLUMN Email TEXT(100);     -- ajouter une colonne
ALTER TABLE Clients ALTER COLUMN Ville TEXT(80);    -- changer le type/la taille
ALTER TABLE Clients DROP COLUMN Email;              -- supprimer une colonne
```

On y gère également les **contraintes et clés** :

```sql
-- Ajouter une contrainte d'unicité
ALTER TABLE Clients ADD CONSTRAINT idx_nom UNIQUE (NomClient);

-- Ajouter une clé étrangère
ALTER TABLE Commandes ADD CONSTRAINT fk_client
    FOREIGN KEY (IDClient) REFERENCES Clients (IDClient);

-- Supprimer une contrainte
ALTER TABLE Clients DROP CONSTRAINT idx_nom;
```

---

## CREATE INDEX et DROP INDEX

Les index, déterminants pour les performances (chapitre 18.6), se créent et se suppriment aussi par DDL :

```sql
CREATE INDEX idx_ville ON Clients (Ville);              -- index simple
CREATE UNIQUE INDEX idx_email ON Clients (Email);       -- index unique
CREATE INDEX idx_nom ON Clients (Nom, Prenom);          -- index multi-colonnes
DROP INDEX idx_ville ON Clients;                        -- suppression
```

Des options complémentaires existent (`WITH PRIMARY` pour un index de clé primaire, `WITH DISALLOW NULL`, `WITH IGNORE NULL`) pour affiner le comportement de l'index.

---

## DROP TABLE : supprimer une table

`DROP TABLE` supprime **intégralement** une table — sa structure **et** ses données :

```sql
DROP TABLE ClientsTemp;
```

```vba
CurrentDb.Execute "DROP TABLE ClientsTemp;", dbFailOnError
```

> ⚠️ **Opération irréversible** : `DROP TABLE` détruit définitivement la table et tout son contenu. Contrairement à un `DELETE`, il n'existe pas de filet de sécurité, et — comme nous le verrons — le DDL ne se laisse pas annuler par une transaction. À manier avec une extrême prudence, jamais à l'aveugle.

---

## Exécuter du DDL : particularités à connaître

L'exécution du DDL présente trois spécificités importantes.

**Le mode ANSI-92.** Access connaît deux modes de syntaxe SQL : ANSI-89 (historique, par défaut dans l'interface) et ANSI-92 (plus complet). Certaines fonctionnalités DDL avancées — clauses `DEFAULT`, contraintes `CHECK`, certains types — ne sont pleinement disponibles qu'en mode ANSI-92 ou lorsqu'on exécute le DDL via le fournisseur ACE/OLEDB (donc en ADO). Sur un moteur ACE moderne, `CurrentDb.Execute` traite néanmoins la majorité du DDL courant sans difficulté.

**Le DDL n'est pas transactionnel.** C'est une différence majeure avec le DML. Là où un `UPDATE` ou un `DELETE` peuvent être annulés au sein d'une transaction (chapitre 14), les opérations structurelles (`CREATE`, `ALTER`, `DROP`) sont **validées immédiatement** et **ne peuvent pas être annulées** de façon fiable par un `Rollback`. On ne peut donc pas compter sur une transaction pour « défaire » un `DROP TABLE` malencontreux.

**Les objets en cours d'utilisation.** Modifier ou supprimer une table utilisée par un formulaire ouvert, une requête en cours ou un autre utilisateur provoque une erreur (table verrouillée ou en cours d'utilisation). Il faut s'assurer qu'aucun objet ne dépend de la table avant d'en altérer la structure.

---

## Vérifier l'existence avant de créer ou de supprimer

Pour éviter les erreurs *« la table existe déjà »* ou *« table introuvable »*, on **teste l'existence** de l'objet au préalable, en inspectant la collection `TableDefs` de DAO (chapitre 12.4) ou la collection `AllTables` (chapitre 4.4). Une fonction utilitaire encapsule avantageusement ce contrôle :

```vba
Public Function TableExiste(ByVal nomTable As String) As Boolean
    Dim td As DAO.TableDef
    On Error Resume Next
    Set td = CurrentDb.TableDefs(nomTable)
    TableExiste = (Err.Number = 0)
    On Error GoTo 0
End Function

' Usage : supprimer une table seulement si elle existe
If TableExiste("ClientsTemp") Then
    CurrentDb.Execute "DROP TABLE ClientsTemp;", dbFailOnError
End If
```

---

## DDL-SQL ou objets DAO (TableDef) ?

Voici l'arbitrage central de cette section. Modifier la structure par code peut se faire de deux façons : par **DDL-SQL** (les instructions ci-dessus) ou via les **objets DAO** — `TableDef`, `Field`, `Index`, `Relation` — détaillés au chapitre 12.5.

Le **DDL-SQL** est **concis et portable** : sa syntaxe, proche du SQL standard, se transpose plus aisément vers d'autres SGBD, et une instruction `CREATE TABLE` se lit d'un coup d'œil. C'est l'idéal pour des scripts structurels ponctuels.

Les **objets DAO** offrent en revanche un **contrôle bien plus fin** des propriétés *propres à Access* que le DDL ne sait pas définir : légende (`Caption`), format d'affichage (`Format`), masque de saisie, description, règle et message de validation au niveau du champ, propriétés de liste de choix… Ces attributs nécessitent généralement de manipuler les `Properties` des objets DAO. Dès qu'une table doit être créée avec ces caractéristiques propres à Access, **DAO (chapitre 12.5) est le bon outil** ; pour une création ou une modification structurelle simple et portable, le DDL-SQL suffit et va plus vite à écrire.

---

## Quand utiliser le DDL en pratique ?

Dans une application Access, la structure est le plus souvent conçue dans l'interface graphique, une fois pour toutes. Le DDL par code intervient surtout dans des cas ciblés : la **création de tables temporaires** en cours de traitement, les **scripts d'installation ou de migration** qui bâtissent le schéma au déploiement, l'**ajout d'un index** pour améliorer les performances, ou de rares besoins d'**automatisation structurelle**. En revanche, modifier abondamment le schéma au fil de l'exécution dans une application en production est inhabituel et généralement déconseillé.

---

## Pièges et précautions

Plusieurs précautions s'imposent avec le DDL, en plus de celles déjà évoquées. Les **opérations destructrices** (`DROP TABLE`, `DROP COLUMN`) sont **irréversibles** et non annulables : on vérifie toujours l'existence et l'on encadre par une gestion d'erreurs. Les noms d'objets comportant des **espaces ou des mots réservés** doivent être encadrés de **crochets** `[ ]`. Le **piège des noms d'entiers** (`INTEGER` = 32 bits) reste un classique. Enfin, on garde à l'esprit l'**absence de transaction** : aucune marche arrière automatique n'est possible après une modification structurelle.

---

## En résumé

Le DDL agit sur la **structure** de la base, là où le DML agissait sur les données. `CREATE TABLE` crée une table avec ses colonnes, types et contraintes ; `ALTER TABLE` ajoute, modifie ou supprime colonnes et contraintes ; `CREATE INDEX` gère les index ; `DROP TABLE` supprime — irrémédiablement. Les types Jet SQL diffèrent des libellés d'Access, avec le piège notable d'`INTEGER` (32 bits, à distinguer de `SHORT`). Le DDL s'exécute via `Execute`/`RunSQL`, mais **n'est pas transactionnel** — aucun `Rollback` ne le défait — et certaines fonctionnalités exigent le mode ANSI-92. On vérifie l'existence des objets avant d'agir, et l'on arbitre entre **DDL-SQL** (concis, portable) et **objets DAO** (contrôle complet des propriétés Access, chapitre 12.5) selon le besoin.

Nous avons couvert l'exécution du SQL (11.1), la lecture (11.2), la manipulation des données (11.3) et de la structure (11.4) — toujours, jusqu'ici, avec des requêtes dont les valeurs étaient soit fixes, soit insérées directement dans la chaîne. C'est précisément cette insertion directe qui pose problème dès que des valeurs proviennent de l'extérieur. La section suivante aborde le **SQL paramétré** et la prévention des **injections SQL** — un impératif de sécurité.


⏭️ [11.5. SQL paramétré — prévention des injections SQL](/11-sql-access-vba/05-sql-parametree.md)
