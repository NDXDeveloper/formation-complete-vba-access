🔝 Retour au [Sommaire](/SOMMAIRE.md)

# D. Chaînes de connexion ADO/ODBC pour les SGBD courants

Cette annexe rassemble des modèles de chaînes de connexion prêts à adapter pour les principaux systèmes de gestion de bases de données. Elle complète les **chapitres 10.2** (chaînes ADO pour Access), **10.9** (connexion à des bases externes) et **23.4** (architecture hybride Access + SQL Server via ODBC).

Ce sont des **gabarits de référence** : il suffit de remplacer les valeurs entre chevrons (`<serveur>`, `<base>`, `<utilisateur>`, `<motdepasse>`…) par les vôtres.

## Préalable : le pilote doit être installé

Une chaîne de connexion ne fonctionne que si le **pilote ou le fournisseur qu'elle nomme est réellement installé** sur le poste. Le nom indiqué dans `Driver={…}` ou `Provider=…` doit correspondre exactement à celui présent sur la machine.

Pour vérifier les pilotes ODBC installés : **Administrateur de sources de données ODBC** (`odbcad32.exe`), onglet *Pilotes*.

> ⚠️ **Piège 32/64 bits.** Le pilote ODBC doit avoir la **même architecture** qu'Access. Il existe deux administrateurs ODBC distincts : `C:\Windows\System32\odbcad32.exe` (64 bits) et `C:\Windows\SysWOW64\odbcad32.exe` (32 bits). Un pilote 64 bits est invisible pour un Access 32 bits, et inversement (chapitre 21.7).

## Anatomie d'une chaîne de connexion

Une chaîne est une suite de paires `clé=valeur` séparées par des points-virgules. Les clés sont insensibles à la casse et l'ordre est libre.

```
Cle1=valeur1;Cle2=valeur2;Cle3=valeur3;
```

Deux familles principales :

- **OLE DB** (pour ADO) : commence par `Provider=…`.
- **ODBC** : commence par `Driver={…}` (sans DSN) ou `DSN=…` (avec source configurée).

---

## 1. Access (fichiers .accdb et .mdb)

### OLE DB (ADO) — fournisseur ACE

```
Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Data\GestCom.accdb;
```

Base protégée par mot de passe :

```
Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Data\GestCom.accdb;Jet OLEDB:Database Password=<motdepasse>;
```

Ancien format .mdb (moteur Jet, **32 bits uniquement**) :

```
Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Data\Ancienne.mdb;
```

> `Microsoft.ACE.OLEDB.12.0` est le fournisseur le plus largement compatible (un redistribuable existe). Les versions récentes d'Office installent aussi `Microsoft.ACE.OLEDB.16.0`.

### ODBC — pilote Access

```
Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Data\GestCom.accdb;
```

Pour Excel (rappel utile) :

```
Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=C:\Data\Classeur.xlsx;
```

---

## 2. SQL Server

Le pilote ODBC actuel est **ODBC Driver 18 for SQL Server** ; la version 18.6.2.1 en est la dernière version stable (GA), et le pilote 17 reste maintenu. Côté OLE DB, le fournisseur recommandé est **Microsoft OLE DB Driver for SQL Server** (`MSOLEDBSQL`).

### OLE DB (ADO) — MSOLEDBSQL

Authentification Windows :

```
Provider=MSOLEDBSQL;Server=<serveur>;Database=<base>;Trusted_Connection=yes;
```

Authentification SQL Server :

```
Provider=MSOLEDBSQL;Server=<serveur>;Database=<base>;UID=<utilisateur>;PWD=<motdepasse>;
```

> Les versions récentes peuvent s'enregistrer sous `MSOLEDBSQL19`. Les anciens fournisseurs `SQLOLEDB` et `SQLNCLI11` (SQL Server Native Client) sont **obsolètes**.

### ODBC — Driver 18

Authentification Windows :

```
Driver={ODBC Driver 18 for SQL Server};Server=<serveur>;Database=<base>;Trusted_Connection=yes;Encrypt=yes;TrustServerCertificate=yes;
```

Authentification SQL Server :

```
Driver={ODBC Driver 18 for SQL Server};Server=<serveur>;Database=<base>;UID=<utilisateur>;PWD=<motdepasse>;Encrypt=yes;TrustServerCertificate=yes;
```

Instance nommée ou port spécifique :

```
Server=<serveur>\<instance>;      (instance nommée)
Server=<serveur>,1433;             (port explicite)
```

> ⚠️ **Chiffrement par défaut.** Depuis le pilote ODBC 18, `Encrypt` vaut `yes` par défaut (ce n'était pas le cas du pilote 17). Si le serveur utilise un certificat auto-signé, la connexion échoue tant que l'on n'ajoute pas `TrustServerCertificate=yes` (ou un certificat valide). C'est une cause d'échec très fréquente lors de la migration vers le pilote 18.

### DAO — table liée ou requête Pass-Through

Pour une table liée (`TableDef.Connect`) ou une requête Pass-Through (`QueryDef.Connect`), la chaîne ODBC est **préfixée par `ODBC;`** :

```
ODBC;DRIVER={ODBC Driver 18 for SQL Server};SERVER=<serveur>;DATABASE=<base>;Trusted_Connection=Yes;Encrypt=yes;TrustServerCertificate=yes;
```

---

## 3. PostgreSQL

Pilote ODBC : **psqlODBC** (« PostgreSQL Unicode » ou « PostgreSQL ANSI »).

```
Driver={PostgreSQL Unicode};Server=<serveur>;Port=5432;Database=<base>;Uid=<utilisateur>;Pwd=<motdepasse>;
```

Variante 64 bits selon l'installation :

```
Driver={PostgreSQL Unicode(x64)};Server=<serveur>;Port=5432;Database=<base>;Uid=<utilisateur>;Pwd=<motdepasse>;
```

En OLE DB, on passe généralement par le fournisseur OLE DB pour pilotes ODBC :

```
Provider=MSDASQL;Driver={PostgreSQL Unicode};Server=<serveur>;Port=5432;Database=<base>;Uid=<utilisateur>;Pwd=<motdepasse>;
```

---

## 4. MySQL

Pilote ODBC : **MySQL Connector/ODBC**. La branche actuelle est la 9.x (Connector/ODBC 9.6), mais le nom historique `MySQL ODBC 8.0 Unicode Driver` reste très répandu.

```
Driver={MySQL ODBC 8.0 Unicode Driver};Server=<serveur>;Port=3306;Database=<base>;User=<utilisateur>;Password=<motdepasse>;Option=3;
```

Avec un pilote récent et un jeu de caractères explicite :

```
Driver={MySQL ODBC 9.0 Unicode Driver};Server=<serveur>;Port=3306;Database=<base>;User=<utilisateur>;Password=<motdepasse>;CharSet=utf8mb4;
```

> Le nom exact (`8.0`, `9.0`…, *Unicode* ou *ANSI*) doit correspondre au pilote installé, visible dans l'administrateur ODBC. Privilégier la version *Unicode*.

---

## 5. MariaDB

Pilote ODBC : **MariaDB Connector/ODBC**.

```
Driver={MariaDB ODBC 3.2 Driver};Server=<serveur>;Port=3306;Database=<base>;Uid=<utilisateur>;Pwd=<motdepasse>;
```

> Le pilote ODBC MySQL fonctionne souvent aussi avec MariaDB, les deux moteurs partageant une large compatibilité protocolaire (chapitre 23.7).

---

## 6. Oracle

### OLE DB — Oracle Provider for OLE DB

```
Provider=OraOLEDB.Oracle;Data Source=<NomTNS>;User Id=<utilisateur>;Password=<motdepasse>;
```

Connexion directe (EZConnect, sans entrée TNS) :

```
Provider=OraOLEDB.Oracle;Data Source=//<serveur>:1521/<service>;User Id=<utilisateur>;Password=<motdepasse>;
```

### ODBC — pilote Oracle

```
Driver={Oracle in OraClient19Home1};Dbq=<NomTNS>;Uid=<utilisateur>;Pwd=<motdepasse>;
```

> Le nom du pilote (`Oracle in OraClient19Home1`…) dépend de la version du client Oracle installé. L'ancien pilote « Microsoft ODBC for Oracle » et le fournisseur `MSDAORA` sont **obsolètes**.

---

## 7. SQLite

Pilote ODBC tiers (SQLite ODBC Driver) :

```
Driver={SQLite3 ODBC Driver};Database=C:\Data\base.db;
```

---

## DSN ou sans DSN (DSN-less)

| Approche | Exemple | Caractéristiques |
|---|---|---|
| **DSN** (machine/utilisateur) | `DSN=<nomDSN>;UID=<u>;PWD=<p>;` | Configuré dans l'administrateur ODBC. Nom court et stable, mais à créer sur **chaque poste**. |
| **DSN fichier** | `FileDSN=C:\Conf\maSource.dsn;` | Paramètres stockés dans un fichier `.dsn` partageable. |
| **Sans DSN (DSN-less)** | Chaîne `Driver={…};…` complète dans le code | Aucune configuration par poste à prévoir — **recommandé pour le déploiement**, à condition que le pilote soit installé. |

---

## Où ces chaînes s'utilisent en VBA Access (rappels)

- **ADO** — ouverture d'une connexion (chapitre 10.3) :
  ```vba
  Dim cn As Object: Set cn = CreateObject("ADODB.Connection")
  cn.Open "Provider=MSOLEDBSQL;Server=SRV1;Database=GestCom;Trusted_Connection=yes;"
  ```
- **Table liée DAO** — propriété `Connect` puis `RefreshLink` (chapitres 12.7 et 21.4).
- **Requête Pass-Through** — propriété `Connect` du `QueryDef` (chapitre 11.9).
- **Liaison par `DoCmd`** :
  ```vba
  DoCmd.TransferDatabase acLink, "ODBC Database", _
      "ODBC;DRIVER={ODBC Driver 18 for SQL Server};SERVER=SRV1;DATABASE=GestCom;Trusted_Connection=Yes;Encrypt=yes;TrustServerCertificate=yes;", _
      acTable, "dbo.Clients", "Clients"
  ```

---

## Points de vigilance

> ⚠️ **Architecture 32/64 bits.** La première cause d'échec d'une connexion ODBC : un pilote dont l'architecture ne correspond pas à celle d'Access (chapitre 21.7).

> ⚠️ **Pilote ODBC 18 et chiffrement.** Penser à `Encrypt`/`TrustServerCertificate` (voir section SQL Server) lors du passage au pilote 18.

> ⚠️ **Mots de passe en clair.** La chaîne `Connect` d'une table liée est stockée **en clair** dans la base front-end. Préférer l'authentification Windows (`Trusted_Connection=Yes`) ou une saisie à l'exécution ; ne pas y inscrire de mot de passe sensible (chapitre 20.8).

> 💡 **Valeurs contenant des caractères spéciaux.** Si un mot de passe contient `;` ou `}`, encadrer la valeur d'accolades : `PWD={mon;mot}de}passe}` peut nécessiter un échappement particulier — tester, et privilégier des valeurs simples ou l'authentification intégrée.

---

> 💡 **À retenir.** Une chaîne se résume à trois informations : **quel pilote/fournisseur**, **quel serveur/fichier**, **quelle authentification**. Le reste (port, chiffrement, options) est facultatif ou propre au moteur. Pour un référentiel exhaustif et tenu à jour de toutes les variantes, le site **connectionstrings.com** fait autorité. Les noms exacts des pilotes installés se vérifient toujours dans l'administrateur ODBC (`odbcad32.exe`).

⏭️ [E. Référence Jet SQL — syntaxe et fonctions](/annexes/e-reference-jet-sql.md)
