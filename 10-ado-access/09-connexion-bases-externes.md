🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 10.9. Connexion ADO à des bases externes (SQL Server, Oracle, PostgreSQL, MySQL)

Nous arrivons à la vocation première d'ADO, celle qui justifie à elle seule l'existence de ce chapitre : se connecter à des **bases de données externes**. Tout ce que nous avons appris — connexion, recordset, command paramétré, recordsets déconnectés, diagnostic des erreurs — trouve ici son plein emploi. Car la promesse d'ADO, exposée dès la section 10.1, se vérifie maintenant : le **modèle objet reste identique** quelle que soit la source. Se connecter à SQL Server, Oracle, PostgreSQL ou MySQL ne change essentiellement qu'une chose — la **chaîne de connexion** — le reste du code demeurant celui des sections précédentes.

Cette section parcourt les principaux SGBD, leurs chaînes de connexion et leurs prérequis, puis aborde les sujets transversaux : DSN, procédures stockées, et le choix entre ADO et tables liées.

---

## Le principe : un modèle objet, des fournisseurs multiples

Rappelons l'architecture en couches d'ADO (section introductive du chapitre) : votre code parle aux objets ADO, qui s'appuient sur un **fournisseur** chargé de dialoguer avec la source réelle. Changer de base revient donc, pour l'essentiel, à désigner un autre fournisseur dans la chaîne de connexion.

La conséquence est puissante : une fois la connexion établie, **tout fonctionne à l'identique**. Un `Recordset` ouvert sur SQL Server se parcourt exactement comme un recordset ACE (sections 10.4-10.5), un `Command` paramétré s'emploie de la même façon (section 10.6), les erreurs se diagnostiquent via la même collection `Errors` (section 10.8). Le savoir-faire acquis est intégralement réutilisable. Seuls deux éléments diffèrent réellement d'un SGBD à l'autre : la **chaîne de connexion** et le **dialecte SQL** — ce dernier étant traité au chapitre 11 et au chapitre 23.

---

## Deux routes : OLE DB natif ou pont ODBC

Pour atteindre une base externe, deux voies existent.

La première passe par un **fournisseur OLE DB natif**, spécifique au SGBD. C'est généralement la meilleure option : performances optimales et prise en charge complète des fonctionnalités. SQL Server et Oracle disposent de tels fournisseurs.

La seconde passe par le **pont ODBC** : le fournisseur **`MSDASQL`** (*OLE DB Provider for ODBC*) permet à ADO d'exploiter n'importe quel pilote ODBC. C'est la solution universelle, indispensable lorsqu'aucun fournisseur OLE DB natif n'existe — cas de PostgreSQL et MySQL. Détail pratique : `MSDASQL` étant le **fournisseur par défaut** d'ADO, on peut même omettre la clé `Provider` dès lors que la chaîne contient des mots-clés ODBC (`Driver=` ou `DSN=`) ; ADO route alors automatiquement par MSDASQL.

Le tableau ci-dessous résume la voie recommandée pour chaque SGBD courant.

| SGBD | Voie recommandée | Fournisseur / Pilote | Prérequis côté client |
|---|---|---|---|
| **SQL Server** | OLE DB natif | `Provider=MSOLEDBSQL` | Microsoft OLE DB Driver for SQL Server |
| **Oracle** | OLE DB natif | `Provider=OraOLEDB.Oracle` | Oracle Client (ODAC / Instant Client) |
| **PostgreSQL** | ODBC | `Driver={PostgreSQL Unicode}` | Pilote psqlODBC |
| **MySQL / MariaDB** | ODBC | `Driver={MySQL ODBC … Driver}` | MySQL Connector/ODBC |

---

## SQL Server

SQL Server est le compagnon naturel d'Access pour les architectures client-serveur. Plusieurs fournisseurs OLE DB se sont succédé : l'ancien `SQLOLEDB`, puis le *SQL Server Native Client* (`SQLNCLI11`), aujourd'hui dépréciés. Le fournisseur **actuel et recommandé** est **`MSOLEDBSQL`** (*Microsoft OLE DB Driver for SQL Server*).

En **authentification Windows** (intégrée), qui évite de stocker un mot de passe :

```vba
cn.Open "Provider=MSOLEDBSQL;Server=NOM_SERVEUR;" & _
        "Database=MaBase;Trusted_Connection=yes;"
```

En **authentification SQL Server** (compte et mot de passe propres au serveur) :

```vba
cn.Open "Provider=MSOLEDBSQL;Server=NOM_SERVEUR;" & _
        "Database=MaBase;UID=utilisateur;PWD=motdepasse;"
```

Le paramètre `Server` accepte les **instances nommées** (`NOM_SERVEUR\INSTANCE`) et un **port explicite** (`NOM_SERVEUR,1433`). `Trusted_Connection=yes` équivaut à `Integrated Security=SSPI`. Privilégiez l'authentification Windows chaque fois que possible, pour des raisons de sécurité (chapitre 20).

---

## Oracle

Pour Oracle, le fournisseur Microsoft historique `MSDAORA` est obsolète et limité aux anciennes versions. On lui préfère le fournisseur **`OraOLEDB.Oracle`** fourni par Oracle (composant ODAC) :

```vba
cn.Open "Provider=OraOLEDB.Oracle;Data Source=NOM_TNS;" & _
        "User Id=utilisateur;Password=motdepasse;"
```

`Data Source` désigne ici le nom de service défini dans la configuration TNS d'Oracle. Ce fournisseur **exige l'installation du client Oracle** (ODAC ou *Instant Client*) sur chaque poste — un prérequis à ne pas négliger.

---

## PostgreSQL

PostgreSQL ne dispose pas de fournisseur OLE DB officiel de Microsoft. On s'y connecte donc via **ODBC**, à l'aide du pilote **psqlODBC**. En chaîne sans DSN (*DSN-less*) :

```vba
cn.Open "Provider=MSDASQL;Driver={PostgreSQL Unicode};" & _
        "Server=hote;Port=5432;Database=mabase;" & _
        "Uid=utilisateur;Pwd=motdepasse;"
```

Le `Provider=MSDASQL` est explicite ici, mais pourrait être omis (fournisseur par défaut). Le port standard de PostgreSQL est `5432`. Le pilote psqlODBC doit être installé dans la bonne architecture.

---

## MySQL / MariaDB

Comme PostgreSQL, MySQL passe par **ODBC**, via le pilote **MySQL Connector/ODBC** (compatible MariaDB) :

```vba
cn.Open "Provider=MSDASQL;Driver={MySQL ODBC 8.0 Unicode Driver};" & _
        "Server=hote;Port=3306;Database=mabase;" & _
        "Uid=utilisateur;Pwd=motdepasse;"
```

Le port standard de MySQL est `3306`. **Attention** : le nom exact du pilote (`MySQL ODBC 8.0 Unicode Driver`) dépend de la **version installée** ; il faut le vérifier dans l'administrateur de sources de données ODBC de Windows, le numéro de version pouvant différer.

---

## DSN ou DSN-less ?

Les connexions ODBC se déclinent en deux styles, dont le choix conditionne le déploiement.

Une connexion **avec DSN** s'appuie sur un *Data Source Name* préconfiguré dans l'administrateur ODBC de Windows (DSN utilisateur ou système). La chaîne se réduit alors à `DSN=MonDSN;Uid=…;Pwd=…;`. C'est concis et la configuration est centralisée — mais le DSN doit être **créé sur chaque poste**, ce qui alourdit le déploiement.

Une connexion **sans DSN** (*DSN-less*) embarque tous les détails du pilote dans la chaîne, comme dans les exemples ci-dessus. Elle est **autonome et portable** dans le code, sans configuration ODBC préalable — mais le pilote doit néanmoins être installé, et le nom du pilote doit correspondre exactement. Pour la distribution d'une application, le style DSN-less est souvent préféré, car il évite d'avoir à configurer un DSN sur chaque machine.

---

## Le prérequis incontournable : le pilote installé, dans la bonne architecture

C'est le point opérationnel le plus important de cette section, et le piège récurrent du déploiement. **Quelle que soit la voie choisie, le fournisseur ou le pilote correspondant doit être installé sur chaque poste client**, et dans une **architecture compatible** (32 ou 64 bits) avec l'application Access — exactement le même écueil que pour le fournisseur ACE (section 10.2).

Concrètement, cela impose de déployer le *Microsoft OLE DB Driver for SQL Server* pour SQL Server, le *client Oracle* pour Oracle, le pilote *psqlODBC* pour PostgreSQL, ou *Connector/ODBC* pour MySQL — en veillant à la correspondance 32/64 bits. Cette contrainte, lourde de conséquences au déploiement, est approfondie au chapitre 21.7. Le message *« fournisseur non inscrit »* ou *« pilote ODBC introuvable »* en est la manifestation la plus fréquente.

---

## Procédures stockées et paramètres nommés (SQL Server)

C'est ici que se concrétise une nuance laissée en suspens à la section 10.6. Là où le moteur ACE impose des paramètres **positionnels** (`?`), SQL Server prend en charge les **paramètres nommés** (`@nom`) et les **paramètres de sortie** — ce qui donne enfin tout leur sens aux directions `adParamOutput` et `adParamReturnValue`.

Appeler une procédure stockée SQL Server, avec un paramètre d'entrée et un paramètre de sortie, s'écrit ainsi :

```vba
Dim cmd As ADODB.Command
Set cmd = New ADODB.Command
cmd.ActiveConnection = cn
cmd.CommandText = "usp_AjouterClient"
cmd.CommandType = adCmdStoredProc

cmd.Parameters.Append cmd.CreateParameter("@Nom", adVarWChar, adParamInput, 50, "Dupont")
cmd.Parameters.Append cmd.CreateParameter("@NouvelID", adInteger, adParamOutput)

cmd.Execute
Debug.Print "Identifiant généré : " & cmd.Parameters("@NouvelID").Value
```

Pour récupérer l'identité d'une ligne insérée côté SQL Server, on privilégie par ailleurs **`SCOPE_IDENTITY()`** plutôt que `@@IDENTITY` — évoqué à la section 10.5 — car `SCOPE_IDENTITY()` ignore les éventuels déclencheurs et reste fiable :

```vba
Dim rs As ADODB.Recordset
Set rs = cn.Execute("SELECT SCOPE_IDENTITY();")
Debug.Print "Nouvel ID : " & rs.Fields(0).Value
```

---

## ADO ou tables liées ODBC ?

Une question pratique se pose dès qu'Access doit travailler avec un SGBD externe : faut-il utiliser ADO… ou des **tables liées** ? Les deux approches sont valides et **coexistent** souvent dans une même application.

Les **tables liées ODBC** font apparaître les tables distantes comme des tables locales : formulaires, requêtes et code DAO les exploitent de manière transparente, sans connaître leur origine. C'est l'option idéale pour **lier l'interface** à des données serveur avec un minimum de code.

ADO, lui, est l'approche **programmatique** : connexion directe, contrôle total, accès à la demande, sans dépendre d'un lien permanent. On le choisit pour des traitements ciblés, des appels de procédures stockées, ou lorsqu'on ne souhaite pas exposer la table dans l'interface. L'architecture hybride mêlant tables liées et accès direct est détaillée au chapitre 23.4, et les requêtes pass-through — une troisième voie consistant à envoyer du SQL natif au serveur — au chapitre 11.9.

---

## Attention au dialecte SQL

Si le modèle objet ADO est identique partout, **le SQL ne l'est pas**. SQL Server parle le **T-SQL**, Oracle le **PL/SQL**, et chacun a ses fonctions, sa syntaxe de dates, ses spécificités — autant de différences avec le **Jet SQL** d'Access (chapitre 11.8). Une requête qui fonctionne sur ACE peut échouer telle quelle sur SQL Server, et inversement. Adapter le code SQL après un changement de moteur est un sujet à part entière, traité au chapitre 11 et, dans le contexte d'une migration, au chapitre 23.3.

---

## Performance, sécurité et référence

Quelques considérations transversales pour clore. Côté **performance**, les connexions vers des serveurs distants bénéficient du *pooling* (réutilisation des connexions physiques) et d'un `ConnectionTimeout` adapté à la latence réseau (sections 10.3 et chapitre 18.8). Côté **sécurité**, on privilégie l'authentification intégrée, on évite de coder les identifiants en dur, et l'on protège les chaînes de connexion (chapitre 20).

Pour les chaînes de connexion exactes de chaque SGBD, l'**annexe D** constitue le récapitulatif de référence, complétée par le site communautaire *connectionstrings.com*. Les codes d'erreur spécifiques sont quant à eux catalogués en **annexe C**.

---

## En résumé

ADO atteint les bases externes en ne changeant, pour l'essentiel, que la **chaîne de connexion** : le reste du modèle objet demeure identique. Deux voies s'offrent à vous — un **fournisseur OLE DB natif** (recommandé pour SQL Server avec `MSOLEDBSQL`, pour Oracle avec `OraOLEDB.Oracle`) ou le **pont ODBC** via `MSDASQL` (incontournable pour PostgreSQL et MySQL). Le choix entre DSN et DSN-less, et surtout l'**installation du pilote dans la bonne architecture** sur chaque poste, sont les vrais enjeux de déploiement. Face à SQL Server, les **paramètres nommés** et de **sortie** des procédures stockées deviennent disponibles. Enfin, ADO et les **tables liées** ne s'excluent pas : les premières pour le programmatique, les secondes pour lier l'interface — sans oublier que le **dialecte SQL** change avec le moteur.

---

## Conclusion du chapitre 10

Ce chapitre a fait le tour complet d'ADO. Parti de la question du choix entre DAO et ADO (10.1), il a posé les fondations — chaînes de connexion (10.2) et objet `Connection` (10.3) —, exploré le cœur du modèle avec le `Recordset` et ses curseurs (10.4), les opérations CRUD (10.5) et l'objet `Command` paramétré (10.6), puis abordé les capacités avancées : recordsets déconnectés et persistance (10.7), gestion fine des erreurs (10.8) et, pour finir, connexion aux bases externes (10.9).

La leçon d'ensemble mérite d'être rappelée : **ADO ne remplace pas DAO, il le complète**. Pour le moteur ACE local, DAO reste le choix naturel et performant. Pour l'ouverture vers l'extérieur, les recordsets déconnectés et les architectures distribuées, ADO devient irremplaçable. Un développeur Access accompli maîtrise les deux et sait, en connaissance de cause, lequel mobiliser selon le contexte.

Le chapitre suivant approfondit un sujet transversal présent dans tout ce que nous venons de voir, qu'il s'agisse de DAO ou d'ADO : le **SQL dans Access VBA**, des requêtes action aux subtilités du dialecte Jet, en passant par le SQL paramétré et les fonctions de domaine.


⏭️ [11. SQL dans Access VBA](/11-sql-access-vba/README.md)
