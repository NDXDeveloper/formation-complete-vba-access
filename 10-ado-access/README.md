🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 10. ADO – ActiveX Data Objects pour Access

Après les chapitres consacrés à **DAO** (chapitre 9), nous abordons la seconde grande technologie d'accès aux données disponible en VBA Access : **ADO** (*ActiveX Data Objects*). Là où DAO a été conçu sur mesure pour le moteur de base de données natif d'Access (Jet, puis ACE), ADO a été pensé par Microsoft comme une technologie d'accès aux données *universelle*, capable de dialoguer avec n'importe quelle source compatible — d'un fichier `.accdb` local à un serveur SQL Server, Oracle ou PostgreSQL distant, en passant par des sources plus exotiques.

Ce chapitre n'a pas pour but de remplacer DAO dans vos applications Access, mais de vous donner une seconde corde à votre arc, particulièrement précieuse dès que votre application sort du cadre strictement Access pour s'ouvrir à d'autres systèmes de gestion de bases de données.

---

## Pourquoi un chapitre dédié à ADO ?

Dans une application Access « classique », reposant uniquement sur des tables locales ou liées au moteur ACE, DAO reste le choix naturel et le plus performant. On pourrait donc légitimement se demander pourquoi consacrer un chapitre entier à ADO.

La réponse tient en quelques scénarios où ADO devient pertinent, voire incontournable :

La connexion à des bases de données **externes et hétérogènes** est le cas d'usage historique d'ADO. Dès qu'une application Access doit interroger un serveur SQL Server, une base Oracle, MySQL ou PostgreSQL sans passer par des tables liées, ADO offre un modèle de connexion homogène, quel que soit le fournisseur de données sous-jacent.

Les **architectures client-serveur** et les migrations progressives d'Access vers un SGBD plus robuste (voir le chapitre 23 sur la migration) s'appuient fréquemment sur ADO, qui parle nativement le langage des fournisseurs OLE DB et ODBC.

Certaines **fonctionnalités spécifiques** n'existent qu'avec ADO, comme les *recordsets déconnectés* — un jeu d'enregistrements que l'on peut manipuler en mémoire après avoir fermé la connexion — ou la persistance d'un recordset dans un fichier XML.

Enfin, un développeur Access est régulièrement amené à **lire, maintenir ou reprendre du code existant** écrit en ADO. Comprendre ce modèle objet n'est donc pas une option : c'est une compétence de maintenance indispensable.

---

## ADO en bref

ADO est une **bibliothèque COM** (Component Object Model) qui s'appuie elle-même sur la couche **OLE DB**. Cette architecture en couches est la clé de sa polyvalence :

```
Votre code VBA
      │
     ADO          ← interface de programmation simple et homogène
      │
   OLE DB         ← couche d'abstraction bas niveau
      │
 Fournisseurs     ← ACE, SQL Server, Oracle, ODBC...
      │
 Sources de données
```

Concrètement, votre code VBA ne s'adresse jamais directement à la base de données : il dialogue avec les objets ADO, qui transmettent les ordres à un **fournisseur OLE DB** (*provider*). C'est ce fournisseur — choisi via la chaîne de connexion — qui sait réellement parler à la source de données ciblée. Changer de SGBD revient souvent à ne changer que la chaîne de connexion, le reste du code demeurant largement identique.

---

## Le modèle objet ADO en un coup d'œil

Contrairement à DAO et sa hiérarchie profonde (DBEngine → Workspace → Database → Recordset…), le modèle ADO est volontairement **plat et resserré** autour de trois objets centraux :

L'objet **Connection** représente la connexion ouverte vers une source de données. C'est le point d'entrée : il porte la chaîne de connexion, gère l'ouverture et la fermeture, et héberge la collection des erreurs survenues.

L'objet **Command** représente une commande à exécuter — typiquement une requête SQL ou une procédure stockée. Il est particulièrement utile pour les requêtes paramétrées, via sa collection `Parameters`.

L'objet **Recordset** représente un jeu d'enregistrements issu d'une requête. C'est l'équivalent ADO du recordset DAO, mais avec ses propres notions de *curseur* et de *verrouillage* qui lui sont spécifiques.

Autour de ces trois piliers gravitent quelques collections et objets secondaires : la collection `Errors` (rattachée à `Connection`), la collection `Parameters` (rattachée à `Command`), ainsi que `Fields` et `Properties` (rattachées au `Recordset`). L'ensemble forme un modèle compact, rapide à prendre en main une fois les trois objets principaux assimilés.

À titre d'avant-goût, voici le schéma de code que l'on retrouvera tout au long du chapitre — l'ouverture d'une connexion et la lecture d'un jeu d'enregistrements :

```vba
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Set cn = New ADODB.Connection
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CurrentProject.FullName

Set rs = New ADODB.Recordset
rs.Open "SELECT * FROM Clients", cn, adOpenForwardOnly, adLockReadOnly

' ... traitement des enregistrements ...

rs.Close
cn.Close
Set rs = Nothing
Set cn = Nothing
```

Ne vous attardez pas encore sur le détail de chaque mot-clé : tous seront décortiqués dans les sections suivantes. Ce squelette illustre simplement la logique générale d'ADO — *connecter, exécuter, parcourir, libérer*.

---

## Bibliothèque et références

Pour utiliser ADO en *liaison précoce* (early binding), il faut activer la référence à la bibliothèque **Microsoft ActiveX Data Objects** dans l'éditeur VBA (menu *Outils → Références*). Sur les versions récentes de Windows, il s'agit le plus souvent de la version **6.1** (`Microsoft ActiveX Data Objects 6.1 Library`) ; sur des environnements plus anciens, on rencontre encore la version 2.8.

À noter que le préfixe de classe **`ADODB`** (par exemple `ADODB.Connection`) permet de lever toute ambiguïté avec DAO, d'autant qu'Access référence parfois plusieurs bibliothèques d'accès aux données simultanément. Cette précaution évite des conflits de noms classiques sur des objets comme `Recordset`, qui existent à la fois dans DAO et dans ADO.

Le choix entre liaison précoce et liaison tardive (*late binding*), abordé au chapitre 2.6, s'applique pleinement à ADO : la liaison tardive (`CreateObject("ADODB.Connection")`) dispense d'activer la référence et facilite le déploiement sur des postes aux versions hétérogènes, au prix de la perte de l'IntelliSense et d'une légère baisse de performance.

---

## Ce que vous allez apprendre dans ce chapitre

Le chapitre est organisé pour vous mener progressivement de la décision technique initiale jusqu'aux usages les plus avancés d'ADO :

La section **10.1 – DAO vs ADO** pose la question fondamentale : quelle technologie choisir, et selon quels critères ? Elle compare les deux approches pour vous aider à arbitrer en connaissance de cause.

La section **10.2 – Chaînes de connexion** détaille la syntaxe des chaînes de connexion pour Access, notamment le fournisseur `Provider=Microsoft.ACE.OLEDB`, véritable point névralgique de toute connexion ADO.

La section **10.3 – L'objet Connection** explique comment ouvrir, configurer et fermer proprement une connexion, et pourquoi cette discipline conditionne la fiabilité de l'application.

La section **10.4 – L'objet Recordset ADO** introduit les notions propres à ADO que sont les *curseurs* (forward-only, statique, dynamique, jeu de clés) et les *modes de verrouillage*, déterminants pour les performances et le comportement multi-utilisateurs.

La section **10.5 – CRUD avec ADO** couvre les opérations de lecture, modification, ajout et suppression d'enregistrements, c'est-à-dire le cœur de la manipulation des données.

La section **10.6 – L'objet Command et les paramètres** présente la manière propre et sécurisée d'exécuter des requêtes paramétrées, en évitant les écueils de la concaténation de chaînes SQL.

La section **10.7 – Recordsets déconnectés et persistance XML** explore une capacité spécifique à ADO : travailler sur des données en mémoire, indépendamment de toute connexion active, et les sérialiser sur disque.

La section **10.8 – Gestion des erreurs ADO** détaille l'interception des erreurs via la collection `Errors`, dont le fonctionnement diffère sensiblement de la simple gestion par l'objet `Err` de VBA.

Enfin, la section **10.9 – Connexion à des bases externes** met ADO à l'épreuve de son terrain de prédilection : se connecter à SQL Server, Oracle, PostgreSQL ou MySQL depuis une application Access.

---

## Prérequis

Pour aborder ce chapitre dans les meilleures conditions, il est recommandé d'avoir assimilé les fondamentaux de la programmation VBA (chapitre 3), le modèle objet Access (chapitre 4) et, surtout, le chapitre 9 consacré à DAO. La comparaison entre les deux technologies sera d'autant plus parlante que vous maîtrisez déjà la manipulation des recordsets DAO. Une connaissance de base du langage SQL (abordé au chapitre 11) facilitera également la lecture des exemples.

---

## Note de cadrage : ADO n'est pas un remplaçant universel de DAO

Avant de plonger dans le détail, gardons une perspective claire. Pour des opérations purement locales sur le moteur ACE, **DAO demeure généralement plus rapide et mieux intégré** à Access — il accède directement à des fonctionnalités natives qu'ADO ne connaît pas (champs multivalués, propriétés spécifiques au moteur, `RecordsetClone` des formulaires, etc.).

ADO brille en revanche dès que l'on franchit la frontière du « tout Access » : sources hétérogènes, architectures distribuées, recordsets déconnectés. L'objectif de ce chapitre n'est donc pas de vous convertir à ADO, mais de vous rendre **capable de choisir la bonne technologie selon le contexte** — et c'est précisément par cette question de choix que s'ouvre la première section.


⏭️ [10.1. DAO vs ADO — quelle technologie choisir ?](/10-ado-access/01-dao-vs-ado.md)
