🔝 Retour au [Sommaire](/SOMMAIRE.md)

# Chapitre 9 — DAO : Data Access Objects

> **Niveau :** Intermédiaire  
> **Prérequis conseillés :** rappels fondamentaux VBA (chapitre 3) et modèle objet Access (chapitre 4). Une connaissance de base du SQL est utile mais non indispensable.

## Introduction

DAO — *Data Access Objects* — est la bibliothèque d'accès aux données **native** d'Access. Elle constitue l'interface de programmation directe vers le moteur de base de données qui anime Access : historiquement le moteur Jet, aujourd'hui le moteur ACE (*Access Connectivity Engine*). Dès que l'on ouvre une table, exécute une requête, parcourt des résultats ou écrit un enregistrement en VBA, c'est le plus souvent DAO qui travaille en coulisses, que le développeur en ait conscience ou non.

Il est important de bien situer DAO par rapport aux outils déjà rencontrés. Là où `DoCmd` automatise l'interface utilisateur (ouvrir un formulaire, lancer une requête, exporter un état) et où le SQL exprime des opérations *ensemblistes* portant sur des groupes entiers d'enregistrements, DAO offre un contrôle fin, **enregistrement par enregistrement**, sur les données. C'est l'outil de prédilection dès qu'un traitement doit lire, comparer, transformer ou écrire des lignes de façon procédurale, avec une logique métier qui dépasse ce qu'une simple instruction SQL peut exprimer.

## Le moteur derrière Access : de Jet à ACE

DAO est indissociable du moteur qui stocke et gère physiquement les données d'Access. Pendant longtemps, ce moteur s'est appelé Jet (*Joint Engine Technology*) et travaillait avec les fichiers `.mdb`. À partir d'Access 2007 et de l'apparition du format `.accdb`, Microsoft a fait évoluer ce moteur vers ACE, tout en conservant une compatibilité ascendante. ACE est essentiellement Jet enrichi de nouvelles capacités : champs multivalués, champs pièce jointe, types de données supplémentaires.

Ce détail historique n'est pas anecdotique. DAO existe en réalité sous deux formes : l'ancienne bibliothèque *Microsoft DAO 3.6 Object Library*, liée à Jet et aux fichiers `.mdb`, et la bibliothèque moderne *Microsoft Office Access database engine Object Library*, liée à ACE et indispensable pour exploiter les fonctionnalités introduites avec `.accdb`. C'est cette seconde version que vous utiliserez dans toute application Access récente.

On notera au passage que DAO a connu une période d'éclipse : à la fin des années 1990, Microsoft a mis en avant ADO comme technologie d'accès aux données plus universelle, et DAO a semblé un temps abandonné. Le format `.accdb` et le moteur ACE l'ont réhabilité : pour manipuler des données Access natives, DAO est aujourd'hui clairement la technologie recommandée.

## Pourquoi apprendre DAO

Trois raisons principales justifient l'investissement.

D'abord, la **performance**. DAO étant l'interface native du moteur ACE, il est généralement plus rapide qu'ADO pour accéder à des données Access locales, sans la couche d'abstraction OLE DB qu'ADO impose.

Ensuite, l'**accès complet aux spécificités d'Access**. Certaines fonctionnalités propres à Access — propriétés personnalisées des objets, champs multivalués, champs pièce jointe — ne sont pleinement accessibles que par DAO. ADO les ignore ou les expose de façon incomplète.

Enfin, l'**omniprésence dans le code existant**. La grande majorité des applications Access en production reposent sur DAO. Savoir lire, maintenir et faire évoluer ce code est une compétence incontournable pour tout développeur amené à reprendre des projets Access.

## La hiérarchie objet DAO en un coup d'œil

DAO s'organise en une arborescence d'objets imbriqués. Au sommet se trouve `DBEngine`, qui représente le moteur lui-même ; il contient des espaces de travail (`Workspaces`), eux-mêmes contenant les bases ouvertes (`Databases`). C'est au niveau de l'objet `Database` que se rattachent les collections que vous manipulerez le plus :

```
DBEngine
└─ Workspaces
   └─ Workspace
      └─ Databases
         └─ Database
            ├─ Recordsets      ' ← cœur de ce chapitre
            ├─ TableDefs        ' ← structure des tables   → chapitre 12
            ├─ QueryDefs        ' ← requêtes sauvegardées  → chapitre 12
            ├─ Relations        ' ← relations entre tables → chapitre 12
            └─ Containers / Documents
```

Dans la pratique, on n'a presque jamais besoin de remonter jusqu'à `DBEngine` ou `Workspace`. La fonction `CurrentDb()` fournit directement un objet `Database` pointant sur la base courante, ce qui constitue le point d'entrée habituel de tout traitement DAO.

## Un premier contact avec DAO

Pour donner une idée concrète du style de programmation que ce chapitre va détailler, voici l'ossature minimale d'un parcours de données en DAO : on obtient la base courante, on ouvre un jeu d'enregistrements à partir d'une requête, on le parcourt jusqu'à la fin, puis on libère proprement les objets.

```vba
Dim db As DAO.Database
Dim rs As DAO.Recordset

Set db = CurrentDb()
Set rs = db.OpenRecordset("SELECT NomClient FROM tblClients", dbOpenSnapshot)

Do While Not rs.EOF
    Debug.Print rs!NomClient
    rs.MoveNext
Loop

rs.Close
Set rs = Nothing
Set db = Nothing
```

Chacune des notions effleurées ici — le type de `Recordset` choisi (`dbOpenSnapshot`), la condition d'arrêt `EOF`, le déplacement `MoveNext`, et la clôture rigoureuse des objets en fin de traitement — fait l'objet d'une section dédiée dans ce chapitre.

## DAO ou ADO ?

La question du choix entre DAO et ADO revient inévitablement. En résumé : DAO est optimisé et complet pour les données natives Access (ACE/Jet), tandis qu'ADO est plus universel et historiquement privilégié pour dialoguer avec des sources externes comme SQL Server, Oracle ou PostgreSQL via OLE DB. Pour du travail purement Access, DAO est le choix par défaut. La comparaison détaillée, avec critères de décision, est traitée au début du [chapitre 10](/10-ado-access/01-dao-vs-ado.md).

## Avant de commencer : activer et qualifier la référence

Dans une application Access moderne, la référence à la bibliothèque DAO (*Microsoft Office Access database engine Object Library*) est normalement activée par défaut. Pour le vérifier, ouvrez l'éditeur VBA (Alt+F11), puis **Outils → Références**, et assurez-vous que la bibliothèque est cochée. Le numéro de version qui l'accompagne dépend de votre version d'Access.

Une bonne pratique consiste à recourir à la **liaison précoce** (déclarer ses variables avec leur type explicite : `Dim rs As DAO.Recordset`) et à toujours **préfixer** les types par `DAO.`. C'est d'autant plus important qu'ADO expose lui aussi un objet `Recordset` : en l'absence de préfixe, une ambiguïté de nommage peut conduire VBA à instancier le mauvais objet. Cette distinction entre liaisons précoce et tardive est développée au [chapitre 2](/02-interface-environnement/06-early-vs-late-binding.md).

## Périmètre du chapitre

Ce chapitre se concentre sur la **manipulation des données** : ouvrir des bases, parcourir des `Recordset`, lire et écrire des champs, rechercher, trier et filtrer, puis libérer les ressources. Les **objets structurels** de DAO — création et modification de tables (`TableDef`), de requêtes sauvegardées (`QueryDef`) et de relations (`Relation`) par code — ne sont volontairement pas traités ici, mais au [chapitre 12](/12-querydefs-tabledefs/README.md). La section 9.1 introduit toutefois ces objets sur le plan conceptuel, afin de présenter l'architecture DAO dans son ensemble.

## Plan du chapitre

- **9.1** — [Architecture DAO — DBEngine, Workspace, Database, TableDef, QueryDef](/09-dao-data-access-objects/01-architecture-dao.md) : vue d'ensemble de la hiérarchie et du rôle de chaque objet.
- **9.2** — [Ouverture et gestion d'une Database avec `CurrentDb()`](/09-dao-data-access-objects/02-ouverture-database.md) : le point d'entrée de tout traitement DAO.
- **9.3** — [Recordset DAO — types (Table, Dynaset, Snapshot, Forward-only)](/09-dao-data-access-objects/03-recordset-types.md) : choisir le bon type selon l'usage et la performance recherchée.
- **9.4** — [Navigation dans un Recordset (`MoveFirst`, `MoveNext`, `EOF`, `BOF`)](/09-dao-data-access-objects/04-navigation-recordset.md) : se déplacer correctement dans un jeu d'enregistrements.
- **9.5** — [Lecture et modification des champs (`Fields`)](/09-dao-data-access-objects/05-lecture-modification-champs.md) : accéder aux valeurs et aux propriétés des colonnes.
- **9.6** — [Ajout d'enregistrements (`AddNew`/`Update`)](/09-dao-data-access-objects/06-ajout-enregistrements.md) : insérer de nouvelles lignes par code.
- **9.7** — [Modification d'enregistrements (`Edit`/`Update`)](/09-dao-data-access-objects/07-modification-enregistrements.md) : mettre à jour des lignes existantes.
- **9.8** — [Suppression d'enregistrements (`Delete`)](/09-dao-data-access-objects/08-suppression-enregistrements.md) : retirer des lignes en toute sécurité.
- **9.9** — [Recherche dans un Recordset (`FindFirst`, `FindNext`, `Seek`)](/09-dao-data-access-objects/09-recherche-recordset.md) : localiser un enregistrement précis.
- **9.10** — [Tri et filtrage côté Recordset](/09-dao-data-access-objects/10-tri-filtrage-recordset.md) : ordonner et restreindre les données en mémoire.
- **9.11** — [Clôture et libération des objets DAO](/09-dao-data-access-objects/11-cloture-liberation-dao.md) : éviter fuites mémoire et verrous résiduels.
- **9.12** — [RecordsetClone — manipuler les enregistrements d'un formulaire](/09-dao-data-access-objects/12-recordsetclone.md) : agir sur les données affichées sans perturber l'interface.
- **9.13** — [Champs multivalués (MVF) et champs Pièce jointe — `Recordset2` et `Field2`](/09-dao-data-access-objects/13-champs-multivalues-pieces-jointes.md) : exploiter les types de données propres à ACE.

## En résumé

DAO est le moyen le plus direct, le plus rapide et le plus complet d'accéder par programmation aux données natives d'Access. Maîtriser l'objet `Recordset` et son cycle de vie — ouverture, navigation, lecture/écriture, clôture — constitue le socle de la quasi-totalité des traitements de données en VBA Access. C'est précisément l'objet de ce chapitre.

⏭️ [9.1. Architecture DAO — DBEngine, Workspace, Database, TableDef, QueryDef](/09-dao-data-access-objects/01-architecture-dao.md)
