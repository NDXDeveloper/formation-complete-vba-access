🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 12. QueryDefs et TableDefs

Le chapitre 9 nous a appris à manipuler les **données** d'Access avec les recordsets DAO, et le chapitre 11 à maîtriser le **langage SQL**. Ce chapitre aborde une troisième facette, complémentaire : les **objets structurels** de DAO. Au lieu de manipuler le *contenu* de la base, nous allons manipuler ses *objets* eux-mêmes — les requêtes sauvegardées, la structure des tables, les relations, les tables liées.

C'est précisément la transition annoncée à la fin du chapitre 11 : passer du SQL *écrit* aux objets DAO qui *représentent et manipulent* les requêtes et les tables de la base. Là où l'on tapait une instruction `CREATE TABLE` ou une requête `SELECT`, on va désormais agir sur les objets `QueryDef`, `TableDef`, `Field`, `Index` et `Relation` qui composent la définition même de l'application.

---

## Les deux visages de DAO : données et structure

DAO possède deux dimensions distinctes, qu'il est éclairant de distinguer. La première, étudiée au chapitre 9, est l'**accès aux données** : le `Recordset` lit, parcourt et modifie le contenu des tables. La seconde, objet de ce chapitre, est la **définition de la structure** : les objets `TableDef`, `QueryDef` et `Relation` décrivent et façonnent l'ossature de la base.

Ce chapitre **complète donc le portrait de DAO** entamé au chapitre 9. Après avoir vu comment lire et écrire des enregistrements, nous voyons comment inspecter, créer et modifier les objets qui les contiennent et les interrogent.

---

## Le pendant objet du SQL et du DDL

Une autre façon de situer ce chapitre est de le voir comme le **versant objet** de ce que le chapitre 11 abordait par le langage.

Un **`QueryDef`** est l'objet DAO qui représente une **requête sauvegardée** : son texte SQL réside dans sa propriété `SQL`. Là où le chapitre 11 exécutait du SQL, souvent en ligne, ce chapitre manipule les requêtes sauvegardées elles-mêmes — les lire, les créer, les modifier, les supprimer, et surtout les **paramétrer** (un point déjà sollicité aux sections 11.5 et 11.6).

De même, les objets **`TableDef`**, **`Field`** et **`Index`** sont la représentation objet de la **structure des tables** — le pendant du DDL vu à la section 11.4. Rappelons l'arbitrage qui y était posé : le DDL-SQL est concis et portable, tandis que les objets DAO offrent un **contrôle complet des propriétés propres à Access** (légendes, formats, descriptions, masques de saisie) que le DDL ne sait pas définir. C'est ici que cette approche objet est pleinement développée.

À titre d'aperçu, voici l'esprit de ce chapitre — *inspecter* la structure par code plutôt que manipuler des données :

```vba
Dim db As DAO.Database
Dim td As DAO.TableDef

Set db = CurrentDb
For Each td In db.TableDefs
    ' On ignore les tables système et masquées
    If Left$(td.Name, 4) <> "MSys" And Left$(td.Name, 1) <> "~" Then
        Debug.Print td.Name & " — " & td.Fields.Count & " champ(s)"
    End If
Next td
```

Ce simple parcours énumère les tables de la base et leur nombre de champs : on lit la *structure*, non le contenu. C'est tout l'objet du chapitre.

---

## Pourquoi manipuler structure et requêtes par code ?

Plusieurs besoins concrets justifient cette manipulation programmatique.

L'**automatisation du schéma** : bâtir ou faire évoluer la structure d'une base par script, lors d'une installation ou d'une migration. L'**inspection de la structure** : parcourir tables et champs pour générer de la documentation, alimenter une interface dynamique, ou bâtir des outils génériques pilotés par les métadonnées. La **gestion dynamique des requêtes** : modifier le SQL d'une requête sauvegardée selon les circonstances, ou administrer des requêtes paramétrées. La **gestion des relations** et de l'intégrité référentielle par code. La **gestion des tables liées**, enfin — notamment la **reliaison** vers un back-end déplacé, opération vitale pour les applications scindées (chapitres 15.9 et 21.4). Sans oublier la possibilité de définir des **propriétés propres à Access** hors de portée du DDL.

---

## Le modèle objet structurel de DAO

Ce chapitre met à profit la hiérarchie d'objets introduite au chapitre 9.1. Pour mémoire, la structure s'articule ainsi :

```
DBEngine → Workspace → Database
                          ├── QueryDefs   (requêtes sauvegardées → SQL, Parameters, Fields)
                          ├── TableDefs    (tables → Fields, Indexes, Connect si liée)
                          └── Relations    (relations entre tables, intégrité référentielle)
```

Chaque collection contient des objets dotés de leurs propres sous-collections : un `QueryDef` expose ses `Parameters` et ses `Fields`, un `TableDef` ses `Fields` et ses `Indexes`. C'est en naviguant et en agissant sur ces collections que l'on inspecte et façonne la base.

---

## Ce que vous allez apprendre dans ce chapitre

Le chapitre suit la progression naturelle : d'abord les requêtes, puis les tables, puis les relations et les objets liés.

La section **12.1** présente la collection **`QueryDefs`** et l'accès aux requêtes sauvegardées (lecture de leur SQL, énumération). La section **12.2** détaille comment **créer, modifier et supprimer** des requêtes par code. La section **12.3** est consacrée aux **requêtes paramétrées** — leur définition et leur exécution par code, prolongeant le paramétrage abordé au chapitre 11.

La section **12.4** introduit la collection **`TableDefs`** pour **inspecter** la structure des tables. La section **12.5** explique comment **créer et modifier des tables** par code — champs, index, clés —, l'approche objet annoncée au chapitre 11.4. La section **12.6** traite de la collection **`Relations`** et de la gestion des relations entre tables.

La section **12.7** aborde la gestion programmatique des **tables liées** (*linked tables*), un sujet clé pour les applications réparties. Enfin, la section **12.8** présente les **propriétés personnalisées** des objets base de données, pour étendre les métadonnées par code.

---

## Pourquoi DAO (et non ADO) pour la structure ?

Le choix de DAO pour ce chapitre n'est pas arbitraire. Comme établi au chapitre 10.1, DAO est le **spécialiste du moteur ACE** : c'est lui qui accède le plus complètement à la structure native d'Access. ADO dispose certes d'un pendant structurel — la bibliothèque **ADOX** —, mais celui-ci est moins complet et moins intégré aux spécificités d'Access (champs avec leurs propriétés, relations, propriétés personnalisées). Pour inspecter et façonner la structure d'une base Access, **DAO est l'outil de référence**, et c'est lui que ce chapitre emploie de bout en bout.

---

## Prérequis

Ce chapitre suppose maîtrisés les **fondamentaux de DAO** (chapitre 9, en particulier l'architecture du chapitre 9.1) et une bonne familiarité avec les **collections** (chapitre 3.7), puisqu'on énumère constamment `QueryDefs`, `TableDefs` et `Relations`. La connaissance du **SQL** (chapitre 11) est nécessaire pour les requêtes sauvegardées, dont le contenu est précisément du SQL, et pour comprendre le lien avec le paramétrage. Les **fondamentaux VBA** (chapitre 3) restent évidemment requis.

---

## Note de cadrage et ressources

Ce chapitre est le **complément structurel** du chapitre 11 : là où celui-ci écrivait du SQL, celui-ci manipule les objets qui le portent et la structure qu'il décrit. Les deux se répondent — on pourra avoir l'un et l'autre à l'esprit, notamment pour les requêtes sauvegardées (chapitre 11.2 sur leur précompilation, chapitre 11.11 sur leur empilement) et pour la création de tables (chapitre 11.4 sur le DDL). Pour la syntaxe du SQL contenu dans les `QueryDef`, l'annexe E (référence Jet SQL) demeure utile.

---

La porte d'entrée vers la structure d'une base est sa collection de requêtes sauvegardées. Avant de créer ou de modifier quoi que ce soit, apprenons à y accéder et à en explorer le contenu — c'est par là que commence le chapitre.


⏭️ [12.1. Collection QueryDefs — accès et manipulation des requêtes sauvegardées](/12-querydefs-tabledefs/01-collection-querydefs.md)
