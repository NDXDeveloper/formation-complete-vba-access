🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.11 Clôture et libération des objets DAO

## Introduction

Plusieurs sections de ce chapitre ont conclu leurs exemples par `rs.Close` et `Set rs = Nothing`, en renvoyant ici l'explication systématique. C'est l'objet de cette section : comprendre pourquoi et comment libérer proprement les objets DAO. Cette discipline n'est pas une formalité décorative. Un objet `Recordset` ou `Database` mobilise des ressources — mémoire, curseurs, verrous — et une libération négligée a des conséquences concrètes : fuites mémoire, verrous qui persistent, base de données qui reste « en cours d'utilisation » et qu'on ne peut plus compacter ni ouvrir en exclusif. Apprendre à clôturer correctement, **y compris lorsqu'une erreur survient**, fait partie intégrante d'un code DAO robuste.

## Deux opérations distinctes

Libérer un objet recouvre en réalité deux opérations différentes, qu'il ne faut pas confondre.

`Close` **ferme** l'objet et libère les ressources tenues côté moteur : curseurs, verrous, structures internes. C'est une action qui agit sur le moteur ACE.

`Set obj = Nothing` **libère la référence** VBA vers l'objet : le compteur de références diminue, et la mémoire associée peut être restituée. C'est une action qui agit côté VBA.

Les deux sont complémentaires. Pour un objet que l'on a ouvert et qui possède une méthode `Close`, la règle est donc : **`Close` d'abord, puis `Set = Nothing`**.

## Quoi fermer, quoi seulement libérer

Tous les objets DAO ne se traitent pas de la même façon.

Un `Recordset` se ferme par `Close`, puis se libère par `Set = Nothing`.

Pour une `Database`, tout dépend de son origine. Si elle a été obtenue par **`CurrentDb()`**, on se contente de `Set db = Nothing` : il ne faut **pas** appeler `Close` (point établi à la section [9.2](/09-dao-data-access-objects/02-ouverture-database.md)). Si elle a été ouverte explicitement par **`OpenDatabase`**, alors `Close` est obligatoire, suivi de `Set = Nothing` :

```vba
Dim dbExterne As DAO.Database
Set dbExterne = DBEngine.OpenDatabase("C:\Donnees\Archive.accdb")
' ... traitement ...
dbExterne.Close                ' obligatoire : base ouverte explicitement
Set dbExterne = Nothing
```

Les objets de structure se libèrent généralement par simple `Set = Nothing` (un `QueryDef` dispose toutefois d'une méthode `Close`). Quant à un `Workspace`, on ne ferme que celui qu'on a soi-même créé, jamais le `Workspace` par défaut.

## L'ordre de libération : enfant avant parent

Les objets se libèrent dans l'**ordre inverse** de leur obtention : l'objet dépendant avant celui dont il dépend. Concrètement, un `Recordset` ouvert à partir d'une `Database` se ferme et se libère **avant** la `Database`. Cette logique « dernier ouvert, premier fermé » évite de libérer un objet parent dont un objet enfant dépend encore.

## Le motif de nettoyage de base

Dans un cas simple, sans gestion d'erreurs, le nettoyage se présente ainsi :

```vba
Dim db As DAO.Database
Dim rs As DAO.Recordset
Set db = CurrentDb()
Set rs = db.OpenRecordset("tblClients", dbOpenDynaset)

' ... traitement ...

rs.Close                 ' fermer le Recordset
Set rs = Nothing         ' libérer sa référence
Set db = Nothing         ' libérer la Database (pas de Close avec CurrentDb)
```

Ce motif est correct… tant qu'aucune erreur ne survient entre l'ouverture et le nettoyage. Or c'est précisément là que réside le danger.

## Le nettoyage robuste : libérer même en cas d'erreur

Si une erreur se produit au milieu du traitement, l'exécution saute vers le gestionnaire d'erreurs et les lignes `Close` / `Set = Nothing` placées en fin de procédure sont **purement et simplement ignorées** : les objets fuient. La solution consiste à regrouper le nettoyage dans une étiquette de sortie atteinte **dans tous les cas**, succès comme erreur.

```vba
Sub TraitementRobuste()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    On Error GoTo GestionErreur

    Set db = CurrentDb()
    Set rs = db.OpenRecordset("tblClients", dbOpenDynaset)

    ' ... traitement ...

Sortie:
    On Error Resume Next            ' une erreur de nettoyage ne doit rien masquer
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

GestionErreur:
    ' journalisation / signalement de l'erreur (chapitre 13)
    Resume Sortie
End Sub
```

Trois précautions méritent d'être soulignées dans ce motif. Le test `If Not rs Is Nothing Then` évite une erreur si le `Recordset` n'a jamais été ouvert — par exemple si l'erreur est survenue **avant** l'`OpenRecordset`. Le `On Error Resume Next` placé au début de la section de nettoyage garantit qu'une éventuelle erreur lors de la fermeture ne fera pas dérailler le nettoyage lui-même. Enfin, le gestionnaire d'erreurs se termine par `Resume Sortie`, de sorte que le nettoyage est exécuté que l'on soit arrivé là par le chemin normal ou par une erreur. La structure générale de gestion des erreurs est traitée au chapitre [13](/13-gestion-erreurs/README.md), et en particulier à la section [13.2](/13-gestion-erreurs/02-on-error-goto.md).

## Faut-il toujours faire `Set = Nothing` ?

La question est légitime, car VBA libère **automatiquement** les variables objet locales lorsque la procédure se termine : leur portée s'achève, et les références sont restituées. Pour une variable strictement locale libérée tout à la fin d'une procédure, `Set = Nothing` est donc, en théorie, redondant.

Deux nuances rendent néanmoins la libération explicite recommandable. D'une part, `Close` n'est **pas** automatique et reste important : on ferme toujours explicitement les `Recordset` (et les bases ouvertes par `OpenDatabase`) afin de libérer les verrous côté moteur **sans attendre** la fin de portée. D'autre part, pour toute variable de durée de vie plus longue — niveau module, variable globale ou statique — la libération par `Set = Nothing` est **indispensable**, car la portée ne suffit pas à les restituer. En pratique, on retiendra une règle simple : toujours `Close` ce qui doit l'être, et faire `Set = Nothing` par hygiène et par cohérence, ce qui devient obligatoire dès que la variable dépasse la portée d'une procédure.

## Conséquence concrète : le fichier de verrouillage

L'enjeu de la libération devient tangible avec le **fichier de verrouillage** d'Access (`.laccdb` pour un fichier `.accdb`, `.ldb` pour un `.mdb`). Tant qu'un `Recordset` reste ouvert, des verrous y sont consignés. Une référence de `Recordset` qui fuit maintient la base « en cours d'utilisation » : le fichier de verrouillage ne se libère pas, ce qui peut empêcher le compactage de la base ou son ouverture en mode exclusif. Beaucoup de problèmes mystérieux de ce type — « impossible de compacter », « la base est utilisée par un autre processus » alors qu'on est seul — trouvent leur origine dans des objets DAO non libérés. Ces aspects sont approfondis au chapitre [15](/15-multi-utilisateurs/README.md) pour le verrouillage, et à la section [18.7](/18-optimisation-performance/07-compactage-automatique.md) pour le compactage.

## Pièges courants

Les erreurs récurrentes sont les suivantes. Oublier purement et simplement de fermer les `Recordset`. Placer le nettoyage uniquement sur le chemin de succès, où il sera sauté en cas d'erreur. Libérer les objets dans le mauvais ordre, le parent avant l'enfant. Appeler `Close` sur la `Database` issue de `CurrentDb()`. Tenter de fermer un objet déjà fermé. Et omettre le test `Is Nothing` dans un gestionnaire d'erreurs, ce qui provoque une erreur lorsque l'objet n'avait jamais été ouvert.

## Points clés à retenir

Libérer un objet DAO recouvre deux opérations : `Close`, qui restitue les ressources du moteur, et `Set = Nothing`, qui libère la référence VBA. On ferme et libère les objets dans l'ordre inverse de leur obtention, l'enfant avant le parent. Un `Recordset` se ferme toujours par `Close` puis `Set = Nothing` ; une `Database` issue de `CurrentDb()` se libère par `Set = Nothing` seul, tandis qu'une base ouverte par `OpenDatabase` doit être fermée. Surtout, le nettoyage doit s'exécuter **même en cas d'erreur** : on le regroupe dans une étiquette de sortie atteinte par tous les chemins, en protégeant chaque fermeture par un test `Is Nothing`. Une libération négligée ne reste pas sans conséquence : verrous persistants, fichier de verrouillage bloqué, base impossible à compacter.

⏭️ [9.12. RecordsetClone — manipuler les enregistrements d'un formulaire](/09-dao-data-access-objects/12-recordsetclone.md)
