🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 1.2. Différences fondamentales entre VBA Excel et VBA Access

La section précédente l'a souligné : le langage VBA est commun à toutes les applications Office, mais ce qu'il permet de faire dépend de l'application hôte. Cette section approfondit ce point pour les nombreux développeurs qui découvrent Access après avoir pratiqué VBA dans Excel. Le réflexe naturel — transposer directement ses habitudes — conduit souvent à des contresens, car les deux environnements reposent sur des logiques profondément différentes.

Bonne nouvelle d'emblée : tout ce que vous savez du *langage* lui-même reste valable. La difficulté, si l'on n'y prend pas garde, vient de tout le reste — la façon de penser les données, de les atteindre, de construire l'interface et de réagir aux actions de l'utilisateur.

## Un langage commun, deux univers

VBA est strictement le même langage dans Excel et dans Access : même éditeur (le VBE), même runtime, même syntaxe, mêmes structures de contrôle, mêmes fonctions intrinsèques (`Len`, `Mid`, `Format`, `IIf`, `Now`…). Une boucle `For…Next`, une `Function`, une gestion d'erreurs `On Error GoTo` s'écrivent à l'identique.

Ce qui change, c'est le **modèle objet** que ce langage pilote. Or, en VBA, l'essentiel d'un programme consiste justement à manipuler les objets de l'application hôte. C'est pourquoi, malgré un socle identique, écrire du VBA pour Access ressemble assez peu à en écrire pour Excel. On peut résumer la situation ainsi : **le vocabulaire et la grammaire sont les mêmes, mais on ne parle pas du tout des mêmes choses.**

## La différence de fond : tableur contre base de données relationnelle

Toutes les autres différences découlent de celle-ci. Excel est un **tableur** : une grille libre de cellules où données et présentation se mêlent, sans structure imposée. On y place ce que l'on veut, où on veut, et les formules relient les cellules entre elles.

Access est une **base de données relationnelle** : les données vivent dans des tables au schéma défini (colonnes typées, clés, relations, contraintes d'intégrité), normalisées pour éviter les redondances. Surtout, la présentation est **séparée** des données : les tables stockent, les formulaires et les états affichent. Une même donnée peut alimenter plusieurs formulaires et plusieurs états sans jamais être dupliquée.

Ce changement de paradigme impose un changement de réflexes. Là où l'on pensait « telle valeur dans telle cellule », on pense désormais « tel champ de tel enregistrement ». Là où l'on parcourait une plage ligne par ligne, on raisonne en **ensembles** d'enregistrements. C'est sans doute l'ajustement mental le plus important pour qui vient d'Excel.

## L'unité de travail : Range contre Recordset

Dans Excel, l'unité de manipulation est la **cellule** (ou la plage, via l'objet `Range`). On y accède le plus souvent par **position** : `Cells(2, 3)`, `Range("A1:C10")`. Pour traiter des données, le réflexe est de boucler sur des lignes et des colonnes.

Dans Access, cette notion de position n'existe pas. Les données n'ont pas d'« adresse » de type A1. On les manipule par l'intermédiaire d'un **Recordset** (DAO ou ADO), c'est-à-dire un jeu d'enregistrements, dont on lit les champs **par leur nom**. Et, le plus souvent, on évite même de parcourir les enregistrements un à un : on confie le travail à une requête SQL, qui opère sur des lots entiers en une seule instruction.

Le contraste se voit bien sur une tâche simple — additionner une colonne de montants.

```vba
' Réflexe Excel : on parcourt les cellules une à une
Dim total As Double, i As Long
For i = 2 To 100
    total = total + Worksheets("Ventes").Cells(i, "C").Value
Next i
```

```vba
' Réflexe Access : une approche ensembliste fait tout le travail
Dim total As Currency
total = DSum("Montant", "tblVentes")
```

La version Access ne décrit pas *comment* parcourir les données : elle décrit *quoi* obtenir. C'est tout l'esprit du travail ensembliste, étranger à la logique cellulaire d'Excel.

## Deux modèles objet distincts

Puisque l'essentiel du VBA consiste à piloter le modèle objet, il faut bien mesurer à quel point celui d'Access diffère de celui d'Excel.

| | Excel | Access |
|---|---|---|
| Sommet | `Application` | `Application` |
| Conteneur de données | `Workbook` → `Worksheet` | `CurrentDb` → `TableDef` / requêtes |
| Donnée élémentaire | `Range` / `Cells` | `Recordset` → `Field` |
| Interface | `Worksheet`, `UserForm` | `Form`, `Report`, contrôles |
| Automatisation | objet métier directement | `DoCmd` |

Concrètement, là où Excel expose `Workbooks`, `Worksheets` et `Range`, Access expose des collections comme `Forms`, `Reports`, `AllForms`, `AllTables`, l'objet `CurrentDb` pour les données, et l'objet `DoCmd` pour les actions. On ne retrouve donc presque aucun des objets familiers d'Excel ; il faut apprendre une nouvelle hiérarchie (traitée en détail au chapitre 4).

La manière de référencer un élément d'interface illustre bien ce dépaysement :

```vba
' Excel
Worksheets("Feuil1").Range("A1").Value = "Bonjour"

' Access
Forms!frmClients!txtNom.Value = "Bonjour"   ' depuis l'extérieur du formulaire
Me.txtNom = "Bonjour"                        ' depuis le code du formulaire lui-même
```

## DoCmd : un objet central propre à Access

Access possède un objet sans réel équivalent dans Excel : **`DoCmd`**. Il sert de point d'entrée à une grande partie de l'automatisation — ouvrir un formulaire, lancer un état, exécuter une requête, importer ou exporter des données, naviguer entre enregistrements.

```vba
DoCmd.OpenForm "frmClients"
DoCmd.OpenReport "rptFactures", acViewPreview
DoCmd.RunSQL "UPDATE tblTarifs SET PrixHT = PrixHT * 1.02"
```

Dans Excel, ces opérations passeraient directement par le modèle objet (`Workbooks.Open`, etc.) ; il n'existe pas d'objet centralisant ainsi les commandes. Découvrir `DoCmd` et son rôle est l'une des premières étapes de l'apprentissage de VBA pour Access ; le chapitre 5 lui est entièrement consacré.

## L'absence d'enregistreur de macros dans Access

C'est une surprise fréquente — et de taille — pour qui vient d'Excel. Dans Excel, l'**enregistreur de macros** observe vos actions et génère le code VBA correspondant. C'est un formidable outil d'apprentissage : on enregistre, on lit le code produit, on s'en inspire.

**Access ne possède pas d'enregistreur de macros.** Et le mot « macro » y désigne tout autre chose : non pas du VBA enregistré, mais des objets d'automatisation distincts, composés d'actions prédéfinies (voir sections 1.1 et 1.3). Conséquence pratique : dans Access, on ne peut pas « enregistrer pour apprendre ». Il faut écrire le code soi-même, donc connaître le modèle objet dès le départ. C'est l'une des raisons pour lesquelles VBA pour Access est réputé plus exigeant à aborder — et pourquoi une bonne compréhension du modèle objet (chapitre 4) y est si précieuse.

## Une programmation profondément événementielle

Dans Excel, beaucoup de solutions sont **procédurales** : l'utilisateur lance une macro, celle-ci s'exécute du début à la fin. Les événements existent (`Worksheet_Change`, `Workbook_Open`), mais ils occupent souvent une place secondaire.

Dans Access, c'est l'inverse : l'application tout entière est **pilotée par les événements**. Chaque formulaire et chaque contrôle expose un modèle d'événements riche — `Current` (arrivée sur un enregistrement), `BeforeUpdate` et `AfterUpdate` (avant/après validation), `Click`, `GotFocus`, `LostFocus`, et bien d'autres. Le développement consiste largement à placer la bonne logique dans le bon événement : valider une saisie avant l'enregistrement, recalculer un champ quand un autre change, charger des données à l'ouverture d'un formulaire.

Cette orientation événementielle change la façon de concevoir un programme. On n'écrit plus tant un « script qui se déroule » qu'un ensemble de réactions déclenchées par l'utilisateur et par le cycle de vie des formulaires (détaillé au chapitre 8).

## Formulaires liés aux données contre UserForms indépendants

Dans Excel, l'interface de saisie, c'est avant tout la grille ; pour des écrans personnalisés, on utilise des **UserForms**, largement **indépendants** des données : un contrôle peut tout au plus être lié à une cellule, et c'est le plus souvent au code de transférer manuellement les valeurs dans un sens puis dans l'autre.

Dans Access, le formulaire est l'élément central de l'interface, et il est généralement **lié aux données**. On définit une source d'enregistrements (`RecordSource`) pour le formulaire et une source de contrôle (`ControlSource`) pour chaque contrôle : dès lors, l'affichage et la mise à jour des données se font **automatiquement**, sans code de transfert. Saisir dans une zone de texte liée modifie directement le champ correspondant de l'enregistrement courant.

Cette notion de **liaison** (binding) n'a quasiment pas d'équivalent dans Excel. C'est pourtant un pilier d'Access : elle explique pourquoi tant de choses fonctionnent « toutes seules » et pourquoi on écrit, au final, beaucoup moins de code de plomberie qu'on ne le ferait avec des UserForms Excel.

## La persistance des données : un changement de réflexe

Dans Excel, les modifications restent en mémoire tant que l'on n'a pas **enregistré le classeur** ; tant que l'on n'a pas sauvegardé, on peut tout annuler en fermant sans enregistrer.

Dans Access, la logique est différente et déroute souvent au début : les modifications de données sont **écrites automatiquement** dans la table. Dès que l'on quitte un enregistrement (ou que l'on ferme le formulaire), la modification est validée et persistée. Il n'y a pas de « sauvegarde du document » pour les données : le fichier `.accdb` n'est pas un document que l'on enregistre comme un classeur. (Les transactions, abordées au chapitre 14, permettent de regrouper et, au besoin, d'annuler un ensemble de modifications, mais c'est une démarche volontaire.)

Le réflexe à acquérir : dans Access, on ne « sauvegarde » pas ses données, elles le sont déjà. Ce qui doit être enregistré explicitement, ce sont les modifications de *structure* (un formulaire ou un état que l'on vient de modifier en conception), pas les données elles-mêmes.

## L'accès aux données : DAO et ADO au cœur d'Access

Dans Excel, les données sont dans les cellules, à portée de main ; il n'existe pas de couche d'accès aux données native (on peut interroger des sources externes via ADO, mais ce n'est pas l'usage central).

Dans Access, l'accès programmatique aux données passe par des bibliothèques dédiées : **DAO** (*Data Access Objects*), native et utilisée par défaut, et **ADO** (*ActiveX Data Objects*), plus orientée sources externes. L'objet `CurrentDb` est la porte d'entrée la plus courante vers la base ouverte. Maîtriser les recordsets, leur navigation et leurs opérations (lecture, ajout, modification, suppression) est une compétence centrale en VBA Access — sans véritable contrepartie côté Excel. Les chapitres 9 (DAO) et 10 (ADO) y sont consacrés.

## Le multi-utilisateur, natif dans Access

Excel est essentiellement un outil **mono-document** : un classeur, généralement un utilisateur à la fois. Access, lui, est conçu dès l'origine pour le **multi-utilisateur**, typiquement selon une architecture front-end / back-end (interface distribuée à chacun, données partagées sur le réseau).

Cela introduit des préoccupations absentes du monde Excel : verrouillage des enregistrements, gestion des conflits de mise à jour, génération fiable de numéros séquentiels en environnement concurrent. Ces sujets, qui n'ont pas d'équivalent en VBA Excel, font l'objet du chapitre 15.

## Des réflexes de performance différents

Les deux environnements ont aussi leurs bonnes pratiques propres en matière de performance.

Côté Excel, on cherche surtout à **éviter les allers-retours avec la grille** : lire et écrire en bloc via des tableaux plutôt que cellule par cellule, désactiver le rafraîchissement de l'écran (`Application.ScreenUpdating = False`) et le recalcul automatique.

Côté Access, le principe directeur est de **privilégier les opérations ensemblistes** : une requête SQL qui met à jour des milliers d'enregistrements en une instruction sera bien plus efficace qu'une boucle parcourant un recordset. On désactive aussi les messages système (`DoCmd.SetWarnings False`) et, au besoin, le rafraîchissement de l'interface (`Application.Echo False`). Ces techniques sont approfondies au chapitre 18.

## Tableau de synthèse Excel ↔ Access

| Aspect | VBA dans Excel | VBA dans Access |
|---|---|---|
| Paradigme | Tableur (grille de cellules) | Base de données relationnelle |
| Unité de données | Cellule / plage (`Range`) | Enregistrement / champ (`Recordset`) |
| Adressage des données | Par position (A1, `Cells(i, j)`) | Par nom de champ, via SQL / recordset |
| Modèle objet principal | `Application` → `Workbook` → `Worksheet` → `Range` | `Application` → `Forms` / `Reports` / `CurrentDb` / `DoCmd` |
| Objet d'automatisation central | (pas d'équivalent direct) | `DoCmd` |
| Enregistreur de macros | Oui (génère du VBA) | Non |
| Sens du mot « macro » | Code VBA enregistré | Objet d'actions, distinct du VBA |
| Programmation événementielle | Présente mais secondaire | Centrale et omniprésente |
| Interface utilisateur | La grille + UserForms (indépendants) | Formulaires liés aux données + états |
| Liaison aux données | Quasi inexistante | Native (`RecordSource` / `ControlSource`) |
| Persistance des données | Sauvegarde explicite du classeur | Enregistrements écrits automatiquement |
| Accès aux données | Données dans les cellules | DAO / ADO + SQL (via `CurrentDb`) |
| Recalcul automatique | Oui (moteur de calcul) | Non |
| Multi-utilisateur | Mono-document | Natif (front-end / back-end) |
| Volumétrie | ~1 million de lignes par feuille | Données relationnelles indexées (jusqu'à 2 Go par fichier) |

## À retenir

- Le **langage VBA et l'éditeur sont identiques** ; ce qui change radicalement, c'est le modèle objet et le paradigme.
- Excel raisonne en **cellules et en positions** ; Access raisonne en **enregistrements, champs et ensembles** (SQL).
- On ne manipule pas les données d'Access cellule par cellule : on passe par des **recordsets** (DAO/ADO) et par **SQL**.
- Access est **profondément événementiel** et repose sur des **formulaires liés aux données**, là où Excel mêle données et présentation dans la grille.
- Les données Access sont **persistées automatiquement** : il n'y a pas de « sauvegarde du document » comme dans Excel.
- Access **n'a pas d'enregistreur de macros** ; il faut écrire le code, ce qui suppose de connaître le modèle objet dès le départ.

---


⏭️ [1.3. Macros Access vs code VBA — quand choisir quoi](/01-introduction-vba-access/03-macros-vs-vba.md)
