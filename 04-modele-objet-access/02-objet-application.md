🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 4.2. L'objet Application

Au sommet du modèle objet trône l'objet **`Application`**. Il représente l'**instance d'Access en cours d'exécution** et constitue le point d'accès à tout le reste. Comme l'a souligné la section 4.1, il est **implicite** dans le code : on écrit `Forms` ou `DLookup(...)` sans le préfixer. Cette section détaille ce qu'il expose, et un usage moins connu : piloter Access depuis une autre application.

## Ce que représente l'objet Application

`Application` est la **racine** du versant application du modèle objet (section 4.1). Il correspond à l'application Access elle-même, telle qu'elle tourne à un instant donné. À travers lui, on accède aux formulaires et états ouverts, au projet et aux données courants, au moteur DAO, aux options, et à quantité de méthodes utilitaires.

Parce qu'il est la racine, il est **implicite** dans le code écrit *à l'intérieur* d'Access : les deux écritures suivantes sont équivalentes, et l'on emploie naturellement la forme courte.

```vba
Debug.Print Application.Forms.Count
Debug.Print Forms.Count              ' identique : Application est sous-entendu
```

## Les principales propriétés

L'objet `Application` expose de nombreuses propriétés. Les plus utiles :

| Propriété | Rôle | Détail |
|---|---|---|
| `CurrentProject` | objets de l'application courante | section 4.3 |
| `CurrentData` | objets de données courants | section 4.3 |
| `Forms` / `Reports` | formulaires / états **ouverts** | chapitres 6 et 7 |
| `Modules` | modules de code ouverts | chapitre 2 |
| `DBEngine` | racine du modèle DAO | sections 4.5 et 4.6 |
| `Screen` | objet actif (formulaire, contrôle…) | section 4.7 |
| `DoCmd` | objet d'automatisation | chapitre 5 |
| `VBE` | l'éditeur Visual Basic (extensibilité) | chapitre 2 |
| `References` | bibliothèques référencées | section 2.5 |
| `Version` | version d'Access (ex. « 16.0 ») | — |
| `Name` | « Microsoft Access » | — |
| `Visible` | Access est-il visible | (utile en automation) |
| `CurrentObjectName` / `CurrentObjectType` | nom et type de l'objet actif | — |

```vba
Debug.Print Application.Version             ' ex. "16.0"
Debug.Print Application.CurrentObjectName   ' nom de l'objet actif
Debug.Print Application.CurrentProject.Name ' nom du projet courant
```

## Les principales méthodes

Côté méthodes, `Application` en propose un éventail tout aussi riche.

| Méthode | Rôle | Détail |
|---|---|---|
| `DLookup`, `DSum`, `DCount`, `DMax`, `DMin`, `DAvg`… | fonctions de **domaine** | section 11.10 |
| `Eval` | évaluer une expression fournie en chaîne | — |
| `Run` | exécuter une procédure par son **nom** | — |
| `GetOption` / `SetOption` | lire / écrire une **option** d'Access | lien section 1.4 |
| `Echo` | suspendre / rétablir le rafraîchissement de l'écran | chapitre 18 |
| `Quit` | quitter Access | (voir aussi `DoCmd.Quit`, ch. 5) |
| `CompactRepair` | compacter et réparer une base | section 18.7 |
| `RefreshDatabaseWindow` | rafraîchir le volet de navigation | — |
| `FollowHyperlink` | ouvrir une URL ou un fichier | — |
| `OpenCurrentDatabase` / `CloseCurrentDatabase` / `NewCurrentDatabase` | gérer la base courante | (surtout en automation) |

Trois de ces méthodes méritent un éclairage particulier.

**Les fonctions de domaine sont des méthodes d'`Application`.** `DLookup`, `DCount`, `DSum`, etc. — omniprésentes sous Access — sont en réalité des méthodes de l'objet `Application`. On les appelle directement, justement parce que `Application` est implicite. Elles sont étudiées en section 11.10.

```vba
Dim nb As Long
nb = DCount("*", "tblClients")     ' équivaut à Application.DCount("*", "tblClients")
```

**Lire et écrire les options par code, avec `GetOption`/`SetOption`.** Les réglages du menu *Fichier > Options* (section 1.4) sont accessibles par programmation, ce qui permet d'adapter le comportement d'Access à la volée.

```vba
Debug.Print Application.GetOption("Show Status Bar")
Application.SetOption "Confirm Action Queries", False
```

**Exécuter une procédure dynamiquement, avec `Run`.** `Application.Run` lance une procédure désignée par son **nom** (sous forme de chaîne), avec d'éventuels arguments — pratique pour des appels dynamiques ou vers une base bibliothèque.

```vba
Application.Run "modUtilitaires.Recalculer"
```

Citons enfin `Echo`, qui suspend le rafraîchissement de l'écran le temps d'un traitement lourd (un grand classique d'optimisation, chapitre 18), et `Quit`, qui ferme Access :

```vba
Application.Echo False         ' fige l'affichage
' ... traitement intensif ...
Application.Echo True          ' rétablit l'affichage

Application.Quit acQuitSaveNone   ' quitter sans enregistrer
```

## Application comme cible d'automation

Jusqu'ici, `Application` désignait **notre propre** instance d'Access. Mais l'objet a un second usage, à l'envers : **piloter Access depuis une autre application** — par exemple depuis Excel ou Word. On crée alors une instance d'`Access.Application`, on lui ouvre une base, on la manipule, puis on la referme.

```vba
' Depuis une AUTRE application : piloter Access (liaison tardive)
Dim accApp As Object
Set accApp = CreateObject("Access.Application")
accApp.OpenCurrentDatabase "C:\Apps\Gestion.accdb"
accApp.Visible = True
' ... manipulations via accApp ...
accApp.Quit
Set accApp = Nothing
```

La distinction est importante : **à l'intérieur** d'Access, `Application` est implicite et désigne l'instance courante ; **depuis l'extérieur**, il faut **créer et référencer** explicitement une instance d'`Access.Application`. Ce scénario relève de l'automation inter-applications (chapitre 22) et du choix entre liaison précoce et tardive (section 2.6).

## À retenir

- L'objet **`Application`** est la **racine** du modèle objet Access ; il représente l'instance en cours et est **implicite** dans le code (`Forms`, `DLookup(...)`…).
- Il expose les grandes branches du modèle (**`CurrentProject`**, **`CurrentData`**, `Forms`, `Reports`, `DBEngine`, `Screen`, `DoCmd`…) ainsi que des propriétés d'information (`Version`, `CurrentObjectName`…).
- Les **fonctions de domaine** (`DLookup`, `DCount`…) sont des **méthodes d'`Application`** (section 11.10) ; `GetOption`/`SetOption` lisent et écrivent les **options** par code (section 1.4) ; `Run` exécute une procédure par son nom ; `Echo` optimise l'affichage (chapitre 18).
- Le même objet sert à **piloter Access depuis l'extérieur** via `CreateObject("Access.Application")` (chapitre 22) — à distinguer de l'usage interne implicite.

---


⏭️ [4.3. CurrentProject et CurrentData](/04-modele-objet-access/03-currentproject-currentdata.md)
