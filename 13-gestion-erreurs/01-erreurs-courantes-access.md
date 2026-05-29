🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13.1. Erreurs courantes en VBA Access — catalogue et causes

## Introduction

Avant d'apprendre à intercepter et traiter les erreurs, il faut savoir les reconnaître. Un développeur Access expérimenté ne lit pas le message d'erreur dans son intégralité à chaque incident : il identifie immédiatement, au numéro et à quelques mots-clés, de quelle erreur il s'agit, de quelle couche elle provient et quelle en est la cause la plus probable. Ce réflexe ne s'acquiert qu'au contact répété d'un nombre finalement restreint d'erreurs : la grande majorité des incidents rencontrés en production se résument à quelques dizaines de codes récurrents.

Cette section constitue un catalogue de référence de ces erreurs. Elle ne vise pas l'exhaustivité — Access et ses moteurs définissent plusieurs centaines de codes — mais la pertinence : on y trouve les erreurs que vous rencontrerez réellement, classées par origine, avec pour chacune sa cause typique. L'objectif est double : accélérer le diagnostic lorsqu'une erreur survient, et nourrir l'écriture d'un code défensif capable de prévenir les situations les plus courantes.

## Les trois grandes familles d'erreurs

Toute erreur en VBA Access appartient à l'une de ces trois familles, qu'il est essentiel de distinguer car elles ne se traitent pas de la même manière.

### Les erreurs de compilation

Ce sont les erreurs détectées **avant** l'exécution, lorsque VBA analyse la syntaxe et la cohérence du code. Une variable mal orthographiée alors que `Option Explicit` est actif, une parenthèse manquante, un appel à une procédure inexistante : autant d'anomalies signalées dès la compilation, qui empêchent le code de s'exécuter. Point important pour ce chapitre : **les instructions `On Error` n'ont aucune prise sur les erreurs de compilation**. Elles se corrigent à l'écriture, pas au moment de l'exécution. Compiler régulièrement le projet (menu *Débogage → Compiler*) permet de les débusquer en amont.

### Les erreurs d'exécution

Ce sont les erreurs qui surviennent **pendant** le déroulement du programme, lorsqu'une instruction par ailleurs syntaxiquement correcte se heurte à une situation impossible : diviser par zéro, ouvrir un fichier absent, insérer une clé en double, utiliser un objet non initialisé. Ce sont elles qui font l'objet de l'ensemble de ce chapitre, car ce sont les seules que l'on peut intercepter et traiter par code grâce au mécanisme `On Error`. Chaque erreur d'exécution est identifiée par un **numéro** et accompagnée d'une **description**.

### Les erreurs de logique

Ce sont les plus insidieuses : le code s'exécute sans déclencher la moindre erreur, mais produit un résultat faux. Un total calculé à tort, un enregistrement mis à jour au mauvais endroit, une condition inversée. Aucun mécanisme de gestion d'erreur ne les détecte, car techniquement aucune erreur ne se produit. Elles relèvent du débogage et des tests (chapitre 19), non de la gestion d'erreurs proprement dite, et ne sont mentionnées ici que pour mémoire.

## Comment lire une erreur d'exécution

Lorsqu'une erreur d'exécution survient, deux informations sont disponibles : le **numéro** (`Err.Number`) et la **description** (`Err.Description`). Il est crucial de comprendre que ces deux éléments n'ont pas la même valeur diagnostique.

La description est rédigée pour l'utilisateur, dépend de la langue d'installation d'Access et reste parfois vague — la tristement célèbre « Erreur définie par l'application ou par l'objet » en est le meilleur exemple. Le numéro, lui, est **stable et indépendant de la langue** : c'est lui qui constitue l'identifiant fiable de l'erreur. C'est donc toujours sur le numéro qu'il faut raisonner pour diagnostiquer et pour écrire du code de traitement conditionnel.

Le numéro renseigne également sur l'origine de l'erreur, ce qui oriente immédiatement la recherche de la cause :

| Plage de numéros | Origine | Couche concernée |
|---|---|---|
| 1 à 1000 (environ) | Moteur d'exécution VBA | Le langage lui-même |
| 2000 à 2999 | Application Microsoft Access | L'interface et le modèle objet Access |
| 3000 à 3999 | Moteur de base de données (Jet/ACE), DAO | Le moteur de données |

Cette correspondance n'est pas absolue, mais elle constitue un repère très utile : une erreur 3xxx pointe vers la base de données, une erreur 2xxx vers l'interface Access, une erreur à deux ou trois chiffres vers le code VBA lui-même.

## Catalogue des erreurs d'exécution VBA

Ces erreurs sont communes à tout l'environnement VBA (on les rencontre aussi sous Excel ou Word). Elles trahissent généralement un problème dans la logique du code plutôt que dans les données.

| N° | Message (approximatif) | Cause typique |
|---|---|---|
| 5 | Appel de procédure ou argument incorrect | Argument hors de la plage admise (ex. `Mid` avec une position négative), fonction appelée dans un contexte invalide |
| 6 | Dépassement de capacité (*Overflow*) | Valeur trop grande pour le type cible — par exemple le produit de deux `Integer` affecté à un `Integer`, ou un compteur qui dépasse 32 767 |
| 9 | L'indice n'appartient pas à la sélection | Accès à un élément de tableau hors bornes, ou nom inexistant dans une collection (ex. `Forms("FrmInexistant")`) |
| 11 | Division par zéro | Dénominateur valant 0, ou valeur `Null` convertie en 0 dans un calcul |
| 13 | Incompatibilité de type | Affectation ou comparaison entre types inconciliables — typiquement une valeur `Null` ou un texte traité comme un nombre |
| 52 | Nom ou numéro de fichier incorrect | Numéro de canal `#` invalide ou chemin mal formé lors d'un accès fichier |
| 53 | Fichier introuvable | Fichier absent au chemin indiqué — souvent un chemin codé en dur qui n'existe pas sur le poste cible |
| 70 | Autorisation refusée | Fichier en lecture seule, déjà ouvert par un autre processus, ou droits NTFS insuffisants |
| 75 | Erreur d'accès au chemin d'accès / au fichier | Tentative d'écriture dans un emplacement protégé, fichier verrouillé |
| 76 | Chemin d'accès introuvable | Dossier inexistant — lecteur réseau déconnecté, chemin UNC erroné |
| 91 | Variable objet ou variable de bloc With non définie | Variable objet utilisée sans `Set` préalable, ou objet déjà libéré (valeur `Nothing`) |
| 94 | Utilisation incorrecte de Null | Valeur `Null` affectée à une variable qui ne l'accepte pas (`String`, `Long`, etc.) |
| 424 | Objet requis | Référence à un objet inexistant ou mal qualifié dans une expression |
| 438 | Cet objet ne gère pas cette propriété ou cette méthode | Appel d'un membre inexistant sur l'objet — faute de frappe, mauvais type d'objet, ou mauvaise version de bibliothèque référencée |
| 457 | Cette clé est déjà associée à un élément de cette collection | Ajout d'une clé en double dans un objet `Collection` ou un `Scripting.Dictionary` |
| 462 | L'ordinateur serveur distant n'existe pas ou n'est pas disponible | En Automation, objet d'une autre application mal qualifié (ex. `Cells` au lieu de `xlApp.Cells`) ou instance fermée prématurément |

L'erreur **91** mérite une attention particulière : c'est probablement la plus fréquente en VBA, et elle signale presque toujours un `Set` oublié, un objet qui n'a pas pu être créé, ou une fonction censée renvoyer un objet qui a renvoyé `Nothing`.

## Catalogue des erreurs du moteur de base de données (DAO / ACE)

Ces erreurs (numéros 3xxx) proviennent du moteur de données. Elles concernent les opérations sur les tables, les requêtes et les enregistrements, et sont au cœur des applications Access. Beaucoup d'entre elles ne dépendent pas du code mais de l'**état des données** ou de l'**environnement** (réseau, droits, concurrence).

| N° | Message (approximatif) | Cause typique |
|---|---|---|
| 3021 | Aucun enregistrement en cours | Lecture d'un champ alors que le recordset est vide ou positionné sur EOF / BOF |
| 3022 | Les modifications créeraient des valeurs en double dans l'index, la clé primaire ou la relation | Insertion ou mise à jour violant une clé primaire ou un index unique |
| 3027 | Mise à jour impossible. Base de données ou objet en lecture seule | Base ouverte en lecture seule, requête non modifiable, ou fichier `.accdb` en lecture seule |
| 3043 | Erreur de disque ou de réseau | Connexion réseau vers le back-end perdue pendant une opération |
| 3044 | `<chemin>` n'est pas un chemin d'accès valide | Chemin du back-end introuvable — table liée pointant vers un emplacement déplacé |
| 3045 | Utilisation impossible ; fichier déjà utilisé | Tentative d'ouverture **exclusive** d'une base déjà ouverte par un autre utilisateur |
| 3050 | Verrouillage du fichier impossible | Impossible de créer ou d'écrire le fichier de verrouillage `.laccdb` — droits insuffisants sur le dossier du back-end |
| 3051 | Le moteur de base de données ne peut pas ouvrir le fichier | Permissions insuffisantes sur le fichier ou son dossier, partage réseau inaccessible |
| 3061 | Trop peu de paramètres. N attendu(s) | Un nom de champ ou de paramètre n'est pas reconnu — le plus souvent une **faute de frappe** dans un nom de champ, que le moteur interprète alors comme un paramètre à renseigner |
| 3075 | Erreur de syntaxe dans l'expression | SQL mal formé — apostrophe non échappée, parenthèse manquante, séparateur de date incorrect |
| 3078 | Le moteur ne trouve pas la table ou la requête source | Table ou requête inexistante ou mal orthographiée dans le SQL |
| 3146 | ODBC — échec de l'appel | Échec d'une opération sur une table liée ODBC — serveur indisponible, requête refusée par le serveur, délai dépassé |
| 3167 | L'enregistrement est supprimé | L'enregistrement courant a été supprimé par un autre utilisateur depuis sa lecture |
| 3186 / 3188 | Enregistrement actuellement verrouillé par l'utilisateur `<x>` | Conflit de verrouillage en environnement multi-utilisateur |
| 3197 | Un autre utilisateur a modifié les mêmes données (conflit d'écriture) | Deux utilisateurs ont modifié le même enregistrement (verrouillage optimiste) |
| 3200 | Impossible de supprimer ou de modifier l'enregistrement : la table inclut des enregistrements liés | Suppression d'un enregistrement parent référencé par des enfants, sans suppression en cascade |
| 3201 | Impossible d'ajouter ou de modifier : un enregistrement lié est requis | Insertion d'un enfant sans parent existant — clé étrangère orpheline |

Deux cas reviennent constamment en production. L'erreur **3061** (« Trop peu de paramètres ») n'a presque jamais pour cause un véritable paramètre manquant : neuf fois sur dix, il s'agit d'un nom de champ mal orthographié dans une requête, que le moteur ne reconnaît pas et prend pour un paramètre. Les erreurs **3022, 3200 et 3201** forment quant à elles le trio de l'intégrité des données : elles signalent respectivement une clé en double et les deux faces de l'intégrité référentielle (impossible d'orpheliner un enfant, impossible de supprimer un parent encore référencé).

## Catalogue des erreurs spécifiques à l'interface Access

Ces erreurs (numéros 2xxx) sont propres à Access. Elles surviennent lors de la manipulation des formulaires, des états, des contrôles et de l'exécution d'actions `DoCmd`. Beaucoup sont liées au **contexte** : une action demandée dans un mode ou une fenêtre où elle n'a pas de sens.

| N° | Message (approximatif) | Cause typique |
|---|---|---|
| 2001 | Vous avez annulé l'opération précédente | Opération interrompue, souvent en cascade à la suite d'une autre erreur |
| 2046 | La commande ou l'action n'est pas disponible actuellement | Action `DoCmd` / `RunCommand` lancée dans un contexte incompatible — aucun objet actif, mode inadapté |
| 2105 | Vous ne pouvez pas atteindre l'enregistrement spécifié | `GoToRecord acNewRec` alors que `AllowAdditions = False`, ou déplacement au-delà du dernier enregistrement |
| 2110 | Access ne peut pas placer le focus sur le contrôle `<x>` | `SetFocus` sur un contrôle masqué (`Visible = False`) ou désactivé (`Enabled = False`) |
| 2113 | La valeur entrée n'est pas valide pour ce champ | Saisie incompatible avec le type ou le masque du champ |
| 2169 | Vous ne pouvez pas enregistrer cet enregistrement actuellement | Échec d'enregistrement — fréquemment à la suite d'un `Cancel` dans `BeforeUpdate` ou d'une règle de validation non respectée |
| 2237 | Le texte entré ne correspond à aucun élément de la liste | Saisie hors liste dans une zone de liste déroulante avec `LimitToList = True` |
| 2448 | Vous ne pouvez pas affecter de valeur à cet objet | Affectation à un contrôle non lié, calculé, ou à une propriété en lecture seule dans le contexte courant |
| 2450 | Access ne trouve pas le formulaire référencé | Référence `Forms!NomFormulaire` à un formulaire fermé ou inexistant |
| 2465 | Access ne trouve pas le champ auquel votre expression fait référence | Nom de champ ou de contrôle inexistant dans une expression — faute de frappe, ou champ absent du `RecordSource` |
| 2467 | L'expression fait référence à un objet fermé ou inexistant | Accès à un objet (formulaire, contrôle) déjà fermé |
| 2486 | Vous ne pouvez pas exécuter cette action actuellement | Action lancée dans un contexte ou un mode incompatible |
| 2501 | L'action `<nom>` a été annulée | Action `DoCmd` (`OpenForm`, `OpenReport`...) interrompue — typiquement un état sans données qui se ferme via l'événement `NoData`, un `Cancel = True` dans un événement, ou l'utilisateur qui annule une boîte de dialogue |

L'erreur **2501** est sans doute la plus déroutante pour le débutant, car elle n'est pas toujours le signe d'un véritable problème : elle survient légitimement lorsqu'un état est annulé parce qu'il ne contient aucune donnée, ou lorsqu'un traitement est volontairement interrompu par `Cancel`. C'est l'archétype de l'erreur « attendue » qu'un bon gestionnaire d'erreurs doit savoir reconnaître et ignorer silencieusement.

## Le cas de « l'erreur définie par l'application ou par l'objet »

Parmi toutes les descriptions possibles, « Erreur définie par l'application ou par l'objet » est la plus frustrante car la moins informative. Elle apparaît lorsqu'un objet — Access, une bibliothèque, ou un serveur d'Automation — déclenche une erreur qui ne correspond à aucun code VBA standard doté d'une description claire.

Face à cette description, la consultation du seul `Err.Number` ne suffit généralement pas. Le diagnostic repose alors sur trois éléments complémentaires : la propriété `Err.Source` (qui indique quel composant a levé l'erreur), la **ligne exacte** où l'erreur s'est produite, et le contexte applicatif (quelle opération était en cours). Pour les erreurs issues de DAO ou d'ADO, l'information réellement utile ne se trouve d'ailleurs pas dans l'objet `Err` mais dans les collections d'erreurs spécifiques de ces technologies — un point développé dans la section [13.5](/13-gestion-erreurs/05-erreurs-dao-ado.md).

## Pour mémoire : les erreurs de compilation les plus fréquentes

Bien qu'elles échappent au mécanisme `On Error`, ces erreurs valent d'être connues car elles se rencontrent très souvent à l'écriture, et certaines surprennent (notamment celles liées aux références manquantes lors d'un changement de poste ou de version d'Office).

| Message (approximatif) | Cause typique |
|---|---|
| Variable non définie | `Option Explicit` actif et variable utilisée sans déclaration — généralement une faute de frappe |
| Sub ou Function non définie | Appel d'une procédure inexistante ou mal orthographiée |
| Type défini par l'utilisateur non défini | Référence à une bibliothèque manquante (mention `MANQUANT` dans la fenêtre *Références*) — fréquent après un changement de poste ou de version d'Office |
| Membre de méthode ou de données introuvable | Membre (`.xxx`) inexistant sur l'objet, détecté dès la compilation en liaison précoce |
| Argument non facultatif | Procédure appelée sans fournir un argument obligatoire |
| Incompatibilité de type ByRef | Type de l'argument transmis différent du type attendu par le paramètre `ByRef` |
| Attendu : fin d'instruction | Erreur de syntaxe — instruction mal formée |

## En résumé

La gestion d'erreurs commence par la reconnaissance. Trois familles coexistent : les erreurs de **compilation** (corrigées à l'écriture, hors de portée de `On Error`), les erreurs d'**exécution** (le cœur de ce chapitre, identifiées par un numéro stable), et les erreurs de **logique** (silencieuses, relevant du débogage).

Pour les erreurs d'exécution, le numéro prime sur la description : il est fiable, indépendant de la langue, et sa plage indique l'origine — VBA pour les petits numéros, Access pour les 2xxx, le moteur de données pour les 3xxx. Quelques codes concentrent l'essentiel des incidents réels : l'**erreur 91** (objet non initialisé) côté VBA, les **erreurs 3022, 3061, 3200 et 3201** côté données, et les **erreurs 2465 et 2501** côté interface. Les connaître permet de diagnostiquer en quelques secondes ce qui, autrement, conduirait à de longues recherches.

Maintenant que ces erreurs sont identifiées, la section suivante construit le mécanisme qui permet de les intercepter de façon systématique et fiable.

---


⏭️ [13.2. On Error GoTo — structure robuste de gestion d'erreurs](/13-gestion-erreurs/02-on-error-goto.md)
