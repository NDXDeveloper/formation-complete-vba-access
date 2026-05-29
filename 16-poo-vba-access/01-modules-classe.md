🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 16.1. Modules de classe — création et instanciation

Toute la programmation orientée objet en VBA repose sur une brique unique : le **module de classe**. Comprendre ce qu'il est, comment le créer, comment en tirer des objets et comment ces objets naissent et meurent constitue le socle de tout ce qui suit dans ce chapitre. Cette section pose ce socle, en laissant volontairement de côté les propriétés et méthodes (objets des [sections 16.2](02-proprietes-get-let-set.md) et [16.3](03-methodes-encapsulation.md)) pour se concentrer sur la classe elle-même et son instanciation.

## Qu'est-ce qu'un module de classe ?

Un **module de classe** définit une **classe** : un modèle, un plan de construction à partir duquel on crée des **objets** (les instances). La classe décrit ce que contiendront ses objets (leur état interne) et ce qu'ils sauront faire (leurs comportements) ; chaque objet créé à partir d'elle est une entité indépendante, dotée de sa propre copie des données.

L'analogie classique est celle du moule et des pièces : la classe est le moule, les objets sont les pièces qu'il produit. On définit le moule une fois, et l'on fabrique autant de pièces que nécessaire, chacune distincte des autres.

## Module de classe ou module standard ?

La distinction avec un **module standard** est fondamentale.

Un module standard contient du code et des données **globaux** : ses procédures publiques sont appelables partout, ses variables existent en un seul exemplaire partagé, et il ne s'instancie pas. C'est l'outil idéal pour des fonctions utilitaires ou des aides applicatives transversales.

Un module de classe, lui, définit un **type** que l'on **instancie**. Chaque `New` crée un objet indépendant possédant son propre état. C'est l'outil pour modéliser des entités porteuses de données — un client, une commande, un dépôt de données.

En résumé : un module standard, c'est « un seul exemplaire, partout » ; un module de classe, c'est « autant d'exemplaires que l'on veut, chacun avec ses propres données ».

À noter que les **modules de formulaire et d'état sont eux-mêmes des modules de classe** (cf. [section 2.3](/02-interface-environnement/03-types-modules.md)) : c'est précisément pour cette raison qu'un formulaire peut être ouvert en plusieurs instances.

## Créer un module de classe

Dans l'éditeur VBA, on insère un module de classe par le menu **Insertion → Module de classe**. Le nouveau module apparaît dans l'explorateur de projets sous la rubrique « Modules de classe ».

On le nomme via la fenêtre **Propriétés** (touche F4), à travers sa propriété `(Name)`. Ce nom est celui que l'on utilisera pour déclarer des variables du type correspondant. Une convention répandue préfixe le nom par `cls` (par exemple `clsClient`), selon les conventions de nommage abordées à la [section 24.1](/24-bonnes-pratiques-ressources/01-conventions-nommage-access.md).

## Anatomie d'une classe minimale

Voici une classe réduite à l'essentiel. Elle déclare un état interne privé et définit ses deux événements de cycle de vie — sans propriété ni méthode publique pour l'instant.

```vba
' --- Module de classe nommé : clsClient ---
Option Compare Database
Option Explicit

' État interne (privé) — l'encapsulation est détaillée en 16.3
Private m_Id As Long
Private m_Nom As String

' Appelé automatiquement à la création de l'objet (New)
Private Sub Class_Initialize()
    Debug.Print "clsClient : objet créé"
End Sub

' Appelé automatiquement à la destruction de l'objet
Private Sub Class_Terminate()
    Debug.Print "clsClient : objet détruit"
End Sub
```

`Option Explicit` est, comme toujours, fortement recommandé. Les champs `m_Id` et `m_Nom` sont **privés** : ils constituent l'état interne de chaque objet, que l'on exposera ensuite de manière contrôlée via des propriétés. Les membres publics qu'on ajoutera (procédures et propriétés) formeront l'**interface** de la classe — ce que le monde extérieur peut manipuler.

## Instancier un objet : New

Disposer d'une classe ne suffit pas ; il faut en créer des instances. Cela se fait en deux temps : déclarer une variable du type de la classe, puis lui affecter un nouvel objet avec le mot-clé **`New`**.

```vba
' --- Dans un module standard ou un module de formulaire ---
Public Sub DemoInstanciation()
    Dim c As clsClient

    Set c = New clsClient        ' déclenche Class_Initialize → "objet créé"

    ' ... utilisation de l'objet (propriétés et méthodes : sections 16.2 et 16.3) ...

    Set c = Nothing              ' libère la référence → Class_Terminate → "objet détruit"
End Sub
```

Deux points méritent l'attention. D'abord, l'affectation d'un objet exige le mot-clé **`Set`** : un objet est un **type référence**, et `Set` indique qu'on affecte une référence, non une valeur. Omettre `Set` provoque une erreur. Ensuite, c'est l'instruction `Set c = New clsClient` qui crée réellement l'objet et déclenche son `Class_Initialize`.

## Les deux styles de déclaration — et lequel éviter

VBA autorise deux façons d'écrire l'instanciation, mais elles ne se valent pas.

La forme **explicite**, recommandée, sépare la déclaration de la création :

```vba
Dim c As clsClient
Set c = New clsClient
```

La forme **auto-instanciée** combine les deux :

```vba
Dim c As New clsClient        ' à éviter
```

Cette seconde forme est déconseillée. Avec elle, l'objet est créé automatiquement à sa première utilisation, ce qui fait **perdre le contrôle du moment de création**. Surtout, le test `If c Is Nothing` devient peu fiable : le simple fait d'évaluer la variable la recrée, si bien qu'elle n'est jamais perçue comme `Nothing`. La forme explicite évite ces surprises et rend le code prévisible : on crée l'objet quand on le décide, et l'on peut tester son existence sans effet de bord.

## Le cycle de vie : Initialize et Terminate

Deux événements jalonnent la vie d'un objet, et constituent ce que VBA offre de plus proche d'un constructeur et d'un destructeur.

`Class_Initialize` est appelé **à la création** de l'objet (au `New`). On y place l'initialisation : valeurs par défaut des champs, ouverture éventuelle de ressources. Une limite importante, déjà signalée dans l'introduction du chapitre : `Class_Initialize` **ne reçoit aucun argument**. VBA ne connaît pas de constructeur paramétré ; on ne peut donc pas écrire `New clsClient(42, "Dupont")`.

`Class_Terminate` est appelé **à la destruction** de l'objet. On y place le nettoyage : fermeture des recordsets, libération des ressources détenues par l'objet.

## Comptage de références et destruction

Quand un objet est-il détruit ? VBA gère les objets par **comptage de références**. Un objet vit tant qu'au moins une variable le référence ; il est détruit — et son `Class_Terminate` déclenché — lorsque sa **dernière référence** disparaît, soit parce qu'on l'a mise à `Nothing`, soit parce qu'elle est sortie de portée (cf. [section 3.4](/03-rappels-fondamentaux/04-portee-variables.md)).

Plusieurs variables peuvent référencer **le même objet**, ce qui illustre bien le mécanisme :

```vba
Public Sub DemoReferences()
    Dim c As clsClient
    Dim d As clsClient

    Set c = New clsClient        ' un seul objet est créé
    Set d = c                    ' d et c référencent LE MÊME objet (pas une copie)

    ' Modifier l'objet via d affecte aussi c : c'est le même objet.

    Set c = Nothing              ' il reste une référence (d) : l'objet N'EST PAS détruit
    Set d = Nothing              ' dernière référence libérée → Class_Terminate
End Sub
```

Ici, `Set d = c` ne duplique pas l'objet : il copie la **référence**. Les deux variables désignent la même instance, et l'objet ne sera détruit qu'après libération de la dernière des deux. Mettre explicitement les variables objet à `Nothing` lorsqu'on a fini de les utiliser est une bonne pratique, surtout lorsque l'objet détient des ressources, car elle rend la destruction **déterministe**.

## Initialiser un objet avec des valeurs

Puisque `Class_Initialize` n'accepte pas d'arguments, comment créer un objet déjà renseigné ? Trois approches s'offrent, complémentaires :

- **affecter les propriétés après création** : `Set c = New clsClient` puis `c.Id = 42`, `c.Nom = "Dupont"` (les propriétés sont l'objet de la section 16.2) ;
- **prévoir une méthode d'initialisation** explicite, par exemple `c.Init 42, "Dupont"`, appelée juste après le `New` ;
- **recourir à une fonction fabrique** qui crée et configure l'objet en une seule opération, puis le retourne — patron étudié à la [section 16.7](07-singleton-factory-observer.md).

Le choix dépend du contexte ; la fabrique est particulièrement utile lorsqu'un objet ne doit jamais exister dans un état non initialisé.

## Points de vigilance

- **`Set` est obligatoire** pour affecter un objet : les objets sont des types référence.
- **Préférer `Set x = New …` à `Dim x As New …`** : la forme auto-instanciée fait perdre le contrôle de la création et fausse le test `Is Nothing`.
- **`Class_Initialize` n'a pas d'argument** : initialiser après création, via une méthode `Init` ou une fabrique.
- **Une affectation `Set d = c` partage l'objet**, elle ne le copie pas : modifier l'un modifie l'autre.
- **L'objet n'est détruit qu'à sa dernière référence** : libérer toutes les variables qui le pointent.
- **Mettre les objets à `Nothing` après usage**, surtout s'ils détiennent des ressources, pour une destruction déterministe.

## En résumé

Un **module de classe** définit un type d'objet ; contrairement à un module standard, qui n'existe qu'en un exemplaire global, il s'**instancie** autant de fois que voulu, chaque objet possédant son propre état. On crée un module de classe par Insertion → Module de classe, on le nomme via la fenêtre Propriétés, puis on l'instancie en déclarant une variable de son type et en lui affectant `Set x = New MaClasse` — la forme explicite étant préférable à l'auto-instanciation `Dim x As New MaClasse`. La vie de l'objet est encadrée par `Class_Initialize` (à la création, sans argument) et `Class_Terminate` (à la destruction), VBA gérant la durée de vie par **comptage de références** : l'objet n'est détruit qu'une fois sa dernière référence libérée. Faute de constructeur paramétré, on initialise un objet après sa création, par affectation de propriétés, par une méthode `Init` ou par une fabrique.

La section suivante donne à ces objets une interface contrôlée : les [propriétés (Property Get, Let, Set)](02-proprietes-get-let-set.md).

⏭️ [16.2. Propriétés (Property Get, Let, Set)](/16-poo-vba-access/02-proprietes-get-let-set.md)
