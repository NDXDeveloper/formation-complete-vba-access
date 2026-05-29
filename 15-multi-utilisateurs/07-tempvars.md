🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15.7. TempVars — variables globales de session

Une application partagée a besoin de conserver, tout au long d'une session, un certain nombre d'informations propres à l'utilisateur en cours : son identité après authentification, son rôle, le chemin du back-end, quelques paramètres chargés au démarrage. Les **TempVars** offrent un moyen propre et fiable de stocker ces valeurs globalement. Introduites avec Access 2007, elles présentent par rapport aux variables globales VBA classiques deux atouts décisifs, ainsi que des limites précises qu'il faut connaître. Cette section les détaille.

## Qu'est-ce qu'une TempVar ?

La collection **`TempVars`** (variables temporaires) regroupe des objets **`TempVar`**, chacun défini par un nom et une valeur. On y stocke des valeurs nommées accessibles depuis tout le code de l'application. Par défaut, un objet `TempVar` reste en mémoire **jusqu'à la fermeture de la base** ; il survit donc au passage d'un formulaire à l'autre, d'un module à l'autre. Les TempVars offrent un moyen commode d'**échanger des données entre procédures VBA et macros** — un point sur lequel nous reviendrons.

## « Globales de session », et non partagées entre utilisateurs

Il faut lever d'emblée une ambiguïté du terme « globales ». Une TempVar est globale **au sein d'une seule session, d'une seule instance** d'Access : elle vit dans la mémoire du front-end de l'utilisateur courant. Elle n'est **pas partagée** entre utilisateurs, ni entre instances. Deux utilisateurs travaillant sur leur propre copie du front-end ont chacun leur propre jeu de TempVars, indépendantes.

Autrement dit, les TempVars servent à mémoriser l'état **d'une session utilisateur**, pas à faire communiquer les utilisateurs entre eux. Pour un état réellement partagé (savoir qui est connecté, par exemple), il faut une **table** — c'est l'objet de la [section 15.8](08-gestion-sessions-utilisateurs.md). Cette distinction est essentielle dans le contexte multi-utilisateur.

## Pourquoi des TempVars plutôt que des variables globales VBA ?

Deux raisons justifient leur usage.

**Elles survivent à une réinitialisation du code.** C'est l'argument principal. Lorsqu'une **erreur non gérée** survient en VBA, l'état du projet est réinitialisé et les variables globales (`Public`) perdent leur valeur. Les TempVars, elles, **conservent leur valeur quoi qu'il arrive** : elles sont dépendables là où une variable globale ne l'est pas. Le mot important est « non gérée » : avec une gestion d'erreurs rigoureuse (chapitre 13), les variables globales survivent aussi. À noter également une nuance : ce risque de réinitialisation concerne les versions **non compilées** (`.accdb`) ; dans une version **compilée** (`.accde`), les variables globales ne sont plus réinitialisées par une erreur non gérée. Les TempVars restent néanmoins le filet de sécurité le plus simple.

**Elles sont accessibles depuis les macros, les requêtes et les sources de contrôle.** Une variable globale VBA ne l'est pas directement : il faudrait l'envelopper dans une fonction publique. Une TempVar se référence partout — dans une requête, une expression, une propriété `ControlSource` — via la syntaxe `TempVars!NomVariable` (ou `[TempVars]![NomVariable]`). C'est un avantage considérable pour une application mêlant code et macros.

En contrepartie, elles offrent **peu de sûreté de typage** : une TempVar est de type `Variant`, sans contrôle de type ni assistance à la saisie (IntelliSense). Pour pallier cela, une bonne pratique consiste à les encapsuler derrière des propriétés d'un module de classe, qui apportent typage et lisibilité.

## Manipuler les TempVars

La création et la mise à jour se font par affectation directe (qui crée la variable si elle n'existe pas) ou par la méthode `Add` ; côté macro, par l'action `SetTempVar`.

```vba
' Créer ou mettre à jour
TempVars!UtilisateurID = 42
TempVars!Role = "Gestionnaire"
TempVars.Add "CheminBackEnd", "\\Serveur\Partage\MaBase_be.accdb"   ' forme explicite

' Lire
Dim idUtil As Long
idUtil = Nz(TempVars!UtilisateurID, 0)

' Tester l'existence : une TempVar inexistante vaut Null
If IsNull(TempVars!Role) Then
    MsgBox "Aucun rôle défini : utilisateur non authentifié."
End If

' Supprimer
TempVars.Remove "Role"        ' une seule (action macro : RemoveTempVar)
TempVars.RemoveAll            ' toutes (action macro : RemoveAllTempVars)
```

On peut référencer un objet de la collection par son nom ou par son numéro d'ordre. La méthode `Remove` (ou l'action `RemoveTempVar`) supprime une variable, `RemoveAll` (ou `RemoveAllTempVars`) les supprime toutes — typiquement à la déconnexion. Dans une requête ou une expression, on écrit simplement `[TempVars]![NomVariable]`.

## Les limites à connaître

Trois contraintes encadrent les TempVars.

**Un maximum de 255.** La collection peut contenir jusqu'à 255 objets `TempVar`. C'est largement suffisant en pratique, mais la limite existe.

**Données textuelles ou numériques uniquement — pas d'objets.** Une `TempVar` ne peut stocker que des données simples (texte, nombre, date, booléen). Elle **ne peut pas stocker d'objet** — ni recordset, ni contrôle, ni formulaire, ni tableau, ni collection ou dictionnaire. Toute tentative déclenche l'**erreur d'exécution 32538 : « TempVars can only store data. They cannot store objects. »**

Cette limite est à l'origine d'un piège classique lors de la **conversion d'une macro en VBA**. En macro, l'action de définition ne manipule que des valeurs ; mais en VBA, écrire `TempVars.Add "x", Me.txtNom` transmet l'objet contrôle entier et provoque l'erreur 32538. Il faut alors stocker explicitement la **valeur scalaire** :

```vba
TempVars.Add "Ctrl", Me.txtNom         ' ERREUR 32538 : on passe l'objet contrôle
TempVars.Add "Nom", Me.txtNom.Value    ' CORRECT : on passe la valeur
```

**Nettoyer en fin d'usage.** Une TempVar non supprimée reste en mémoire jusqu'à la fermeture de la base ; il est recommandé de supprimer celles dont on n'a plus besoin, et de vider la collection à la déconnexion.

## Usages typiques en multi-utilisateur

Dans une application partagée, les TempVars hébergent naturellement l'état de la session de l'utilisateur courant :

- l'**identité et le rôle** de l'utilisateur après authentification, réutilisés ensuite pour filtrer les données, tracer les modifications ou appliquer des droits (voir [section 20.6](/20-securite-protection/06-gestion-droits-utilisateurs.md)) ;
- le **chemin du back-end** ou d'autres informations de connexion ;
- des **paramètres applicatifs** chargés une fois au démarrage ;
- le passage de valeurs **entre formulaires** sans couplage fort, en complément de `OpenArgs` ([section 6.6](/06-formulaires/06-openargs.md)) et des techniques de communication entre formulaires ([section 6.9](/06-formulaires/09-communication-formulaires.md)).

## TempVars et autres mécanismes d'état

Le choix du bon support dépend de la portée souhaitée. Les **variables globales VBA** peuvent stocker des objets et sont légèrement plus rapides, mais sont fragiles (réinitialisation) et invisibles aux macros et requêtes. Une **table de paramètres** persiste au-delà de la session et peut être partagée entre utilisateurs ; c'est le bon choix pour la configuration durable. Les **propriétés personnalisées de la base** ([section 12.8](/12-querydefs-tabledefs/08-proprietes-personnalisees.md)) conviennent aussi à une configuration applicative persistante. Les **TempVars**, enfin, occupent une niche précise : l'état **transitoire, propre à la session courante**, dépendable et accessible partout.

## Points de vigilance

- **Globales de session, pas partagées.** Chaque instance a ses propres TempVars ; pour un état partagé entre utilisateurs, utiliser une table (15.8).
- **Pas d'objets** : seulement texte/nombre/date. Stocker un objet lève l'erreur 32538 ; passer `.Value` lors d'une conversion macro → VBA.
- **Type Variant, sans contrôle de type** : envisager une encapsulation par classe pour la robustesse.
- **Maximum 255** TempVars.
- **Survie aux réinitialisations** surtout pertinente en `.accdb` ; en `.accde`, les variables globales survivent aussi aux erreurs non gérées.
- **Nettoyer** les TempVars inutiles et vider la collection à la déconnexion.

## En résumé

Les **TempVars** sont des variables nommées de type `Variant`, regroupées dans une collection limitée à **255** objets, qui restent en mémoire jusqu'à la fermeture de la base. Leur intérêt majeur tient à deux propriétés : elles **survivent à une réinitialisation du code** provoquée par une erreur non gérée (là où une variable globale VBA perd sa valeur, du moins en version non compilée), et elles sont **accessibles depuis les macros, les requêtes et les sources de contrôle** via `TempVars!Nom`. Leurs limites : elles ne stockent que des données simples — **jamais d'objets**, sous peine de l'erreur 32538 — et constituent un état **propre à chaque session**, jamais partagé entre utilisateurs. Elles sont donc l'outil idéal pour mémoriser l'identité, le rôle et les paramètres de la session courante, en complément d'une table pour tout ce qui doit être partagé ou persistant.

La section suivante traite justement de cet état partagé entre utilisateurs : la [gestion des sessions utilisateurs (LDB et table de sessions)](08-gestion-sessions-utilisateurs.md).

⏭️ [15.8. Gestion des sessions utilisateurs (LDB et table de sessions)](/15-multi-utilisateurs/08-gestion-sessions-utilisateurs.md)
