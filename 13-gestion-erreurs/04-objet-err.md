🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13.4. Objet Err — numéro, description, source

## Introduction

Les deux sections précédentes ont montré comment intercepter une erreur ; encore faut-il savoir **de quelle erreur il s'agit**. Cette information est centralisée dans un objet unique, toujours disponible, que VBA met à jour à chaque incident : l'objet `Err`. C'est lui que l'on interroge dans un gestionnaire pour connaître le numéro de l'erreur, son libellé et sa provenance ; c'est lui que l'on manipule pour générer ses propres erreurs ou pour remettre l'état d'erreur à zéro.

Bien connaître l'objet `Err` est indispensable à toute gestion d'erreurs sérieuse. Au-delà de ses propriétés, il faut comprendre son **cycle de vie** — quand il est renseigné, quand il est effacé — car une méconnaissance de ce cycle conduit à des diagnostics erronés ou à la perte pure et simple de l'information sur l'erreur. Cette section détaille les propriétés de `Err`, ses méthodes, et les pratiques qui en garantissent une exploitation fiable.

## Qu'est-ce que l'objet Err ?

`Err` est un objet **intrinsèque et global** : il existe en permanence, sans qu'il soit nécessaire de le déclarer ou de l'instancier, et il est accessible depuis n'importe quel point du code. Il conserve les informations relatives à la **dernière erreur d'exécution survenue**, et à elle seule.

Ce dernier point est fondamental : il n'existe qu'un seul objet `Err` pour l'ensemble de la session VBA. Il ne mémorise pas un historique des erreurs, mais uniquement la plus récente. Chaque nouvelle erreur écrase les informations de la précédente. C'est la raison pour laquelle, comme on le verra, il faut souvent capturer ses valeurs sans délai, avant qu'une autre opération ne vienne les remplacer.

Lorsqu'aucune erreur n'est en cours, la propriété `Err.Number` vaut 0. Tester `If Err.Number <> 0` est donc la manière canonique de savoir si une erreur s'est produite.

## Les propriétés principales

L'objet `Err` expose plusieurs propriétés. Trois d'entre elles concentrent l'essentiel de l'usage courant.

| Propriété | Type | Contenu |
|---|---|---|
| `Number` | Long | Numéro de l'erreur ; vaut 0 lorsqu'aucune erreur n'est en cours |
| `Description` | String | Texte descriptif de l'erreur |
| `Source` | String | Nom du composant (projet, application, bibliothèque) à l'origine de l'erreur |

### Number

`Err.Number` est la propriété la plus importante. Elle renvoie le numéro de l'erreur — un entier `Long` — qui constitue, comme l'a établi la section [13.1](/13-gestion-erreurs/01-erreurs-courantes-access.md), l'**identifiant fiable et indépendant de la langue** de l'erreur. C'est sur ce numéro que repose tout traitement conditionnel : c'est lui que l'on examine dans un `Select Case` pour réagir différemment selon la nature de l'incident.

```vba
If Err.Number = 3022 Then
    MsgBox "Cet enregistrement existe déjà."
End If
```

On notera qu'`Err` employé seul, sans propriété, renvoie également le numéro — un héritage des anciennes versions de Basic. Cette forme abrégée est à éviter : `Err.Number` est explicite et sans ambiguïté.

### Description

`Err.Description` renvoie le texte décrivant l'erreur, tel qu'il serait affiché dans la boîte de dialogue système. C'est ce libellé que l'on présente généralement à l'utilisateur ou que l'on consigne dans un journal.

Sa valeur diagnostique reste cependant limitée : le texte dépend de la langue d'installation et peut être vague. Pour les erreurs issues du moteur de données (DAO, ADO), il est même souvent générique, le détail réel étant rangé dans des collections spécifiques étudiées à la section [13.5](/13-gestion-erreurs/05-erreurs-dao-ado.md). La description complète utilement le numéro, mais ne le remplace jamais comme base de raisonnement.

### Source

`Err.Source` indique **quel composant a généré l'erreur**. Sa portée varie selon l'origine de l'incident. Pour une erreur d'exécution VBA ordinaire au sein du projet, `Source` contient généralement le nom du projet VBA et apporte peu d'information. En revanche, elle prend tout son sens dans deux situations : lors d'une opération d'**Automation**, où elle identifie l'application externe responsable (par exemple le nom de l'application Excel ou Word), et pour les **erreurs personnalisées**, où le développeur renseigne lui-même cette propriété afin d'indiquer la procédure ou le module d'origine (voir section [13.8](/13-gestion-erreurs/08-erreurs-personnalisees-raise.md)).

## Les propriétés complémentaires

D'autres propriétés, d'usage plus spécialisé, complètent l'objet `Err`.

### LastDLLError

`Err.LastDLLError` est destinée aux appels d'**API Windows** réalisés via l'instruction `Declare`. Lorsqu'une fonction d'une DLL échoue, elle ne déclenche pas d'erreur VBA : elle renvoie un code de retour qu'il appartient au développeur de tester. Le code d'erreur système correspondant (celui que retournerait la fonction Windows `GetLastError`) est alors disponible dans `Err.LastDLLError`. Cette propriété est centrale dans le chapitre consacré à l'intégration avec l'API Windows (chapitre 22) ; elle est mentionnée ici pour mémoire, car elle fait partie intégrante de l'objet `Err`.

### HelpFile et HelpContext

`Err.HelpFile` et `Err.HelpContext` permettent d'associer à une erreur une rubrique d'aide contextuelle : le chemin d'un fichier d'aide et l'identifiant de la rubrique à y ouvrir. Ces propriétés, héritées d'une époque où les applications étaient livrées avec des fichiers d'aide compilés, sont aujourd'hui rarement employées. Elles restent disponibles, notamment dans le cadre de `Err.Raise`, mais ne concernent que des contextes très spécifiques.

## Le cycle de vie de l'objet Err

Comprendre **quand** l'objet `Err` est renseigné et **quand** il est remis à zéro est la clé d'une exploitation correcte. Une erreur de raisonnement sur ce cycle est la cause la plus fréquente des diagnostics faussés.

L'objet `Err` est **renseigné** dès qu'une erreur d'exécution survient : ses propriétés reflètent alors cette erreur.

Il est **réinitialisé** (sa propriété `Number` repasse à 0) automatiquement dans les cas suivants :

- l'exécution de **toute instruction `On Error`** — `On Error GoTo`, `On Error Resume Next`, `On Error GoTo 0` et `On Error GoTo -1` ;
- l'exécution de **toute instruction `Resume`**, sous l'une quelconque de ses formes ;
- une **sortie de procédure** (`Exit Sub`, `Exit Function`, `Exit Property`) ou la fin naturelle de la procédure ;
- un appel explicite à **`Err.Clear`**.

Enfin, il est **écrasé** dès qu'une nouvelle erreur se produit, l'ancienne information étant alors définitivement perdue.

Un point mérite d'être souligné, car il prête à confusion : un simple **appel de procédure ne réinitialise pas `Err` par lui-même**. En revanche, la procédure appelée le fera presque toujours, dès lors qu'elle contient sa propre instruction `On Error` — ce qui est le cas de toute procédure correctement protégée. En pratique, il faut donc considérer que tout appel à une autre routine de l'application risque d'effacer l'objet `Err`. Cette réalité conduit directement à la bonne pratique exposée ci-dessous.

## Capturer les valeurs de Err dès l'entrée du gestionnaire

De ce qui précède découle une règle professionnelle essentielle : **dans un gestionnaire d'erreurs, capturer les propriétés de `Err` dans des variables locales dès la première instruction**, avant tout appel ou toute opération susceptible de les réinitialiser.

```vba
GestionErreur:
    Dim lngNumero As Long
    Dim strDescription As String
    Dim strSource As String

    ' Capture immédiate, avant toute opération réinitialisant Err
    lngNumero = Err.Number
    strDescription = Err.Description
    strSource = Err.Source

    ' À partir d'ici, on peut appeler d'autres procédures sans risque :
    ' même si JournaliserErreur réinitialise Err, on conserve nos copies
    JournaliserErreur lngNumero, strDescription, strSource
    MsgBox "Erreur " & lngNumero & " : " & strDescription, vbCritical

    Resume SortieProcedure
End Sub
```

Sans cette capture, l'appel à `JournaliserErreur` — qui contient assurément sa propre instruction `On Error` — remettrait `Err.Number` à 0. Toute référence ultérieure à `Err.Number` ou `Err.Description` au sein du gestionnaire renverrait alors des valeurs vides, et l'information sur l'erreur réelle serait perdue. Cette précaution, anodine en apparence, devient indispensable dès que le gestionnaire fait appel à du code externe — ce qui sera systématiquement le cas avec une gestion centralisée (section [13.7](/13-gestion-erreurs/07-module-centralise-erreurs.md)).

## Err.Clear : réinitialiser explicitement

La méthode `Err.Clear` remet l'objet `Err` à son état initial : `Number` repasse à 0, et les autres propriétés sont vidées. Bien que de nombreuses instructions (`Resume`, `On Error`, `Exit Sub`...) effacent `Err` implicitement, `Err.Clear` est utile lorsqu'on souhaite réinitialiser l'objet **sans** modifier la gestion d'erreurs en cours ni reprendre l'exécution.

Son usage le plus courant intervient avec `On Error Resume Next`, lorsque l'on enchaîne plusieurs vérifications : entre deux tests de `Err.Number`, `Err.Clear` garantit qu'une erreur antérieure non effacée ne provoquera pas de faux positif lors de la vérification suivante.

```vba
On Error Resume Next

operation1
If Err.Number <> 0 Then GererProbleme1
Err.Clear                          ' on repart d'un état propre

operation2
If Err.Number <> 0 Then GererProbleme2

On Error GoTo 0
```

## Err.Raise : un aperçu

L'objet `Err` ne sert pas seulement à consulter les erreurs : il permet aussi d'en **déclencher**. La méthode `Err.Raise` génère une erreur d'exécution, exactement comme si VBA l'avait rencontrée, ce qui permet de signaler une situation anormale propre à la logique métier de l'application.

```vba
Err.Raise vbObjectError + 1000, "ModuleClient", "Le client est inactif."
```

On remarque ici la constante `vbObjectError`, qui sert de base aux numéros d'erreurs personnalisés afin de les distinguer des codes système. Le mécanisme complet de `Err.Raise` — choix des numéros, renseignement de la source, propagation — fait l'objet de la section [13.8](/13-gestion-erreurs/08-erreurs-personnalisees-raise.md). Il suffit ici de retenir que la création d'erreurs sur mesure passe par cette méthode de l'objet `Err`.

## Obtenir le texte d'une erreur en dehors d'un gestionnaire

Il est parfois utile de connaître le libellé associé à un numéro d'erreur sans qu'une erreur ne se soit réellement produite — par exemple pour constituer une table de correspondance ou enrichir un journal. Deux fonctions le permettent.

La fonction `Error$` (ou `Error`) renvoie le texte standard correspondant à un numéro d'erreur VBA :

```vba
Debug.Print Error$(11)        ' affiche « Division par zéro »
```

Pour les erreurs **spécifiques à Access** (la plage 2xxx), la méthode `Application.AccessError` est plus fiable, car elle connaît les libellés propres à Access que `Error$` ignore :

```vba
Debug.Print Application.AccessError(2501)   ' libellé de l'erreur Access 2501
```

Ces fonctions sont complémentaires de `Err.Description` : cette dernière renseigne sur l'erreur *en cours*, tandis qu'`Error$` et `AccessError` permettent d'obtenir le libellé d'un numéro *arbitraire*, à tout moment.

## Erl n'est pas une propriété de Err

Une confusion fréquente mérite d'être levée : la fonction `Erl`, qui renvoie le numéro de la ligne où une erreur s'est produite (voir section [13.2](/13-gestion-erreurs/02-on-error-goto.md)), **n'est pas une propriété de l'objet `Err`**. C'est une fonction indépendante, héritée du Basic historique, dépourvue de lien syntaxique avec `Err`. On écrit donc `Erl` et non `Err.Line` ou `Err.Erl`. Cette distinction, sans conséquence pratique majeure, évite des erreurs de syntaxe lors de l'écriture des gestionnaires.

## En résumé

L'objet `Err` est un objet global et intrinsèque qui conserve les informations sur la dernière erreur d'exécution survenue — et sur elle seule. Ses trois propriétés essentielles sont `Number` (l'identifiant fiable, base de tout traitement conditionnel), `Description` (le libellé, utile mais secondaire) et `Source` (le composant d'origine, surtout pertinent en Automation et pour les erreurs personnalisées). Des propriétés complémentaires, comme `LastDLLError` pour les appels d'API, complètent l'objet.

La maîtrise de son **cycle de vie** est déterminante : `Err` est renseigné à la survenue d'une erreur, et réinitialisé par toute instruction `On Error`, tout `Resume`, toute sortie de procédure ou un appel à `Err.Clear` — un appel de procédure protégée l'effaçant lui aussi de fait. De là découle la règle d'or des gestionnaires : **capturer les valeurs de `Err` dans des variables locales dès la première instruction du gestionnaire**, avant tout appel externe. La méthode `Err.Clear` permet une réinitialisation explicite, et `Err.Raise` ouvre la voie aux erreurs personnalisées. Enfin, `Error$` et `Application.AccessError` fournissent les libellés à partir d'un numéro arbitraire, et la fonction `Erl`, bien que liée à la gestion d'erreurs, demeure indépendante de l'objet `Err`.

La section suivante quitte le cadre des erreurs VBA génériques pour aborder les erreurs propres au moteur de données : l'interception spécifique des erreurs DAO et ADO, dont le détail réside justement au-delà de l'objet `Err`.

---


⏭️ [13.5. Erreurs DAO et ADO — interception spécifique](/13-gestion-erreurs/05-erreurs-dao-ado.md)
