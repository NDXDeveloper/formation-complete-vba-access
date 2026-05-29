🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15.4. Détection et gestion des conflits de verrouillage

La [section 15.3](03-strategies-verrouillage.md) a présenté le verrouillage du point de vue de celui qui pose le verrou. Cette section adopte le point de vue inverse : celui de l'utilisateur qui se heurte à un enregistrement **déjà verrouillé** par un autre. C'est le *conflit de verrouillage* — distinct du conflit de mise à jour traité au [chapitre 14](/14-transactions/05-conflits-mise-a-jour.md). On y voit comment le détecter, l'identifier, et le gérer proprement plutôt que de laisser le message générique d'Access surgir.

## Conflit de verrouillage ou conflit de mise à jour ?

Il est essentiel de ne pas confondre les deux situations, car elles appellent des réponses différentes.

Le **conflit de mise à jour** (section 14.5) survient en verrouillage optimiste : la donnée a *changé* depuis sa lecture, et la collision est détectée à l'enregistrement, sous la forme de l'erreur **3197**. Personne ne « bloque » l'enregistrement ; c'est son contenu qui a évolué.

Le **conflit de verrouillage** (cette section) survient lorsqu'un autre utilisateur *détient actuellement un verrou* sur l'enregistrement ou la page, et que l'on tente de l'éditer. On ne peut tout simplement pas commencer (ou terminer) la modification tant que le verrou n'est pas libéré. C'est typiquement le revers du verrouillage pessimiste.

Une nuance rassurante : tant qu'un utilisateur détient un verrou de page ou d'enregistrement, un autre peut **lire** ces données sans aucun conflit ; seule une tentative d'**édition** sur la même page provoque l'erreur. Lire n'est jamais bloqué ; seul écrire l'est.

## Quand survient un conflit de verrouillage

Plusieurs configurations le déclenchent :

- un autre utilisateur édite l'enregistrement (ou un voisin de la même page, en verrouillage de page) en mode pessimiste, et l'on tente de l'éditer à son tour ;
- une opération réclame un verrou de table déjà détenue par un autre processus ;
- la base, ou la table, est ouverte en mode **exclusif** par une autre session ;
- deux instances d'Access tournent sur le même poste et se disputent les mêmes données.

À noter que scinder l'application en front-end / back-end ([section 15.1](01-architecture-front-end-back-end.md)) réduit nettement la fréquence de ces conflits, en isolant l'essentiel de l'activité dans des front-ends séparés.

## Les codes d'erreur à intercepter

Le moteur signale ces situations par des erreurs interceptables. Les principales :

- **3260** — « Couldn't update; currently locked by user *nom* on machine *machine*. » C'est l'erreur la plus utile : elle indique qu'un autre utilisateur détient le verrou de la page visée, **et nomme cet utilisateur et son poste**.
- **3218** — « Couldn't update; currently locked. » Variante générique : l'enregistrement est verrouillé, sans identification.
- **3186** — « Couldn't save; currently locked by user *nom* on machine *machine*. » Équivalent au moment de la sauvegarde.
- **3188** — verrou détenu par une autre session **sur le même poste** (typiquement deux instances de l'application).
- **3211** — le moteur n'a pas pu verrouiller une table car elle est déjà utilisée par une autre personne ou un autre processus.
- **3008 / 3045** — la base ou la table est ouverte (ou ne peut être ouverte) en mode **exclusif** par une autre session.

Et, pour mémoire, à ne pas confondre : **3197** relève du conflit de *mise à jour* (donnée modifiée), traité en 14.5.

Une distinction stratégique s'impose au sein de cette liste. Un verrou de page ou d'enregistrement (3260, 3218, 3186, 3188) est en général **transitoire** : il sera libéré dès que l'autre utilisateur aura terminé son édition ; réessayer a donc du sens. À l'inverse, un conflit d'ouverture exclusive (3008, 3045) est **durable** tant que l'autre session reste ouverte exclusivement : inutile de réessayer en boucle, mieux vaut informer l'utilisateur immédiatement.

## Détecter et identifier

La détection passe par le numéro d'erreur, qu'on teste dans le gestionnaire. Pour l'erreur 3260, la propriété `Err.Description` contient le nom de l'utilisateur et de la machine qui détiennent le verrou — information qu'on peut afficher à l'utilisateur bloqué. Attention toutefois : extraire ces éléments en analysant la chaîne du message est fragile, car ce message est localisé et peut varier. Pour connaître de façon fiable qui est connecté, on s'appuiera plutôt sur le fichier de verrouillage ou une table de sessions, comme exposé à la [section 15.8](08-gestion-sessions-utilisateurs.md).

## Stratégies de gestion

Plutôt que de subir le message générique d'Access, on choisit une réponse adaptée. Quatre approches se combinent.

**1. Le réessai avec délai (*wait-and-retry*).** C'est la technique de référence pour un verrou transitoire : on intercepte l'erreur, on patiente un court instant, puis on retente l'opération un nombre limité de fois. Avant de réessayer, il est utile d'appeler `DBEngine.Idle dbRefreshCache`, qui laisse le moteur rafraîchir son cache et libérer les verrous de lecture. Après quelques tentatives infructueuses, on renonce et l'on informe l'utilisateur.

**2. Rafraîchir puis retenter.** Lorsque le conflit survient au `Edit`, il est souvent pertinent de rafraîchir la vue des données avec leur état courant avant de retenter le `Edit` une seconde fois.

**3. Informer l'utilisateur.** Pour un conflit durable (ouverture exclusive) ou après épuisement des tentatives, un message clair — idéalement nommant l'utilisateur bloquant grâce à l'erreur 3260 — vaut mieux qu'un réessai stérile.

**4. Réduire les conflits par conception.** La meilleure gestion reste la prévention : éditions courtes, verrouillage optimiste là où c'est pertinent (pour ne pas détenir de verrou pendant l'édition), transactions brèves (rappel : une transaction tient les verrous, cf. 15.3), et architecture front-end / back-end.

## Le patron de réessai

L'exemple suivant encapsule la logique de réessai pour l'obtention d'un verrou en édition. Il distingue le succès, le verrou transitoire (que l'on retente) et toute autre erreur (que l'on propage).

```vba
Public Function EditerAvecReessai(rs As DAO.Recordset) As Boolean
    Const MAX_TENTATIVES As Integer = 5
    Const DELAI_MS As Long = 400
    Dim tentative As Integer
    Dim n As Long, d As String

    For tentative = 1 To MAX_TENTATIVES
        On Error Resume Next
        Err.Clear
        rs.Edit
        Select Case Err.Number
            Case 0                              ' verrou obtenu
                On Error GoTo 0
                EditerAvecReessai = True
                Exit Function
            Case 3260, 3218, 3186, 3188         ' verrou transitoire : patienter et réessayer
                DBEngine.Idle dbRefreshCache
                Attendre DELAI_MS               ' pause (ex. via l'API Sleep, cf. 22.2)
            Case Else                           ' autre erreur : propager
                n = Err.Number: d = Err.Description
                On Error GoTo 0
                Err.Raise n, , d
        End Select
    Next tentative

    On Error GoTo 0
    EditerAvecReessai = False                    ' échec après plusieurs tentatives
End Function
```

La routine `Attendre` réalise une simple pause de quelques centaines de millisecondes — par exemple à l'aide de l'API `Sleep` (voir [section 22.2](/22-api-windows-integration-office/02-api-courantes.md)). On évitera un délai trop long, qui dégraderait l'expérience, comme un délai nul, qui saturerait le réseau de tentatives. Reconnaître la **famille** des erreurs de verrou se prête bien à une centralisation dans un module de gestion d'erreurs (cf. [section 13.7](/13-gestion-erreurs/07-module-centralise-erreurs.md)).

## Le cas des formulaires

Sur un formulaire lié, un conflit de verrouillage peut se produire alors qu'**aucun code VBA n'est en cours d'exécution** — par exemple lorsque l'utilisateur modifie directement un contrôle. Dans ce cas, `On Error` n'a aucune prise : c'est l'événement `Error` du formulaire qui reçoit l'erreur, et c'est là qu'il faut intervenir pour substituer un message ou un comportement personnalisé au dialogue standard. Ce mécanisme, lié à la propriété `RecordLocks`, est traité à la [section 15.6](06-recordlocks-formulaires.md).

## Points de vigilance

- **Verrou ≠ modification concurrente.** 3260/3218/3186/3188 signalent un verrou détenu ; 3197 signale une donnée modifiée (14.5). Réponses différentes.
- **Lire ne bloque pas.** Seule l'édition d'une donnée verrouillée déclenche un conflit ; la consultation reste possible.
- **Transitoire vs durable.** Réessayer un verrou d'enregistrement transitoire ; ne pas réessayer une ouverture exclusive — informer.
- **Ne pas analyser le message d'erreur** pour identifier l'utilisateur de façon durable : il est localisé. Préférer le suivi de sessions (15.8).
- **Borne le réessai.** Un nombre maximal de tentatives et un délai raisonnable évitent les boucles infinies et la saturation réseau.
- **Sur formulaire, c'est `Form_Error`** qui capte le conflit, pas `On Error`.

## En résumé

Un conflit de verrouillage survient lorsqu'un autre utilisateur détient un verrou sur l'enregistrement ou la page que l'on souhaite éditer — situation distincte du conflit de mise à jour (3197) où la donnée a simplement changé. La consultation d'une donnée verrouillée ne pose jamais problème ; seule son édition déclenche une erreur. Les codes à intercepter sont principalement **3260** (verrouillé par un utilisateur nommé), **3218** (verrouillé, générique), **3186** et **3188**, ainsi que **3211** (table en usage) et **3008/3045** (ouverture exclusive). On distingue les verrous transitoires, que l'on traite par un **réessai avec délai** appuyé sur `DBEngine.Idle dbRefreshCache`, des conflits durables d'ouverture exclusive, pour lesquels on informe directement l'utilisateur. Sur un formulaire lié, ces conflits se gèrent dans l'événement `Error`. Et la meilleure stratégie demeure préventive : éditions et transactions courtes, verrouillage optimiste et architecture front-end / back-end.

La section suivante traite d'un problème concret et récurrent du multi-utilisateur, où verrouillage et concurrence se conjuguent : la [génération de numéros séquentiels fiables](05-numeros-sequentiels-multi-utilisateurs.md).

⏭️ [15.5. Génération de numéros séquentiels fiables en multi-utilisateur (DMax+1, table de compteurs)](/15-multi-utilisateurs/05-numeros-sequentiels-multi-utilisateurs.md)
