🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 19.7. Débogage à distance et techniques sur poste utilisateur

Voici le cas le plus difficile : un défaut qui ne se manifeste **que** chez l'utilisateur, et que l'on ne reproduit pas en développement. Le « ça marche sur ma machine » est le symptôme classique — et c'est justement dans les différences entre les deux machines que se cache le problème. Sans accès à l'éditeur, le débogage interactif des sections précédentes ne s'applique plus ; il faut basculer d'une approche d'**observation** vers une approche d'**instrumentation** et de diagnostic à distance.

## Pourquoi le débogage interactif ne s'applique plus

Plusieurs obstacles se cumulent. Une application déployée est souvent un **ACCDE**, dont le code source est compilé et retiré : ni points d'arrêt, ni pas à pas (voir le [chapitre 20.3](../20-securite-protection/03-compilation-accde.md)). Sous **Access Runtime**, il n'y a tout simplement pas d'éditeur ([21.2](../21-deploiement-distribution/02-access-runtime.md)). Même sur une base ACCDB, on n'est pas physiquement présent pour surveiller la fenêtre Exécution immédiate, et la sortie de `Debug.Print` y est de toute façon invisible (voir la [section 19.2](02-debug-print-assert.md)). Surtout, le défaut dépend fréquemment de facteurs **propres au poste** : les données de l'utilisateur, sa version d'Office, son architecture 32 ou 64 bits, ses paramètres régionaux, ses droits, ou la séquence d'actions qu'il a suivie.

## Changer d'approche : instrumenter plutôt qu'observer

Puisqu'on ne peut pas observer en direct, l'application doit **enregistrer d'elle-même** de quoi reconstituer ce qui s'est passé. La fondation est donc le couple gestion d'erreurs robuste et journalisation : chaque procédure intercepte ses erreurs et les consigne avec leur contexte. Ce dispositif, présenté aux chapitres [13.7](../13-gestion-erreurs/07-module-centralise-erreurs.md) (module centralisé d'erreurs) et [19.5](05-journalisation-evenements.md) (journalisation), est ici l'outil principal — pas l'accessoire.

## Journaliser pour reconstituer

Une entrée de journal n'est utile que si elle contient de quoi diagnostiquer. Pour une erreur, on consigne son **numéro** et sa **description**, la **procédure** où elle s'est produite, les **valeurs clés** des variables en jeu, et l'**action** en cours — en plus du contexte standard (horodatage, utilisateur, poste, version). En jalonnant les opérations majeures de traces d'entrée et de sortie, le journal restitue en outre le **chemin** réellement parcouru, à défaut de pile d'appels.

```vba
GestionErreur:
    LogErreur "Err " & Err.Number & " (" & Err.Description & ")" & _
              " — étape : " & etapeCourante & _
              " — id : " & idEnCours, "EnregistrerCommande"
    Resume Suite
```

> ℹ️ La fonction `Erl` renvoie le numéro de la ligne où l'erreur est survenue, à condition que le code soit numéroté — une pratique que certains réservent aux versions de production pour localiser précisément les erreurs. Voir l'objet `Err` au [chapitre 13.4](../13-gestion-erreurs/04-objet-err.md).

## Capturer l'environnement

Comme le défaut tient souvent à l'environnement, une routine de **rapport de diagnostic** — déclenchable par l'utilisateur (un bouton « Générer un rapport de diagnostic ») — recueille les informations à transmettre au support : version d'Access, architecture, paramètres régionaux, utilisateur, poste, chemin du back-end, et l'état des **références**.

```vba
Public Sub GenererRapportDiagnostic()
    On Error Resume Next                      ' un rapport ne doit jamais échouer
    Dim chemin As String, f As Integer, ref As Object, bitness As String

    #If Win64 Then
        bitness = "64 bits"
    #Else
        bitness = "32 bits"
    #End If

    chemin = Environ$("LOCALAPPDATA") & "\MonAppli\diagnostic.txt"

    f = FreeFile
    Open chemin For Output As #f
    Print #f, "=== Rapport de diagnostic ==="
    Print #f, "Date         : " & Now
    Print #f, "Utilisateur  : " & Environ$("USERNAME")
    Print #f, "Poste        : " & Environ$("COMPUTERNAME")
    Print #f, "Version appli : " & VERSION_APPLI
    Print #f, "Access       : " & Application.Version
    Print #f, "Architecture : " & bitness
    Print #f, "Back-end     : " & CheminBackEnd()
    Print #f, ""
    Print #f, "--- Références ---"
    For Each ref In Application.References
        Dim fp As String
        fp = "(indisponible)"
        fp = ref.FullPath
        Print #f, ref.Name & " " & ref.Major & "." & ref.Minor & _
                  IIf(ref.IsBroken, "  *** ROMPUE ***", "") & "  " & fp
    Next ref
    Close #f
End Sub
```

Les **références rompues** sont une cause classique de défaut en production : une bibliothèque absente ou d'une version différente sur le poste de l'utilisateur provoque des erreurs de compilation ou des fonctions introuvables. Les énumérer en signalant `IsBroken` est souvent révélateur. De même, les **paramètres régionaux** (format des dates, séparateur décimal) expliquent quantité de comportements divergents d'un poste à l'autre.

> ℹ️ Les différences 32/64 bits sont traitées au [chapitre 21.7](../21-deploiement-distribution/07-32-bits-vs-64-bits.md), les références COM au [chapitre 2.5](../02-interface-environnement/05-references-bibliotheques-com.md), la localisation du SQL au [chapitre 11.7](../11-sql-access-vba/07-localisation-formatage-sql.md). Le rapport peut être transmis automatiquement par courriel via Outlook ([22.5](../22-api-windows-integration-office/05-automation-outlook.md)) ou déposé dans un dossier partagé.

## Activer la trace détaillée à distance

Le seuil de journalisation gagne à être **configurable** plutôt que figé dans une constante : en le lisant depuis la configuration (clé de registre, table de paramètres), le support peut le relever à `DEBUG` sur le poste concerné, sans redéployer l'application, pour capturer en détail l'opération défaillante — puis le rabaisser ensuite. On obtient ainsi des traces fines, ciblées, au moment voulu.

> ℹ️ La configuration peut s'appuyer sur le registre ([22.9](../22-api-windows-integration-office/09-registre-windows.md)) ou une table de paramètres ; les niveaux de journalisation sont décrits à la [section 19.5](05-journalisation-evenements.md).

## Reproduire les conditions

Le journal et le rapport de diagnostic permettent souvent de **rejouer le problème en développement**. On récupère une copie des données de l'utilisateur — anonymisée si elles sont sensibles (voir la [section 19.6](06-simulation-donnees-test.md)) —, car la donnée est fréquemment la différence déterminante, et l'on reproduit autant que possible son environnement (version, architecture, paramètres régionaux). Journal, données et contexte réunis suffisent généralement à reproduire, donc à corriger.

## Déboguer sur la machine réelle

Lorsqu'on peut accéder au poste — en personne ou via une prise de contrôle à distance —, deux possibilités s'ouvrent. On observe l'utilisateur reproduire le défaut en partageant son écran. Ou bien l'on déploie temporairement une version **ACCDB** (avec code source) sur la machine, afin d'y retrouver l'éditeur et ses outils interactifs : conserver une telle version « de débogage » à côté de l'ACCDE livré facilite grandement ces interventions.

> ℹ️ Voir la mise à jour du front-end ([21.3](../21-deploiement-distribution/03-mise-a-jour-front-end.md)) et la compatibilité entre versions d'Access ([21.6](../21-deploiement-distribution/06-compatibilite-versions-access.md)).

## `MsgBox` en dernier recours

Quand aucun autre moyen n'est disponible sur une machine qu'on ne peut pas instrumenter, une `MsgBox` temporaire placée à un point clé permet de faire apparaître une valeur à l'écran, que l'utilisateur lit ou capture. C'est rudimentaire, intrusif et bloquant — et à retirer impérativement ensuite —, mais c'est parfois le chemin le plus court pour confirmer une valeur. On le réserve au dépannage ponctuel, jamais à une instrumentation durable, pour laquelle la journalisation reste la bonne réponse.

## Points clés à retenir

- Un défaut propre au poste de l'utilisateur ne se débogue pas en interactif : pas d'éditeur sur un ACCDE ou un Runtime, et l'on n'est pas présent. La cause tient souvent aux différences d'environnement et de données.
- L'approche bascule de l'observation à l'instrumentation : l'application doit enregistrer de quoi reconstituer l'incident, via une gestion d'erreurs robuste et une journalisation riche en contexte.
- Une routine de rapport de diagnostic capture version, architecture, paramètres régionaux, chemin du back-end et état des références — les références rompues et les réglages locaux étant des coupables fréquents.
- Un seuil de journalisation configurable permet de relever la verbosité à distance, le temps de capturer l'opération fautive.
- Le journal, une copie (anonymisée) des données et le contexte permettent de reproduire le problème en développement.
- Sur la machine réelle, la prise de contrôle à distance ou le déploiement d'une version ACCDB de débogage rendent les outils interactifs à nouveau disponibles ; la `MsgBox` n'est qu'un dépannage de dernier recours.

---


⏭️ [20. Sécurité et protection](/20-securite-protection/README.md)
