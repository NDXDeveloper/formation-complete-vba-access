🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 19.5. Journalisation d'événements applicatifs

La fenêtre Exécution immédiate est éphémère et réservée au développement (voir la [section 19.2](02-debug-print-assert.md)). En production, on ne peut pas surveiller l'éditeur : pour comprendre un problème qui ne survient que chez l'utilisateur, il faut une trace **persistante** de ce qu'a fait l'application. La journalisation enregistre ce qui s'est passé, à quel moment, dans quel ordre et dans quel contexte — de quoi reconstituer la séquence d'événements qui a mené à un incident. C'est le pendant, durable et exploitable en production, de la trace de débogage.

## Pourquoi journaliser

Un journal répond à des besoins que le débogage interactif ne couvre pas. Il fournit un **enregistrement durable**, consultable après coup, indispensable pour diagnostiquer un défaut sur le poste d'un utilisateur (voir la [section 19.7](07-debogage-distance.md)). Il permet de **reconstituer l'enchaînement** des opérations, là où une simple capture d'erreur ne montre que le point d'arrivée. Il facilite le **support** (« envoyez-moi votre journal »). Et il dépasse la seule gestion des erreurs : on y consigne aussi le déroulement normal de l'application, ce qui aide à comprendre son usage et à repérer les anomalies récurrentes.

## Quoi journaliser — et à quel niveau

### Les événements à tracer

On journalise les jalons utiles : le **cycle de vie** (démarrage, fermeture, connexion, version), les **opérations significatives** (import ou export lancé puis terminé, traitement par lot de N enregistrements, génération d'état, compactage), les **avertissements** (anomalies récupérées, comportements inhabituels mais gérés) et les **erreurs** interceptées. Chaque entrée porte son contexte : horodatage, utilisateur, poste, et la procédure ou le module concerné.

À l'inverse, on ne journalise **jamais** de données sensibles — mots de passe, informations personnelles — pour des raisons de confidentialité et de sécurité.

### Les niveaux de gravité

Une bonne journalisation est **hiérarchisée** par niveau : `DEBUG` (trace fine), `INFO` (déroulement normal), `WARN` (avertissement), `ERROR` (erreur). Un **seuil minimal** configurable contrôle la verbosité : en développement, on journalise à partir de `DEBUG` ; en production, à partir de `INFO`, ce qui écarte d'emblée le bruit et son coût. Ce filtrage par niveau est le moyen le plus simple de garder une journalisation détaillée hors de la production.

## Où journaliser : choisir la destination

Deux destinations principales s'offrent à nous, aux propriétés opposées.

Une **table** est facile à interroger, trier et filtrer, et — si elle réside dans le back-end — centralise les événements de tous les utilisateurs. En contrepartie, écrire dans la base ajoute de la charge, crée de la contention sur une table partagée, et devient inutilisable précisément quand c'est la base qui défaille. La journalisation des erreurs en table est traitée au [chapitre 13.6](../13-gestion-erreurs/06-journalisation-erreurs-table.md).

Un **fichier texte local**, par utilisateur, est au contraire **robuste** : il fonctionne même si la base est indisponible ou corrompue, n'engendre aucune contention, et se transmet aisément pour le support. C'est le choix le plus fiable pour diagnostiquer en production — notamment les problèmes de démarrage ou ceux qui mettent en cause la base elle-même.

En pratique, on combine souvent les deux : un fichier local toujours actif (fiabilité), complété d'une table pour les événements que l'on souhaite agréger. (Le journal d'événements Windows constitue une troisième voie, possible via API, mais généralement disproportionnée pour une application Access.)

> ℹ️ Le `FileSystemObject` est traité au [chapitre 22.10](../22-api-windows-integration-office/10-filesystemobject.md).

## Un module de journalisation

On centralise la journalisation dans un module unique. L'exemple suivant écrit dans un fichier local, filtre par niveau, et — principe cardinal — **ne plante jamais l'application**, même en cas d'échec d'écriture.

```vba
' ============================================================
'  modLog — journalisation d'événements applicatifs
' ============================================================
Option Compare Database
Option Explicit

Public Enum NiveauLog
    nlvDebug = 0
    nlvInfo = 1
    nlvAvert = 2
    nlvErreur = 3
End Enum

Private Const NIVEAU_MINI As Long = nlvInfo     ' seuil : on ignore en dessous

Public Sub Journaliser(ByVal niveau As NiveauLog, ByVal message As String, _
                       Optional ByVal source As String = "")
    If niveau < NIVEAU_MINI Then Exit Sub        ' filtrage avant toute écriture

    On Error Resume Next                          ' la journalisation ne doit jamais échouer bruyamment
    Dim ligne As String
    ligne = Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & _
            Etiquette(niveau) & vbTab & _
            Environ$("USERNAME") & vbTab & _
            source & vbTab & message
    EcrireFichier ligne
End Sub

' Raccourcis de confort
Public Sub LogInfo(ByVal msg As String, Optional ByVal src As String = "")
    Journaliser nlvInfo, msg, src
End Sub
Public Sub LogAvert(ByVal msg As String, Optional ByVal src As String = "")
    Journaliser nlvAvert, msg, src
End Sub
Public Sub LogErreur(ByVal msg As String, Optional ByVal src As String = "")
    Journaliser nlvErreur, msg, src
End Sub

' ---------- Mécanique interne ----------

Private Sub EcrireFichier(ByVal ligne As String)
    Dim chemin As String, f As Integer
    chemin = CheminLog()
    AssurerDossier chemin
    RoterSiVolumineux chemin
    f = FreeFile
    Open chemin For Append As #f          ' crée le fichier s'il n'existe pas
    Print #f, ligne
    Close #f
End Sub

Private Function CheminLog() As String
    ' Fichier local par utilisateur : robuste et sans contention
    CheminLog = Environ$("LOCALAPPDATA") & "\MonAppli\journal.log"
End Function

Private Sub AssurerDossier(ByVal cheminFichier As String)
    Dim fso As Object, dossier As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    dossier = fso.GetParentFolderName(cheminFichier)
    If Not fso.FolderExists(dossier) Then fso.CreateFolder dossier
End Sub

Private Sub RoterSiVolumineux(ByVal chemin As String)
    Const TAILLE_MAX As Long = 2& * 1024 * 1024     ' 2 Mo
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(chemin) Then
        If fso.GetFile(chemin).Size > TAILLE_MAX Then
            If fso.FileExists(chemin & ".1") Then fso.DeleteFile chemin & ".1"
            fso.MoveFile chemin, chemin & ".1"       ' conserve le journal précédent
        End If
    End If
End Sub

Private Function Etiquette(ByVal n As NiveauLog) As String
    Select Case n
        Case nlvDebug:  Etiquette = "DEBUG"
        Case nlvInfo:   Etiquette = "INFO"
        Case nlvAvert:  Etiquette = "WARN"
        Case nlvErreur: Etiquette = "ERROR"
    End Select
End Function
```

L'usage est alors limpide, et les entrées séparées par tabulation se réimportent sans peine dans un tableur ou une table pour analyse.

```vba
LogInfo "Démarrage v" & VERSION_APPLI
LogInfo "Import terminé : " & n & " ligne(s)", "ImporterVentes"
LogAvert "Modèle introuvable, valeur par défaut appliquée", "Config"
```

## Intégrer à la gestion des erreurs

Le journal et la gestion des erreurs travaillent de concert : un gestionnaire d'erreurs centralisé consigne chaque erreur interceptée, avec son numéro, sa description et sa source.

```vba
GestionErreur:
    LogErreur Err.Number & " - " & Err.Description, "EnregistrerCommande"
    Resume Suite
```

> ℹ️ Voir l'objet `Err` ([13.4](../13-gestion-erreurs/04-objet-err.md)) et le module de gestion centralisée des erreurs ([13.7](../13-gestion-erreurs/07-module-centralise-erreurs.md)).

## Maîtriser la taille : rotation et purge

Un journal croît indéfiniment si on n'y prend garde. Pour un fichier, on **rote** : on le renomme lorsqu'il dépasse une taille (le module ci-dessus en conserve une version précédente), et l'on peut garder plusieurs anciens fichiers. Pour une table, on **purge** périodiquement les entrées anciennes — d'autant qu'une table de journal très active contribue au gonflement de la base (voir les sections [18.7](../18-optimisation-performance/07-compactage-automatique.md) et [18.9](../18-optimisation-performance/09-limites-taille-2go.md)).

## Performance et fiabilité

La journalisation a un coût (écriture fichier ou base). Deux principes le maîtrisent. D'abord, ne pas journaliser **dans une boucle serrée** au niveau `INFO` : on réserve le détail au niveau `DEBUG` (filtré en production) et l'on journalise plutôt des agrégats (« 1 000 lignes traitées ») que chaque ligne — ce qui rejoint la mise en garde de la [section 18.1](../18-optimisation-performance/01-profilage-mesure-performances.md). Ensuite, le seuil de niveau écarte l'écriture **avant** toute entrée/sortie, rendant les traces fines quasi gratuites en production.

Côté fiabilité, le journaliseur doit être **non bloquant** : une défaillance d'écriture ne doit jamais interrompre l'application (d'où le `On Error Resume Next`). Et un format **cohérent et analysable** (séparateurs réguliers) permet d'exploiter les journaux par import et tri.

## Confidentialité, et la différence avec l'audit

Au-delà de l'exclusion des données sensibles déjà évoquée, la journalisation se distingue de l'**audit**. Le journal applicatif sert le diagnostic et l'exploitation ; l'audit — qui a modifié quoi, et quand, à des fins de sécurité ou de conformité — répond à des exigences plus strictes (exhaustivité, résistance à l'altération). Ce sont deux dispositifs distincts.

> ℹ️ L'audit des accès et des modifications est traité au [chapitre 20.7](../20-securite-protection/07-audit-acces-modifications.md), le chiffrement des données sensibles au [chapitre 20.8](../20-securite-protection/08-chiffrement-donnees-sensibles.md).

## Points clés à retenir

- En production, où l'éditeur est inaccessible, un journal persistant permet de reconstituer ce qui s'est passé et de diagnostiquer sur le poste de l'utilisateur.
- On journalise le cycle de vie, les opérations significatives, les avertissements et les erreurs, avec leur contexte — jamais de données sensibles.
- Une journalisation hiérarchisée par niveau (`DEBUG`/`INFO`/`WARN`/`ERROR`), avec un seuil configurable, contrôle la verbosité.
- Le fichier texte local est la destination la plus robuste (il survit à une base défaillante) ; une table centralise mais ajoute charge et contention. Les deux se combinent.
- Un module de journalisation centralisé, **non bloquant**, au format analysable, est la bonne structure ; il s'intègre au gestionnaire d'erreurs.
- On maîtrise la taille par rotation (fichier) ou purge (table), et le coût par le seuil de niveau et l'évitement des traces en boucle.
- Le journal applicatif n'est pas un journal d'audit : ce dernier obéit à des exigences propres.

---


⏭️ [19.6. Simulation de données et environnements de test](/19-debogage-tests/06-simulation-donnees-test.md)
