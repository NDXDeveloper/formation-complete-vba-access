🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 21.3. Mise à jour du front-end — stratégies d'auto-update

Dans une architecture front-end / back-end, chaque utilisateur possède sa propre copie locale du front-end. Cette organisation, indispensable au bon fonctionnement multi-utilisateur, a une conséquence directe sur la maintenance : chaque fois que vous publiez une nouvelle version de l'application, ce sont autant de copies locales qu'il faut remplacer. Sur un parc de quelques dizaines ou centaines de postes, passer manuellement de machine en machine est tout simplement impraticable.

L'**auto-update** désigne l'ensemble des mécanismes permettant qu'une nouvelle version du front-end soit déployée automatiquement sur le poste de l'utilisateur, sans intervention manuelle. C'est l'une des fonctionnalités qui distingue une application Access amateur d'une application réellement exploitable en production.

## Pourquoi mettre à jour le front-end, et non le back-end

Il faut d'abord bien comprendre ce qui change d'une version à l'autre. Dans le modèle front-end / back-end :

- Le **front-end** contient les formulaires, états, requêtes et code VBA. C'est lui qui évolue à chaque livraison : correction d'un bug, nouvel écran, amélioration d'un traitement. C'est donc lui, et lui seul, que l'auto-update doit remplacer.
- Le **back-end** contient les données. Il reste en place, partagé, et ne doit surtout pas être écrasé par une mise à jour, sous peine de détruire les données des utilisateurs.

L'auto-update concerne par conséquent exclusivement le front-end. Les modifications de la *structure* du back-end (ajout d'une table, d'un champ, modification d'un index) constituent un sujet distinct et plus délicat — celui des migrations de schéma — qui exige des précautions spécifiques et ne se résume jamais à une simple copie de fichier. Une bonne pratique consiste à coordonner les deux : faire en sorte qu'une nouvelle version du front-end soit déployée *en même temps* que les évolutions de schéma qu'elle suppose, et à empêcher un front-end obsolète de travailler sur un back-end déjà modifié (point abordé plus bas).

## La contrainte technique fondamentale

Toute stratégie d'auto-update doit composer avec une contrainte incontournable : **un fichier Access ouvert ne peut pas s'écraser lui-même.** Tant que le front-end est en cours d'exécution, son fichier est verrouillé par Access (matérialisé par le fichier de verrouillage `.laccdb`), et il est impossible de le remplacer par une version plus récente.

Cette contrainte explique pourquoi on ne peut pas se contenter, dans le code du front-end, de copier la nouvelle version par-dessus l'ancienne : le fichier en cours d'exécution est précisément celui que l'on voudrait remplacer. Toutes les stratégies qui suivent contournent cet obstacle, soit en confiant la copie à un **processus tiers** (un lanceur), soit en différant la copie **après la fermeture** du front-end.

## Suivre les versions : numérotation et stockage

Avant de déclencher une mise à jour, il faut pouvoir répondre à une question simple : la copie locale est-elle à jour ? Cela suppose de comparer la version locale à une version de référence.

### Comparaison par date de fichier

L'approche la plus simple et la plus robuste consiste à comparer les **dates de dernière modification** des fichiers : si le front-end présent dans le dossier de déploiement réseau est plus récent que la copie locale, c'est qu'une mise à jour est disponible. Cette méthode ne nécessite aucune gestion explicite de numéro de version et s'appuie directement sur le système de fichiers :

```vba
' Renvoie True si la version réseau est plus récente que la copie locale
Public Function MiseAJourDisponible(ByVal cheminReseau As String, _
                                    ByVal cheminLocal As String) As Boolean
    MiseAJourDisponible = False
    If Dir(cheminReseau) = "" Then Exit Function       ' source introuvable
    If Dir(cheminLocal) = "" Then                       ' aucune copie locale
        MiseAJourDisponible = True
        Exit Function
    End If
    MiseAJourDisponible = (FileDateTime(cheminReseau) > FileDateTime(cheminLocal))
End Function
```

### Comparaison par numéro de version explicite

Lorsque l'on souhaite un contrôle plus fin (par exemple imposer une version minimale, ou afficher la version à l'utilisateur), on numérote explicitement l'application. La version locale est alors une valeur **compilée dans le front-end**, fixée au moment de la génération de l'ACCDE :

```vba
' Version du front-end, à incrémenter à chaque livraison
Public Const APP_VERSION As String = "1.4.2"
```

La version de référence est, elle, stockée à un emplacement que l'on peut faire évoluer sans recompiler — typiquement un petit fichier texte dans le dossier de déploiement réseau :

```vba
' Lit la version de référence depuis un fichier texte sur le réseau
Public Function LireVersionReference() As String
    Const FICHIER_VERSION As String = "\\Serveur\Partage\Deploiement\version.txt"
    Dim numFichier As Integer
    Dim ligne As String

    LireVersionReference = ""
    If Dir(FICHIER_VERSION) = "" Then Exit Function

    numFichier = FreeFile
    Open FICHIER_VERSION For Input As #numFichier
    Line Input #numFichier, ligne
    Close #numFichier

    LireVersionReference = Trim$(ligne)
End Function
```

Un point de vigilance : comparer des numéros de version sous forme de chaînes est trompeur, car `"1.4.10"` est considéré comme *inférieur* à `"1.4.2"` dans un tri alphabétique. Si vous comparez des versions multi-parties, convertissez chaque segment en nombre, ou adoptez un simple **numéro de build entier** croissant, beaucoup plus facile à comparer de façon fiable. En pratique, la comparaison par date de fichier évite entièrement ce piège et reste préférable dans la plupart des cas.

## Stratégie 1 : le lanceur (launcher)

C'est l'approche la plus répandue et la plus robuste. Le principe est de séparer le rôle de *mise à jour* du rôle d'*exécution* : l'utilisateur ne lance jamais directement le front-end, mais un petit **lanceur** distinct, qui se charge de mettre à jour la copie locale avant de la démarrer.

Le déroulement est le suivant :

1. L'utilisateur clique sur le raccourci du lanceur (et non sur le front-end).
2. Le lanceur compare la version disponible sur le réseau à la copie locale.
3. Si le réseau est plus récent — ou si aucune copie locale n'existe — il copie le front-end réseau vers le dossier local.
4. Le lanceur démarre la copie locale, le cas échéant en mode runtime.

L'atout décisif de cette approche est que **le lanceur est un processus séparé** : il effectue la copie alors que le front-end n'est pas encore ouvert, ce qui élimine la contrainte du fichier verrouillé. De plus, le lanceur évolue très rarement ; une fois installé sur les postes, il n'a en général plus jamais besoin d'être modifié.

Le lanceur peut prendre plusieurs formes : un fichier de commandes (`.cmd`), un script VBScript (`.vbs`), une petite base Access dédiée, ou un exécutable compilé. Voici un exemple complet sous forme de VBScript, qui ne dépend d'aucune installation particulière :

```vbscript
' MAJ_Lanceur.vbs — met à jour la copie locale puis lance l'application
Option Explicit

Const SRC_FE     = "\\Serveur\Partage\Deploiement\MonApp.accde"
Const ACCESS_EXE = """C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE"""

Dim fso, sh, dossierLocal, feLocal

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")

' Dossier local inscriptible, propre à l'utilisateur
dossierLocal = sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\MonApp"
If Not fso.FolderExists(dossierLocal) Then fso.CreateFolder dossierLocal
feLocal = dossierLocal & "\MonApp.accde"

' Copier si la copie locale est absente ou plus ancienne que la version réseau
If Not fso.FileExists(feLocal) Then
    fso.CopyFile SRC_FE, feLocal, True
ElseIf fso.GetFile(SRC_FE).DateLastModified > fso.GetFile(feLocal).DateLastModified Then
    fso.CopyFile SRC_FE, feLocal, True
End If

' Lancer la copie locale en mode runtime
sh.Run ACCESS_EXE & " """ & feLocal & """ /runtime", 1, False
```

Le raccourci distribué aux utilisateurs pointe alors vers ce script plutôt que vers le front-end. L'opération de copie repose ici sur le `FileSystemObject`, dont la manipulation est détaillée à la section 22.10.

## Stratégie 2 : auto-mise à jour avec relance externe

Lorsqu'on ne souhaite pas dépendre d'un lanceur externe, le front-end peut gérer lui-même sa mise à jour, à condition de déléguer la copie à un processus tiers qui s'exécutera *après* sa fermeture.

Le déroulement est le suivant :

1. Au démarrage, le front-end vérifie si une version plus récente existe sur le réseau.
2. Si c'est le cas, il écrit un petit script de mise à jour dans un dossier temporaire et le lance.
3. Le front-end se ferme (`DoCmd.Quit`).
4. Le script attend que le fichier ne soit plus verrouillé, copie la nouvelle version par-dessus l'ancienne, puis relance l'application.

Le contrôle de mise à jour s'insère dans la séquence de démarrage du front-end (macro `AutoExec` ou événement d'ouverture du formulaire de démarrage, dont l'ordre est décrit au chapitre 8) :

```vba
' À appeler au démarrage du front-end
Public Sub VerifierMiseAJour()
    Const SRC_FE As String = "\\Serveur\Partage\Deploiement\MonApp.accde"
    Dim cheminLocal As String

    cheminLocal = CurrentProject.FullName   ' le front-end actuellement ouvert

    If MiseAJourDisponible(SRC_FE, cheminLocal) Then
        If MsgBox("Une mise à jour est disponible. L'installer maintenant ?", _
                  vbQuestion + vbYesNo, "Mise à jour") = vbYes Then
            LancerMiseAJourEtQuitter SRC_FE, cheminLocal
        End If
    End If
End Sub
```

La procédure de relance génère un fichier de commandes qui boucle tant que la copie échoue — c'est-à-dire tant qu'Access détient encore le verrou sur le fichier — puis relance l'application une fois la copie réussie :

```vba
Private Sub LancerMiseAJourEtQuitter(ByVal source As String, ByVal cible As String)
    Dim cheminScript As String
    Dim numFichier As Integer

    cheminScript = Environ$("TEMP") & "\maj_monapp.cmd"

    numFichier = FreeFile
    Open cheminScript For Output As #numFichier
    Print #numFichier, "@echo off"
    Print #numFichier, "rem Attendre la libération du fichier puis copier"
    Print #numFichier, ":boucle"
    Print #numFichier, "timeout /t 1 /nobreak >nul"
    Print #numFichier, "copy /y """ & source & """ """ & cible & """ >nul"
    Print #numFichier, "if errorlevel 1 goto boucle"
    Print #numFichier, "start """" """ & cible & """ /runtime"
    Print #numFichier, "del ""%~f0"""
    Close #numFichier

    Shell "cmd /c """ & cheminScript & """", vbHide
    DoCmd.Quit
End Sub
```

La boucle `if errorlevel 1 goto boucle` est l'élément clé : tant que le front-end n'est pas complètement fermé, la copie échoue et le script réessaie ; dès qu'Access libère le verrou, la copie aboutit et l'application est relancée. Pour une relance parfaitement fiable, on peut substituer à la ligne `start` un appel explicite à `MSACCESS.EXE` suivi du chemin et du commutateur `/runtime`, plutôt que de reposer sur l'association de fichier.

## Comparaison des deux stratégies

Le lanceur et l'auto-mise à jour répondent au même besoin par des chemins différents :

- Le **lanceur** est plus simple à raisonner et plus robuste : la copie a lieu avant toute ouverture, donc sans contrainte de verrou. Il exige en contrepartie de distribuer un raccourci spécifique et de déployer initialement le lanceur sur les postes. C'est l'approche recommandée par défaut.
- L'**auto-mise à jour interne** évite tout composant externe et concentre la logique dans le front-end, mais elle est plus délicate à mettre au point (gestion du verrou, relance) et impose une fermeture-réouverture de l'application visible par l'utilisateur.

Dans les deux cas, la copie de référence du front-end réside dans un dossier de déploiement réseau, dont la gestion des chemins (UNC) est traitée à la section 21.5.

## Où copier le front-end sur le poste

Quel que soit le mécanisme retenu, la copie locale doit être placée dans un **dossier inscriptible par l'utilisateur**, sans privilège administrateur. Le dossier `%LOCALAPPDATA%` (typiquement `C:\Utilisateurs\<nom>\AppData\Local`) est le choix naturel : il est propre à chaque utilisateur et accessible en écriture.

À l'inverse, le dossier `Program Files` ne convient pas : il est protégé en écriture pour les utilisateurs standard, et toute tentative d'y copier une mise à jour échouera sur la plupart des postes correctement administrés. Placer le front-end sous `Program Files` est une cause fréquente d'échec silencieux des mises à jour.

## Mise à jour obligatoire et cohérence front-end / back-end

Une mise à jour peut être **facultative** (l'utilisateur choisit de l'installer) ou **obligatoire** (l'application refuse de démarrer tant qu'elle n'est pas à jour). Le second cas est indispensable lorsqu'une évolution du schéma du back-end rend les anciennes versions du front-end incompatibles : laisser tourner un front-end obsolète sur un back-end déjà modifié exposerait à des erreurs, voire à des corruptions de données.

Le motif le plus sûr consiste à faire **déclarer par le back-end la version minimale de front-end attendue**, par exemple dans une petite table de paramètres. Au démarrage, le front-end compare sa propre version (la constante `APP_VERSION`) à cette version minimale, et bloque l'accès — en proposant la mise à jour — si elle est insuffisante :

```vba
' Bloque l'accès si le front-end est antérieur à la version exigée par le back-end
Public Function FrontEndCompatible() As Boolean
    Dim versionExigee As String
    versionExigee = DLookup("Valeur", "tblParametres", "Cle = 'VersionFEMini'")
    FrontEndCompatible = (APP_VERSION >= versionExigee)   ' cf. remarque sur la comparaison
End Function
```

Ce verrou garantit que front-end et back-end restent toujours cohérents, et qu'une mise à jour de structure ne peut jamais être contournée par un poste resté en arrière. La fonction `DLookup` employée ici est décrite à la section 11.10.

## Articulation avec le reste du déploiement

L'auto-update s'inscrit dans la chaîne globale du déploiement et s'appuie sur plusieurs autres sections du chapitre :

- Le livrable copié est l'**ACCDE** préparé selon la section 21.1, lancé en mode runtime (section 21.2).
- Après la copie d'un front-end neuf, il peut être nécessaire de **rétablir les liaisons vers le back-end** selon le contexte du poste, ce qui relève de la reliaison automatique des tables (section 21.4).
- Le dossier de déploiement réseau et les **chemins UNC** font l'objet de la section 21.5.
- L'**installation initiale** sur un poste vierge — qui précède toute mise à jour — est traitée par l'installation automatisée (section 21.8).

L'auto-update gère les *évolutions* d'une application déjà installée ; il ne dispense pas d'une procédure d'installation de premier déploiement.

## Points de vigilance

Plusieurs erreurs reviennent fréquemment dans la mise en place d'un auto-update :

- **Tenter d'écraser le front-end ouvert depuis son propre code.** Le fichier étant verrouillé, l'opération est impossible : il faut toujours passer par un processus tiers ou différer la copie après fermeture.
- **Copier la mise à jour dans un dossier non inscriptible** (`Program Files`), provoquant un échec silencieux sur les postes des utilisateurs standard.
- **Comparer les versions sous forme de chaînes** sans tenir compte des segments numériques, ce qui conduit à des comparaisons fausses ; préférer la date de fichier ou un numéro de build entier.
- **Oublier la cohérence front-end / back-end** lors d'une évolution de schéma, et laisser des postes obsolètes travailler sur des données dont la structure a changé.
- **Écraser le back-end** par mégarde lors d'une mise à jour : seul le front-end doit être remplacé, jamais le fichier de données.

⏭️ [21.4. Reliaison automatique des tables back-end (re-link)](/21-deploiement-distribution/04-reliaison-automatique-tables.md)
