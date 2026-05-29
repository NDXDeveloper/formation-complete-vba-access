🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 21.8. Installation automatisée (scripts, paramètres de ligne de commande)

L'installation automatisée est l'aboutissement du déploiement : c'est l'étape qui met en place l'application sur un poste **vierge**, en réunissant tout ce qui a été préparé au fil des sections précédentes. Automatiser cette installation est ce qui rend possible un déploiement à grande échelle : plutôt que de configurer chaque poste à la main, on exécute un script — éventuellement de façon silencieuse et centralisée — qui accomplit l'ensemble des opérations de façon reproductible.

Cette section décrit ce qu'une installation doit accomplir, présente les paramètres de ligne de commande d'Access utiles au lancement, et montre comment assembler les différentes pièces en un script d'installation cohérent.

## Installation initiale et mise à jour : deux étapes distinctes

Il faut d'abord bien situer l'installation par rapport à l'auto-update vu à la section 21.3. Les deux répondent à des moments différents du cycle de vie :

- L'**installation initiale** (objet de cette section) configure un poste qui ne dispose encore de rien : elle copie l'application, crée les raccourcis, met en place la configuration et les autorisations.
- La **mise à jour** (section 21.3) intervient ensuite, sur un poste déjà installé, pour y déployer les nouvelles versions du front-end.

Les deux sont complémentaires, et une approche élégante consiste à les articuler : l'installation met en place un **lanceur** (section 21.3), et c'est ce lanceur qui, à chaque démarrage ultérieur, assure automatiquement les mises à jour. L'installation initiale prépare le terrain ; l'auto-update entretient l'application dans la durée.

## Ce qu'une installation doit accomplir

Installer une application Access sur un poste vierge suppose d'enchaîner plusieurs opérations, dont chacune renvoie à une section antérieure du chapitre :

- **S'assurer de la présence du moteur d'exécution.** Si le poste ne dispose pas d'un Access complet compatible, le moteur Runtime (section 21.2) doit être installé, dans la bonne bitness (section 21.7).
- **Copier le front-end ou le lanceur localement**, dans un dossier inscriptible par l'utilisateur (section 21.3).
- **Établir la première liaison vers le back-end** et mémoriser son chemin UNC (sections 21.4 et 21.5). En pratique, cette reliaison est le plus souvent réalisée par l'application elle-même au premier démarrage.
- **Créer un raccourci de lancement** qui ouvre l'application en mode runtime (sections 21.1 et 21.2).
- **Autoriser l'emplacement d'installation** afin que le code VBA ne soit pas bloqué par la sécurité d'Access (section 20.5).

L'installation automatisée consiste à coder cet enchaînement de manière à pouvoir le rejouer à l'identique sur chaque poste.

## Les paramètres de ligne de commande d'Access

Le lancement d'une application repose sur l'exécutable `MSACCESS.EXE`, auquel on passe le chemin de la base suivi d'éventuels commutateurs. Les plus utiles au déploiement sont :

- **`"chemin"`** : le chemin de la base à ouvrir, premier argument (entre guillemets s'il contient des espaces).
- **`/runtime`** : ouvre l'application en mode runtime (section 21.2).
- **`/cmd`** : transmet une chaîne récupérable dans le code par la fonction `Command()` ; doit être le **dernier** commutateur de la ligne.
- **`/excl`** : ouvre la base en accès exclusif ; **`/ro`** : en lecture seule.
- **`/x macro`** : exécute la macro indiquée au démarrage.
- **`/compact [cible]`** : compacte et répare la base (compactage abordé à la section 18.7).
- **`/nostartup`** : démarre sans afficher l'écran d'accueil ; **`/profile profil`** : utilise un profil de paramètres particulier.

Les anciens commutateurs **`/user`**, **`/pwd`** et **`/wrkgrp`**, hérités de la sécurité par groupe de travail du moteur Jet, ne s'appliquent pas au format `.accdb` et relèvent du legacy.

### Passer des paramètres à l'application : /cmd et Command()

Le commutateur `/cmd` est particulièrement précieux : il permet de transmettre une information à l'application au moment de son lancement, que le code lit ensuite via la fonction `Command()`. On l'emploie typiquement pour sélectionner un environnement (production, test) ou pointer une configuration, sans modifier l'application elle-même :

```vba
' Au démarrage : lecture du paramètre transmis via /cmd
Public Sub TraiterParametreDemarrage()
    Dim parametre As String
    parametre = LCase$(Trim$(Command()))

    Select Case parametre
        Case "prod"
            ' Configuration de production
        Case "test"
            ' Configuration de test (autre back-end, journalisation accrue...)
        Case Else
            ' Comportement par défaut
    End Select
End Sub
```

La ligne de lancement correspondante place `/cmd` en dernier :

```
"C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE" "C:\App\MonApp.accde" /runtime /cmd prod
```

## Créer le raccourci de lancement

Le raccourci est ce sur quoi l'utilisateur clique réellement. Sa cible est l'exécutable `MSACCESS.EXE`, ses arguments le chemin de l'application et le commutateur `/runtime` — ou, mieux, le **lanceur** afin que chaque démarrage déclenche la vérification de mise à jour. Un script crée ce raccourci via l'objet `WScript.Shell` :

```vbscript
' Création d'un raccourci de lancement sur le Bureau
Dim sh, lnk, accessExe, dossierLocal, feLocal

Set sh = CreateObject("WScript.Shell")
accessExe    = "C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE"
dossierLocal = sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\MonApp"
feLocal      = dossierLocal & "\MonApp.accde"

Set lnk = sh.CreateShortcut(sh.SpecialFolders("Desktop") & "\MonApp.lnk")
lnk.TargetPath       = accessExe
lnk.Arguments        = """" & feLocal & """ /runtime"
lnk.WorkingDirectory = dossierLocal
lnk.Description       = "Lancer MonApp"
lnk.Save
```

Le chemin de `MSACCESS.EXE` variant selon la version, la bitness et le mode d'installation, il est plus robuste de **résoudre dynamiquement** ce chemin plutôt que de le coder en dur : la valeur par défaut de la clé de registre `App Paths\MSACCESS.EXE` (sous `HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths`) fournit l'emplacement exact de l'exécutable. La lecture du registre est traitée à la section 22.9.

## Approuver l'emplacement : les emplacements de confiance

Une étape souvent oubliée, et pourtant décisive, concerne la **sécurité des macros**. Access bloque par défaut l'exécution du code VBA provenant d'un emplacement non approuvé (Centre de gestion de la confidentialité, section 20.5). Si le dossier d'installation n'est pas déclaré de confiance, l'application peut s'ouvrir sans que son code ne s'exécute, voire afficher des avertissements bloquants.

Deux solutions existent. La première consiste à ajouter le dossier d'installation à la liste des **emplacements de confiance**, information stockée dans le registre sous une clé propre à la version (le « 16.0 » commun à Access 2016 à 2024 et Microsoft 365, voir section 21.6) :

```
HKCU\Software\Microsoft\Office\16.0\Access\Security\Trusted Locations\MonApp\
    Path             = (le dossier d'installation)
    AllowSubFolders  = 1
    Description       = MonApp
```

Si l'emplacement approuvé est un chemin **réseau** (UNC), il faut en outre activer la valeur `AllowNetworkLocations` sous la même branche, les emplacements de confiance réseau étant désactivés par défaut. Approuver un dossier **local** comme dans l'exemple ci-dessous ne requiert pas ce réglage.

La seconde solution consiste à **signer numériquement** le code par un certificat de confiance (section 20.4) : une fois l'éditeur approuvé, le code s'exécute quel que soit son emplacement. Les deux approches sont complémentaires, le choix dépendant du contexte de l'organisation.

## Installer le moteur d'exécution

Lorsque le poste ne dispose d'aucun Access compatible, l'installation du moteur **Runtime** (section 21.2) constitue un prérequis. Cette opération nécessite généralement des privilèges d'administration et est souvent traitée comme une étape distincte du script applicatif. Le programme d'installation du Runtime peut s'exécuter en mode silencieux à l'aide de commutateurs adaptés, et pour les versions récentes via l'outil de déploiement d'Office et un fichier de configuration. La bitness du Runtime doit impérativement correspondre à celle du livrable et de l'environnement (section 21.7).

## Assembler le tout : un script d'installation

Les différentes opérations se réunissent en un script unique, idéalement **idempotent** (sans dommage s'il est relancé) et exécutable de façon **silencieuse** pour permettre un déploiement de masse. L'exemple suivant crée le dossier local, copie le front-end, déclare l'emplacement de confiance et installe le raccourci ; la reliaison vers le back-end est laissée au premier démarrage de l'application (section 21.4) :

```vbscript
' install.vbs — installation initiale de MonApp sur un poste
Option Explicit

Const SRC_FE     = "\\Serveur\AppPartage\Deploiement\MonApp.accde"
Const ACCESS_EXE = "C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE"

Dim fso, sh, dossierLocal, feLocal, cleConfiance, lnk

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")

' 1) Créer le dossier local inscriptible
dossierLocal = sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\MonApp"
If Not fso.FolderExists(dossierLocal) Then fso.CreateFolder dossierLocal
feLocal = dossierLocal & "\MonApp.accde"

' 2) Copier le front-end depuis le réseau
fso.CopyFile SRC_FE, feLocal, True

' 3) Déclarer le dossier comme emplacement de confiance (par utilisateur)
cleConfiance = "HKCU\Software\Microsoft\Office\16.0\Access\Security\" & _
               "Trusted Locations\MonApp\"
sh.RegWrite cleConfiance & "Path", dossierLocal, "REG_SZ"
sh.RegWrite cleConfiance & "AllowSubFolders", 1, "REG_DWORD"
sh.RegWrite cleConfiance & "Description", "MonApp", "REG_SZ"

' 4) Créer le raccourci de lancement sur le Bureau
Set lnk = sh.CreateShortcut(sh.SpecialFolders("Desktop") & "\MonApp.lnk")
lnk.TargetPath       = ACCESS_EXE
lnk.Arguments        = """" & feLocal & """ /runtime"
lnk.WorkingDirectory = dossierLocal
lnk.Description       = "Lancer MonApp"
lnk.Save

' La reliaison vers le back-end est assurée au premier démarrage (cf. 21.4)
```

Ce script repose sur le `FileSystemObject` pour la copie (section 22.10) et sur les écritures de registre (section 22.9) pour l'emplacement de confiance. Comme l'écriture du registre `HKCU` et le raccourci du Bureau sont **propres à l'utilisateur**, un déploiement couvrant plusieurs comptes doit s'exécuter dans le contexte de chacun, ou recourir à une approche par stratégie de groupe ciblant la machine.

## Choisir l'outil : scripts ou programme d'installation ?

Plusieurs niveaux d'outillage sont envisageables selon l'ampleur et l'exigence du déploiement.

Pour un déploiement interne sur réseau local, les **scripts** — fichiers de commandes, VBScript ou PowerShell — sont légers et n'exigent aucun outillage particulier ; PowerShell offre aujourd'hui les capacités les plus complètes. Pour une distribution plus aboutie, dotée d'une désinstallation propre, d'une gestion des prérequis et d'entrées dans le menu Démarrer, un véritable **programme d'installation** construit avec un outil dédié (de type Inno Setup, WiX ou NSIS) est préférable. Enfin, à l'échelle de l'entreprise, le déploiement de masse passe le plus souvent par les outils d'administration centralisée — **stratégie de groupe, Configuration Manager ou Intune** — qui exécutent l'installation sur l'ensemble du parc sans intervention locale.

## Articulation avec le reste du déploiement

L'installation automatisée fédère l'ensemble des sujets du chapitre, dont elle constitue la synthèse opérationnelle :

- Elle déploie le livrable **ACCDE** (section 21.1) et le fait ouvrir en **mode runtime** (sections 21.1 et 21.2).
- Elle installe au besoin le moteur **Runtime** (section 21.2), dans la **bitness** appropriée (section 21.7).
- Elle prépare l'**auto-update** ultérieur via le lanceur (section 21.3).
- Elle laisse la **reliaison** au premier démarrage (section 21.4), avec des chemins **UNC** (section 21.5).
- Elle approuve l'emplacement via les **emplacements de confiance** ou la **signature numérique** (sections 20.5 et 20.4), et s'appuie sur le **registre** et le **système de fichiers** (sections 22.9 et 22.10).

## Points de vigilance

Plusieurs erreurs reviennent fréquemment dans l'automatisation de l'installation :

- **Oublier d'approuver l'emplacement d'installation**, ce qui aboutit à une application dont le code VBA ne s'exécute pas ou qui affiche des avertissements de sécurité.
- **Coder en dur le chemin de `MSACCESS.EXE`**, alors qu'il varie selon la version, la bitness et le mode d'installation ; mieux vaut le résoudre via le registre.
- **Installer dans un dossier non inscriptible** (`Program Files`), empêchant les mises à jour ultérieures du front-end (section 21.3).
- **Négliger la nature par utilisateur** de certaines opérations (emplacement de confiance `HKCU`, raccourci du Bureau) lors d'un déploiement multi-comptes.
- **Confondre installation initiale et mise à jour**, et tenter de tout faire reposer sur un même mécanisme alors que les deux étapes ont des rôles distincts.
- **Déployer un Runtime ou un ACCDE de mauvaise bitness** par rapport à l'environnement du poste (section 21.7).

⏭️ [22. API Windows et intégration Office](/22-api-windows-integration-office/README.md)
