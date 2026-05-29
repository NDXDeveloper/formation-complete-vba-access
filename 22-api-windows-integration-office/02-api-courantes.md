🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 22.2. API courantes — GetUserName, GetComputerName, ShellExecute, Sleep

Forts des mécanismes de déclaration et d'appel posés à la section 21.1, nous pouvons désormais cataloguer quelques fonctions de l'API Windows d'usage très fréquent. Pour chacune, cette section fournit la déclaration portable, une fonction d'enveloppe prête à l'emploi, et les particularités d'appel à connaître. Lorsqu'une solution native existe en VBA, elle est signalée, conformément au principe énoncé à la section 22.1 : ne descendre au niveau de l'API que lorsqu'elle apporte une réelle valeur.

Les quatre fonctions du titre couvrent des besoins parmi les plus répandus : identifier l'utilisateur et la machine, ouvrir un fichier ou lancer un programme, et temporiser un traitement.

## GetUserName — le nom de l'utilisateur Windows

`GetUserName`, exposée par `advapi32.dll`, renvoie le nom du compte Windows sous lequel l'application s'exécute. Sa déclaration ne comporte ni handle ni pointeur sensible à la bitness ; un simple `PtrSafe` suffit :

```vba
Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" _
    Alias "GetUserNameA" ( _
    ByVal lpBuffer As String, _
    ByRef nSize As Long) As Long
```

L'appel suit le motif du tampon de chaîne décrit à la section 22.1 :

```vba
Public Function NomUtilisateur() As String
    Dim buffer As String
    Dim taille As Long

    buffer = String$(256, vbNullChar)
    taille = Len(buffer)

    If GetUserName(buffer, taille) <> 0 Then
        ' Pour GetUserName, 'taille' inclut le terminateur nul
        NomUtilisateur = Left$(buffer, taille - 1)
    End If
End Function
```

**Alternative native :** VBA propose `Environ("USERNAME")`, plus simple, qui lit la variable d'environnement correspondante. La différence est subtile : `GetUserName` interroge directement le contexte de sécurité, tandis que `Environ` lit une variable qui, en théorie, pourrait avoir été modifiée dans l'environnement du processus. Pour la plupart des applications, `Environ("USERNAME")` est parfaitement suffisant et préférable par sa simplicité.

## GetComputerName — le nom de la machine

`GetComputerName`, exposée cette fois par `kernel32.dll`, renvoie le nom de l'ordinateur. Sa déclaration est de structure identique à la précédente :

```vba
Private Declare PtrSafe Function GetComputerName Lib "kernel32.dll" _
    Alias "GetComputerNameA" ( _
    ByVal lpBuffer As String, _
    ByRef nSize As Long) As Long
```

Mais son appel révèle une différence de **sémantique de la longueur** qui illustre exactement la mise en garde de la section 22.1 : à la sortie, `GetComputerName` place dans `nSize` le nombre de caractères copiés **hors** terminateur nul, là où `GetUserName` l'**incluait**. La troncature se fait donc sans le `- 1` :

```vba
Public Function NomOrdinateur() As String
    Dim buffer As String
    Dim taille As Long

    buffer = String$(256, vbNullChar)
    taille = Len(buffer)

    If GetComputerName(buffer, taille) <> 0 Then
        ' Pour GetComputerName, 'taille' exclut le terminateur nul
        NomOrdinateur = Left$(buffer, taille)
    End If
End Function
```

Cette divergence, pour deux fonctions pourtant jumelles en apparence, rappelle qu'il faut toujours consulter la documentation propre à chaque API plutôt que de présumer un comportement uniforme.

**Alternative native :** ici encore, `Environ("COMPUTERNAME")` fournit la même information de façon plus directe.

## ShellExecute — ouvrir un fichier, une URL ou lancer un programme

`ShellExecute`, exposée par `shell32.dll`, est l'une des API les plus polyvalentes. Sa force tient à ce qu'elle s'appuie sur les **associations de fichiers** du système : elle ouvre une URL dans le navigateur par défaut, un PDF dans le lecteur par défaut, un dossier dans l'explorateur, ou lance un programme — selon la nature de la cible et le verbe demandé. C'est ce qui la distingue de la fonction VBA native `Shell`, laquelle ne sait qu'exécuter un programme et ignore les associations de documents.

`ShellExecute` manipule un handle de fenêtre et renvoie un handle d'instance : elle est donc **sensible à la bitness** et appelle la compilation conditionnelle décrite à la section 21.7 :

```vba
#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As LongPtr
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" ( _
        ByVal hwnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long
#End If
```

Les paramètres sont, dans l'ordre : le handle de la fenêtre parente (0 si aucune), le **verbe** (`"open"`, `"print"`, `"edit"`, `"explore"`, ou `vbNullString` pour l'action par défaut), la **cible** (fichier, dossier, URL ou programme), les **paramètres** éventuels si la cible est un exécutable, le **répertoire de travail**, et le **mode d'affichage** de la fenêtre. La fonction d'enveloppe suivante en simplifie l'usage :

```vba
Public Function OuvrirAvecShell(ByVal cible As String, _
                                Optional ByVal parametres As String = vbNullString, _
                                Optional ByVal verbe As String = "open") As Boolean
    #If VBA7 Then
        Dim resultat As LongPtr
    #Else
        Dim resultat As Long
    #End If
    Const SW_SHOWNORMAL As Long = 1

    resultat = ShellExecute(0, verbe, cible, parametres, vbNullString, SW_SHOWNORMAL)
    OuvrirAvecShell = (resultat > 32)   ' succès si la valeur de retour dépasse 32
End Function
```

La valeur de retour mérite une attention particulière : `ShellExecute` renvoie une valeur **supérieure à 32 en cas de succès**, et une valeur inférieure ou égale à 32 en cas d'échec, cette dernière étant alors un **code d'erreur** (cible introuvable, aucune association, etc.). D'où le test `> 32` plutôt qu'un test sur zéro.

Quelques usages typiques :

- Ouvrir une URL : `OuvrirAvecShell "https://www.exemple.fr"`.
- Ouvrir un document dans son application par défaut : `OuvrirAvecShell "C:\docs\rapport.pdf"`.
- Ouvrir un dossier dans l'explorateur : `OuvrirAvecShell "C:\docs"`.
- Lancer un programme avec un argument : `OuvrirAvecShell "notepad.exe", "C:\fichier.txt"`.
- Imprimer un document : `OuvrirAvecShell "C:\docs\rapport.pdf", , "print"`.

Le mode d'affichage repose sur les constantes `SW_` (notamment `SW_HIDE = 0`, `SW_SHOWNORMAL = 1`, `SW_SHOWMAXIMIZED = 3`), qui contrôlent l'état de la fenêtre ouverte.

## Sleep — suspendre l'exécution

`Sleep`, exposée par `kernel32`, suspend l'exécution pendant une durée exprimée en millisecondes. Elle comble un manque réel, VBA ne disposant dans Access d'aucune fonction native de temporisation à la milliseconde. Sa déclaration est minimale :

```vba
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
```

Son emploi est immédiat : `Sleep 500` met l'exécution en pause durant un demi-seconde.

Une mise en garde est toutefois **essentielle** : `Sleep` bloque le thread courant, ce qui **fige entièrement Access** pendant toute la durée de la pause. L'interface ne répond plus, aucun événement n'est traité, et l'application paraît bloquée. Pour de très courtes pauses, c'est sans conséquence ; mais une longue temporisation par `Sleep` dans le thread de l'interface donne l'impression d'une application plantée. Lorsqu'il faut attendre tout en gardant Access réactif, une autre approche s'impose, présentée ci-dessous.

## GetTickCount — mesurer le temps écoulé et attendre sans figer

`GetTickCount`, également exposée par `kernel32`, renvoie le nombre de millisecondes écoulées depuis le démarrage du système. Sans paramètre et renvoyant un simple entier, sa déclaration est triviale :

```vba
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
```

Elle sert classiquement à **mesurer une durée** par différence entre deux relevés, et permet surtout de construire une attente **réactive**, alternative au figement de `Sleep`, en cédant la main à Windows à chaque tour de boucle :

```vba
' Pause de 'duree' millisecondes en laissant Access réactif
Public Sub PauseReactive(ByVal duree As Long)
    Dim debut As Long
    debut = GetTickCount()
    Do While (GetTickCount() - debut) < duree
        DoEvents          ' laisse Windows traiter les messages en attente
    Loop
End Sub
```

Le compromis est clair : `Sleep` est économe mais fige l'application, tandis que cette boucle reste réactive au prix d'une consommation de processeur (boucle active). Pour équilibrer les deux, on peut combiner un `Sleep` de quelques millisecondes et un `DoEvents` à chaque itération, obtenant ainsi une attente à la fois réactive et peu coûteuse.

Deux réserves sur `GetTickCount` : renvoyant un entier 32 bits, sa valeur finit par déborder après plusieurs jours de fonctionnement continu, ce qui la réserve à la mesure de **courtes durées** ; pour des mesures longues, la variante `GetTickCount64` est préférable.

## D'autres API fréquemment utiles

Au-delà de ces fonctions, d'autres API reviennent régulièrement et obéissent aux mêmes principes de déclaration et d'appel. On peut citer `GetSystemMetrics` (`user32`), qui fournit les dimensions de l'écran — utile pour centrer un formulaire —, `FindWindow` (`user32`), qui retrouve le handle d'une fenêtre par sa classe ou son titre, ou encore `GetTempPath` (`kernel32`), qui renvoie le dossier temporaire du système. Leurs déclarations portables figurent à l'annexe H, qui constitue le point de référence pour reprendre des déclarations déjà vérifiées.

## Préférer le natif quand c'est possible

Ce catalogue ne doit pas faire oublier le principe directeur de la section 22.1. Pour le nom de l'utilisateur ou de la machine, `Environ` rend dans la plupart des cas l'appel d'API superflu. Pour bien des opérations de fichiers, le `FileSystemObject` (section 22.10) offre une interface de plus haut niveau et plus sûre. L'API n'est pleinement justifiée que là où elle apporte une capacité réellement absente du natif — comme `ShellExecute` et son recours aux associations de fichiers, ou `Sleep` pour la temporisation.

## Articulation avec le reste du chapitre et de la formation

Cette section prolonge et applique les fondations posées ailleurs :

- Les **mécanismes de déclaration et d'appel** sont détaillés à la section 22.1.
- La **compatibilité 32/64 bits**, mobilisée par `ShellExecute`, est traitée à la section 21.7.
- Les **déclarations portables** prêtes à l'emploi sont rassemblées à l'annexe H.
- La **gestion d'erreurs** et les **implications de sécurité** relèvent respectivement du chapitre 13 et de la section 20.5.
- Le **système de fichiers**, alternative de haut niveau à bon nombre d'opérations, est traité à la section 22.10.

## Points de vigilance

Plusieurs erreurs reviennent fréquemment avec ces API :

- **Utiliser un `Sleep` long dans le thread de l'interface**, ce qui fige Access et donne l'impression d'un plantage ; préférer une attente réactive lorsque c'est nécessaire.
- **Confondre la sémantique de la longueur** entre `GetUserName` (terminateur nul inclus) et `GetComputerName` (exclu), et tronquer le résultat d'un caractère de trop ou de trop peu.
- **Tester `ShellExecute` sur zéro** au lieu de vérifier que le retour dépasse 32.
- **Omettre la compilation conditionnelle de `ShellExecute`**, et provoquer une incompatibilité 32/64 bits (section 21.7).
- **Recourir à l'API là où le natif suffit**, par exemple appeler `GetUserName` quand `Environ("USERNAME")` aurait fait l'affaire.

⏭️ [22.3. Automation avec Excel — export et import de données](/22-api-windows-integration-office/03-automation-excel.md)
