🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 22.1. Déclaration et appel d'API Windows depuis Access (PtrSafe)

L'API Windows est la bibliothèque de fonctions du système d'exploitation : des milliers de fonctions, exposées par des DLL système, qui pilotent tout ce que fait Windows. VBA sait les appeler directement, à condition de les **déclarer** au préalable. C'est ce mécanisme qui ouvre à Access des capacités absentes de son propre modèle objet — connaître l'utilisateur connecté, lancer un programme, lire le registre, et bien d'autres — et qui sert de fondation à l'ensemble des techniques d'API de ce chapitre.

Appeler l'API est puissant mais exigeant : on y manipule des fonctions de bas niveau, sans le filet de sécurité dont bénéficie le reste de VBA. Une déclaration erronée peut faire planter Access sans avertissement. Cette section détaille l'anatomie complète d'une déclaration et d'un appel, afin de poser ces fondations avec la rigueur qu'elles réclament.

## Une déclaration, pas une référence

Une première particularité mérite d'être soulignée d'emblée. Contrairement à l'automation des applications Office, abordée plus loin dans ce chapitre, l'appel d'une fonction d'API **ne nécessite aucune référence** à activer (chapitre 2.5) : il suffit de déclarer la fonction par une instruction `Declare`, et VBA établit le lien avec la DLL au moment de l'appel. C'est à la fois plus léger et plus direct, mais cela reporte sur le développeur l'entière responsabilité de l'exactitude de la déclaration.

## L'instruction Declare : anatomie

### La syntaxe générale

Une déclaration d'API suit ce gabarit :

```vba
[Public | Private] Declare [PtrSafe] {Sub | Function} nom Lib "bibliothèque" _
    [Alias "nomRéel"] ([liste_arguments]) [As type]
```

Chaque élément a un rôle précis :

- `Public` ou `Private` fixe la **portée** de la déclaration.
- `PtrSafe` atteste la **compatibilité 64 bits** de la déclaration (voir ci-dessous).
- `Sub` ou `Function` selon que la fonction renvoie ou non une valeur exploitable.
- `nom` est le nom sous lequel vous appellerez la fonction en VBA.
- `Lib` indique la **DLL** qui contient la fonction.
- `Alias` précise le **nom réellement exporté** par la DLL, lorsqu'il diffère du nom VBA.
- la liste d'arguments et le type de retour, soumis à des règles strictes développées plus loin.

### Portée et emplacement de la déclaration

La portée obéit à une contrainte importante : dans un **module standard**, une déclaration peut être `Public` ou `Private` ; mais dans un **module de classe ou de formulaire**, elle doit obligatoirement être `Private`. Une déclaration `Public` y provoquerait une erreur.

La bonne pratique, déjà recommandée à la section 21.7, consiste à **centraliser toutes les déclarations d'API dans un module standard dédié**. On dispose ainsi d'un point unique pour les vérifier, les maintenir et garantir leur compatibilité 32/64 bits.

### PtrSafe et la compatibilité 32/64 bits

Le mot-clé `PtrSafe` est requis par VBA7 (Office 2010 et ultérieur) et indispensable en 64 bits : une déclaration qui en serait dépourvue provoquerait, en 64 bits, une erreur de compilation désactivant l'intégralité du projet. Ce sujet — `PtrSafe`, le type `LongPtr` pour les pointeurs et handles, la compilation conditionnelle — a été traité en détail à la section 21.7, à laquelle il convient de se reporter. On retiendra ici l'essentiel : **toute déclaration doit comporter `PtrSafe`, et tout handle ou pointeur doit être typé `LongPtr`.**

### Lib : la bibliothèque d'origine

`Lib` désigne la DLL hébergeant la fonction. Les DLL système se trouvant dans le répertoire système de Windows, **aucun chemin n'est nécessaire** : il suffit du nom. Les bibliothèques les plus fréquemment sollicitées sont :

- **`kernel32`** : noyau du système (temporisation, informations machine, mémoire).
- **`user32`** : interface et fenêtres.
- **`advapi32`** : registre, sécurité, informations utilisateur.
- **`shell32`** : opérations du shell (ouverture de fichiers, exécution).
- **`gdi32`** : dessin et graphismes.
- **`winmm`** : multimédia.

L'extension `.dll` est facultative pour ces DLL système (`"kernel32"` et `"kernel32.dll"` sont équivalents).

### Alias : le nom réellement exporté

`Alias` permet de dissocier le nom utilisé en VBA du nom réellement exporté par la DLL. On y recourt dans plusieurs cas : lorsque le nom réel n'est pas un identifiant VBA valide, lorsqu'on souhaite renommer la fonction pour son propre confort, et surtout — cas le plus fréquent — pour choisir entre les variantes ANSI et Unicode d'une fonction de chaîne, comme l'explique la section suivante.

## ANSI ou Unicode : les variantes A et W

La plupart des fonctions de l'API qui manipulent du texte existent en **deux variantes** : une version ANSI suffixée `A` (par exemple `GetUserNameA`) et une version Unicode suffixée `W` (par exemple `GetUserNameW`). Il n'existe pas de fonction au nom « nu » : le nom sans suffixe n'est qu'une commodité du langage C, résolue à la compilation vers l'une ou l'autre variante.

En VBA, le marshaling d'un paramètre déclaré `As String` vers une fonction `A` effectue automatiquement la conversion entre le texte interne (Unicode) et l'ANSI attendu par la fonction. C'est pourquoi, pour des chaînes simples, **on déclare la variante `A`** via `Alias`, en laissant VBA gérer la conversion. L'emploi de la variante `W` avec un paramètre `As String` est plus délicat et suppose une manipulation manuelle des pointeurs (via `StrPtr` et des tableaux d'octets). Dans la pratique courante, la variante `A` est donc le choix par défaut.

## Passage des paramètres : ByVal, ByRef et le cas des chaînes

VBA passe ses arguments **par référence** (`ByRef`) par défaut, alors que l'API attend des conventions précises. La correspondance doit donc être établie explicitement :

- Les **valeurs scalaires** que la fonction reçoit « par valeur » en C — un entier, un handle — se déclarent `ByVal`.
- Les paramètres que la fonction **remplit** (un entier de sortie passé par pointeur) se déclarent `ByRef`, de sorte que l'API écrive dans votre variable.

Le cas des **chaînes** est la principale source de confusion, et mérite une attention particulière. Un paramètre de type chaîne se déclare presque toujours **`ByVal ... As String`** — ce qui paraît contre-intuitif. La raison tient au marshaling : lorsqu'une chaîne est passée `ByVal`, VBA transmet à la fonction un **pointeur vers le tampon de caractères**, ce qu'attend précisément une API recevant un `LPSTR`. Une chaîne passée `ByRef` transmettrait un pointeur vers un pointeur, incorrect pour la plupart des fonctions C. La règle à mémoriser est donc : **les chaînes destinées à l'API sont `ByVal ... As String`.**

Un corollaire utile : pour transmettre un pointeur **nul** là où une chaîne est optionnelle, on passe la constante `vbNullString` — et non une chaîne vide `""`. La première produit un véritable pointeur nul, la seconde un pointeur vers une chaîne vide, ce qui n'est pas équivalent du point de vue de l'API.

## Correspondance des types Windows et VBA

Traduire correctement les types est au cœur d'une déclaration réussie. La correspondance usuelle est la suivante :

| Type Windows | Type VBA | Remarque |
|---|---|---|
| `DWORD`, `UINT`, `INT`, `LONG` | `Long` | entier 32 bits |
| `BOOL` | `Long` | 0 = faux, valeur non nulle = vrai |
| `WORD` | `Integer` | entier 16 bits |
| `BYTE` | `Byte` | |
| `HANDLE`, `HWND`, `HDC`, `HKEY`, pointeurs | `LongPtr` | **sensible à la bitness** (section 21.7) |
| `LPSTR`, `LPCSTR` (chaîne ANSI) | `ByVal ... As String` | avec la variante `A` |
| pointeur vers une structure | `ByRef ...` (type défini par l'utilisateur) | |
| pointeur vers un entier de sortie | `ByRef ... As Long` | l'API y écrit le résultat |

Le seul type réellement sensible à l'architecture est `LongPtr`, employé pour les handles et les pointeurs : c'est lui qui garantit la compatibilité entre 32 et 64 bits.

## Le motif du tampon de chaîne

De nombreuses fonctions qui **renvoient** du texte ne créent pas la chaîne elles-mêmes : elles écrivent dans un tampon que l'appelant doit avoir **pré-alloué**, et auquel il communique sa taille. Le motif est récurrent : on alloue une chaîne d'une longueur suffisante, on transmet cette longueur, l'API remplit le tampon, puis on tronque le résultat au texte réellement renvoyé.

Considérons la déclaration de `GetUserName`, qui récupère le nom de l'utilisateur Windows (cette fonction et d'autres seront cataloguées à la section 22.2) :

```vba
' Récupère le nom de l'utilisateur Windows connecté
Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" _
    Alias "GetUserNameA" ( _
    ByVal lpBuffer As String, _
    ByRef nSize As Long) As Long
```

Son appel illustre le motif du tampon :

```vba
Public Function NomUtilisateurWindows() As String
    Dim buffer As String
    Dim taille As Long

    buffer = String$(256, vbNullChar)   ' pré-allouer le tampon
    taille = Len(buffer)                 ' communiquer sa taille à l'API

    If GetUserName(buffer, taille) <> 0 Then     ' retour non nul = succès
        ' Ici, 'taille' contient le nombre de caractères, terminateur nul compris
        NomUtilisateurWindows = Left$(buffer, taille - 1)
    Else
        NomUtilisateurWindows = ""
    End If
End Function
```

Deux points sont à retenir. D'une part, **le tampon doit être assez grand** : une allocation insuffisante conduit à un texte tronqué, voire à une erreur. D'autre part, **la sémantique exacte de la longueur varie d'une fonction à l'autre** : certaines comptent le terminateur nul dans la longueur retournée, d'autres non, et certaines exigent simplement de tronquer au premier caractère nul. Il faut donc se référer à la documentation de chaque fonction — les cas concrets des API courantes sont détaillés à la section 22.2.

## Appeler la fonction et interpréter le résultat

Un point fondamental distingue l'appel d'API de l'appel d'une procédure VBA : **une fonction d'API ne déclenche pas d'erreur d'exécution VBA en cas d'échec.** Elle se contente de renvoyer une valeur, qu'il revient au code d'interpréter selon le contrat de la fonction — souvent zéro pour un échec et une valeur non nulle pour un succès, mais les conventions varient. Ignorer cette valeur de retour, c'est s'exposer à poursuivre comme si tout allait bien alors que l'appel a échoué.

Pour obtenir le **code d'erreur du système** après un appel infructueux, VBA expose `Err.LastDllError`, qui restitue la dernière erreur signalée par Windows :

```vba
If GetUserName(buffer, taille) = 0 Then
    Debug.Print "Échec de GetUserName, code système : " & Err.LastDllError
End If
```

`Err.LastDllError` n'est exploitable qu'**immédiatement après l'appel** concerné, et seulement pour les fonctions qui renseignent l'erreur du dernier appel. La gestion d'erreurs propre à VBA (chapitre 13) encadre quant à elle le code environnant, mais ne capte pas en elle-même les échecs internes d'une API.

## As Any : souplesse et danger

Le pseudo-type `As Any` désactive la vérification de type sur un paramètre, autorisant à y passer des valeurs de natures différentes — typiquement une chaîne ou un pointeur nul `ByVal 0&` selon les cas. Cette souplesse est utile pour les fonctions dont un paramètre peut alternativement recevoir un pointeur valide ou la valeur nulle.

Elle a toutefois un revers majeur : en supprimant tout contrôle, `As Any` reporte sur le développeur l'entière responsabilité de passer un argument de la bonne forme. Une erreur n'est plus détectée à la compilation et se traduit, à l'exécution, par un comportement indéfini ou un plantage. `As Any` est donc une technique avancée, à réserver aux cas qui l'exigent réellement et à manier avec la plus grande prudence.

## Trouver et traduire une signature

Les signatures des fonctions de l'API sont publiées dans la documentation Windows officielle, sous forme de prototypes en langage C. Le travail du développeur consiste à **traduire** fidèlement ce prototype en déclaration VBA, en appliquant les correspondances de types et les conventions de passage exposées plus haut. Cette traduction est l'étape la plus délicate : c'est elle qui détermine la justesse, et donc la sûreté, de l'appel.

Pour éviter les erreurs de traduction, l'**annexe H** de cette formation rassemble des déclarations d'API Windows déjà vérifiées et compatibles 32/64 bits, prêtes à être reprises. Partir d'une déclaration éprouvée est toujours préférable à une retranscription manuelle, qui multiplie les occasions de se tromper.

## Bonnes pratiques et risques

L'appel d'API s'exécutant sans filet de sécurité, quelques précautions sont impératives :

- **Enregistrer son travail avant de tester** un nouvel appel d'API : une déclaration incorrecte peut provoquer la fermeture brutale d'Access et la perte des modifications non sauvegardées.
- **Centraliser les déclarations** dans un module standard dédié, pour leur maintenance et leur cohérence.
- **Respecter la bitness** en suivant scrupuleusement les conventions de la section 21.7, et tester sur les deux architectures présentes sur le parc.
- **Privilégier les solutions natives** lorsqu'elles existent : si VBA, le modèle objet d'Access ou un objet de plus haut niveau (tel le `FileSystemObject`, section 22.10) couvre le besoin, mieux vaut les utiliser que descendre au niveau de l'API.
- **Vérifier systématiquement la valeur de retour** de chaque appel plutôt que de présumer son succès.

## Articulation avec le reste du chapitre et de la formation

Cette section fonde l'ensemble des techniques d'API du chapitre et s'articule avec plusieurs autres :

- Les **API courantes** prêtes à l'emploi (nom d'utilisateur, nom de machine, exécution, temporisation) sont cataloguées à la section 22.2.
- La **compatibilité 32/64 bits**, `PtrSafe` et `LongPtr`, est traitée en profondeur à la section 21.7.
- Les **déclarations portables** déjà vérifiées figurent à l'annexe H.
- La **gestion d'erreurs** environnante relève du chapitre 13, et les **implications de sécurité** de l'appel de code système, de la section 20.5.
- Les sections sur le **registre** (22.9) et le **système de fichiers** (22.10) mettent en œuvre, respectivement, des appels d'API et des objets de plus haut niveau.

## Points de vigilance

Plusieurs erreurs reviennent fréquemment dans la déclaration et l'appel d'API :

- **Omettre `PtrSafe`** ou typer un handle en `Long` plutôt qu'en `LongPtr`, avec les conséquences décrites à la section 21.7.
- **Déclarer une chaîne en `ByRef`** au lieu de `ByVal ... As String`, ce qui transmet à l'API un pointeur incorrect.
- **Confondre `vbNullString` et `""`** pour passer un pointeur nul à un paramètre chaîne.
- **Sous-dimensionner un tampon** de chaîne, ou se tromper sur la sémantique de la longueur retournée.
- **Ignorer la valeur de retour** de la fonction et poursuivre comme si l'appel avait réussi.
- **Recourir à `As Any` sans nécessité**, et perdre tout contrôle de type sur un paramètre sensible.
- **Tester un nouvel appel sans avoir enregistré**, et risquer la perte de travail en cas de plantage.

⏭️ [22.2. API courantes — GetUserName, GetComputerName, ShellExecute, Sleep](/22-api-windows-integration-office/02-api-courantes.md)
