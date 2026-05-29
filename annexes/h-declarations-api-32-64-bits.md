🔝 Retour au [Sommaire](/SOMMAIRE.md)

# H. Déclarations API Windows compatibles 32/64 bits

Cette annexe fournit une bibliothèque de déclarations d'API Windows prêtes à utiliser, compatibles avec les versions **32 et 64 bits** d'Access. Elle complète les **chapitres 22.1 et 22.2** (déclaration et appel d'API, API courantes) ainsi que le **chapitre 21.7** (différences 32/64 bits).

## Le modèle de compatibilité (VBA7, PtrSafe, LongPtr)

Depuis Office 2010, VBA existe en deux architectures (32 et 64 bits) et introduit un modèle de déclaration unifié. Toutes les versions ciblées par cette formation (2016, 2019, 2021, M365) reposent sur **VBA7**.

| Élément | Rôle |
|---|---|
| `PtrSafe` | Mot-clé ajouté à `Declare` indiquant que la déclaration est compatible 64 bits. **Obligatoire en 64 bits**, recommandé partout. |
| `LongPtr` | Type entier de la **taille d'un pointeur** : 4 octets en 32 bits, 8 octets en 64 bits. À utiliser pour tous les **handles et pointeurs** (`hwnd`, `hDC`, handles de processus…). |
| `LongLong` | Entier 8 octets, **disponible uniquement en 64 bits**. À éviter au profit de `LongPtr` pour rester portable. |

**Constantes de compilation conditionnelle** utiles :

- `VBA7` : `True` sur Office 2010 et ultérieur (VBA disposant de `PtrSafe`/`LongPtr`).
- `Win64` : `True` uniquement sur Office **64 bits** (indépendamment du système d'exploitation).

---

## La règle d'or : une seule déclaration suffit

> 💡 **Sur VBA7, une déclaration `Declare PtrSafe … LongPtr …` fonctionne telle quelle en 32 **et** en 64 bits.** En 32 bits, `LongPtr` est traité comme `Long` ; en 64 bits, comme `LongLong`. Il n'est donc **pas** nécessaire d'écrire deux versions séparées dans la grande majorité des cas.

La compilation conditionnelle par bitness (`#If Win64`) ne sert que pour les rares API dont la signature diffère réellement entre les deux architectures.

### Conversion d'une ancienne déclaration

Pour rendre compatible une déclaration héritée (32 bits) :

1. Ajouter `PtrSafe` après `Declare`.
2. Remplacer par `LongPtr` chaque **handle** ou **pointeur** (anciennement `Long`) — y compris la **valeur de retour** si c'est un handle.
3. Laisser en `Long` les entiers « plats » (drapeaux, tailles, compteurs).
4. Déclarer en `As LongPtr` les **variables VBA** qui reçoivent un handle.

---

## Modèle de rétrocompatibilité (Office 2007 et antérieur)

Nécessaire **uniquement** pour faire cohabiter le code avec des versions pré-2010 (sans `PtrSafe`). Sinon, la version `PtrSafe` seule suffit.

```vba
#If VBA7 Then
    ' Office 2010+ (32 ou 64 bits)
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    ' Office 2007 et antérieur (32 bits uniquement)
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If
```

---

## Où placer les déclarations

- Dans un **module standard**, en tête de module (sous `Option Explicit`).
- Préférer `Private` pour éviter les collisions de noms ; `Public` est possible **uniquement dans un module standard** (interdit dans un module de classe ou de formulaire, où `Private` est obligatoire).

---

## Bibliothèque de déclarations courantes

Toutes les déclarations ci-dessous sont en forme `PtrSafe` (compatibles 32/64 bits sur VBA7).

### Système et utilisateur

```vba
Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, ByRef nSize As Long) As Long

Private Declare PtrSafe Function GetComputerName Lib "kernel32.dll" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, ByRef nSize As Long) As Long

Private Declare PtrSafe Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" _
    (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
```

### Pause et minutage

```vba
Private Declare PtrSafe Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Declare PtrSafe Function GetTickCount Lib "kernel32.dll" () As Long

Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long

' Minutage haute résolution (profilage, chapitre 18.1)
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32.dll" _
    (ByRef lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32.dll" _
    (ByRef lpFrequency As Currency) As Long
```

> Le type `Currency` (8 octets) sert de conteneur pour le compteur 64 bits, sur les deux architectures. Pour un **temps écoulé**, diviser l'écart de compteur par la fréquence : la mise à l'échelle interne de `Currency` s'annule dans le rapport.

### Gestion des fenêtres

```vba
Private Declare PtrSafe Function FindWindow Lib "user32.dll" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Private Declare PtrSafe Function GetForegroundWindow Lib "user32.dll" () As LongPtr

Private Declare PtrSafe Function SetForegroundWindow Lib "user32.dll" _
    (ByVal hwnd As LongPtr) As Long

Private Declare PtrSafe Function ShowWindow Lib "user32.dll" _
    (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
```

### Lancement d'applications

```vba
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, _
     ByVal lpParameters As String, ByVal lpDirectory As String, _
     ByVal nShowCmd As Long) As LongPtr

' Exécuter et attendre la fin d'un processus
Private Declare PtrSafe Function OpenProcess Lib "kernel32.dll" _
    (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
     ByVal dwProcessId As Long) As LongPtr
Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32.dll" _
    (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32.dll" _
    (ByVal hObject As LongPtr) As Long
```

### Écran et interface

```vba
Private Declare PtrSafe Function GetSystemMetrics Lib "user32.dll" _
    (ByVal nIndex As Long) As Long
' nIndex : SM_CXSCREEN = 0 (largeur), SM_CYSCREEN = 1 (hauteur)

Private Declare PtrSafe Function MessageBeep Lib "user32.dll" _
    (ByVal wType As Long) As Long
```

### Fichiers de configuration (.ini)

```vba
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32.dll" _
    Alias "WritePrivateProfileStringA" _
    (ByVal lpAppName As String, ByVal lpKeyName As String, _
     ByVal lpString As String, ByVal lpFileName As String) As Long

Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32.dll" _
    Alias "GetPrivateProfileStringA" _
    (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
     ByVal lpReturnedString As String, ByVal nSize As Long, _
     ByVal lpFileName As String) As Long
```

---

## Exemples d'appel (rappels d'utilisation)

Récupérer le nom de l'utilisateur Windows (motif classique du tampon pré-alloué) :

```vba
Public Function NomUtilisateurWindows() As String
    Dim sTampon As String * 255
    Dim lTaille As Long
    lTaille = 255
    If GetUserName(sTampon, lTaille) <> 0 Then
        ' lTaille inclut le caractère nul final
        NomUtilisateurWindows = Left$(sTampon, lTaille - 1)
    End If
End Function
```

Ouvrir une URL ou un fichier avec son application associée (`vbNullString` = pointeur de chaîne nul, `0&` = handle nul) :

```vba
ShellExecute 0&, "open", "https://www.example.com", vbNullString, vbNullString, 1
```

---

## Points de vigilance

> ⚠️ **`PtrSafe` obligatoire en 64 bits.** Une déclaration sans `PtrSafe` provoque une erreur de compilation sur Access 64 bits. L'ajouter systématiquement.

> ⚠️ **Ne jamais laisser `Long` à la place d'un handle en 64 bits.** Un handle ou pointeur déclaré `Long` est tronqué à 32 bits : plantage ou corruption mémoire. Utiliser `LongPtr` pour tout handle/pointeur, et pour les variables VBA qui les reçoivent.

> ⚠️ **`LongLong` indisponible en 32 bits.** Préférer `LongPtr`, qui s'adapte aux deux architectures, sauf nécessité explicite d'un entier 64 bits.

> ⚠️ **Versions `A` (ANSI) vs `W` (Unicode).** Les déclarations ci-dessus utilisent les variantes `…A` (ANSI), qui attendent des chaînes `String` classiques. C'est le choix le plus simple et le plus courant en VBA Access.

> 💡 **`GetTickCount` déborde après ~49 jours** (compteur 32 bits). Pour des durées longues, l'API `GetTickCount64` renvoie un compteur plus large ; pour du profilage précis, préférer `QueryPerformanceCounter`.

---

> 💡 **À retenir.** La compatibilité 32/64 bits tient en trois gestes : **`PtrSafe`** sur la déclaration, **`LongPtr`** pour les handles et pointeurs, **`Long`** pour les entiers ordinaires. Sur VBA7, **une seule déclaration couvre les deux architectures** — le `#If VBA7`/`#If Win64` ne sert qu'à la rétrocompatibilité avec Office 2007 ou aux rares API réellement divergentes. La théorie de la déclaration et de l'appel d'API figure aux chapitres 22.1 et 22.2.

⏭️ [I. Raccourcis clavier de l'éditeur VBA](/annexes/i-raccourcis-clavier-editeur-vba.md)
