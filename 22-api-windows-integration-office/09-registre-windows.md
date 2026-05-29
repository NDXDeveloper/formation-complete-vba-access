🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 22.9. Lecture/écriture du registre Windows

Le registre Windows est la base de configuration centrale du système. Une application Access l'utilise couramment pour **persister ses propres paramètres** entre deux sessions — chemin du back-end pour la reliaison (section 21.4), version installée (section 21.3), dernière configuration utilisée — et pour **lire des informations système** : localisation de l'exécutable Access ou déclaration d'un emplacement de confiance (section 21.8). Ce sujet, sollicité à plusieurs reprises au chapitre 21, fait ici l'objet d'un traitement complet.

Trois approches coexistent, de la plus simple à la plus puissante, et le choix dépend du besoin.

## Le registre en bref

Le registre s'organise en **ruches** (hives), désignées par des racines : `HKEY_CURRENT_USER` (HKCU) pour les réglages propres à l'utilisateur, `HKEY_LOCAL_MACHINE` (HKLM) pour ceux de la machine, `HKEY_CLASSES_ROOT` (HKCR) pour les associations de fichiers, et quelques autres. Chaque ruche contient une arborescence de **clés** — comparables à des dossiers — et chaque clé renferme des **valeurs**, paires nom/donnée, dont une valeur **par défaut** (sans nom). Les données sont typées : `REG_SZ` (chaîne), `REG_DWORD` (entier), `REG_BINARY` (binaire), `REG_EXPAND_SZ` (chaîne avec variables d'environnement), `REG_MULTI_SZ` (chaînes multiples).

Une distinction est déterminante pour le déploiement : **HKCU est inscriptible par l'utilisateur** sans privilège particulier, tandis que **l'écriture dans HKLM exige des droits d'administrateur**. C'est la raison pour laquelle, à la section 21.8, les emplacements de confiance étaient déclarés sous HKCU. Enfin, le registre étant critique pour le système, toute écriture imprudente — surtout dans HKLM — peut le déstabiliser : on n'écrit que sous ses propres clés, jamais à l'aveugle dans des clés système inconnues.

## Trois approches selon le besoin

- Les **fonctions natives de VBA** (`SaveSetting`, `GetSetting`…) : les plus simples, mais cantonnées à une sous-arborescence fixe ; idéales pour les seuls réglages de l'application.
- **WScript.Shell** (`RegRead`, `RegWrite`, `RegDelete`) : l'outil polyvalent, capable d'accéder à n'importe quelle clé ; le choix par défaut pour lire ou écrire une valeur isolée.
- L'**API advapi32** : l'approche de bas niveau, réservée aux cas avancés tels que l'énumération des clés ; verbeuse, mais incontournable lorsque les précédentes ne suffisent pas.

## Les fonctions natives de VBA : SaveSetting et GetSetting

VBA intègre des fonctions dédiées au stockage des réglages d'une application : `SaveSetting`, `GetSetting`, `GetAllSettings` et `DeleteSetting`. Leur emploi est immédiat :

```vba
' Enregistrer un paramètre de l'application
SaveSetting "MonApp", "Configuration", "CheminBackEnd", "\\Serveur\Donnees\be.accdb"

' Lire un paramètre, avec une valeur par défaut si la clé est absente
Dim chemin As String
chemin = GetSetting("MonApp", "Configuration", "CheminBackEnd", "")

' Supprimer un paramètre
DeleteSetting "MonApp", "Configuration", "CheminBackEnd"
```

Leur simplicité a une contrepartie : elles écrivent et lisent **uniquement** sous la branche fixe `HKEY_CURRENT_USER\Software\VB and VBA Program Settings\`. Elles conviennent donc parfaitement pour persister les réglages **propres à l'application**, mais ne permettent **pas** d'accéder à une clé arbitraire ni à une clé système — pour cela, il faut recourir à WScript.Shell. À noter qu'elles dispensent de toute gestion d'erreur en lecture, `GetSetting` renvoyant la valeur par défaut fournie lorsque la clé n'existe pas.

## WScript.Shell : accéder à n'importe quelle clé

L'objet `WScript.Shell` permet de lire, écrire et supprimer **n'importe quelle valeur** du registre, en spécifiant le chemin complet préfixé de la racine (HKCU, HKLM, HKCR…). C'est l'approche employée à la section 21.8.

### Lire une valeur

`RegRead` renvoie la valeur, dans son type d'origine (chaîne, nombre, tableau pour les types multiples). Mais elle **déclenche une erreur si la clé ou la valeur n'existe pas** : la lecture doit donc être encadrée (chapitre 13). La fonction d'enveloppe suivante renvoie une valeur par défaut en cas d'absence :

```vba
' Lit une valeur du registre ; renvoie la valeur par défaut si absente
Public Function LireRegistre(ByVal chemin As String, _
                             Optional ByVal defaut As Variant = "") As Variant
    Dim sh As Object
    On Error GoTo Absent
    Set sh = CreateObject("WScript.Shell")
    LireRegistre = sh.RegRead(chemin)
    Exit Function
Absent:
    LireRegistre = defaut
End Function
```

Une subtilité de chemin doit être connue : un chemin **terminé par une barre oblique inverse** désigne la **valeur par défaut** de la clé, tandis qu'un chemin se terminant par un nom désigne la **valeur nommée** correspondante.

### Écrire une valeur

`RegWrite` écrit une valeur en précisant son type, et **crée au passage les clés intermédiaires** manquantes :

```vba
' Écrit une valeur dans le registre
Public Sub EcrireRegistre(ByVal chemin As String, ByVal valeur As Variant, _
                          Optional ByVal typeValeur As String = "REG_SZ")
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    sh.RegWrite chemin, valeur, typeValeur
End Sub
```

`RegWrite` prend en charge les types `REG_SZ`, `REG_EXPAND_SZ`, `REG_DWORD` et `REG_BINARY` — mais **pas** `REG_MULTI_SZ`, qui relève de l'API. La suppression, enfin, s'effectue par `RegDelete` : un chemin sans barre finale supprime une **valeur**, un chemin avec barre finale supprime une **clé** (vide).

### Deux exemples concrets

Localiser l'exécutable Access — usage évoqué à la section 21.8 — consiste à lire la **valeur par défaut** de sa clé App Paths, d'où la barre finale :

```vba
Dim cheminAccess As String
cheminAccess = LireRegistre( _
    "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE\")
```

Déclarer un emplacement de confiance (section 21.8) illustre l'écriture sous HKCU, qui ne requiert aucun droit d'administrateur, avec deux types de valeurs :

```vba
Const BASE As String = "HKCU\Software\Microsoft\Office\16.0\Access\Security\" & _
                       "Trusted Locations\MonApp\"
EcrireRegistre BASE & "Path", "C:\Utilisateurs\Public\MonApp", "REG_SZ"
EcrireRegistre BASE & "AllowSubFolders", 1, "REG_DWORD"
```

Le « 16.0 » de ce chemin est le numéro de version commun à Access 2016 à 2024 et Microsoft 365 (section 21.6). C'est d'ailleurs en descendant au **numéro de build** dans le registre que l'on peut, lorsque c'est nécessaire, distinguer ces versions que `Application.Version` ne sait pas départager.

## La redirection 32/64 bits

Un piège propre aux environnements 64 bits mérite une attention particulière, en lien direct avec la section 21.7. Sous un Windows 64 bits, le registre maintient des **vues séparées** pour les logiciels 32 et 64 bits sous une partie de `HKLM\Software` : un processus **32 bits** voit automatiquement une vue redirigée sous `HKLM\SOFTWARE\WOW6432Node\…`. La **bitness d'Access** détermine donc ce qu'il lit et écrit sous cette branche. La ruche HKCU, elle, n'est pour l'essentiel pas concernée par cette redirection.

WScript.Shell agissant dans la vue du processus appelant, un Access 32 bits et un Access 64 bits ne voient pas nécessairement les mêmes valeurs sous `HKLM\Software`. Pour forcer explicitement l'accès à la vue 64 ou 32 bits, il faut recourir à l'API et à ses indicateurs dédiés (`KEY_WOW64_64KEY` ou `KEY_WOW64_32KEY`).

## L'approche API pour les cas avancés

Lorsque WScript.Shell ne suffit pas, l'API `advapi32` offre un contrôle total via des fonctions telles que `RegOpenKeyEx`, `RegQueryValueEx`, `RegSetValueEx` et `RegCloseKey`. Trois besoins la rendent incontournable : **énumérer** les sous-clés ou les valeurs d'une clé — ce que WScript.Shell ne sait pas faire et qui exige `RegEnumKeyEx` ou `RegEnumValue` —, **forcer une vue 32 ou 64 bits** précise, et manipuler le type `REG_MULTI_SZ`.

Cette approche est en contrepartie verbeuse et délicate, comme tout appel d'API (section 22.1) ; ses déclarations portables figurent à l'annexe H. Pour l'immense majorité des besoins — lire ou écrire une valeur isolée —, **WScript.Shell demeure préférable**, et l'on ne descend à l'API que pour l'énumération ou le contrôle fin de la bitness.

## Bonnes pratiques et précautions

Quelques principes encadrent l'usage du registre depuis Access :

- **Privilégier HKCU** pour les réglages de l'application : aucun droit d'administrateur n'est requis, et les paramètres restent propres à chaque utilisateur.
- **Réserver HKLM aux installations** disposant des droits adéquats (section 21.8), et ne jamais y écrire sans nécessité.
- **N'écrire que sous ses propres clés**, jamais à l'aveugle dans des clés système, le registre étant critique pour le fonctionnement de Windows.
- **Encadrer toute lecture** par une gestion d'erreur, `RegRead` échouant sur une clé absente.
- **Tenir compte de la redirection 32/64 bits** lorsqu'on lit ou écrit sous `HKLM\Software`.

## Articulation avec le reste du chapitre et de la formation

Cette section répond aux usages du registre disséminés dans le chapitre 21 et s'appuie sur d'autres parties de la formation :

- Le stockage du **chemin du back-end** sert la reliaison (section 21.4), et celui des **réglages et de la version**, l'auto-update (section 21.3).
- Les **emplacements de confiance** et la **localisation de MSACCESS.EXE** relèvent de la section 21.8.
- La distinction fine des **versions d'Office** par le build renvoie à la section 21.6.
- La **redirection 32/64 bits** prolonge la section 21.7.
- L'**approche API** mobilise les principes de la section 22.1 et les déclarations de l'annexe H, et la **gestion d'erreurs** relève du chapitre 13.

## Points de vigilance

Plusieurs erreurs reviennent fréquemment dans l'usage du registre :

- **Croire que `SaveSetting` accède à une clé arbitraire**, alors qu'il est cantonné à une branche fixe de HKCU.
- **Lire avec `RegRead` sans gestion d'erreur**, et provoquer un plantage sur une clé absente.
- **Écrire dans HKLM sans droits d'administrateur**, ce qui échoue sur un poste correctement administré.
- **Ignorer la redirection `WOW6432Node`** sous Windows 64 bits, et lire ou écrire dans la mauvaise vue selon la bitness d'Access.
- **Confondre valeur par défaut et valeur nommée**, c'est-à-dire la présence ou l'absence de la barre oblique finale dans le chemin.
- **Tenter d'écrire un `REG_MULTI_SZ` avec `RegWrite`**, type qu'il ne prend pas en charge.
- **Modifier des clés système à l'aveugle**, au risque de déstabiliser Windows.

⏭️ [22.10. Manipulation du système de fichiers (FileSystemObject)](/22-api-windows-integration-office/10-filesystemobject.md)
