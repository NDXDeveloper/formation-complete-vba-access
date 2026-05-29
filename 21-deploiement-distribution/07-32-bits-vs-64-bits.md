🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 21.7. Différences 32 bits / 64 bits — déclarations API et compatibilité

La bitness — le caractère 32 ou 64 bits d'une installation d'Access — constitue, avec la version (section 21.6), l'un des deux grands axes de compatibilité d'un déploiement. C'est aussi l'un des plus pernicieux, car ses effets ne se manifestent souvent qu'à l'exécution sur le poste de l'utilisateur, sous forme d'erreurs de compilation bloquantes ou de plantages, alors que tout fonctionnait parfaitement sur le poste du développeur.

L'essentiel des difficultés se concentre sur deux points : les **déclarations d'API Windows** dans le code VBA, et les **contrôles ActiveX**. Cette section explique pourquoi la bitness importe pour le déploiement et montre comment écrire un code qui s'exécute indifféremment sur les deux architectures.

## Deux architectures, indépendantes de Windows

Access, comme l'ensemble d'Office, existe en deux éditions : **32 bits** et **64 bits**. Cette bitness est celle du logiciel Office, et elle est **indépendante de celle du système d'exploitation**. Un Windows 64 bits — la quasi-totalité des postes aujourd'hui — peut faire tourner aussi bien un Office 32 bits qu'un Office 64 bits. Le seul cas exclu est l'Office 64 bits sur un Windows 32 bits, ce dernier ayant pratiquement disparu.

L'avantage principal du 64 bits tient à la mémoire qu'il peut adresser, nettement supérieure à celle du 32 bits. Cet écart n'est déterminant que pour les traitements très gourmands en mémoire ; pour la plupart des applications Access, le choix de la bitness se joue ailleurs, sur les questions de compatibilité abordées ci-dessous.

## Pourquoi la bitness complique le déploiement

Trois conséquences pratiques découlent de cette dualité, et toutes pèsent sur le déploiement.

D'abord, comme établi aux sections 21.1, 21.2 et 21.6, **un fichier ACCDE doit correspondre à la bitness du poste cible** : un ACCDE 32 bits ne s'ouvre pas sous un Access 64 bits, et inversement. Un parc mêlant les deux architectures impose donc de produire et de distribuer deux livrables distincts.

Ensuite, **les déclarations d'API Windows diffèrent entre les deux architectures.** C'est le point central de cette section : un code d'API rédigé pour le 32 bits provoque, en 64 bits, une erreur de compilation qui désactive l'ensemble du projet VBA.

Enfin, **de nombreux contrôles ActiveX anciens n'existent qu'en 32 bits** et refusent de se charger sous Office 64 bits, ce qui peut purement et simplement empêcher une application de fonctionner.

À cela s'ajoute une contrainte structurelle déjà évoquée à la section 21.2 : un Office 32 bits et un Office 64 bits **ne cohabitent pas** sur une même machine. Un poste est donc nécessairement de l'une ou l'autre architecture, jamais des deux.

## VBA7, PtrSafe et LongPtr : les fondations

La prise en charge du 64 bits est apparue avec **VBA7**, le moteur VBA introduit à partir d'Office 2010. Toutes les versions visées par cette formation (2016 et ultérieures) reposent sur VBA7, dans ses éditions 32 et 64 bits. VBA7 a introduit deux nouveautés indispensables à la portabilité des déclarations d'API :

- Le mot-clé **`PtrSafe`**, qui doit accompagner toute instruction `Declare` pour indiquer qu'elle a été vérifiée comme compatible 64 bits.
- Le type **`LongPtr`**, un type dont la taille s'adapte à l'architecture : 4 octets (l'équivalent d'un `Long`) en 32 bits, 8 octets en 64 bits. C'est le type à utiliser pour tout ce qui représente un **pointeur** ou un **handle** (descripteur de fenêtre, de périphérique, etc.).

Il existe par ailleurs un type **`LongLong`**, entier 64 bits véritable, mais qui n'est disponible **qu'en 64 bits** : l'employer dans du code 32 bits provoque une erreur de compilation. C'est donc `LongPtr`, et non `LongLong`, qu'il faut utiliser pour rédiger des déclarations portables.

## L'enjeu central : les déclarations d'API Windows

### Le mot-clé PtrSafe

En 64 bits, toute instruction `Declare` dépourvue de `PtrSafe` provoque une erreur de compilation signalant que le code doit être mis à jour pour les systèmes 64 bits. Comme toute erreur de compilation en VBA, elle **désactive l'intégralité du projet** : aucune procédure ne s'exécute plus, y compris celles qui n'ont aucun rapport avec l'API en cause. Un seul `Declare` oublié suffit ainsi à rendre l'application inutilisable sur un poste 64 bits.

### Le type LongPtr pour les pointeurs et les handles

Au-delà du simple ajout de `PtrSafe`, il faut **typer correctement les pointeurs et les handles**. En 64 bits, ces valeurs occupent 8 octets ; les déclarer en `Long` (4 octets) tronque la valeur et conduit à des plantages ou à des corruptions de mémoire, parfois de façon erratique et difficile à diagnostiquer. Tout paramètre ou valeur de retour représentant un handle ou un pointeur doit donc être déclaré en `LongPtr`. En revanche, les paramètres représentant de simples entiers conservent le type `Long` sur les deux architectures.

### Une déclaration portable type

La technique consiste à fournir, au moyen de la compilation conditionnelle, une version de la déclaration pour VBA7 (avec `PtrSafe` et `LongPtr`) et une version pour le VBA antérieur. Voici l'exemple de `ShellExecute` (API également présentée à la section 22.2), qui illustre bien le mélange de `LongPtr` pour les handles et de `Long` pour les entiers simples :

```vba
#If VBA7 Then
    ' Office 2010 et ultérieur (32 ou 64 bits) : PtrSafe requis
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As LongPtr
#Else
    ' Office 2007 et antérieur (32 bits uniquement)
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

Dans la branche VBA7, `hwnd` (un handle de fenêtre) et la valeur de retour (un handle d'instance) sont en `LongPtr` et s'ajustent automatiquement à l'architecture, tandis que `nShowCmd` (un simple entier) reste en `Long`. Pour un parc limité aux versions 2016 et ultérieures, la seule branche VBA7 suffit en pratique, puisque le VBA antérieur n'y est plus présent ; conserver la branche `#Else` reste néanmoins une bonne habitude de portabilité.

### Les variables qui reçoivent les résultats

La vigilance sur le type ne s'arrête pas à la déclaration : la **variable qui reçoit un résultat d'API de type `LongPtr` doit elle aussi être déclarée `LongPtr`**, et non `Long`. Déclarer cette variable en `Long` rétablit précisément le défaut de troncature que `LongPtr` cherchait à éviter :

```vba
Dim resultat As LongPtr
resultat = ShellExecute(0, "open", "https://exemple.fr", vbNullString, vbNullString, 1)
```

## Compilation conditionnelle : VBA7 et Win64

Deux constantes de compilation servent à écrire du code adaptatif, et il importe de ne pas les confondre :

- **`VBA7`** indique la *version du moteur VBA* : elle vaut `True` sur Office 2010 et ultérieur, **quelle que soit la bitness**. C'est elle qui sert à décider d'inclure ou non `PtrSafe`.
- **`Win64`** indique l'*architecture* : elle vaut `True` uniquement en 64 bits. C'est elle qui sert lorsqu'on a réellement besoin de code différent selon la bitness, par exemple pour employer `LongLong`.

La confusion classique consiste à croire que `VBA7` signale le 64 bits : il n'en est rien. Un Office 2016 **32 bits** a `VBA7 = True` mais `Win64 = False`.

```vba
#If Win64 Then
    ' Code spécifique au 64 bits (ex. usage direct de LongLong)
#Else
    ' Code spécifique au 32 bits
#End If
```

Dans la pratique, l'usage correct de `LongPtr` rend rarement nécessaire le recours à `#If Win64` : c'est `#If VBA7` qui suffit dans l'immense majorité des déclarations d'API, `LongPtr` absorbant à lui seul la différence de largeur entre architectures.

## Les contrôles ActiveX, frein fréquent au 64 bits

Une cause majeure de blocage tient aux **contrôles ActiveX**. Beaucoup de contrôles tiers anciens — calendriers, grilles, composants divers — n'ont jamais été déclinés en 64 bits. Sous Office 64 bits, ils refusent de se charger, et un formulaire qui en dépend devient inutilisable.

C'est précisément cette dépendance qui a longtemps maintenu de nombreuses organisations sur Office 32 bits. Avant de viser le 64 bits, il faut donc inventorier les contrôles ActiveX employés par l'application (chapitre 8.5) et s'assurer qu'ils existent dans cette architecture, ou les remplacer par des solutions natives. Une application reposant sur un contrôle ActiveX exclusivement 32 bits ne pourra tout simplement pas être déployée sur un parc 64 bits sans refonte.

## Quelle bitness choisir ?

Historiquement, Microsoft recommandait par défaut l'installation **32 bits** d'Office, essentiellement pour préserver la compatibilité avec les compléments, contrôles ActiveX et anciennes solutions conçus pour cette architecture. Cette recommandation a évolué : pour les installations récentes sur Windows 64 bits, le **64 bits est désormais l'édition installée par défaut**.

Du point de vue du développeur, la conséquence est double. D'une part, **on ne maîtrise pas toujours la bitness des postes utilisateurs**, qui peut varier au sein d'un même parc ; il faut donc rédiger un code portable, indépendant de l'architecture. D'autre part, **l'ACCDE devant correspondre à la bitness du poste**, un parc mixte impose de produire les deux livrables et de déployer le bon sur chaque poste, en cohérence avec la logique de détection abordée à la section 21.2.

## Au-delà des API : la mémoire

La différence d'architecture a aussi un effet sur la mémoire adressable. Une application Access 32 bits est bornée par une limite de mémoire de processus relativement basse, là où le 64 bits peut en exploiter bien davantage. Cela peut compter pour des traitements manipulant de très grands volumes en mémoire, des tableaux massifs ou des automations Office portant sur de gros fichiers.

Il faut toutefois bien distinguer cette **limite de mémoire du processus** de la **limite de taille du fichier `.accdb`**, fixée à 2 Go indépendamment de la bitness : ce sont deux contraintes différentes, et passer en 64 bits n'augmente en rien la taille maximale d'une base Access. La limite des 2 Go du fichier et ses stratégies de contournement sont traitées à la section 18.9.

## Bonnes pratiques

Quelques principes garantissent un code résistant aux deux architectures :

- **Rédiger toutes les déclarations d'API en forme portable** (`PtrSafe`, `LongPtr`, compilation conditionnelle), de sorte qu'un même code source se compile et s'exécute en 32 comme en 64 bits.
- **Centraliser les déclarations d'API** dans un module unique, ce qui facilite leur vérification et leur mise à jour. L'annexe H rassemble des déclarations d'API Windows compatibles 32/64 bits prêtes à l'emploi.
- **Tester l'application sur les deux architectures** réellement présentes sur le parc, plutôt que de présumer qu'un fonctionnement correct en 32 bits vaut pour le 64 bits.
- **Inventorier les dépendances ActiveX** et vérifier leur disponibilité en 64 bits avant tout déploiement sur cette architecture.

## Articulation avec le reste du déploiement

La question de la bitness recoupe de nombreux points de la formation :

- La contrainte de bitness de l'**ACCDE** a été posée lors du packaging et du runtime (sections 21.1 et 21.2), et la détection du poste cible y est abordée.
- La distinction entre **version et bitness** est traitée à la section 21.6.
- Les **contrôles ActiveX** relèvent de la section 8.5.
- La **déclaration et l'appel d'API** au sens général font l'objet des sections 22.1 et 22.2, et les déclarations portables prêtes à l'emploi de l'**annexe H**.
- La **limite de 2 Go** du fichier, à ne pas confondre avec la mémoire de processus, est traitée à la section 18.9.

## Points de vigilance

Plusieurs erreurs reviennent fréquemment sur le sujet de la bitness :

- **Oublier `PtrSafe`** sur une déclaration `Declare`, ce qui provoque en 64 bits une erreur de compilation désactivant tout le projet.
- **Déclarer un handle ou un pointeur en `Long`** au lieu de `LongPtr`, entraînant troncature, plantages ou corruptions difficiles à diagnostiquer.
- **Confondre `VBA7` et `Win64`** dans la compilation conditionnelle, et croire à tort que `VBA7` signale le 64 bits.
- **Employer `LongLong` dans du code censé tourner aussi en 32 bits**, ce qui empêche la compilation sur cette architecture.
- **Distribuer un ACCDE de mauvaise bitness** par rapport au poste cible, le rendant impossible à ouvrir.
- **Déployer en 64 bits une application dépendant d'un contrôle ActiveX 32 bits**, qui refusera alors de se charger.
- **Confondre la limite mémoire du 32 bits avec la limite de 2 Go du fichier `.accdb`**, et croire qu'un passage au 64 bits repousse cette dernière.

⏭️ [21.8. Installation automatisée (scripts, paramètres de ligne de commande)](/21-deploiement-distribution/08-installation-automatisee.md)
